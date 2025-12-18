VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm075012 
   BorderStyle     =   1  '單線固定
   Caption         =   "庭期資料維護"
   ClientHeight    =   5964
   ClientLeft      =   648
   ClientTop       =   516
   ClientWidth     =   8952
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form6"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5964
   ScaleWidth      =   8952
   Begin VB.CommandButton Command4 
      Caption         =   "其他出庭律師(&L)"
      Height          =   375
      Left            =   5184
      TabIndex        =   61
      Top             =   5070
      Width           =   1485
   End
   Begin VB.CommandButton cmdQueryFile 
      Caption         =   "卷宗區"
      Height          =   345
      Left            =   6840
      Style           =   1  '圖片外觀
      TabIndex        =   54
      Top             =   1890
      Width           =   915
   End
   Begin VB.CommandButton cmdNote 
      Caption         =   "上傳電子筆錄"
      Height          =   375
      Left            =   1860
      Style           =   1  '圖片外觀
      TabIndex        =   53
      Top             =   5070
      Width           =   1485
   End
   Begin VB.CommandButton cmdBrief 
      Caption         =   "上傳開庭紀要"
      Height          =   375
      Left            =   240
      Style           =   1  '圖片外觀
      TabIndex        =   52
      Top             =   5070
      Width           =   1485
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "上傳開庭通知"
      Height          =   375
      Left            =   3510
      TabIndex        =   51
      Top             =   5070
      Width           =   1485
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  '平面
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   1860
      Left            =   5790
      TabIndex        =   44
      Top             =   2250
      Width           =   3135
      Begin VB.CommandButton Command1 
         Caption         =   "E-Mail開庭紀要(&E)"
         Height          =   375
         Left            =   1065
         TabIndex        =   45
         Top             =   120
         Width           =   1635
      End
      Begin MSForms.ListBox lstMailCC 
         Height          =   315
         Left            =   1110
         TabIndex        =   14
         Top             =   570
         Width           =   1905
         VariousPropertyBits=   746586139
         ScrollBars      =   2
         DisplayStyle    =   2
         Size            =   "3360;556"
         MatchEntry      =   0
         ListStyle       =   1
         MultiSelect     =   1
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         Caption         =   "E-Mail副本："
         Height          =   210
         Index           =   10
         Left            =   15
         TabIndex        =   47
         Top             =   570
         Width           =   1020
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         Caption         =   "(可複選)"
         Height          =   255
         Index           =   14
         Left            =   240
         TabIndex        =   46
         Top             =   795
         Width           =   795
      End
   End
   Begin VB.TextBox txtPaperNum 
      Height          =   300
      Left            =   990
      MaxLength       =   9
      TabIndex        =   0
      Top             =   757
      Width           =   1575
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7575
      Top             =   75
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
            Picture         =   "frm075012.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075012.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075012.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075012.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075012.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075012.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075012.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075012.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075012.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075012.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075012.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbar 
      Align           =   1  '對齊表單上方
      Height          =   576
      Left            =   0
      TabIndex        =   36
      Top             =   0
      Width           =   8952
      _ExtentX        =   15790
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
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   8370
      Top             =   1140
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8340
      Top             =   660
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSForms.Label lbePerson 
      Height          =   255
      Left            =   3030
      TabIndex        =   60
      Top             =   3332
      Width           =   1500
      BackColor       =   -2147483637
      VariousPropertyBits=   27
      Size            =   "2646;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cboPerson 
      Height          =   285
      Left            =   990
      TabIndex        =   6
      Top             =   3317
      Width           =   1935
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "3413;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   300
      Index           =   1
      Left            =   4740
      TabIndex        =   5
      Top             =   2990
      Width           =   945
      VariousPropertyBits=   671105051
      MaxLength       =   7
      Size            =   "1667;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   300
      Index           =   0
      Left            =   990
      TabIndex        =   1
      Top             =   2352
      Width           =   300
      VariousPropertyBits=   671105051
      MaxLength       =   1
      Size            =   "529;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   0
      Left            =   990
      TabIndex        =   7
      Top             =   3628
      Width           =   1095
      VariousPropertyBits=   671105051
      MaxLength       =   10
      Size            =   "1931;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   1
      Left            =   2145
      TabIndex        =   8
      Top             =   3628
      Width           =   1095
      VariousPropertyBits=   671105051
      MaxLength       =   10
      Size            =   "1931;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   2
      Left            =   3300
      TabIndex        =   9
      Top             =   3628
      Width           =   1095
      VariousPropertyBits=   671105051
      MaxLength       =   10
      Size            =   "1931;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text2 
      Height          =   300
      Index           =   0
      Left            =   990
      TabIndex        =   10
      Top             =   3947
      Width           =   1095
      VariousPropertyBits=   671105051
      MaxLength       =   10
      Size            =   "1931;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text2 
      Height          =   300
      Index           =   1
      Left            =   2145
      TabIndex        =   11
      Top             =   3947
      Width           =   1095
      VariousPropertyBits=   671105051
      MaxLength       =   10
      Size            =   "1931;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text2 
      Height          =   300
      Index           =   2
      Left            =   3300
      TabIndex        =   12
      Top             =   3947
      Width           =   1095
      VariousPropertyBits=   671105051
      MaxLength       =   10
      Size            =   "1931;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   300
      Index           =   4
      Left            =   990
      TabIndex        =   2
      Top             =   2671
      Width           =   300
      VariousPropertyBits=   671105051
      MaxLength       =   1
      Size            =   "529;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   660
      Index           =   7
      Left            =   990
      TabIndex        =   13
      Top             =   4290
      Width           =   7575
      VariousPropertyBits=   -1466941413
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "13361;1164"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   300
      Index           =   6
      Left            =   2790
      TabIndex        =   4
      Top             =   2990
      Width           =   585
      VariousPropertyBits=   671105051
      MaxLength       =   4
      Size            =   "1032;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   300
      Index           =   5
      Left            =   990
      TabIndex        =   3
      Top             =   2990
      Width           =   945
      VariousPropertyBits=   671105051
      MaxLength       =   7
      Size            =   "1667;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label4 
      Height          =   255
      Left            =   6600
      TabIndex        =   59
      Top             =   780
      Width           =   1500
      BackColor       =   -2147483637
      VariousPropertyBits=   27
      Size            =   "2646;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cboCaseName 
      Height          =   285
      Left            =   990
      TabIndex        =   58
      Top             =   1403
      Width           =   7245
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "12779;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbeCusName 
      Height          =   255
      Left            =   2130
      TabIndex        =   57
      Top             =   1099
      Width           =   6435
      BackColor       =   -2147483637
      VariousPropertyBits=   27
      Size            =   "11351;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label UIDname 
      Height          =   255
      Left            =   5220
      TabIndex        =   56
      Top             =   5580
      Width           =   900
      VariousPropertyBits=   27
      Caption         =   "UIDname"
      Size            =   "1587;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label IDname 
      Height          =   255
      Left            =   1350
      TabIndex        =   55
      Top             =   5580
      Width           =   900
      VariousPropertyBits=   27
      Caption         =   "IDname"
      Size            =   "1587;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label13 
      Caption         =   "取消庭期日期："
      Height          =   255
      Left            =   3450
      TabIndex        =   50
      Top             =   3013
      Width           =   1260
   End
   Begin VB.Label Label12 
      Caption         =   "開  庭  別：         (1.民事庭 2.偵查庭 3.刑事庭 4.刑附民庭 5.行政庭 6.調解庭)"
      Height          =   255
      Left            =   50
      TabIndex        =   49
      Top             =   2375
      Width           =   5955
   End
   Begin VB.Label Label10 
      Caption         =   "智權人員："
      Height          =   255
      Left            =   5625
      TabIndex        =   48
      Top             =   780
      Width           =   930
   End
   Begin VB.Label lblNum 
      Height          =   255
      Left            =   4050
      TabIndex        =   43
      Top             =   1737
      Width           =   735
   End
   Begin VB.Label lblRevDate 
      Height          =   255
      Left            =   1005
      TabIndex        =   42
      Top             =   1737
      Width           =   975
   End
   Begin VB.Label lblLNum 
      Height          =   255
      Left            =   4080
      TabIndex        =   41
      Top             =   2056
      Width           =   1335
   End
   Begin VB.Label lblLawNum 
      Height          =   255
      Left            =   1005
      TabIndex        =   40
      Top             =   2056
      Width           =   2175
   End
   Begin VB.Label Label19 
      Caption         =   "檢  察  官："
      Height          =   255
      Left            =   50
      TabIndex        =   39
      Top             =   3970
      Width           =   920
   End
   Begin VB.Label Label17 
      Caption         =   "法        官："
      Height          =   255
      Left            =   50
      TabIndex        =   38
      Top             =   3651
      Width           =   920
   End
   Begin VB.Label Label15 
      Caption         =   "股別："
      Height          =   255
      Left            =   3435
      TabIndex        =   37
      Top             =   2056
      Width           =   615
   End
   Begin VB.Label UTM 
      Height          =   255
      Left            =   7245
      TabIndex        =   35
      Top             =   5580
      Width           =   825
   End
   Begin VB.Label UDT 
      Height          =   255
      Left            =   6285
      TabIndex        =   34
      Top             =   5580
      Width           =   735
   End
   Begin VB.Label Label29 
      Caption         =   "UpdateID ："
      Height          =   255
      Left            =   4245
      TabIndex        =   33
      Top             =   5580
      Width           =   975
   End
   Begin VB.Label CTM 
      Height          =   255
      Left            =   3285
      TabIndex        =   32
      Top             =   5580
      Width           =   855
   End
   Begin VB.Label CDT 
      Height          =   255
      Left            =   2325
      TabIndex        =   31
      Top             =   5580
      Width           =   735
   End
   Begin VB.Label Label24 
      Caption         =   "CreateID ："
      Height          =   255
      Left            =   285
      TabIndex        =   30
      Top             =   5580
      Width           =   975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      X1              =   6510
      X2              =   15060
      Y1              =   1650
      Y2              =   1650
   End
   Begin VB.Label Label9 
      Caption         =   "案件名稱："
      Height          =   255
      Left            =   50
      TabIndex        =   29
      Top             =   1418
      Width           =   975
   End
   Begin VB.Label lbeGov 
      Height          =   255
      Left            =   4830
      TabIndex        =   28
      Top             =   1737
      Width           =   3135
   End
   Begin VB.Label lbeCaseNum 
      Height          =   255
      Left            =   3705
      TabIndex        =   17
      Top             =   780
      Width           =   1815
   End
   Begin VB.Label lbeCus 
      Height          =   255
      Left            =   990
      TabIndex        =   16
      Top             =   1099
      Width           =   1110
   End
   Begin VB.Label Label6 
      Caption         =   "當  事  人："
      Height          =   255
      Left            =   50
      TabIndex        =   27
      Top             =   1099
      Width           =   915
   End
   Begin VB.Label Label5 
      Caption         =   "開庭種類：         (1.偵查 2.審理 3.言詞辯論 4.調查 5.調解)"
      Height          =   255
      Left            =   50
      TabIndex        =   26
      Top             =   2694
      Width           =   5040
   End
   Begin VB.Label Label3 
      Caption         =   "開庭日期："
      Height          =   255
      Left            =   50
      TabIndex        =   25
      Top             =   3013
      Width           =   900
   End
   Begin VB.Label Label18 
      Caption         =   "備        註："
      Height          =   255
      Left            =   45
      TabIndex        =   24
      Top             =   4350
      Width           =   920
   End
   Begin VB.Label Label11 
      Caption         =   "時間："
      Height          =   255
      Left            =   2145
      TabIndex        =   23
      Top             =   3013
      Width           =   615
   End
   Begin VB.Label Label8 
      Caption         =   "法院案號："
      Height          =   255
      Left            =   50
      TabIndex        =   22
      Top             =   2056
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "開庭人員："
      Height          =   255
      Left            =   50
      TabIndex        =   21
      Top             =   3332
      Width           =   900
   End
   Begin VB.Label Label2 
      Caption         =   "本所案號：  "
      Height          =   255
      Left            =   2820
      TabIndex        =   20
      Top             =   780
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "收  文  號："
      Height          =   255
      Index           =   0
      Left            =   50
      TabIndex        =   19
      Top             =   780
      Width           =   975
   End
   Begin VB.Label Label21 
      Caption         =   "收  受  日："
      Height          =   255
      Left            =   50
      TabIndex        =   18
      Top             =   1737
      Width           =   975
   End
   Begin VB.Label Label25 
      Caption         =   "機關代號："
      Height          =   255
      Left            =   3075
      TabIndex        =   15
      Top             =   1737
      Width           =   975
   End
End
Attribute VB_Name = "frm075012"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/17 改成Form2.0 ; cboCaseName、lbeCusName、Label4、Text(index)、Text1(index)、Text2(index)、cboPerson、lbePerson、lstMailCC
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/3 日期欄已修改
Option Explicit

Dim blnIsSave As Boolean, PaperNum As String, LcTmp As String
Dim blnIsSearch As Boolean, intSaveKind As Integer, blnisEdit As Boolean, Today As String
Dim lc01 As String, lc02 As String, lc03 As String, lc04 As String, blnIsNew As Boolean
Dim blnIsCancel As Boolean
Dim m_CP01 As String
Dim m_CP02 As String
Dim m_CP03 As String
Dim m_CP04 As String
Dim m_DataList() As String
Dim m_DataCount As Integer
Dim m_IndexNow As Integer
Dim m_EDIT As Integer
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
Dim m_stat As Integer
Dim m_StrTo As String, m_CP13 As String, m_cdp10 As String 'Add By Sindy 2011/6/14
'Dim m_cdp16 As String 'Add By Sindy 2011/6/14
'Added By Lydia 2015/10/30 上傳開庭通知
'Public m_strSaveFiles As String '新增附件
Public m_strSaveFileType As String '欲新增的檔案種類:1.開庭通知 2.開庭紀要 3.電子筆錄 Modify By Sindy 2016/6/24
Dim m_AttachPath As String '附件暫存區
Public m_strSaveFilesOA As String '新增開庭通知附件 Add By Sindy 2016/6/24
Public m_strSaveFilesBRIEF As String '新增開庭紀要附件 Add By Sindy 2016/6/24
Public m_strSaveFilesNOTE As String '新增開庭紀要附件 Add By Sindy 2016/6/24
Dim m_CP43 As String, m_CP14 As String, m_CP29 As String 'Add By Sindy 2020/6/10
Dim m_CL02 As String 'Added by Lydia 2024/07/29 收文號之出庭律師
Dim m_CP10 As String 'Added by Lydia 2025/03/19 案件性質

Private Sub cboPerson_Click()
Dim strTemp As String
Dim nPos As Integer
Dim strPerson As String
   
   nPos = 0
   strPerson = ""
   If cboPerson <> "" Then
      nPos = InStr(cboPerson.Text, ",")
      If nPos <> 0 Then
         strPerson = Left(cboPerson.Text, nPos - 1)
      Else
         strPerson = cboPerson
      End If
      If strPerson <> "" Then
         If ClsPDGetStaff(strPerson, strTemp) Then
            lbePerson = strTemp
         End If
      End If
   End If
End Sub

'Add By Sindy 2010/11/26
'Modified by Lydia 2021/09/17 改成Form 2.0
'Private Sub cboPerson_KeyPress(KeyAscii As Integer)
Private Sub cboPerson_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub cboPerson_Validate(Cancel As Boolean)
Dim strTemp As String
Dim nPos As Integer
Dim strPerson As String
   
   nPos = 0
   If cboPerson <> "" Then
      nPos = InStr(cboPerson.Text, ",")
      If nPos <> 0 Then
         strPerson = Left(cboPerson.Text, nPos - 1)
      Else
         strPerson = cboPerson
      End If
      If ClsPDGetStaff(strPerson, strTemp) Then
         lbePerson = strTemp
      Else
         Cancel = True
      End If
   End If
   If Cancel Then
      cboPerson.SelStart = 0
      cboPerson.SelLength = Len(cboPerson)
   End If
End Sub

'Add By Sindy 2016/6/27 查看卷宗區
Private Sub cmdQueryFile_Click()
   Screen.MousePointer = vbHourglass
   frm100101_L.m_strKey = txtPaperNum.Text
   frm100101_L.SetParent Me
   If frm100101_L.QueryData = True Then
      frm100101_L.Show
      Me.Hide
   Else
      Unload frm100101_L
   End If
   Screen.MousePointer = vbDefault
End Sub

'Add By Sindy 2011/6/9 E-Mail開庭紀要
Private Sub Command1_Click()
Dim strToCC As String, i As Long, s As Variant
Dim strSubject As String, strContent As String
Dim strFiles As String
Dim adoRst As ADODB.Recordset
   
   'Modify By Sindy 2016/6/27 改抓卷宗區
   strSql = "SELECT cpp01,cpp02 FROM casepaperpdf WHERE cpp01='" & txtPaperNum.Text & "' and instr(upper(cpp02),upper('.BRIEF.')) > 0"
   intI = 1
   Set adoRst = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      '正本
      If Trim(m_StrTo) = "" Then
         MsgBox "收件人空白，無法寄送！"
         Exit Sub
      End If
      '副本
      For i = 0 To lstMailCC.ListCount - 1
         If lstMailCC.Selected(i) = True Then
            If strToCC = "" Then
               strToCC = Left(Trim(lstMailCC.List(i)), 5)
            Else
               strToCC = strToCC & ";" & Left(Trim(lstMailCC.List(i)), 5)
            End If
         End If
      Next
      
      Screen.MousePointer = vbHourglass
      '附件
      strFiles = ""
      adoRst.MoveFirst
      Do While Not adoRst.EOF
         If GetAttachFile(adoRst.Fields("cpp01"), adoRst.Fields("cpp02"), m_AttachPath & "\" & adoRst.Fields("cpp02")) = True Then
            strFiles = strFiles & "*" & m_AttachPath & "\" & adoRst.Fields("cpp02")
         End If
         adoRst.MoveNext
      Loop
      strFiles = Mid(strFiles, 2)
      strSubject = lbeCaseNum & "之" & ChangeTStringToTDateString(Text(5)) & " 開庭紀要 !"
      strContent = "本所案號：" + lbeCaseNum + vbCrLf + _
                   "開庭紀要,如附件！" + vbCrLf
      '寄收件人及副本,含附件
      'Modify By Sindy 2020/6/10
      '1. 原只發EMAIL給案件最新智權人員，請修改為若有案源則發給案源介紹人(可能多個)，無案源才發給案件最新智權人員。
      '2. 加發副本給該C類收文號之相關總收文號之承辦人、協辦人員及所有的出庭律師(以收文號讀取caselawer)。
      PUB_SendMail strUserNum, m_StrTo, "", strSubject, strContent, "", strFiles, , , , strToCC
      s = MsgBox("郵件已送出", , "MAIL!!")
      Screen.MousePointer = vbDefault
      Me.Command1.SetFocus
   Else
      MsgBox "無開庭紀要附件！"
   End If
   
   Set adoRst = Nothing
End Sub

Private Function GetAttachFile(ByVal strCP09 As String, ByRef pFileName As String, Optional pSavePath As String) As Boolean
   Dim stAttPath As String
   
On Error GoTo ErrHnd

   If pSavePath = "" Then
      If Dir(m_AttachPath, vbDirectory) = "" Then
         MkDir m_AttachPath
      End If
      stAttPath = m_AttachPath & "\" & pFileName
   Else
      '傳完整的檔案路徑:路徑+檔名
      If InStr(pSavePath, m_AttachPath) > 0 Then
         If Dir(m_AttachPath, vbDirectory) = "" Then
            MkDir m_AttachPath
         End If
      End If
      stAttPath = pSavePath
   End If
   
   GetAttachFile = PUB_GetAttachFile_CPP(strCP09, pFileName, stAttPath, True)

   Exit Function
   
ErrHnd:
   If Err.Number = 70 Then
      MsgBox ChgSQL(pFileName) & "檔案已開啟！", vbCritical
   Else
      MsgBox Err.Description, vbCritical
   End If
End Function

Private Sub Form_Load()
'*****************
   m_bInsert = IsUserHasRightOfFunction("frm075011", strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction("frm075011", strEdit, False)
   m_bDelete = IsUserHasRightOfFunction("frm075011", strDel, False)
   m_bQuery = IsUserHasRightOfFunction("frm075011", strFind, False)
'****************

   'Added by Lydia 2021/09/17 如果一開始將ListBox拉到需要的大小，字型會自動放大；所以畫面預設為一列高度，Form_Load才放大到需要的大小
   lstMailCC.Height = 1125
   lstMailCC.Width = 1900
   'end 2021/09/17
   
   MoveFormToCenter Me
   Cleartxt
   blnIsSave = False
   blnisEdit = False
   blnIsNew = False
   TxtCanTUse
   CmdEnabled
   Today = ChangeWStringToTString(GetTodayDate)
   GetStaff
   GetData
   
'**********************
   If m_bInsert Then
       tlbar.Buttons(1).Enabled = True
   Else
       tlbar.Buttons(1).Enabled = False
   End If
   If m_bUpdate Then
       tlbar.Buttons(2).Enabled = True
   Else
       tlbar.Buttons(2).Enabled = False
   End If
   If m_bDelete Then
       tlbar.Buttons(3).Enabled = True
   Else
       tlbar.Buttons(3).Enabled = False
   End If
'************************
   
'   'Add By Sindy 2011/6/9
'   '自動預設收件人為CP13智權人員,若為離職狀態時,則設為主管
'   If GetStaffName(m_CP13, False) <> "" Then
'      m_strTo = m_CP13
'   Else
'      '直屬主管
'      strSql = "SELECT st52 FROM staff WHERE st01='" & m_CP13 & "' "
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'      If intI = 1 Then
'         m_strTo = "" & RsTemp.Fields("st52")
'      End If
'   End If

   m_AttachPath = App.path & "\" & strUserNum 'Added by Lydia 2015/10/30
End Sub

' 按下按鍵
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

   Select Case KeyCode
      Case vbKeyF2:        ' 新增
         If m_bInsert Then
           New_Click
           txtPaperNum.SetFocus
           m_EDIT = 1
         End If
      Case vbKeyF3:        '修改
        If m_bUpdate Then
           m_EDIT = 2
           Edit_Click
        End If
      Case vbKeyF4:        '查詢
           m_EDIT = 4
           txtPaperNum.SetFocus
           QueryCourtyardPeriod
      Case vbKeyF5:        '刪除
        If m_bDelete Then
           Delete_Click
           m_EDIT = 0
        End If
      Case vbKeyHome:      '第一筆
          If m_EDIT = 0 Then
           QueryData (m_DataList(1))
           m_IndexNow = 1
           txtPaperNum.Text = m_DataList(m_IndexNow)
           Chktoolbar
         End If
       Case vbKeyPageUp:    '上一筆
           If m_EDIT = 0 Then
            m_IndexNow = m_IndexNow - 1
            QueryData (m_DataList(m_IndexNow))
            txtPaperNum.Text = m_DataList(m_IndexNow)
            tlbar.Buttons(8).Enabled = True
            tlbar.Buttons(9).Enabled = True
            Chktoolbar
          End If
      Case vbKeyPageDown:  '下一筆
          If m_EDIT = 0 Then
            m_IndexNow = m_IndexNow + 1
            QueryData (m_DataList(m_IndexNow))
            txtPaperNum.Text = m_DataList(m_IndexNow)
            tlbar.Buttons(6).Enabled = True
            tlbar.Buttons(7).Enabled = True
            Chktoolbar
          End If
      Case vbKeyEnd:       '最後一筆
          If m_EDIT = 0 Then
            QueryData (m_DataList(m_DataCount))
            m_IndexNow = m_DataCount
            txtPaperNum.Text = m_DataList(m_IndexNow)
            Chktoolbar
          End If
      Case vbKeyReturn:
       If m_EDIT <> 0 Then
        If txtPaperNum.Text = "" Then
            MsgBox "收文號不可空白!", vbExclamation, "庭期資料維護"
            txtPaperNum.SetFocus
            m_stat = 1
            Exit Sub
        End If
            cmdok_Click
          If m_stat <> 1 Then
             m_EDIT = 0
             m_stat = 0
          End If
       End If
      Case vbKeyF9:        '確定
        If m_EDIT <> 0 Then
        If txtPaperNum.Text = "" Then
            MsgBox "收文號不可空白!", vbExclamation, "庭期資料維護"
            txtPaperNum.SetFocus
            m_stat = 1
            Exit Sub
         End If
         cmdok_Click
         If m_stat <> 1 Then
            m_EDIT = 0
            m_stat = 0
         End If
        End If
      Case vbKeyF10:       '取消
      If m_EDIT <> 0 Or m_stat = 1 Then
        
       If MsgBox("你並未存檔,確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
          Exit Sub
       End If
        m_EDIT = 0
         m_stat = 0
       CmdEnabled
       blnisEdit = False
       blnIsNew = False
       blnIsCancel = True
       If blnIsSearch Then
          blnIsSearch = False
          tlbar.Buttons(11).Enabled = False
       End If
       TxtCanTUse
       If m_DataCount = 0 Then
          Cleartxt
       Else
          txtPaperNum.Text = m_DataList(m_IndexNow)
          QueryData (m_DataList(m_IndexNow))
          Chktoolbar
       End If
       End If
      Case vbKeyEscape:    '離開
            Unload Me
            m_EDIT = 0
            frm075011.Show
   End Select
   
   If m_stat = 1 Then
      Exit Sub
   End If
   
      ' Ken 90.07.16 -- Start
   If KeyCode = vbKeyReturn Or KeyCode = vbKeyF9 Or _
      KeyCode = vbKeyF10 Then
      
      If m_bInsert Then
          tlbar.Buttons(1).Enabled = True
      Else
          tlbar.Buttons(1).Enabled = False
      End If
      If m_bUpdate Then
          tlbar.Buttons(2).Enabled = True
      Else
          tlbar.Buttons(2).Enabled = False
      End If
      If m_bDelete Then
          tlbar.Buttons(3).Enabled = True
      Else
          tlbar.Buttons(3).Enabled = False
      End If
   End If
   ' Ken 90.07.16 -- End

End Sub

Private Sub Form_Unload(Cancel As Integer)
   frm075011.SearchData
   Erase m_DataList
   m_DataCount = 0
   'Add By Cheng 2002/07/18
   Set frm075012 = Nothing
End Sub

Private Sub lbeCus_Change()
 'Dim StrCusName As String
 '  If Len(lbeCus) > 8 Then
 '     If objPublicData.GetCustomer(0, lbeCus, StrCusName) Then lbeCusName = StrCusName
 '  End If
End Sub

Private Sub txtPaperNum_GotFocus()
  'edit by nickc 2007/06/11  切換輸入法改用API
  'txtPaperNum.IMEMode = 2
  CloseIme
End Sub

Private Sub txtPaperNum_KeyPress(KeyAscii As Integer)
  KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtPaperNum_LostFocus()
  txtPaperNum.Text = UCase(txtPaperNum.Text)
  If m_EDIT = 0 Then
     Exit Sub
  End If
  If txtPaperNum.Text = "" Then
     MsgBox "收文號不可空白!", vbExclamation, "庭期資料維護"
     txtPaperNum.SetFocus
     Exit Sub
  End If
  If Len(txtPaperNum.Text) <> 9 Then
     MsgBox "收文號輸入錯誤!", vbExclamation, "庭期資料維護"
     txtPaperNum.SetFocus
     Exit Sub
  End If
  If ChkPaperNum(txtPaperNum.Text) = False Then
     MsgBox "收文號輸入錯誤!", vbExclamation, "庭期資料維護"
     txtPaperNum.SetFocus
     InverseTextBox txtPaperNum
     Exit Sub
  End If
  ReadCaseprogress
End Sub

'檢查收文號
Private Function ChkPaperNum(strNum As String) As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   strSql = "SELECT CP09 FROM CASEPROGRESS WHERE " & _
            "CP01 ='" & lc01 & "' AND CP02 ='" & lc02 & "'" & _
            " AND CP03 ='" & lc03 & "' AND CP04 ='" & lc04 & "'" & _
            " AND CP09 ='" & strNum & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If Not rsTmp.EOF And Not rsTmp.BOF Then
      ChkPaperNum = True
   Else
      ChkPaperNum = False
   End If
End Function

'Modified by Lydia 2021/09/17 改成Form 2.0
'Private Sub Text_KeyPress(Index As Integer, KeyAscii As Integer)
Private Sub Text_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text_Validate(Index As Integer, Cancel As Boolean)
Dim strTemp As String

   If m_EDIT <> 1 And m_EDIT <> 2 Then Exit Sub
   If Not (blnisEdit Or blnIsNew) Then Exit Sub
   Select Case Index
'      Case 0
'          If Text(0) <> "" Then
'          If CheckIsTaiwanDate(Text(0)) Then
'      '       If Val(Today) - Val(Text(0)) < 0 Then
'      '           MsgBox "收受日大於系統日", vbCritical
'      '           Cancel = True
'      '       End If
'          Else
'              Cancel = True
'              DataErrorMessage 5, "收受日"
'          End If
'          End If
'      Case 1
'          If Text(1) <> "" Then
'             'edit by nickc 2007/02/07 不用 dll 了
'             'If objLawDll.GetGovName(Text(1), strTemp) Then
'             If ClsPDGetGovName(Text(1), strTemp) Then
'                lbeGov = strTemp
'             Else
'               Cancel = True
'             End If
'          End If
'          If Text(1) = "" Then lbeGov = ""
      Case 4
          If Text(Index) <> "" Then
               'Modify by Amy +5.調解
               If Not (Text(Index) = "1" Or Text(Index) = "2" Or Text(Index) = "3" Or Text(Index) = "4" Or Text(Index) = "5") Then
                  DataErrorMessage 1, "開庭種類"
                  Cancel = True
                End If
          End If
      'Add By Sindy 2011/10/20
      Case 0
          If Text(Index) <> "" Then
               'Modify by Amy 2018/01/24 +6.調解庭
               If Not (Text(Index) = "1" Or Text(Index) = "2" Or Text(Index) = "3" Or Text(Index) = "4" Or Text(Index) = "5" Or Text(Index) = "6") Then
                  DataErrorMessage 1, "開庭別"
                  Cancel = True
                End If
          End If
      Case 2
        If Text(Index) <> "" Then Text(Index) = UCase(Text(Index))
           
      Case 5, 1
          If Text(Index) <> "" Then
              If CheckIsTaiwanDate(Text(Index)) Then
      '           If Val(Today) - Val(Text(index)) > 0 Then
      '               MsgBox "開庭日期小於系統日", vbCritical
      '               Cancel = True
      '           End If
              Else
                 Cancel = True
              End If
          End If
      Case 6
             If Text(Index) <> "" Then
              If Len(Text(Index)) = 4 Then
                If Not ChkTime(Text(Index)) Then Cancel = True
              Else
                 MsgBox "時間輸入格式錯誤!", vbExclamation
                 Cancel = True
              End If
           End If
      Case 7
          If Text(7).Text <> "" Then
             If CheckLengthIsOK(Text(7), 2000) = False Then
                Cancel = True
             End If
          End If
   
   End Select
   If Not (blnisEdit Or blnIsNew) Then Cancel = False
   
   If Cancel Then TextInverse Text(Index)

End Sub

Private Sub Text_GotFocus(Index As Integer)
   Select Case Index
          Case 7
              'edit by nickc 2007/06/11  切換輸入法改用API
              'Text(Index).IMEMode = 1
              OpenIme
          Case Else
              'edit by nickc 2007/06/11  切換輸入法改用API
              'Text(Index).IMEMode = 2
              CloseIme
   End Select
   TextInverse Text(Index)
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   'edit by nickc 2007/06/11  切換輸入法改用API
   'Text1(Index).IMEMode = 1
   OpenIme
   TextInverse Text1(Index)
End Sub

'Modified by Lydia 2021/09/17 改成Form 2.0
'Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_GotFocus(Index As Integer)
   'edit by nickc 2007/06/11  切換輸入法改用API
   'Text2(Index).IMEMode = 1
   OpenIme
   TextInverse Text2(Index)
End Sub

'Modified by Lydia 2021/09/17 改成Form 2.0
'Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
Private Sub Text2_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub QueryCourtyardPeriod()
   blnIsNew = True
   Cleartxt
   TxtCanUse
   CmdUnabled
   intSaveKind = 4
End Sub

Private Sub tlbar_ButtonClick(ByVal Button As MSComctlLib.Button)

   Select Case Button.Index
       Case 1
            New_Click
            txtPaperNum.SetFocus
            m_EDIT = 1
       Case 2
   '      'Add By Sindy 2011/6/15
   '      If m_cdp10 <> strUserNum and Pub_StrUserSt03<>"M51" Then
   '         MsgBox "無權限可修改資料!", vbInformation, "庭期資料維護"
   '         Exit Sub
   '      End If
         
          m_EDIT = 2
          Edit_Click
       Case 3
          Delete_Click
          m_EDIT = 0
       Case 4
          m_EDIT = 4
          txtPaperNum.SetFocus
          QueryCourtyardPeriod
       Case 6
         If m_EDIT = 0 Then
          QueryData (m_DataList(1))
          m_IndexNow = 1
          txtPaperNum.Text = m_DataList(m_IndexNow)
          Chktoolbar
         End If
   '       rsTemp.MoveFirst
   '       PutDataInObject
       Case 7
          If m_EDIT = 0 Then
          m_IndexNow = m_IndexNow - 1
          QueryData (m_DataList(m_IndexNow))
          txtPaperNum.Text = m_DataList(m_IndexNow)
          tlbar.Buttons(8).Enabled = True
          tlbar.Buttons(9).Enabled = True
          Chktoolbar
          End If
   '       rsTemp.MovePrevious
   '       If rsTemp.BOF Then
   '          rsTemp.MoveFirst
   '          PutDataInObject
   '          DataErrorMessage (6)
   '       End If
   '       PutDataInObject
       Case 8
   '       rsTemp.MoveNext
   '       If rsTemp.EOF Then
   '          rsTemp.MoveLast
   '          PutDataInObject
   '          DataErrorMessage (7)
   '       End If
   '       PutDataInObject
        If m_EDIT = 0 Then
          m_IndexNow = m_IndexNow + 1
          QueryData (m_DataList(m_IndexNow))
          txtPaperNum.Text = m_DataList(m_IndexNow)
          tlbar.Buttons(6).Enabled = True
          tlbar.Buttons(7).Enabled = True
          Chktoolbar
        End If
       Case 9
        If m_EDIT = 0 Then
          QueryData (m_DataList(m_DataCount))
          m_IndexNow = m_DataCount
          txtPaperNum.Text = m_DataList(m_IndexNow)
          Chktoolbar
        End If
   '       rsTemp.MoveLast
   '       PutDataInObject
       Case 11
         If txtPaperNum.Text = "" Then
            MsgBox "收文號不可空白!", vbExclamation, "庭期資料維護"
            txtPaperNum.SetFocus
            Exit Sub
          End If
          cmdok_Click
          If m_stat <> 1 Then
             m_stat = 0
             m_EDIT = 0
          End If
       Case 12
         
          If MsgBox("你並未存檔,確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
             Exit Sub
          End If
           m_EDIT = 0
          m_stat = 0
          CmdEnabled
          blnisEdit = False
          blnIsNew = False
          blnIsCancel = True
          If blnIsSearch Then
             blnIsSearch = False
             tlbar.Buttons(11).Enabled = False
          End If
          TxtCanTUse
          If m_DataCount = 0 Then
             Cleartxt
          Else
             txtPaperNum.Text = m_DataList(m_IndexNow)
             QueryData (m_DataList(m_IndexNow))
             Chktoolbar
          End If
          'PutDataInObject
       Case 14
           m_EDIT = 0
           Unload Me
          ' frm075011.cmdSearch
           frm075011.Show
   End Select

   If m_stat = 1 Then
      Exit Sub
   End If

   '*********************
   If Button.Index <> 14 And Button.Index <> 1 And _
      Button.Index <> 2 And Button.Index <> 3 And _
      Button.Index <> 4 Then

      If m_bInsert Then
          tlbar.Buttons(1).Enabled = True
      Else
          tlbar.Buttons(1).Enabled = False
      End If
      If m_bUpdate Then
          tlbar.Buttons(2).Enabled = True
      Else
          tlbar.Buttons(2).Enabled = False
      End If
      If m_bDelete Then
          tlbar.Buttons(3).Enabled = True
      Else
          tlbar.Buttons(3).Enabled = False
      End If
   End If
'****************************

End Sub

Private Sub CmdEnabled()
Dim i As Integer
   
   tlbar.Buttons(11).Enabled = False
   tlbar.Buttons(12).Enabled = False
   For i = 1 To 9
      tlbar.Buttons(i).Enabled = True
   Next
   tlbar.Buttons(14).Enabled = True
   'Add By Sindy 2011/6/15
   Command1.Enabled = True
   lstMailCC.Enabled = True
'   cmdOpenAtt.Enabled = True
'   cmdAddAtt.Enabled = False
'   cmdRemAtt.Enabled = False
   '2011/6/15 End
   cmdFile.Enabled = False 'Added by Lydia 2015/10/30
   'Add By Sindy 2016/6/27
   cmdBrief.Enabled = False
   cmdNote.Enabled = False
   '2016/6/27 END
   Command4.Enabled = False 'Added by Lydia 2024/07/29
End Sub

Private Sub CmdUnabled()
Dim i As Integer
   
   tlbar.Buttons(11).Enabled = True
   tlbar.Buttons(12).Enabled = True
   For i = 1 To 9
      tlbar.Buttons(i).Enabled = False
   Next
   tlbar.Buttons(14).Enabled = False
   tlbar.Buttons(4).Enabled = False
   'Add By Sindy 2011/6/15
   Command1.Enabled = False
   lstMailCC.Enabled = False
'   cmdOpenAtt.Enabled = False
   If m_EDIT = 4 Then '查詢
'      cmdAddAtt.Enabled = False
'      cmdRemAtt.Enabled = False
   Else
'      If m_cdp10 = strUserNum Or Pub_StrUserSt03 = "M51" Then
'         cmdAddAtt.Enabled = True
         'Modify By Sindy 2011/8/11
         'cmdRemAtt.Enabled = True
'         If Pub_StrUserSt03 = "M51" Then cmdRemAtt.Enabled = True
'      End If
      cmdFile.Enabled = True 'Added by Lydia 2015/10/30
      'Add By Sindy 2016/6/27
      cmdBrief.Enabled = True
      cmdNote.Enabled = True
      '2016/6/27 END
      Command4.Enabled = True 'Added by Lydia 2024/07/29
   End If
   '2011/6/15 End
End Sub

Private Sub TxtCanTUse()
Dim i As Integer
   
   cboPerson.Locked = True
   txtPaperNum.Locked = True
   For i = 0 To 2
      Text1(i).Locked = True
      Text2(i).Locked = True
   Next
   For i = 0 To 7
      If i <> 2 And i <> 3 Then
         Text(i).Locked = True
      End If
   Next
End Sub

Private Sub TxtCanUse()
Dim i As Integer
   
   txtPaperNum.Locked = False
   cboPerson.Locked = False
   For i = 0 To 2
      Text1(i).Locked = False
      Text2(i).Locked = False
   Next
   For i = 0 To 7
      If i <> 2 And i <> 3 Then
         Text(i).Locked = False
      End If
   Next
End Sub

Private Sub GetData()
Dim num As String, i As Integer, n As Integer
Dim nRow As Integer
   n = 1
   'With frm075011.MSHFlexGrid1
   '    For i = 1 To .Rows - 1
   '    .Col = 0
   '    .Row = i
   '     If .Text = "v" Then
   '        .Col = 2
   '        If n = 1 Then
   '           PaperNum = "'" + .Text + "'"
   '           n = n + 1
   '        Else:
   '           PaperNum = PaperNum + "," + "'" + .Text + "'"
   '           n = n + 1
   '        End If
   '     End If
   '    Next
   'End With
   With frm075011.MSHFlexGrid1
      m_DataCount = 0
      For nRow = 1 To .Rows - 1
         .col = 0
         .row = nRow
         If .Text = "v" Then
            ReDim Preserve m_DataList(nRow)
            m_DataList(n) = .TextMatrix(nRow, 2)
            m_DataCount = m_DataCount + 1
            n = n + 1
         End If
      Next nRow
   End With
   lc01 = frm075011.txtCP01
   lc02 = frm075011.txtCP02
   lc03 = IIf(frm075011.txtCP03 = "", "0", frm075011.txtCP03.Text)
   lc04 = IIf(frm075011.txtCP04 = "", "00", frm075011.txtCP04.Text)
   'Added by Lydia 2015/10/30
   m_CP01 = lc01: m_CP02 = lc02
   m_CP03 = lc03: m_CP04 = lc04
   
   lbeCaseNum = GiveSymbol(lc01, lc02, lc03, lc04, LcTmp)
   lbeCus = frm075011.lbeCus
   lbeCusName = frm075011.lbeCusName
   
   'Modified by Lydia 2021/09/17
   'For i = 0 To frm075011.cboCaseName.ListCount
   For i = 0 To frm075011.cboCaseName.ListCount - 1
      cboCaseName.AddItem frm075011.cboCaseName.List(i)
   Next
   cboCaseName.ListIndex = 0
   If m_DataCount <> 0 Then
      txtPaperNum.Text = m_DataList(1)
      QueryData (m_DataList(1))
      m_IndexNow = 1
      Chktoolbar
   End If
   'Added by Lydia 2024/07/29
   m_CL02 = "": Me.Tag = ""
   If txtPaperNum <> "" Then
      strSql = "select CL02 from CaseLawer where CL01='" & txtPaperNum & "' order by CL02"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         m_CL02 = RsTemp.GetString(adClipString, , , ",")
      End If
   End If
   'end 2024/07/29
End Sub

Private Sub Chktoolbar()
    If m_IndexNow = m_DataCount Then
       tlbar.Buttons(6).Enabled = True
       tlbar.Buttons(7).Enabled = True
       tlbar.Buttons(8).Enabled = False
       tlbar.Buttons(9).Enabled = False
    ElseIf m_IndexNow = 1 Then
       tlbar.Buttons(6).Enabled = False
       tlbar.Buttons(7).Enabled = False
       tlbar.Buttons(8).Enabled = True
       tlbar.Buttons(9).Enabled = True
    End If
    If m_DataCount = 1 Then
       tlbar.Buttons(6).Enabled = False
       tlbar.Buttons(7).Enabled = False
       tlbar.Buttons(8).Enabled = False
       tlbar.Buttons(9).Enabled = False
    End If
End Sub

Private Function ChkInsertData() As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim intI As Integer
  
  intI = 0
  strSql = "SELECT * FROM COURTYARDPERIOD WHERE CDP01 ='" & txtPaperNum.Text & "'"
  rsTmp.CursorLocation = adUseClient
  rsTmp.Open strSql, cnnConnection, adOpenDynamic, adLockReadOnly
 
  If rsTmp.EOF = False Then
     ChkInsertData = True
     MsgBox "已有資料,不可重複新增!", vbInformation, "庭期資料維護"
  Else
     ChkInsertData = False
  End If
  rsTmp.Close
  Set rsTmp = Nothing
End Function

'存檔
Private Function insertdata() As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim strLaw As String
Dim nPos As Integer
Dim strPerson As String
Dim iErr As Integer, sErrMsg As String
   
   insertdata = True
On Error Resume Next
   
   'SetList lstAtt, txtCR(9) 'Add By Sindy 2011/6/15
   'Modify By Sindy 2011/6/15 +CDP16
   'Modify By Sindy 2011/10/20 +CDP17+CDP18
   'Modify By Sindy 2016/6/27 取消,CDP16
   strSql = "INSERT INTO COURTYARDPERIOD (CDP01,CDP02,CDP03,CDP04,CDP05,CDP06,CDP07,CDP08,CDP09," & _
           "CDP10,CDP11,CDP12,CDP17,CDP18) values ("
   '收文號
   If txtPaperNum.Text <> "" Then
      strSql = strSql & "'" & txtPaperNum.Text & "'"
   Else
       strSql = strSql & "NULL"
   End If
   strSql = strSql & ","
   '開庭人員
   If cboPerson.Text <> "" Then
      nPos = 0
      nPos = InStr(cboPerson.Text, ",")
      If nPos <> 0 Then
         strPerson = Left(cboPerson.Text, nPos - 1)
         strSql = strSql & "'" & strPerson & "'"
      End If
   Else
       strSql = strSql & "NULL"
   End If
   strSql = strSql & ","
   '開庭日期
   If Text(5).Text <> "" Then
      strSql = strSql & "'" & ChangeTStringToWString(Text(5).Text) & "'"
   Else
       strSql = strSql & "NULL"
   End If
   strSql = strSql & ","
   '時間
   If Text(6).Text <> "" Then
      strSql = strSql & "'" & Text(6).Text & "'"
   Else
       strSql = strSql & "NULL"
   End If
   strSql = strSql & ","
   '機關代號
   If lblNum.Caption <> "" Then
      strSql = strSql & "'" & lblNum.Caption & "'"
   Else
       strSql = strSql & "NULL"
   End If
   strSql = strSql & ","
   '開庭種類
   If Text(4).Text <> "" Then
      strSql = strSql & "'" & Text(4).Text & "'"
   Else
       strSql = strSql & "NULL"
   End If
   strSql = strSql & ","
   '備註
   If Text(7).Text <> "" Then
      strSql = strSql & "'" & Text(7).Text & "'"
   Else
       strSql = strSql & "NULL"
   End If
   strSql = strSql & ","
   
   strLaw = Text1(0) + IIf(Text1(1) = "", "", "," + Text1(1)) + IIf(Text1(2) = "", "", "," + Text1(2))
   '法官
   If strLaw <> "" Then
      strSql = strSql & "'" & strLaw & "'"
   Else
       strSql = strSql & "NULL"
   End If
   strSql = strSql & ","
   
   strLaw = ""
   strLaw = Text2(0) + IIf(Text2(1) = "", "", "," + Text2(1)) + IIf(Text2(2) = "", "", "," + Text2(2))
   '檢察官
   If strLaw <> "" Then
      strSql = strSql & "'" & strLaw & "'"
   Else
       strSql = strSql & "NULL"
   End If
   strSql = strSql & ","
   
   If strUserNum <> "" Then
      strSql = strSql & "'" & strUserNum & "'"
   Else
       strSql = strSql & "NULL"
   End If
   strSql = strSql & ","
   
   strSql = strSql & "'" & ChangeWStringToTString(GetTodayDate) & "'"
   strSql = strSql & ","
   strSql = strSql & "'" & Format(time, "HHMM") & "'"
   
   'Add By Sindy 2011/6/15
'   '附件檔名
'   strSql = strSql & ","
'   If txtCR(9) <> "" Then
'      strSql = strSql & "'" & txtCR(9) & "'"
'   Else
'      strSql = strSql & "NULL"
'   End If
   
   'Add By Sindy 2011/10/20
   strSql = strSql & ","
   '開庭別
   If Text(0).Text <> "" Then
      strSql = strSql & "'" & Text(0).Text & "'"
   Else
       strSql = strSql & "NULL"
   End If
   strSql = strSql & ","
   '取消庭期日期
   If Text(1).Text <> "" Then
      strSql = strSql & "'" & Text(1).Text & "'"
   Else
       strSql = strSql & "NULL"
   End If
   '2011/10/20 End
   
   strSql = strSql & ")"
           
   Pub_SeekTbLog strSql 'Add By Sindy 2011/6/15
   cnnConnection.Execute strSql
   
   cnnConnection.CommitTrans
   
'   'Add By Sindy 2011/6/24
'   '上傳附件檔
'   If UploadAtt(txtPaperNum.Text, iErr, sErrMsg) = False Then
'      GoTo ErrHand
'   End If
   'Modify By Sindy 2016/6/27 上傳開庭紀要附件存到卷宗區
   If m_strSaveFilesBRIEF <> "" Then
       If PUB_UpdReplyFile(m_strSaveFilesBRIEF, txtPaperNum.Text, m_CP01, m_CP02, m_CP03, m_CP04, 通知開庭, , UCase("BRIEF")) = False Then Exit Function
       '刪除匯入來源檔
       Call PUB_DelPCOrgFile(m_strSaveFilesBRIEF): m_strSaveFilesBRIEF = ""
   End If
   '2016/6/27 END
   
   'Added by Lydia 2015/10/30 上傳開庭通知附件存在卷宗區
   If m_strSaveFilesOA <> "" Then
       If IsRecordExist(txtPaperNum) = False Then
           If PUB_UpdReplyFile(m_strSaveFilesOA, txtPaperNum.Text, m_CP01, m_CP02, m_CP03, m_CP04, 通知開庭, , "OA") = False Then
'              cmdFile.Enabled = False
              Exit Function
           End If
           '刪除匯入來源的回覆單
           Call PUB_DelPCOrgFile(m_strSaveFilesOA): m_strSaveFilesOA = ""
       End If
   End If
   'end 2015/10/30
   
   'Add By Sindy 2016/6/27 上傳電子筆錄附件存到卷宗區
   If m_strSaveFilesNOTE <> "" Then
       If PUB_UpdReplyFile(m_strSaveFilesNOTE, txtPaperNum.Text, m_CP01, m_CP02, m_CP03, m_CP04, 通知開庭, , UCase("NOTE")) = False Then Exit Function
       '刪除匯入來源檔
       Call PUB_DelPCOrgFile(m_strSaveFilesNOTE): m_strSaveFilesNOTE = ""
   End If
   '2016/6/27 END
   
   Exit Function
   
ErrHand:
   insertdata = False
   cnnConnection.RollbackTrans
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   ElseIf iErr <> 0 Then
      MsgBox sErrMsg, vbCritical
   End If
End Function

Private Sub Cleartxt()
Dim i As Integer
   
   cboPerson = ""
   lbePerson = ""
   lbeGov = ""
   For i = 0 To 7
      If i <> 2 And i <> 3 Then
         Text(i) = ""
      End If
   Next
   For i = 0 To 2
      Text1(i) = ""
      Text2(i) = ""
   Next
   txtPaperNum.Text = ""
   'Add By Sindy 2011/6/14
'   txtCR(9) = ""
'   lstAtt.Clear
   lblRevDate = ""
   lblNum = ""
   lbeGov = ""
   lblLawNum = ""
   lblLNum = ""
   lstMailCC.Clear
   
   'm_strSaveFiles = "" 'Added by Lydia 2015/10/30
   m_strSaveFileType = "" 'Add By Sindy 2016/6/24
   m_strSaveFilesOA = "" 'Add By Sindy 2016/6/24
   m_strSaveFilesBRIEF = "" 'Add By Sindy 2016/6/24
   m_CL02 = "": Me.Tag = "" 'Added by Lydia 2024/07/29
End Sub

Private Sub Edit_Click()
   TxtCanUse
   CmdUnabled
'   Text(0).SetFocus
   blnisEdit = True
   intSaveKind = 2
   txtPaperNum.Locked = True
End Sub

Private Sub Delete_Click()
Dim yn As Integer, Del As Boolean
Dim strSql As String
Dim intFiles As Integer
Dim strTmp() As String
Dim i As Integer
Dim nMode As Integer
Dim strData As String
Dim iErr As Integer, sErrMsg As String
   
   'Add By Sindy 2011/6/15
   If m_cdp10 <> strUserNum And Pub_StrUserSt03 <> "M51" Then
      MsgBox "無權限可刪除資料!", vbInformation, "庭期資料維護"
      Exit Sub
   End If
   
   nMode = 0
   CmdUnabled
   
'   If MsgBox("是否要刪除此筆資料?", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbYes Then
'      strExc(1) = "delete courtyardperiod where cdp01=" & CNULL(txtPaperNum) & " "
'      Del = objLawDll.ExecSQL(1, strExc)
'      If Del Then
'         MsgBox "'" + txtPaperNum + "'  刪除成功"
'         Set rsTemp = Nothing
'         GetRsData
'         rsTemp.MoveFirst
'         PutDataInObject
'      Else
'         MsgBox "'" + txtPaperNum + "'  刪除不成功"
'      End If
'   End If
   
   'Add By Sindy 2011/8/11
   'Modify By Sindy 2016/6/27 改判斷卷宗區
   'strSql = "SELECT * FROM COURTYARDPERIOD WHERE CDP01='" & txtPaperNum.Text & "'"
   strSql = "SELECT count(*) FROM CasePaperPDF WHERE CPP01='" & txtPaperNum.Text & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   intFiles = 0
   If intI = 1 Then
      intFiles = Val(RsTemp(0))
   End If
   '2011/8/11 End
   If MsgBox(IIf(intFiles > 0, "有附件", "") & "是否要刪除此筆資料?", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbYes Then
      strSql = "DELETE FROM COURTYARDPERIOD WHERE CDP01='" & txtPaperNum.Text & "'"
      Pub_SeekTbLog strSql 'Add By Sindy 2011/6/15
      cnnConnection.Execute strSql
      
      'Add By Sindy 2011/8/11
'      '檔案有異動時，移掉的要刪除
'      If strFiles <> "" Then
'         If RemoveAtt(txtPaperNum.Text, strFiles, iErr, sErrMsg) = False Then
'            GoTo ErrHand
'         End If
'      End If
'      '2011/8/11 End
      'Modify By Sindy 2016/6/27
      PUB_DelFtpFile2 txtPaperNum.Text '檔案改放 FTP,必須在DB資料刪除前執行
      strSql = "delete from casepaperpdf where cpp01='" & txtPaperNum.Text & "'"
      cnnConnection.Execute strSql, intI
      '2016/6/27 END
      
      If Err.Number = 0 Then
        ' MsgBox "'" + txtPaperNum + "'  刪除成功"
         If m_DataCount = 1 Then
          '  Cleartxt
          '  Exit Sub
             Unload Me
             frm075011.Show
             Exit Sub
         End If
         m_DataList(m_IndexNow) = ""
         If m_IndexNow = m_DataCount Then
            strData = m_DataList(1)
         Else
            strData = m_DataList(m_IndexNow + 1)
         End If
         
         For i = 1 To m_DataCount
             If nMode = 1 Then
                If i <> m_DataCount Then
                   ReDim Preserve strTmp(i)
                   strTmp(i) = m_DataList(i + 1)
                End If
             Else
                  If m_DataList(i) <> "" Then
                     ReDim Preserve strTmp(i)
                     strTmp(i) = m_DataList(i)
                  ElseIf i <> m_DataCount Then
                     ReDim Preserve strTmp(i)
                     strTmp(i) = m_DataList(i + 1)
                     nMode = 1
                  End If
             End If
         Next i
         Erase m_DataList
         For i = 1 To UBound(strTmp)
             ReDim Preserve m_DataList(i)
             m_DataList(i) = strTmp(i)
         Next
         m_DataCount = UBound(strTmp)
         For i = 1 To m_DataCount
             If m_DataList(i) = strData Then
                m_IndexNow = i
                Exit For
             End If
         Next
         
         QueryData (m_DataList(m_IndexNow))
         txtPaperNum.Text = m_DataList(m_IndexNow)
         CmdEnabled
         TxtCanTUse
         Chktoolbar
      End If
    End If
    
    Exit Sub
    
ErrHand:
   cnnConnection.RollbackTrans
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   ElseIf iErr <> 0 Then
      MsgBox sErrMsg, vbCritical
   End If
End Sub

'修改
Private Function UpdateData() As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim strLaw As String
Dim strLaw1 As String
Dim nPos As Integer
Dim strPerson As String
Dim iErr As Integer, sErrMsg As String, bolRemove As Boolean
Dim arrFile1, ii As Integer
  
   UpdateData = True
  
On Error Resume Next
  
   nPos = 0
   
   If cboPerson.Text <> "" Then
      nPos = InStr(cboPerson.Text, ",")
      If nPos <> 0 Then
         strPerson = Left(cboPerson.Text, nPos - 1)
      End If
   End If
   
   UDT = GetTodayDate
   UTM = Format(time, "HHMM")
   strLaw = Text1(0) + IIf(Text1(1) = "", "", "," + Text1(1)) + IIf(Text1(2) = "", "", "," + Text1(2))
   strLaw1 = Text2(0) + IIf(Text2(1) = "", "", "," + Text2(1)) + IIf(Text2(2) = "", "", "," + Text2(2))
   'SetList lstAtt, txtCR(9) 'Add By Sindy 2011/6/15
   'Modify By Sindy 2011/6/15 +CDP16
   'Modify By Sindy 2011/6/15 +CDP17+CDP18
   'Modify By Sindy 2016/6/27 +"CDP16 ='" & txtCR(9) & "',"
   strSql = "begin user_data.user_enabled:=1; UPDATE COURTYARDPERIOD SET CDP02 ='" & strPerson & "'," & _
           "CDP03 ='" & ChangeTStringToWString(Text(5).Text) & "'," & _
           "CDP04 ='" & Text(6).Text & "'," & _
           "CDP05 ='" & lblNum.Caption & "'," & _
           "CDP06 ='" & Text(4).Text & "'," & _
           "CDP07 ='" & Text(7).Text & "'," & _
           "CDP08 ='" & strLaw & "'," & _
           "CDP09 ='" & strLaw1 & "'," & _
           "CDP13 ='" & strUserNum & "'," & _
           "CDP14 ='" & UDT & "'," & _
           "CDP15 ='" & UTM & "'," & _
           "CDP17 ='" & Text(0).Text & "'," & _
           "CDP18 ='" & ChangeTStringToWString(Text(1).Text) & "' WHERE CDP01 = '" & txtPaperNum.Text & "' ; end ;"
   cnnConnection.BeginTrans
   Pub_SeekTbLog strSql 'Add By Sindy 2011/6/15
   cnnConnection.Execute strSql
   
   'Added by Lydia 2024/07/29
   If Me.Tag <> "" And InStr(Me.Tag, "|") > 0 Then '有點選「出庭律師」
      If PUB_SaveCaseLawer(txtPaperNum, Mid(Me.Tag, InStr(Me.Tag, "|") + 1), m_CL02) = True Then
         m_CL02 = Me.Tag
      End If
   End If
   'end 2024/07/29
   cnnConnection.CommitTrans
   
'   'Add By Sindy 2011/6/24
'   '上傳附件檔
'   If UploadAtt(txtPaperNum.Text, iErr, sErrMsg) = False Then
'      GoTo ErrHand
'   End If
'   '檔案有異動時，移掉的要刪除
'   bolRemove = False
'   If txtCR(9) <> m_cdp16 Then
'      arrFile1 = Split(m_cdp16, ",")
'      For ii = LBound(arrFile1) To UBound(arrFile1)
'         If InStr(txtCR(9) & ",", arrFile1(ii) & ",") > 0 Then
'            arrFile1(ii) = ""
'         Else
'            bolRemove = True
'         End If
'      Next
'      If bolRemove = True Then
'         If RemoveAtt(txtPaperNum.Text, Join(arrFile1, ","), iErr, sErrMsg) = False Then
'            GoTo ErrHand
'         End If
'      End If
'   End If
   'Modify By Sindy 2016/6/27 上傳開庭紀要附件存到卷宗區
   If m_strSaveFilesBRIEF <> "" Then
       If PUB_UpdReplyFile(m_strSaveFilesBRIEF, txtPaperNum.Text, m_CP01, m_CP02, m_CP03, m_CP04, 通知開庭, , UCase("BRIEF")) = False Then Exit Function
       '刪除匯入來源檔
       Call PUB_DelPCOrgFile(m_strSaveFilesBRIEF): m_strSaveFilesBRIEF = ""
   End If
   '2016/6/27 END
   
   'Added by Lydia 2015/10/30 上傳開庭通知附件存在卷宗區
   If m_strSaveFilesOA <> "" Then
       If IsRecordExist(txtPaperNum) = False Then
           If PUB_UpdReplyFile(m_strSaveFilesOA, txtPaperNum.Text, m_CP01, m_CP02, m_CP03, m_CP04, 通知開庭, , "OA") = False Then
              cmdFile.Enabled = False
              Exit Function
           End If
           '刪除匯入來源的回覆單
           Call PUB_DelPCOrgFile(m_strSaveFilesOA): m_strSaveFilesOA = ""
       End If
   End If
   'end 2015/10/30
   
   'Add By Sindy 2016/6/27 上傳電子筆錄附件存到卷宗區
   If m_strSaveFilesNOTE <> "" Then
       If PUB_UpdReplyFile(m_strSaveFilesNOTE, txtPaperNum.Text, m_CP01, m_CP02, m_CP03, m_CP04, 通知開庭, , UCase("NOTE")) = False Then Exit Function
       '刪除匯入來源檔
       Call PUB_DelPCOrgFile(m_strSaveFilesNOTE): m_strSaveFilesNOTE = ""
   End If
   '2016/6/27 END
   
   Exit Function
  
ErrHand:
   UpdateData = False
   cnnConnection.RollbackTrans
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   ElseIf iErr <> 0 Then
      MsgBox sErrMsg, vbCritical
   End If
End Function

Private Sub cmdok_Click()
Dim yn As Integer, i As Integer
Dim n_Find As Integer
 
   n_Find = 0
   If CheckText Then
      m_stat = 1
      Exit Sub
   End If
   If intSaveKind = 1 Then '新增
      If ChkInsertData = False Then
         'Add By Cheng 2002/05/24
         '重新檢查欄位有效性
         If TxtValidate = False Then Exit Sub
         
         If insertdata = True Then
            ReDim Preserve m_DataList(m_DataCount + 1)
            m_DataList(m_DataCount + 1) = txtPaperNum.Text
            QueryData (m_DataList(m_DataCount + 1))
            m_DataCount = m_DataCount + 1
            m_IndexNow = m_DataCount
            CmdEnabled
            TxtCanTUse
            Chktoolbar
         Else
            MsgBox "無法新增!", vbInformation, "庭期資料維護"
            m_stat = 1
            Exit Sub
         End If
       Else
          m_stat = 1
          txtPaperNum.SetFocus
          TextInverse txtPaperNum
          Exit Sub
       End If
   ElseIf intSaveKind = 2 Then '修改
      'Add By Cheng 2002/05/24
      '重新檢查欄位有效性
      If TxtValidate = False Then Exit Sub
      
       If UpdateData = True Then
            QueryData (m_DataList(m_IndexNow))
            CmdEnabled
            TxtCanTUse
            Chktoolbar
       Else
          MsgBox "無法修改!", vbInformation, "庭期資料維護"
          m_stat = 1
          Exit Sub
       End If
    ElseIf intSaveKind = 4 Then '查詢
          If txtPaperNum.Text = "" Then
             MsgBox "請輸入收文號!", vbInformation, "庭期資料維護"
             m_stat = 1
             Exit Sub
          End If
          QueryData (txtPaperNum.Text)
                 
          If m_DataCount <> 0 Then
             For i = 1 To m_DataCount
                 If m_DataList(i) = txtPaperNum.Text Then
                    n_Find = 1
                    m_IndexNow = i
                    Exit For
                 End If
              Next i
              If n_Find = 0 Then
                 ReDim Preserve m_DataList(m_DataCount + 1)
                 m_DataList(m_DataCount + 1) = txtPaperNum.Text
                 m_DataCount = m_DataCount + 1
                 m_IndexNow = m_DataCount
              End If
           Else
                 ReDim Preserve m_DataList(m_DataCount + 1)
                 m_DataList(m_DataCount + 1) = txtPaperNum.Text
                 m_DataCount = m_DataCount + 1
                 m_IndexNow = m_DataCount
           End If
            CmdEnabled
            TxtCanTUse
            Chktoolbar
   End If
   m_stat = 0
End Sub

Private Function DisHuman(n As Integer, strData As String) As Boolean
Dim i As Integer, j As Integer, strTemp() As String, t As Integer, DisData As Variant

   j = 1
   If strData <> "" Then strData = strData & "," 'Add By Sindy 2017/12/14
   DisData = Split(strData, ",")
   Select Case n
    Case 1
       For t = 0 To UBound(DisData) - 1
          If Trim(DisData(t)) <> "" Then 'Added by Lydia 2018/05/03 判斷非空白
              Text1(t) = DisData(t)
          End If
       Next
   Case 2
       For t = 0 To UBound(DisData) - 1
          If Trim(DisData(t)) <> "" Then 'Added by Lydia 2018/05/03 判斷非空白
              Text2(t) = DisData(t)
          End If
       Next
   End Select
   DisHuman = True
End Function

Private Function CheckText() As Boolean
'  If m_EDIT = 1 Then
'      If txtPaperNum.Text = "" Then
'         MsgBox "收文號不可空白!", vbExclamation, "庭期資料維護"
'         txtPaperNum.SetFocus
'         m_stat = 1
'         Exit Function
'      End If
'   End If
   If m_EDIT = 4 Then Exit Function
    'If Text(0) = "" Or IsNull(Text(0)) Then MsgBox "收受日不可空白", vbCritical: CheckText = True: Exit Function
    'If Text(1) = "" Or IsNull(Text(1)) Then MsgBox "機關代號不可空白", vbCritical: CheckText = True: Exit Function
    'Add By Sindy 2011/10/20
    If Text(0) = "" Or IsNull(Text(0)) Then
       MsgBox "開庭別不可空白", vbCritical
       CheckText = True
       m_stat = 1
       Text(0).SetFocus
       Exit Function
    End If
    '2011/10/20 End
    If Text(4) = "" Or IsNull(Text(4)) Then
       MsgBox "開庭種類不可空白", vbCritical
       CheckText = True
       m_stat = 1
       Text(4).SetFocus
       Exit Function
    End If
    If Text(5) = "" Or IsNull(Text(5)) Then
        MsgBox "開庭日不可空白", vbCritical
        CheckText = True
        m_stat = 1
        Text(5).SetFocus
        Exit Function
    End If
    If Text(6) = "" Or IsNull(Text(6)) Then
       MsgBox "時間不可空白", vbCritical
       CheckText = True
       Text(6).SetFocus
       Exit Function
    End If
    If cboPerson = "" Then
        MsgBox "請選擇開庭人員", vbCritical
        CheckText = True
        cboPerson.SetFocus
        Exit Function
    End If
End Function

Private Sub New_Click()
   blnIsNew = True
   Cleartxt
   TxtCanUse
   CmdUnabled
   intSaveKind = 1
End Sub

Private Sub QueryData(strNum As String)
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim strTemp As String
   
   'Add By Sindy 2011/6/27
   Me.Enabled = False
   Cleartxt
   txtPaperNum = strNum
   '2011/6/27 End
   strSql = "SELECT * FROM COURTYARDPERIOD WHERE CDP01 ='" & strNum & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If Not rsTmp.EOF Then
      '開庭人員
      If Not IsNull(rsTmp.Fields("CDP02")) Then
         'edit by nickc 2007/02/07 不用 dll 了
         'If objPublicData.GetStaff(rsTmp.Fields("CDP02"), strTemp) Then
         If ClsPDGetStaff(rsTmp.Fields("CDP02"), strTemp) Then
            lbePerson = strTemp
            cboPerson.Text = rsTmp.Fields("CDP02") & "," & strTemp
         End If
      End If
      '開庭日期
      If Not IsNull(rsTmp.Fields("CDP03")) Then
         Text(5).Text = ChangeWStringToTString(rsTmp.Fields("CDP03"))
      End If
      '時間
      If Not IsNull(rsTmp.Fields("CDP04")) Then
         'Modified by Lydia 2018/05/02 +format
         Text(6).Text = Format(rsTmp.Fields("CDP04"), "0000")
      End If
      '機關代號
      If Not IsNull(rsTmp.Fields("CDP05")) Then
         lblNum.Caption = rsTmp.Fields("CDP05")
         'edit by nickc 2007/02/07 不用 dll 了
         'If objLawDll.GetGovName(rsTmp.Fields("CDP05"), lbeGov.Caption) Then
         If ClsPDGetGovName(rsTmp.Fields("CDP05"), lbeGov.Caption) Then
         End If
      End If
      '開庭種類
      If Not IsNull(rsTmp.Fields("CDP06")) Then
         Text(4).Text = rsTmp.Fields("CDP06")
      End If
      '備註
      If Not IsNull(rsTmp.Fields("CDP07")) Then
         Text(7).Text = rsTmp.Fields("CDP07")
      End If
      '法官
      If Not IsNull(rsTmp.Fields("CDP08")) Then
         If DisHuman(1, rsTmp.Fields!cdp08) Then
         End If
      End If
      '檢察官
      If Not IsNull(rsTmp.Fields("CDP09")) Then
         If DisHuman(2, rsTmp.Fields!cdp09) Then
         End If
      End If
      'CreateID
      m_cdp10 = "" 'Add By Sindy 2011/6/15
      If Not IsNull(rsTmp.Fields("CDP10")) Then
         IDname = GetStaffName(rsTmp.Fields("CDP10"), True)
         m_cdp10 = rsTmp.Fields("CDP10") 'Add By Sindy 2011/6/15
      End If
      If Not IsNull(rsTmp.Fields("CDP11")) Then
         CDT = TAIWANDATE(rsTmp.Fields("CDP11"))
         CDT = Format(CDT, "###/##/##")
      End If
      
      If Not IsNull(rsTmp.Fields("CDP12")) Then
         CTM = Format(rsTmp.Fields("CDP12"), "##:##")
      End If
      If Not IsNull(rsTmp.Fields("CDP13")) Then
         UIDname = GetStaffName(rsTmp.Fields("CDP13"), True)
      End If
      If Not IsNull(rsTmp.Fields("CDP14")) Then
         UDT = TAIWANDATE(rsTmp.Fields("CDP14"))
         UDT = Format(UDT, "###/##/##")
      End If
      
      If Not IsNull(rsTmp.Fields("CDP15")) Then
         UTM = Format(rsTmp.Fields("CDP15"), "##:##")
      End If
      
      'Add By Sindy 2011/6/14
'      m_cdp16 = ""
'      If Not IsNull(rsTmp.Fields("CDP16")) Then
'         m_cdp16 = rsTmp.Fields("CDP16")
'         txtCR(9) = rsTmp.Fields("CDP16")
'         SetList lstAtt, rsTmp.Fields("CDP16")
'      End If
      '2011/6/14 End
      
      'Add By Sindy 2011/10/20
      '開庭別
      If Not IsNull(rsTmp.Fields("CDP17")) Then
         Text(0).Text = rsTmp.Fields("CDP17")
      End If
      '取消庭期日期
      If Not IsNull(rsTmp.Fields("CDP18")) Then
         Text(1).Text = ChangeWStringToTString(rsTmp.Fields("CDP18"))
      End If
      '2011/10/20 End
      
      rsTmp.Close
      Set rsTmp = Nothing
      
      ReadCaseprogress
   End If
   Me.Enabled = True
End Sub

'案件進度檔
Private Sub ReadCaseprogress()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim strTemp As String
Dim strNotInStaff As String
  
   strSql = "SELECT * FROM CASEPROGRESS WHERE CP09 ='" & txtPaperNum.Text & "'" & _
            " AND CP01 ='" & lc01 & "' AND CP02 ='" & lc02 & "' AND CP03 ='" & lc03 & "' AND CP04 ='" & lc04 & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If Not rsTmp.EOF Then
      '收受日
      If Not IsNull(rsTmp.Fields("CP05")) Then
         lblRevDate.Caption = ChangeWStringToTString(rsTmp.Fields("CP05"))
      End If
      '機關代號
      If Not IsNull(rsTmp.Fields("CP71")) Then
          lblNum.Caption = rsTmp.Fields("CP71")
          If lblNum.Caption <> "" Then
             'edit by nickc 2007/02/07 不用 dll 了
             'If objLawDll.GetGovName(rsTmp.Fields("CP71"), strTemp) Then
             If ClsPDGetGovName(rsTmp.Fields("CP71"), strTemp) Then
                lbeGov.Caption = strTemp
             End If
          End If
      End If
      '法院案號
      If Not IsNull(rsTmp.Fields("CP35")) Then
          lblLawNum.Caption = rsTmp.Fields("CP35")
      End If
      '股別
      If Not IsNull(rsTmp.Fields("CP30")) Then
         lblLNum.Caption = rsTmp.Fields("CP30")
      End If
      m_CP13 = "" & rsTmp.Fields("CP13") 'Add By Sindy 2011/6/14
      m_CP43 = "" & rsTmp.Fields("CP43") 'Add By Sindy 2020/6/10
      m_CP10 = "" & rsTmp.Fields("CP10") 'Added by Lydia 2025/03/19
   End If
   rsTmp.Close
   Set rsTmp = Nothing
  
   'Modify By Sindy 2011/6/24
   '最新智權人員
   If lc01 = "FCL" Or lc01 = "LIN" Then
      Label4 = GetPrjSalesNM(PUB_GetFCLSalesNo(lc01, lc02, lc03, lc04))
      m_StrTo = PUB_GetFCLSalesNo(lc01, lc02, lc03, lc04)
   Else
      Label4 = GetPrjSalesNM(PUB_GetAKindSalesNo(lc01, lc02, lc03, lc04))
      m_StrTo = PUB_GetAKindSalesNo(lc01, lc02, lc03, lc04)
   End If
   'Modify By Sindy 2020/6/10
   If m_CP43 <> "" Then
      '加發副本給該C類收文號之相關總收文號之承辦人、協辦人員及所有的出庭律師(以收文號讀取caselawer)。
      strExc(0) = "select cp14,cp29 from CaseProgress where CP09='" & m_CP43 & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         '承辦人
         If "" & RsTemp.Fields("CP14") <> "" Then
            lstMailCC.AddItem Trim(RsTemp.Fields("cp14")) & " " & GetPrjSalesNM(Trim(RsTemp.Fields("cp14")))
            lstMailCC.Selected(lstMailCC.ListCount - 1) = True
            strNotInStaff = strNotInStaff & ",'" & RsTemp.Fields("CP14") & "'"
         End If
         '協辦人員
         If "" & RsTemp.Fields("CP29") <> "" Then
            lstMailCC.AddItem Trim(RsTemp.Fields("cp29")) & " " & GetPrjSalesNM(Trim(RsTemp.Fields("cp29")))
            lstMailCC.Selected(lstMailCC.ListCount - 1) = True
            strNotInStaff = strNotInStaff & ",'" & RsTemp.Fields("CP29") & "'"
         End If
      End If
      '其他出庭律師
      strExc(0) = "select cl02 from caselawer where cl01='" & txtPaperNum & "' order by cl02 asc "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         RsTemp.MoveFirst
         Do While Not RsTemp.EOF
            If IsNull(RsTemp.Fields(0).Value) = False Then
               If InStr(strNotInStaff, "" & RsTemp.Fields("cl02")) = 0 Then  'Added by Lydia 2022/12/16 排除重複副本收受者
                  lstMailCC.AddItem Trim(RsTemp.Fields("cl02")) & " " & GetPrjSalesNM(Trim(RsTemp.Fields("cl02")))
                  lstMailCC.Selected(lstMailCC.ListCount - 1) = True
                  strNotInStaff = strNotInStaff & ",'" & Trim(RsTemp.Fields("cl02")) & "'"
               End If 'Added by Lydia 2022/12/16
            End If
            RsTemp.MoveNext
         Loop
      End If
      
      '原只發EMAIL給案件最新智權人員，請修改為若有案源則發給案源介紹人(可能多個)，無案源才發給案件最新智權人員。
      strExc(0) = "select * from LawOfficeSource where LOS06='" & m_CP43 & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         m_StrTo = Replace(RsTemp.Fields("LOS04"), ",", ";")
      End If
   End If
   '2020/6/10 END
   
   'Modify By Sindy 2020/6/11
   '副本：全部員工
   'lstMailCC.Clear
   strSql = "SELECT st01,st02 FROM staff WHERE st04='1' AND st01>'6' AND st01<'F'" & _
            " AND substr(st01,4,1)<>'9' AND substr(st01,1,2)>'63'"
   If strNotInStaff <> "" Then
      strNotInStaff = Mid(strNotInStaff, 2)
      strSql = strSql & " AND st01 not in(" & strNotInStaff & ")"
   End If
   strSql = strSql & " order by st01 asc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      With RsTemp
         RsTemp.MoveFirst
         Do While RsTemp.EOF = False
            lstMailCC.AddItem Trim(RsTemp.Fields("st01")) & " " & Trim(RsTemp.Fields("st02"))
            RsTemp.MoveNext
         Loop
      End With
   End If
   '2020/6/11 End
End Sub

Private Sub GetRsData()
Dim i As Integer
Dim strTmp As String
   
   If PaperNum = "" Then
      strTmp = "cdp01 = ''"
   Else
      strTmp = "where cdp01 in (" + PaperNum + ")"
   End If
   If lc01 = "LA" Then
      strExc(1) = "select cp05,cdp05,cp13,cp35,cp30,cdp01,cdp02,cdp03,cdp04,cdp06,cdp07,cdp08,cdp09" + _
         " from caseprogress,courtyardperiod,hirecase  where " + strTmp + " and cdp01=cp09 and " + _
         ChgCaseprogress(LcTmp) + " and " & ChgHirecase(LcTmp) & " order by cdp01"
   Else
      strExc(1) = "select cp05,cdp05,cp13,cp35,cp30,cdp01,cdp02,cdp03,cdp04,cdp06,cdp07,cdp08,cdp09" + _
         " from caseprogress,courtyardperiod ,lawcase where " + strTmp + " and cdp01=cp09 and " + _
         ChgCaseprogress(LcTmp) + " and " & ChgLawcase(LcTmp) & " order by cdp01"
   End If
   intI = 0
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(1))    'edit by nickc 2007/02/07 不用 dll 了 Set rstemp = objLawDll.ReadRstMsg(intI, strExc(1))
End Sub

Private Sub GetStaff()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String

'modify by sonia 2019/2/14 改用共用Function GetLawerList
'   'modify by sonia 2015/5/6 加外法律師
'   'strSql = "SELECT ST01,ST02 FROM STAFF WHERE ST03 ='L01' AND ST04 = '1' ORDER BY ST01"
'   strSql = "SELECT ST01,ST02 FROM STAFF WHERE (ST03 ='L01' OR ST20='13') AND ST04 = '1' ORDER BY ST01"
'   rsTmp.CursorLocation = adUseClient
'   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'
'   If rsTmp.EOF = False Then
'      Do While rsTmp.EOF = False
'         If Not IsNull(rsTmp.Fields("ST01")) Then
'            cboPerson.AddItem rsTmp.Fields("ST01") & "," & IIf(IsNull(rsTmp.Fields("ST02")), "", rsTmp.Fields("ST02"))
'         End If
'         rsTmp.MoveNext
'      Loop
'   End If
'   rsTmp.Close
'   Set rsTmp = Nothing
Dim i As Integer, varTmp1 As Variant, strTmp As String

   strSql = GetLawerList("1")
   varTmp1 = Split(strSql, ";")
   For i = 0 To UBound(varTmp1)
      strTmp = varTmp1(i)
      cboPerson.AddItem strTmp
   Next
'end 2019/2/14
End Sub

'Add By Cheng 2002/05/24
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

   TxtValidate = False
   For Each objTxt In Text
      If objTxt.Enabled = True Then
         Cancel = False
         Text_Validate objTxt.Index, Cancel
         If Cancel = True Then
            Text(objTxt.Index).SetFocus
            Exit Function
         End If
      End If
   Next
   
   'Added by Lydia 2021/09/17 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
        Exit Function
   End If
    
   TxtValidate = True
End Function

''開啟附件
'Private Sub cmdOpenAtt_Click()
'   If lstAtt.Text = "" Then
'      MsgBox "請選擇欲開啟的附件！"
'   Else
'      PUB_OpenFtpFile txtPaperNum, lstAtt.Text, Winsock1, "2"
'   End If
'End Sub
'
'Private Sub cmdAddAtt_Click()
'Dim stFileName As String
'Dim sFile
'Dim ii As Integer
'Dim fs, f, s
'
'On Error GoTo ErrHnd
'
'   stFileName = "*.*"
'   With CommonDialog1
'      .CancelError = True
'      .FileName = stFileName
'      .Filter = "All Files (*.*)|*.*"
'      .InitDir = PUB_Getdesktop
'      .MaxFileSize = 3000
'      .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
'      .ShowOpen
'      If .FileName <> "" Then
'         If InStr(.FileName, ChrW$(0)) > 0 Then
'            sFile = Split(.FileName, ChrW$(0))
'            For ii = 1 To UBound(sFile)
'               If InStr(sFile(ii), "\") > 0 Then
'                  stFileName = sFile(ii)
'               Else
'                  stFileName = sFile(0) & "\" & sFile(ii)
'               End If
'               Set fs = CreateObject("Scripting.FileSystemObject")
'               Set f = fs.GetFile(stFileName)
'               AddListX lstAtt, stFileName & " (" & -1 * Int(-1 * f.Size / 1024) & " KB)"
'            Next
'         Else
'            stFileName = .FileName
'            Set fs = CreateObject("Scripting.FileSystemObject")
'            Set f = fs.GetFile(stFileName)
'            AddListX lstAtt, stFileName & " (" & -1 * Int(-1 * f.Size / 1024) & " KB)"
'         End If
'         '改上傳到FTP,故只需留檔名
'         txtCR(9) = ComposeAttList(lstAtt)
'      End If
'   End With
'   Exit Sub
'ErrHnd:
'   If Err.Number <> 32755 Then
'      MsgBox Err.Description
'   End If
'End Sub
'
'Private Sub cmdRemAtt_Click()
'   If InStr(lstAtt, "\") = 0 And Pub_StrUserSt03 <> "M51" Then
'      MsgBox "已上傳檔案不可移除！"
'   ElseIf RemoveList(lstAtt) = True Then
'      txtCR(9) = ComposeList(lstAtt)
'      cmdAddAtt.SetFocus
'   End If
'End Sub
'
'Private Sub lstAtt_DblClick()
'   If cmdOpenAtt.Enabled = True Then
'      cmdOpenAtt.Value = True
'   End If
'End Sub
'
'Private Sub SetList(oList As ListBox, p_stList As String)
'Dim arrID
'
'   oList.Clear
'   If p_stList <> "" Then
'      arrID = Split(p_stList, ",")
'      For intI = UBound(arrID) To LBound(arrID) Step -1
'         oList.AddItem arrID(intI), 0
'      Next
'   End If
'End Sub
'
''上傳附件檔
'Private Function UploadAtt(ByVal stKEY As String, Optional iErrNo As Integer, Optional stErrMsg As String) As Boolean
'Dim hOpen As Long
'Dim hConnection As Long
'Dim hDir As Long
'Dim bReturn As Boolean
'Dim dwInternetFlags As Integer
'Dim stDir As String
'Dim stRemoteFile As String
'Dim stLocalFile As String
'Dim stItem As String
'Dim idx As Integer
'Dim iPos As Integer
'Dim IsTimeOut As Boolean
'Dim SeekTimer
'Dim ACT_FTP_IP As String
'Dim arrIP
'Dim ii As Integer
'
'   iErrNo = 0
'   stErrMsg = ""
'
'   stDir = 法務案開庭紀要存放路徑
'   If lstAtt.ListCount > 0 Then
'      For idx = 0 To lstAtt.ListCount - 1
'         stItem = lstAtt.List(idx)
'         iPos = InStr(stItem, "\")
'         If iPos > 0 Then
'            If InStrRev(stItem, " (") > 0 Then
'               stLocalFile = Left(stItem, InStrRev(stItem, " (") - 1)
'            Else
'               stLocalFile = stItem
'            End If
'            stRemoteFile = GetFileName(stLocalFile)
'
'            If hOpen = 0 Then
'               hOpen = InternetOpen("Taie FTP", INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
'               If hOpen = 0 Then
'                  iErrNo = 1
'                  stErrMsg = "網路錯誤！"
'                  GoTo OutPort
'               Else
'                  IsTimeOut = True
'                  If GOOD_FTP_IP <> "" Then
'                     arrIP = Split(GOOD_FTP_IP & ";" & FTP_IP, ";")
'                  Else
'                     arrIP = Split(FTP_IP, ";")
'                  End If
'                  For ii = LBound(arrIP) To UBound(arrIP)
'                     ACT_FTP_IP = arrIP(ii)
'                     If ACT_FTP_IP <> "" Then
'                        '偵測 FTPServer 是否存在
'                        If Winsock1.State Then Winsock1.Close
'                        Winsock1.Connect ACT_FTP_IP, 21
'                        IsTimeOut = False
'                        SeekTimer = Timer
'                        Do While Winsock1.State = 6 And IsTimeOut = False
'                           DoEvents
'                           If Timer - SeekTimer > 1 Then
'                              IsTimeOut = True
'                           End If
'                        Loop
'                        If Winsock1.State Then Winsock1.Close
'                        If IsTimeOut = False Then
'                           Exit For
'                        End If
'                     End If
'                  Next
'
'                  '若是超過時間
'                  If IsTimeOut = True Then
'                     iErrNo = 2
'                     stErrMsg = "無法與FTP Server建立連線！"
'                     GoTo OutPort
'                  Else
'                     GOOD_FTP_IP = ACT_FTP_IP
'                  End If
'
'                  hConnection = InternetConnect(hOpen, ACT_FTP_IP, FTP_Port, _
'                     "pgmid", "pgmpwd", INTERNET_SERVICE_FTP, INTERNET_FLAG_PASSIVE, 0)
'                  If hConnection = 0 Then
'                     iErrNo = 3
'                     stErrMsg = "無法與FTP Server建立連線！"
'                     GoTo OutPort
'                  ElseIf FtpSetCurrentDirectory(hConnection, stDir) = False Then
'                     iErrNo = 4
'                     stErrMsg = "切換至開庭紀要目錄失敗！"
'                     GoTo OutPort
'                  '切換至開庭紀要單號目錄
'                  ElseIf FtpSetCurrentDirectory(hConnection, stKEY) = False Then
'                     hDir = FtpCreateDirectory(hConnection, stKEY)
'                     If hDir = 0 Then
'                        iErrNo = 5
'                        stErrMsg = "建立開庭紀要單號目錄失敗！"
'                        GoTo OutPort
'                     ElseIf FtpSetCurrentDirectory(hConnection, stKEY) = False Then
'                        iErrNo = 6
'                        stErrMsg = "切換至開庭紀要單號目錄失敗！"
'                        GoTo OutPort
'                     End If
'                  End If
'               End If
'            End If
'
'            dwInternetFlags = FTP_TRANSFER_TYPE_BINARY
'            bReturn = FtpPutFile(hConnection, stLocalFile, stRemoteFile, dwInternetFlags, 0)
'            ' Upload successfully
'            If bReturn = False Then
'               iErrNo = 7
'               stErrMsg = "檔案上傳失敗！"
'               GoTo OutPort
'            End If
'         End If
'      Next
'   End If
'   UploadAtt = True
'
'OutPort:
'   If hOpen <> 0 Then InternetCloseHandle (hOpen)
'   If hConnection <> 0 Then InternetCloseHandle (hConnection)
'   If Winsock1.State Then Winsock1.Close
'End Function
'
''刪除附件檔
'Private Function RemoveAtt(ByVal stKEY As String, stFiles As String, Optional iErrNo As Integer, Optional stErrMsg As String) As Boolean
'Dim hOpen As Long
'Dim hConnection As Long
'Dim bReturn As Boolean
'Dim stDir As String
'Dim IsTimeOut As Boolean
'Dim SeekTimer
'Dim ACT_FTP_IP As String
'Dim arrIP
'Dim ii As Integer, jj As Integer
'Dim arrFile
'Dim stRemoteFile As String
'
'   iErrNo = 0
'   stErrMsg = ""
'
'   stDir = 法務案開庭紀要存放路徑
'   arrFile = Split(stFiles, ",")
'   For jj = LBound(arrFile) To UBound(arrFile)
'      If arrFile(jj) <> "" Then
'         If hOpen = 0 Then
'            hOpen = InternetOpen("Taie FTP", INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
'            If hOpen = 0 Then
'               iErrNo = 1
'               stErrMsg = "網路錯誤！"
'               GoTo OutPort
'            Else
'               IsTimeOut = True
'               If GOOD_FTP_IP <> "" Then
'                  arrIP = Split(GOOD_FTP_IP & ";" & FTP_IP, ";")
'               Else
'                  arrIP = Split(FTP_IP, ";")
'               End If
'               For ii = LBound(arrIP) To UBound(arrIP)
'                  ACT_FTP_IP = arrIP(ii)
'                  If ACT_FTP_IP <> "" Then
'                     '偵測 FTPServer 是否存在
'                     If Winsock1.State Then Winsock1.Close
'                     Winsock1.Connect ACT_FTP_IP, 21
'                     IsTimeOut = False
'                     SeekTimer = Timer
'                     Do While Winsock1.State = 6 And IsTimeOut = False
'                        DoEvents
'                        If Timer - SeekTimer > 1 Then
'                           IsTimeOut = True
'                        End If
'                     Loop
'                     If Winsock1.State Then Winsock1.Close
'                     If IsTimeOut = False Then
'                        Exit For
'                     End If
'                  End If
'               Next
'
'               '若是超過時間
'               If IsTimeOut = True Then
'                  iErrNo = 2
'                  stErrMsg = "無法與FTP Server建立連線！"
'                  GoTo OutPort
'               Else
'                  GOOD_FTP_IP = ACT_FTP_IP
'               End If
'
'               hConnection = InternetConnect(hOpen, ACT_FTP_IP, FTP_Port, _
'                  "pgmid", "pgmpwd", INTERNET_SERVICE_FTP, INTERNET_FLAG_PASSIVE, 0)
'               If hConnection = 0 Then
'                  iErrNo = 3
'                  stErrMsg = "無法與FTP Server建立連線！"
'                  GoTo OutPort
'               ElseIf FtpSetCurrentDirectory(hConnection, stDir) = False Then
'                  iErrNo = 4
'                  stErrMsg = "切換至開庭紀要目錄失敗！"
'                  GoTo OutPort
'               '切換至開庭紀要單號目錄
'               ElseIf FtpSetCurrentDirectory(hConnection, stKEY) = False Then
'                  '無法切換時當作已刪除
'                  'iErrNo = 6
'                  'stErrMsg = "切換至開庭紀要單號目錄失敗！"
'                  'GoTo OutPort
'                  Exit For
'               End If
'            End If
'         End If
'         If InStrRev(arrFile(jj), " (") > 0 Then
'            stRemoteFile = Left(arrFile(jj), InStrRev(arrFile(jj), " (") - 1)
'         Else
'            stRemoteFile = arrFile(jj)
'         End If
'         '刪除檔案不控制成功與否
'         bReturn = FtpDeleteFile(hConnection, stRemoteFile)
'      End If
'   Next
'
'   RemoveAtt = True
'
'OutPort:
'   If hOpen <> 0 Then InternetCloseHandle (hOpen)
'   If hConnection <> 0 Then InternetCloseHandle (hConnection)
'   If Winsock1.State Then Winsock1.Close
'End Function
'
'Private Function AddListX(oList As ListBox, stNewItem As String) As Boolean
'Dim idx As Integer, bFound As Boolean, stFileName As String
'
'   If InStr(stNewItem, ",") > 0 Then
'      MsgBox "逗號[,]為系統保留字，請重新命名！", vbExclamation
'      cmdAddAtt.SetFocus
'      Exit Function
'   End If
'   If stNewItem <> "" Then
'      For idx = 0 To oList.ListCount - 1
'         stFileName = GetFileName(oList.List(idx))
'         If GetFileName(stNewItem) = stFileName Then
'            MsgBox "附件[" & stFileName & "]已存在！"
'            AddListX = False
'            bFound = True
'            Exit For
'         End If
'      Next
'      If bFound = False Then
'         oList.AddItem stNewItem, 0
'         AddListX = True
'      End If
'   End If
'End Function
'
''附件
'Private Function ComposeAttList(oList As ListBox) As String
'Dim iPos As Integer, stItem As String, stRtn As String, idx As Integer
'
'   If oList.ListCount > 0 Then
'      stItem = oList.List(0)
'      stRtn = GetFileName(stItem)
'      For idx = 1 To oList.ListCount - 1
'         stItem = oList.List(idx)
'         stRtn = stRtn & "," & GetFileName(stItem)
'      Next
'   End If
'   ComposeAttList = stRtn
'End Function
'
'Private Function RemoveList(oList As ListBox) As Boolean
'Dim ii As Integer
'
'   If oList.ListCount > 0 Then
'      ii = 0
'      Do While ii < oList.ListCount
'         If oList.Selected(ii) = True Then
'            RemoveList = True
'            oList.RemoveItem ii
'            ii = ii - 1
'         End If
'         ii = ii + 1
'      Loop
'   End If
'End Function
'
'Private Function ComposeList(oList As ListBox, Optional p_iOpt As Integer = 0) As String
'Dim iPos As Integer, stItem As String
'
'   strExc(1) = ""
'   If oList.ListCount > 0 Then
'      For intI = 0 To oList.ListCount - 1
'         If p_iOpt = 0 Then
'            iPos = InStr(oList.List(intI), Chr(1))
'            If iPos > 0 Then
'               stItem = Left(oList.List(intI), iPos - 1)
'            Else
'               stItem = oList.List(intI)
'            End If
'         Else
'            stItem = Format(oList.ItemData(intI), "00")
'         End If
'         stItem = GetFileName(stItem) 'Add By Sindy 2012/3/21
'         If intI = 0 Then
'            strExc(1) = stItem
'         Else
'            strExc(1) = strExc(1) & "," & stItem
'         End If
'      Next
'   End If
'   ComposeList = strExc(1)
'End Function
'
'Private Function GetFileName(ByVal FullPath As String) As String
'Dim stItem As String, iPos As Integer
'
'   stItem = FullPath
'   iPos = InStr(stItem, "\")
'   Do While iPos > 0
'      stItem = Mid(stItem, iPos + 1)
'      iPos = InStr(stItem, "\")
'   Loop
'   GetFileName = stItem
'End Function

'Added by Lydia 2015/10/30 上傳開庭通知
Private Sub cmdFile_Click()
   Call frm090801_8.SetParent(Me)
   'frm090801_8.m_strSaveFiles = Me.m_strSaveFiles
   frm090801_8.m_strSaveFiles = Me.m_strSaveFilesOA 'Modify By Sindy 2016/6/24
   frm090801_8.lblCaseNo = m_CP01 & "-" & m_CP02 & "-" & m_CP03 & "-" & m_CP04
   frm090801_8.Caption = "新增開庭通知附件" 'Add By Sindy 2016/6/24
   Me.m_strSaveFileType = "1" 'Add By Sindy 2016/6/24 開庭通知
   frm090801_8.Show vbModal
End Sub
' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal strCP09 As String) As Boolean
Dim adoRst As ADODB.Recordset
   
   IsRecordExist = False
   
   'Modify By Sindy 2016/6/27 +.pdf
   'strSql = "SELECT cpp01 FROM casepaperpdf WHERE cpp01='" & strCP09 & "' and instr(cpp02,'" & 通知開庭 & "') > 0"
   strSql = "SELECT cpp01 FROM casepaperpdf WHERE cpp01='" & strCP09 & "' and instr(upper(cpp02),upper('" & 通知開庭 & ".pdf')) > 0"
   intI = 1
   Set adoRst = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      IsRecordExist = True
      MsgBox "開庭通知附件已存在，無法上傳！"
   End If
   
   Set adoRst = Nothing
End Function
'end 2015/10/30

'Add By Sindy 2016/6/27 上傳開庭紀要
Private Sub cmdBrief_Click()
   Call frm090801_8.SetParent(Me)
   frm090801_8.m_strSaveFiles = Me.m_strSaveFilesBRIEF
   frm090801_8.lblCaseNo = m_CP01 & "-" & m_CP02 & "-" & m_CP03 & "-" & m_CP04
   frm090801_8.Caption = "新增開庭紀要附件"
   Me.m_strSaveFileType = "2" 'Add By Sindy 2016/6/24 開庭紀要
   frm090801_8.bolNotPDF = True '開啟附件的檔案類型*.*
   frm090801_8.Show vbModal
End Sub

'Add By Sindy 2016/6/27 上傳電子筆錄
Private Sub cmdNote_Click()
   Call frm090801_8.SetParent(Me)
   frm090801_8.m_strSaveFiles = Me.m_strSaveFilesNOTE
   frm090801_8.lblCaseNo = m_CP01 & "-" & m_CP02 & "-" & m_CP03 & "-" & m_CP04
   frm090801_8.Caption = "新增電子筆錄附件"
   Me.m_strSaveFileType = "3" 'Add By Sindy 2016/6/24 電子筆錄
   frm090801_8.Show vbModal
End Sub

'Added by Lydia 2024/07/29
Private Sub Command4_Click()
   If m_EDIT = 1 Or txtPaperNum.Text = "" Then
      MsgBox "請先完成新增作業"
   End If
   'Modified by Lydia 2025/03/19 + ,m_CP10
   Call frm071018.SetParent(Me, Me.txtPaperNum, IIf(Me.Tag = "", True, False), Replace(Trim(Left(cboPerson, 6)), ",", ""), m_CP10)
   frm071018.Show vbModal
End Sub
