VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm140114 
   BorderStyle     =   3  '雙線固定對話方塊
   Caption         =   "客戶端平台帳號管理作業"
   ClientHeight    =   5730
   ClientLeft      =   180
   ClientTop       =   1000
   ClientWidth     =   8950
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frm140114.frx":0000
   ScaleHeight     =   5730
   ScaleWidth      =   8950
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7440
      Top             =   30
      _ExtentX        =   988
      _ExtentY        =   988
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140114.frx":060C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140114.frx":0928
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140114.frx":0C44
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140114.frx":0E20
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140114.frx":113C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140114.frx":1458
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140114.frx":1774
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140114.frx":1A90
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140114.frx":1DAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140114.frx":20C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140114.frx":23E4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   520
      Left            =   0
      TabIndex        =   47
      Top             =   0
      Width           =   8950
      _ExtentX        =   15787
      _ExtentY        =   917
      ButtonWidth     =   1076
      ButtonHeight    =   882
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
            Enabled         =   0   'False
            Caption         =   "確定"
            Key             =   "keyOk"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
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
      Height          =   4815
      Left            =   30
      TabIndex        =   48
      Top             =   660
      Width           =   8895
      _ExtentX        =   15681
      _ExtentY        =   8484
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "平台資訊"
      TabPicture(0)   =   "frm140114.frx":2700
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label6"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(7)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label3(4)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label4"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label7"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label8"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label11"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label12"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label13"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label14"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lstUsers(1)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lstAtt"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "textCW12"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "textCW05"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "CommonDialog1"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Frame3"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "cmdOpenAtt(0)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "cmdSaveAtt(0)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "cmdAddAtt(0)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "cmdRemAtt(0)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "cmdSelect(0)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "cmdOpenIE"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "textCW(1)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "textCW(2)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "cboCW03"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "textCW(4)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "textCW(13)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "textCW(16)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Frame4"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Frame5"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "textCW(14)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "textCW(15)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "textCW(17)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "textCW(18)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "cboCW19"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).ControlCount=   39
      TabCaption(1)   =   "帳號資料"
      TabPicture(1)   =   "frm140114.frx":271C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text4"
      Tab(1).Control(1)=   "cmdOK(4)"
      Tab(1).Control(2)=   "cmdOK(3)"
      Tab(1).Control(3)=   "cmdOK(0)"
      Tab(1).Control(4)=   "cmdOK(1)"
      Tab(1).Control(5)=   "cmdOK(2)"
      Tab(1).Control(6)=   "Text2"
      Tab(1).Control(7)=   "Frame1"
      Tab(1).Control(8)=   "Check1"
      Tab(1).Control(9)=   "grd1"
      Tab(1).Control(10)=   "lstUsers(2)"
      Tab(1).Control(11)=   "Label1(2)"
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "多筆查詢"
      TabPicture(2)   =   "frm140114.frx":2738
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Text3"
      Tab(2).Control(1)=   "txt1(4)"
      Tab(2).Control(2)=   "txt1(3)"
      Tab(2).Control(3)=   "txt1(0)"
      Tab(2).Control(4)=   "cmdQuery(0)"
      Tab(2).Control(5)=   "txt1(1)"
      Tab(2).Control(6)=   "MSHFlexGrid1"
      Tab(2).Control(7)=   "txtName"
      Tab(2).Control(8)=   "Label1(13)"
      Tab(2).Control(9)=   "Label1(16)"
      Tab(2).Control(10)=   "Line1"
      Tab(2).Control(11)=   "Label3(2)"
      Tab(2).Control(12)=   "lbl1(1)"
      Tab(2).Control(13)=   "Label1(11)"
      Tab(2).Control(14)=   "Label1(15)"
      Tab(2).Control(15)=   "lbl1(0)"
      Tab(2).ControlCount=   16
      Begin VB.TextBox Text4 
         Appearance      =   0  '平面
         BackColor       =   &H8000000F&
         BorderStyle     =   0  '沒有框線
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   -70110
         TabIndex        =   96
         Text            =   "　　下次更新日期=111111，代表不提醒"
         Top             =   3270
         Width           =   3705
      End
      Begin VB.ComboBox cboCW19 
         Height          =   260
         ItemData        =   "frm140114.frx":2754
         Left            =   6480
         List            =   "frm140114.frx":2756
         TabIndex        =   5
         Text            =   "cboCW19"
         Top             =   630
         Width           =   1305
      End
      Begin VB.TextBox textCW 
         Height          =   270
         Index           =   18
         Left            =   1050
         MaxLength       =   500
         TabIndex        =   8
         Top             =   1470
         Width           =   6990
      End
      Begin VB.TextBox textCW 
         Height          =   270
         Index           =   17
         Left            =   1050
         MaxLength       =   500
         TabIndex        =   7
         Top             =   1200
         Width           =   6990
      End
      Begin VB.TextBox textCW 
         Height          =   270
         Index           =   15
         Left            =   690
         TabIndex        =   90
         Top             =   2310
         Width           =   345
      End
      Begin VB.TextBox textCW 
         Height          =   270
         Index           =   14
         Left            =   360
         TabIndex        =   89
         Top             =   2310
         Width           =   345
      End
      Begin VB.Frame Frame5 
         BorderStyle     =   0  '沒有框線
         Height          =   225
         Left            =   30
         TabIndex        =   87
         Top             =   2700
         Width           =   3735
         Begin VB.CheckBox Check2 
            Caption         =   "網址"
            Height          =   315
            Index           =   2
            Left            =   2790
            TabIndex        =   15
            Top             =   -30
            Width           =   705
         End
         Begin VB.CheckBox Check2 
            Caption         =   "憑證"
            Height          =   315
            Index           =   1
            Left            =   2070
            TabIndex        =   14
            Top             =   -30
            Width           =   705
         End
         Begin VB.CheckBox Check2 
            Caption         =   "帳號密碼"
            Height          =   315
            Index           =   0
            Left            =   1020
            TabIndex        =   13
            Top             =   -30
            Width           =   1035
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "驗證方式："
            Height          =   180
            Left            =   60
            TabIndex        =   88
            Top             =   30
            Width           =   900
         End
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  '沒有框線
         Height          =   225
         Left            =   4620
         TabIndex        =   85
         Top             =   2700
         Width           =   3105
         Begin VB.OptionButton Option1 
            Caption         =   "客戶核准"
            Height          =   255
            Index           =   1
            Left            =   2010
            TabIndex        =   17
            Top             =   0
            Width           =   1065
         End
         Begin VB.OptionButton Option1 
            Caption         =   "本所管理"
            Height          =   255
            Index           =   0
            Left            =   960
            TabIndex        =   16
            Top             =   0
            Width           =   1065
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "管理方式："
            Height          =   180
            Left            =   30
            TabIndex        =   86
            Top             =   30
            Width           =   900
         End
      End
      Begin VB.TextBox textCW 
         Height          =   270
         Index           =   16
         Left            =   5065
         MaxLength       =   2
         TabIndex        =   4
         Top             =   660
         Width           =   400
      End
      Begin VB.TextBox textCW 
         Height          =   270
         Index           =   13
         Left            =   3330
         MaxLength       =   7
         TabIndex        =   3
         Top             =   660
         Width           =   800
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "修改"
         Enabled         =   0   'False
         Height          =   345
         Index           =   4
         Left            =   -73500
         TabIndex        =   36
         Top             =   3060
         Width           =   795
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "刪除"
         Enabled         =   0   'False
         Height          =   345
         Index           =   3
         Left            =   -72660
         TabIndex        =   37
         Top             =   3060
         Width           =   795
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "新增"
         Enabled         =   0   'False
         Height          =   345
         Index           =   0
         Left            =   -74340
         TabIndex        =   35
         Top             =   3060
         Width           =   795
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "確定"
         Enabled         =   0   'False
         Height          =   345
         Index           =   1
         Left            =   -71820
         TabIndex        =   38
         Top             =   3060
         Width           =   795
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "取消"
         Enabled         =   0   'False
         Height          =   345
         Index           =   2
         Left            =   -70980
         TabIndex        =   39
         Top             =   3060
         Width           =   795
      End
      Begin VB.TextBox textCW 
         Height          =   270
         Index           =   4
         Left            =   30
         TabIndex        =   77
         Top             =   2310
         Width           =   345
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  '平面
         BackColor       =   &H8000000F&
         BorderStyle     =   0  '沒有框線
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   -70560
         TabIndex        =   74
         Text            =   $"frm140114.frx":2758
         Top             =   1410
         Width           =   4275
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  '平面
         BackColor       =   &H8000000F&
         BorderStyle     =   0  '沒有框線
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   -74850
         TabIndex        =   73
         Text            =   $"frm140114.frx":2788
         Top             =   630
         Width           =   2235
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   4
         Left            =   -72300
         MaxLength       =   7
         TabIndex        =   44
         Top             =   1320
         Width           =   915
      End
      Begin VB.Frame Frame1 
         Height          =   1665
         Left            =   -74790
         TabIndex        =   63
         Top             =   3090
         Width           =   8505
         Begin VB.TextBox Text1 
            Appearance      =   0  '平面
            BackColor       =   &H8000000F&
            BorderStyle     =   0  '沒有框線
            ForeColor       =   &H000000C0&
            Height          =   225
            Left            =   4680
            TabIndex        =   80
            Text            =   $"frm140114.frx":27A5
            Top             =   0
            Width           =   3705
         End
         Begin VB.Frame Frame2 
            Height          =   735
            Left            =   5220
            TabIndex        =   64
            Top             =   360
            Width           =   2175
            Begin VB.TextBox txtUserNo 
               Height          =   264
               Index           =   0
               Left            =   810
               MaxLength       =   6
               TabIndex        =   30
               Top             =   120
               Width           =   945
            End
            Begin VB.CommandButton cmdAdd 
               Caption         =   "<- 加入"
               Height          =   285
               Index           =   0
               Left            =   45
               TabIndex        =   31
               Top             =   120
               Width           =   735
            End
            Begin VB.CommandButton cmdRemove 
               Caption         =   "移除 ->"
               Height          =   285
               Index           =   0
               Left            =   45
               TabIndex        =   32
               TabStop         =   0   'False
               Top             =   420
               Width           =   735
            End
            Begin VB.Label lblName 
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               Height          =   180
               Index           =   0
               Left            =   840
               TabIndex        =   65
               Top             =   450
               Width           =   1185
            End
         End
         Begin VB.ComboBox cboCD05 
            Height          =   300
            ItemData        =   "frm140114.frx":27D4
            Left            =   1350
            List            =   "frm140114.frx":27D6
            Style           =   2  '單純下拉式
            TabIndex        =   28
            Top             =   1020
            Width           =   1845
         End
         Begin MSForms.TextBox textCD 
            Height          =   285
            Index           =   9
            Left            =   6660
            TabIndex        =   33
            Top             =   1080
            Width           =   1185
            VariousPropertyBits=   671105051
            MaxLength       =   7
            Size            =   "7223;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox textCD 
            Height          =   285
            Index           =   6
            Left            =   3330
            TabIndex        =   78
            Top             =   780
            Width           =   615
            VariousPropertyBits=   671105051
            Size            =   "7223;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox textCD 
            Height          =   285
            Index           =   8
            Left            =   1350
            TabIndex        =   29
            Top             =   1350
            Width           =   2475
            VariousPropertyBits=   671105051
            MaxLength       =   50
            Size            =   "7223;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox textCD 
            Height          =   285
            Index           =   4
            Left            =   1350
            TabIndex        =   27
            Top             =   720
            Width           =   1815
            VariousPropertyBits=   671105051
            MaxLength       =   30
            Size            =   "7223;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox textCD 
            Height          =   285
            Index           =   3
            Left            =   1350
            TabIndex        =   26
            Top             =   420
            Width           =   1815
            VariousPropertyBits=   671105051
            MaxLength       =   60
            Size            =   "7223;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox textCD 
            Height          =   285
            Index           =   7
            Left            =   6660
            TabIndex        =   34
            Top             =   1350
            Width           =   1185
            VariousPropertyBits=   671105051
            MaxLength       =   7
            Size            =   "7223;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.ListBox lstUsers 
            Height          =   1245
            Index           =   0
            Left            =   4020
            TabIndex        =   99
            Top             =   360
            Width           =   1155
            VariousPropertyBits=   746586139
            ScrollBars      =   3
            DisplayStyle    =   2
            Size            =   "2037;2196"
            MatchEntry      =   0
            FontName        =   "新細明體"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "建置日期："
            Height          =   180
            Index           =   6
            Left            =   5730
            TabIndex        =   81
            Top             =   1140
            Width           =   900
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "註　解："
            Height          =   180
            Index           =   5
            Left            =   630
            TabIndex        =   75
            Top             =   1380
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "密　碼："
            Height          =   180
            Index           =   4
            Left            =   630
            TabIndex        =   70
            Top             =   750
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "帳　號："
            Height          =   180
            Index           =   5
            Left            =   630
            TabIndex        =   69
            Top             =   450
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "使用者："
            Height          =   180
            Index           =   6
            Left            =   3330
            TabIndex        =   68
            Top             =   450
            Width           =   720
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "下次更新日期："
            Height          =   180
            Index           =   1
            Left            =   5370
            TabIndex        =   67
            Top             =   1410
            Width           =   1260
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "身份別："
            Height          =   180
            Left            =   630
            TabIndex        =   66
            Top             =   1050
            Width           =   720
         End
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   3
         Left            =   -73350
         MaxLength       =   7
         TabIndex        =   43
         Top             =   1320
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   0
         Left            =   -73800
         MaxLength       =   9
         TabIndex        =   40
         Top             =   390
         Width           =   1095
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "查詢(&Q)"
         Height          =   345
         Index           =   0
         Left            =   -70530
         TabIndex        =   45
         Top             =   510
         Width           =   885
      End
      Begin VB.TextBox txt1 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.5
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   -73800
         MaxLength       =   6
         TabIndex        =   41
         Top             =   690
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "依客戶分別設定帳號資料"
         Height          =   225
         Left            =   -74910
         TabIndex        =   25
         Top             =   1110
         Width           =   2295
      End
      Begin VB.ComboBox cboCW03 
         Height          =   260
         ItemData        =   "frm140114.frx":27D8
         Left            =   1050
         List            =   "frm140114.frx":27DA
         Style           =   2  '單純下拉式
         TabIndex        =   2
         Top             =   630
         Width           =   1300
      End
      Begin VB.TextBox textCW 
         Height          =   270
         Index           =   2
         Left            =   1050
         MaxLength       =   500
         TabIndex        =   6
         Top             =   930
         Width           =   6990
      End
      Begin VB.TextBox textCW 
         Height          =   270
         Index           =   1
         Left            =   1050
         MaxLength       =   4
         TabIndex        =   0
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton cmdOpenIE 
         Caption         =   "進入網站"
         Height          =   255
         Left            =   7830
         TabIndex        =   9
         Top             =   660
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "全選"
         Height          =   255
         Index           =   0
         Left            =   8070
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   3720
         Width           =   735
      End
      Begin VB.CommandButton cmdRemAtt 
         Caption         =   "-> 移除"
         Height          =   255
         Index           =   0
         Left            =   8070
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   4530
         Width           =   735
      End
      Begin VB.CommandButton cmdAddAtt 
         Caption         =   "<- 新增"
         Height          =   255
         Index           =   0
         Left            =   8070
         TabIndex        =   23
         Top             =   4260
         Width           =   735
      End
      Begin VB.CommandButton cmdSaveAtt 
         Caption         =   "下載"
         Height          =   255
         Index           =   0
         Left            =   8070
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   3990
         Width           =   735
      End
      Begin VB.CommandButton cmdOpenAtt 
         Caption         =   "開啟"
         Height          =   255
         Index           =   0
         Left            =   8070
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   3450
         Width           =   735
      End
      Begin VB.Frame Frame3 
         Height          =   765
         Left            =   4620
         TabIndex        =   50
         Top             =   1830
         Width           =   4185
         Begin VB.TextBox txtUserNo 
            Height          =   264
            Index           =   1
            Left            =   840
            MaxLength       =   9
            TabIndex        =   10
            Top             =   150
            Width           =   1005
         End
         Begin VB.CommandButton cmdRemove 
            Caption         =   "移除 ->"
            Height          =   285
            Index           =   1
            Left            =   45
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   420
            Width           =   735
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "<- 加入"
            Height          =   285
            Index           =   1
            Left            =   45
            TabIndex        =   11
            Top             =   120
            Width           =   735
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "（輸入客戶編號）"
            Height          =   180
            Index           =   3
            Left            =   1920
            TabIndex        =   91
            Top             =   210
            Width           =   1440
         End
         Begin VB.Label lblName 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Height          =   180
            Index           =   1
            Left            =   840
            TabIndex        =   51
            Top             =   480
            Width           =   3165
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
         Height          =   1695
         Left            =   -74790
         TabIndex        =   49
         Top             =   1350
         Width           =   8490
         _ExtentX        =   14958
         _ExtentY        =   2981
         _Version        =   393216
         BackColor       =   16777215
         Cols            =   9
         FixedCols       =   0
         HighLight       =   0
         AllowUserResizing=   1
         FormatString    =   "客戶編號 |客戶名稱 |帳號 |密碼 |身份別 |使用者 |建置日期 |下次更新日期 |註解"
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
         _Band(0).Cols   =   9
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   3075
         Left            =   -74880
         TabIndex        =   46
         Top             =   1650
         Width           =   8700
         _ExtentX        =   15346
         _ExtentY        =   5415
         _Version        =   393216
         BackColor       =   16777215
         Cols            =   6
         FixedCols       =   0
         HighLight       =   0
         AllowUserResizing=   1
         FormatString    =   "平台編號  |客戶編號  |客戶名稱  |平台名稱  |平台類別  |網址  "
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
         _Band(0).Cols   =   6
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   510
         Top             =   4110
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSForms.TextBox textCW05 
         Height          =   510
         Left            =   1050
         TabIndex        =   18
         Top             =   2940
         Width           =   7755
         VariousPropertyBits=   -1466939365
         MaxLength       =   200
         ScrollBars      =   3
         Size            =   "13679;900"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCW12 
         Height          =   300
         Left            =   4020
         TabIndex        =   1
         Top             =   330
         Width           =   4785
         VariousPropertyBits=   679495707
         MaxLength       =   50
         Size            =   "8440;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtName 
         Height          =   300
         Left            =   -73800
         TabIndex        =   42
         Top             =   990
         Width           =   2055
         VariousPropertyBits=   679495707
         Size            =   "3625;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ListBox lstAtt 
         Height          =   1275
         Left            =   1050
         TabIndex        =   19
         Top             =   3450
         Width           =   7005
         VariousPropertyBits=   746586139
         ScrollBars      =   3
         DisplayStyle    =   2
         Size            =   "12356;2249"
         MatchEntry      =   0
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ListBox lstUsers 
         Height          =   975
         Index           =   2
         Left            =   -72540
         TabIndex        =   98
         Top             =   360
         Width           =   6195
         VariousPropertyBits=   746586139
         ScrollBars      =   3
         DisplayStyle    =   2
         Size            =   "10927;1720"
         MatchEntry      =   0
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ListBox lstUsers 
         Height          =   920
         Index           =   1
         Left            =   1050
         TabIndex        =   97
         Top             =   1740
         Width           =   3530
         VariousPropertyBits=   746586139
         ScrollBars      =   3
         DisplayStyle    =   2
         Size            =   "6227;1623"
         MatchEntry      =   0
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "性質："
         Height          =   180
         Left            =   5920
         TabIndex        =   95
         Top             =   690
         Width           =   540
      End
      Begin VB.Label Label13 
         Caption         =   "註：網址欄位（快按二下）即可進入網站"
         ForeColor       =   &H000000C0&
         Height          =   915
         Left            =   8070
         TabIndex        =   94
         Top             =   930
         Width           =   765
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "網　址 3 ："
         Height          =   180
         Left            =   90
         TabIndex        =   93
         Top             =   1500
         Width           =   900
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "網　址 2 ："
         Height          =   180
         Left            =   90
         TabIndex        =   92
         Top             =   1230
         Width           =   900
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "更新週期：           (月)"
         Height          =   180
         Left            =   4150
         TabIndex        =   84
         Top             =   690
         Width           =   1695
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "建置日期："
         Height          =   180
         Left            =   2415
         TabIndex        =   83
         Top             =   690
         Width           =   900
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "平台名稱："
         Height          =   180
         Left            =   3120
         TabIndex        =   82
         Top             =   390
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "客　　戶："
         Height          =   180
         Index           =   2
         Left            =   -73440
         TabIndex        =   79
         Top             =   390
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "附件或憑證："
         Height          =   180
         Index           =   4
         Left            =   90
         TabIndex        =   76
         Top             =   3480
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "平台名稱："
         Height          =   180
         Index           =   13
         Left            =   -74730
         TabIndex        =   72
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "( 模糊比對 )"
         Height          =   180
         Index           =   16
         Left            =   -71700
         TabIndex        =   71
         Top             =   1080
         Width           =   930
      End
      Begin VB.Line Line1 
         X1              =   -72420
         X2              =   -72240
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "(系統自動給號)"
         Height          =   180
         Index           =   7
         Left            =   1740
         TabIndex        =   62
         Top             =   390
         Width           =   1200
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "下次更新日期："
         Height          =   180
         Index           =   2
         Left            =   -74730
         TabIndex        =   61
         Top             =   1380
         Width           =   1260
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Height          =   180
         Index           =   1
         Left            =   -72660
         TabIndex        =   60
         Top             =   750
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "客戶編號："
         Height          =   180
         Index           =   11
         Left            =   -74730
         TabIndex        =   59
         Top             =   450
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "使  用  者："
         Height          =   180
         Index           =   15
         Left            =   -74730
         TabIndex        =   58
         Top             =   750
         Width           =   900
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Height          =   180
         Index           =   0
         Left            =   -72660
         TabIndex        =   57
         Top             =   450
         Width           =   1215
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "網　址 1 ："
         Height          =   180
         Left            =   90
         TabIndex        =   56
         Top             =   960
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "備　　註："
         Height          =   180
         Index           =   0
         Left            =   90
         TabIndex        =   55
         Top             =   3000
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "平台類別："
         Height          =   180
         Left            =   90
         TabIndex        =   54
         Top             =   690
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "平台編號："
         Height          =   180
         Index           =   0
         Left            =   90
         TabIndex        =   53
         Top             =   390
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "客　　戶："
         Height          =   180
         Index           =   1
         Left            =   90
         TabIndex        =   52
         Top             =   1770
         Width           =   900
      End
   End
   Begin MSForms.Label Label23 
      Height          =   195
      Left            =   210
      TabIndex        =   100
      Top             =   5505
      Width           =   7905
      VariousPropertyBits=   27
      Caption         =   "CREATE :                                                    UPDATE : "
      Size            =   "13944;344"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
End
Attribute VB_Name = "frm140114"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/5/13 Form2.0已修改
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Created by Sindy 2012/9/12
Option Explicit

' 變數宣告區
Dim m_EditMode As Integer
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
Dim m_FieldList() As FIELDITEM
' 第一筆資料的本所案號
Dim m_FirstKEY As String
' 最後一筆資料的本所案號
Dim m_LastKEY As String
' 目前正在顯示的本所案號
Dim m_CurrKEY As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim tf_CW As Integer
Dim iLstSelRow As Integer  '前次點選的帳號資料列
Dim iCurrSelRow As Integer '目前點選的帳號資料列
Dim m_EditMode2 As Integer '帳號資料編輯狀態
Dim MyArr As Variant
Dim m_AttachPath As String
Dim i As Integer, j As Integer
Dim bolMuchCust As Boolean
Dim m_CurrKeyCD02 As String, m_CurrKeyCD02_Name As String, m_CurrKeyCD03 As String
Dim strCustID As String

'附件
Dim m_FilesRemoved() As String
Dim ii As Integer, jj As Integer
Private Type BrowseInfo
    hwndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As String
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function SHBrowseForFolder Lib "shell32" (lpBI As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long

Private Declare Function SendMessageByNum Lib "user32" _
  Alias "SendMessageA" (ByVal hWnd As Long, ByVal _
  wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Const LB_SETHORIZONTALEXTENT = &H194


Private Function BrowseForFolder(Optional sCaption As String = "請選擇欲儲存的位置", Optional sDefault As String) As String
    Const BIF_RETURNONLYFSDIRS = 1
    Const MAX_PATH = 260
    Dim lPos As Integer, lpIDList As Long, lResult As Long
    Dim sPath As String, tBrowse As BrowseInfo

    With tBrowse
        'Set the owner window
        .hwndOwner = GetActiveWindow        'Me.hWnd in VB
        .lpszTitle = sCaption
        .ulFlags = BIF_RETURNONLYFSDIRS     'Return only if the user selected a directory
    End With

    'Show the dialog
    lpIDList = SHBrowseForFolder(tBrowse)
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        'Get the path from the IDList
        SHGetPathFromIDList lpIDList, sPath
        CoTaskMemFree lpIDList
        lPos = InStr(sPath, vbNullChar)
        If lPos Then
            BrowseForFolder = Left$(sPath, lPos - 1)
            If Right$(BrowseForFolder, 1) <> "\" Then
                BrowseForFolder = BrowseForFolder & "\"
            End If
        End If
    Else
        'User cancelled, return default path
        BrowseForFolder = sDefault
    End If
End Function

Private Sub cboCW19_GotFocus()
    InverseTextBox cboCW19
End Sub

Private Sub cmdok_Click(Index As Integer)
Dim nResponse As Boolean
Dim strTit As String
Dim strMsg As String
Dim bDifference As Boolean
Dim strText As String
   
On Error GoTo ErrHand
   
   bDifference = False
   Select Case Index
      Case 0 '新增
         iLstSelRow = -1
         For i = 1 To GRD1.Rows - 1
            GRD1.row = i
            GRD1.col = 0
            If GRD1.CellBackColor = &HFFC0C0 Then
               iLstSelRow = i
               Exit For
            End If
         Next
         ClearCD
         If GRD1.TextMatrix(GRD1.Rows - 1, 1) <> "" Then
            GRD1.Rows = GRD1.Rows + 1
         End If
         GRD1.row = GRD1.Rows - 1
         grd1_SelChange
         
         cmdOK(0).Enabled = False
         cmdOK(1).Enabled = True
         cmdOK(2).Enabled = True
         cmdOK(3).Enabled = False
         cmdOK(4).Enabled = False
         GRD1.Enabled = False
         m_EditMode2 = 1
         EnableCD True
         textCD(3).SetFocus
         If Check1.Value = False Then
            m_CurrKeyCD02 = textCW(1)
            m_CurrKeyCD02_Name = textCW(1)
         Else
            MyArr = Split(lstUsers(2).List(lstUsers(2).ListIndex), "@")
            m_CurrKeyCD02 = Trim(MyArr(1))
            m_CurrKeyCD02_Name = Trim(MyArr(0))
         End If
         textCD(9) = strSrvDate(2)
      Case 1 '確定
         For i = 1 To GRD1.Rows - 1
            GRD1.row = i
            GRD1.col = 0
            If GRD1.CellBackColor = &HFFC0C0 Then
               If CheckDataValid_CD() = False Then Exit Sub
               '重新檢查欄位有效性
               If TxtValidate_CD = False Then Exit Sub
               Screen.MousePointer = vbHourglass
               GRD1.TextMatrix(i, 1) = m_CurrKeyCD02
               GRD1.TextMatrix(i, 2) = m_CurrKeyCD02_Name
               If m_EditMode2 = 1 Then '新增
                  m_CurrKeyCD03 = textCD(3).Text
               End If
               GRD1.TextMatrix(i, 3) = textCD(3).Text
               GRD1.TextMatrix(i, 4) = textCD(4).Text
               GRD1.TextMatrix(i, 5) = cboCD05.Text
               '轉換使用者姓名
               If ChangeCD06CN(textCD(6).Text, strText) = True Then
                  GRD1.TextMatrix(i, 6) = strText
               Else
                  GRD1.TextMatrix(i, 6) = textCD(6).Text
               End If
               GRD1.TextMatrix(i, 7) = ChangeTStringToTDateString(textCD(9).Text)
               GRD1.TextMatrix(i, 8) = ChangeTStringToTDateString(textCD(7).Text)
               GRD1.TextMatrix(i, 9) = textCD(8).Text
               GRD1.TextMatrix(i, 10) = textCD(6).Text
               cnnConnection.BeginTrans
               bDifference = True
               MyArr = Split(cboCD05, " ") '身份別
               If m_EditMode2 = 1 Then '新增
                  strSql = "insert into custwebid values('" & textCW(1) & "','" & m_CurrKeyCD02 & "'," & _
                           "'" & textCD(3) & "','" & textCD(4) & "','" & Trim(MyArr(0)) & "','" & textCD(6) & "'," & _
                           DBDATE(textCD(7)) & ",'" & textCD(8) & "'," & DBDATE(textCD(9)) & ")"
               ElseIf m_EditMode2 = 2 Then '修改
                  strSql = "update custwebid set " & _
                           "cd04='" & textCD(4) & "'," & _
                           "cd05='" & Trim(MyArr(0)) & "'," & _
                           "cd06='" & textCD(6) & "'," & _
                           "cd07=" & DBDATE(textCD(7)) & "," & _
                           "cd08='" & textCD(8) & "', " & _
                           "cd09=" & DBDATE(textCD(9)) & " " & _
                           "where cd01='" & textCW(1) & "' and cd02='" & m_CurrKeyCD02 & "' and cd03='" & m_CurrKeyCD03 & "'"
               End If
               Pub_SeekTbLog strSql
               cnnConnection.Execute strSql
               '當為修改狀態時,更新帳號資料要一併更新主檔的UpdateID,Date,Time
               If m_EditMode = 2 Then
                  'Modified by Morgan 2023/6/16 時間要存6碼
                  'strSql = "update custweb set " & _
                           "cw09='" & strUserNum & "'," & _
                           "cw10=" & strSrvDate(1) & "," & _
                           "cw11=" & Left(ServerTime, Len(CStr(ServerTime)) - 2) & " " & _
                           "where cw01='" & textCW(1) & "'"
                  strSql = "update custweb set " & _
                           "cw09='" & strUserNum & "'," & _
                           "cw10=" & strSrvDate(1) & "," & _
                           "cw11=to_char(sysdate,'HH24MISS')" & " " & _
                           "where cw01='" & textCW(1) & "'"
                  'end 2023/6/16
                  cnnConnection.Execute strSql
               End If
               cnnConnection.CommitTrans
               GRD1.Refresh
               cmdOK(0).Enabled = True
               cmdOK(1).Enabled = False
               cmdOK(2).Enabled = False
               cmdOK(3).Enabled = True
               cmdOK(4).Enabled = True
               m_EditMode2 = 0
               EnableCD False
               Screen.MousePointer = vbDefault
               Exit For
            End If
         Next
         GRD1.Enabled = True
         
      Case 2 '取消
         If GRD1.TextMatrix(GRD1.Rows - 1, 1) = "" Then
            If GRD1.Rows = 2 Then
               GRD1.Clear
               Call SetGrd(bolMuchCust)
            Else
               GRD1.RemoveItem GRD1.Rows - 1
            End If
         End If

         If iLstSelRow >= 0 Then
            GRD1.row = iLstSelRow
            GRD1.col = 0
            GRD1.CellBackColor = QBColor(15) '顏色還原以便重新讀取資料
            grd1_SelChange
         Else
            ClearCD
         End If

         cmdOK(0).Enabled = True
         cmdOK(1).Enabled = False
         cmdOK(2).Enabled = False
         cmdOK(3).Enabled = True
         cmdOK(4).Enabled = True
         GRD1.Enabled = True
         m_EditMode2 = 0
         EnableCD False
         
      Case 3 '刪除
         For i = 1 To GRD1.Rows - 1
            GRD1.row = i
            GRD1.col = 0
            If GRD1.CellBackColor = &HFFC0C0 Then
               strTit = "詢問"
               strMsg = "是否確定要刪除此筆資料?"
               If MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit) = vbYes Then
                  Screen.MousePointer = vbHourglass
                  cnnConnection.BeginTrans
                  bDifference = True
                  strSql = "delete custwebid " & _
                           "where cd01='" & textCW(1) & "' and cd02='" & m_CurrKeyCD02 & "' and cd03='" & m_CurrKeyCD03 & "'"
                  Pub_SeekTbLog strSql
                  cnnConnection.Execute strSql
                  '當為修改狀態時,更新帳號資料要一併更新主檔的UpdateID,Date,Time
                  If m_EditMode = 2 Then
                     'Modified by Morgan 2023/6/16 時間要存6碼
                     'strSql = "update custweb set " & _
                              "cw09='" & strUserNum & "'," & _
                              "cw10=" & strSrvDate(1) & "," & _
                              "cw11=" & Left(ServerTime, Len(CStr(ServerTime)) - 2) & " " & _
                              "where cw01='" & textCW(1) & "'"
                     strSql = "update custweb set " & _
                              "cw09='" & strUserNum & "'," & _
                              "cw10=" & strSrvDate(1) & "," & _
                              "cw11=to_char(sysdate,'HH24MISS')" & " " & _
                              "where cw01='" & textCW(1) & "'"
                     'end 2023/6/16
                     cnnConnection.Execute strSql
                  End If
                  cnnConnection.CommitTrans
                  If GRD1.Rows = 2 Then
                      GRD1.Clear
                      Call SetGrd(bolMuchCust)
                  Else
                      GRD1.RemoveItem i
                  End If
                  ClearCD
                  GRD1.Refresh
                  Screen.MousePointer = vbDefault
               End If
               Exit For
            End If
         Next i
         m_EditMode2 = 0
         
      Case 4 '修改
         For i = 1 To GRD1.Rows - 1
            GRD1.row = i
            GRD1.col = 0
            If GRD1.CellBackColor = &HFFC0C0 Then
               iLstSelRow = i
               cmdOK(0).Enabled = False
               cmdOK(1).Enabled = True
               cmdOK(2).Enabled = True
               cmdOK(3).Enabled = False
               cmdOK(4).Enabled = False
               GRD1.Enabled = False
               m_EditMode2 = 2
               EnableCD True
               Exit For
            End If
         Next i
      Case Else
   End Select
   Exit Sub

ErrHand:
   Screen.MousePointer = vbDefault
   If bDifference = True Then
      cnnConnection.RollbackTrans
      MsgBox "儲存失敗！" & vbCrLf & Err.Description
   Else
      MsgBox Err.Description
   End If
End Sub

'進入網站
Private Sub cmdOpenIE_Click()
   If Trim(textCW(2)) = "" Then
      MsgBox "請輸入網址！"
      Call textCW_GotFocus(2)
      textCW(2).SetFocus
      Exit Sub
   End If
   Call OpenIE(Trim(textCW(2)))
End Sub

Private Sub Form_Initialize()
   Set rsA = New ADODB.Recordset
   If rsA.State = 1 Then rsA.Close
   rsA.CursorLocation = adUseClient
   rsA.Open "select * from custweb where rownum <2 ", cnnConnection, adOpenStatic, adLockReadOnly
   tf_CW = rsA.Fields.Count
   'Call SetGrd(False)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   'If Me.ActiveControl.Index = 5 And Me.ActiveControl.Name = "textCW" And KeyCode = 13 Then
   
   'Modify By Sindy 2019/1/28
   'If KeyCode = 13 And Not (Me.ActiveControl.Name = "textCW" And Me.ActiveControl.Index = 1) Then
   If KeyCode = 13 Then
      If Not (Me.ActiveControl.Name = "textCW" And Me.ActiveControl.Index = 1) Then
   '2019/1/28 END
         '擋掉Enter動作
         KeyCode = 0
         Exit Sub
      End If
   End If
   
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
         'Modify By Sindy 2021/3/31 薛經理:使用home鍵，會直接跳到第一筆記錄，請取消。(似乎與快速鍵衝突?)
'         If m_bQuery Then
'            If m_EditMode = 0 Then
'               OnAction KeyCode
'               KeyCode = 0
'            End If
'         End If
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
   
   MoveFormToCenter Me
   'Modify By Sindy 2021\5\19
   'm_AttachPath = App.path & "\SeminarAttach"
   m_AttachPath = App.path & "\SeminarAttach\" & strUserNum
   '2021\5\19 END
   ReDim m_FieldList(tf_CW) As FIELDITEM
   
   m_bInsert = IsUserHasRightOfFunction(Me.Name, strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
   m_bDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   m_bQuery = IsUserHasRightOfFunction(Me.Name, strFind, False)
'   m_bOpen = IsUserHasRightOfFunction(Me.Name, strPrint, False)
   
   ReDim m_FilesRemoved(0)
   
   textCW(1).BackColor = &H8000000F
   lblName(0).BackColor = &H8000000F
   lblName(1).BackColor = &H8000000F
   Lbl1(0).BackColor = &H8000000F
   Lbl1(1).BackColor = &H8000000F
   textCW(4).Visible = False
   textCW(14).Visible = False
   textCW(15).Visible = False
   textCD(6).Visible = False
   
   InitialField
   InitialData
   RefreshRange
   ShowFirstRecord
   UpdateToolbarState
   SetCtrlReadOnly True
   'OnAction vbKeyF4 '按查詢
   OnAction vbKeyF10 '按取消
   Me.SSTab1.Tab = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   KillAttach
   Set frm140114 = Nothing
End Sub

Private Sub KillAttach()
On Error Resume Next
   'Modify By Sindy 2021/2/3 刪不掉,改用函數
'   If Dir(m_AttachPath & "\.") <> "" Then
'      Kill m_AttachPath & "\*.*"
'   End If
   'Modify By Sindy 2021\5\19
   'PUB_KillTempFile "SeminarAttach\*.*"
   PUB_KillTempFile "SeminarAttach\" & strUserNum & "\*.*"
   '2021\5\19 END
   '2021/2/3 END
End Sub

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
   
   getGrdColRow GRD1, x, y, nCol, nRow
   GRD1.col = nCol
   GRD1.row = nRow
End Sub

Private Sub grd1_SelChange()
Dim tmpMouseRow
Dim strText As String
   
   GRD1.Visible = False
   tmpMouseRow = GRD1.row
   GRD1.Visible = True
   If tmpMouseRow <> 0 Then
       GRD1.row = tmpMouseRow
       GRD1.col = 0
       If GRD1.CellBackColor = QBColor(15) Then
            GRD1.Visible = False
            For j = 1 To GRD1.Rows - 1
                GRD1.row = j
                For i = 0 To GRD1.Cols - 1
                     GRD1.col = i
                     GRD1.CellBackColor = QBColor(15)
                Next i
            Next j
            GRD1.row = tmpMouseRow
            For i = 0 To GRD1.Cols - 1
                GRD1.col = i
                GRD1.CellBackColor = &HFFC0C0
            Next i
            If m_EditMode <> 0 Then
               'cmdOK(1).Enabled = True
               'cmdOK(2).Enabled = True
            End If
            iCurrSelRow = tmpMouseRow
            m_CurrKeyCD02 = GRD1.TextMatrix(tmpMouseRow, 1)
            m_CurrKeyCD02_Name = GRD1.TextMatrix(tmpMouseRow, 2)
            m_CurrKeyCD03 = GRD1.TextMatrix(tmpMouseRow, 3)
            textCD(3).Text = GRD1.TextMatrix(tmpMouseRow, 3)
            textCD(4).Text = GRD1.TextMatrix(tmpMouseRow, 4)
            SelCombo cboCD05, Left(GRD1.TextMatrix(tmpMouseRow, 5), 1), 1
            '轉換使用者姓名
            textCD(6).Text = GRD1.TextMatrix(tmpMouseRow, 10)
            SetlstUsers 0, textCD(6)
            textCD(9).Text = ChangeTDateStringToTString(GRD1.TextMatrix(tmpMouseRow, 7))
            textCD(7).Text = Trim(ChangeTDateStringToTString(GRD1.TextMatrix(tmpMouseRow, 8))) 'Modify By Sindy 2016/11/18 + Trim : 輸入111111時,顯示會多一個空白在前面,因此Trim掉
            textCD(8).Text = GRD1.TextMatrix(tmpMouseRow, 9)
            GRD1.Visible = True
       End If
   End If
End Sub

Private Sub lstUsers_Click(Index As Integer)
   Select Case Index
      'Added by Morgan 2016/9/8
      Case 1
         If lstUsers(Index).ListIndex >= 0 Then
            If lstUsers(Index).Selected(lstUsers(Index).ListIndex) Then
               'Modify By Sindy 2021/5/12
               'txtUserNo(Index).Text = PUB_Num2Id(lstUsers(Index).ItemData(lstUsers(Index).ListIndex), "1")
               txtUserNo(Index).Text = PUB_GetItemData(lstUsers(Index).Tag, lstUsers(Index).ListIndex)
               '2021/5/12 END
            Else
               txtUserNo(Index).Text = ""
            End If
         Else
            txtUserNo(Index).Text = ""
         End If
      'end 2016/9/8
      Case 2
         If lstUsers(Index).ListIndex < 0 Then
            Call UpdateCtrlData2(textCW(1)) '查詢共同帳號資料
         Else
            MyArr = Split(lstUsers(Index).List(lstUsers(Index).ListIndex), "@")
            m_CurrKeyCD02 = Trim(MyArr(1))
            m_CurrKeyCD02_Name = Trim(MyArr(0))
            Call UpdateCtrlData2(m_CurrKeyCD02) '查詢該客戶的帳號資料
         End If
         Exit Sub
   End Select
End Sub

Private Sub MSHFlexGrid1_DblClick()
Dim tmpMouseRow
Dim strText As String
   
   MSHFlexGrid1.Visible = False
   tmpMouseRow = MSHFlexGrid1.row
   MSHFlexGrid1.Visible = True
   If tmpMouseRow <> 0 Then
       MSHFlexGrid1.row = tmpMouseRow
       MSHFlexGrid1.col = 0
       If MSHFlexGrid1.CellBackColor <> QBColor(15) Then
'            MSHFlexGrid1.Visible = False
'            For j = 1 To MSHFlexGrid1.Rows - 1
'                MSHFlexGrid1.row = j
'                For i = 0 To MSHFlexGrid1.Cols - 1
'                     MSHFlexGrid1.col = i
'                     MSHFlexGrid1.CellBackColor = QBColor(15)
'                Next i
'            Next j
'            MSHFlexGrid1.row = tmpMouseRow
'            For i = 0 To MSHFlexGrid1.Cols - 1
'                MSHFlexGrid1.col = i
'                MSHFlexGrid1.CellBackColor = &HFFC0C0
'            Next i
'            MSHFlexGrid1.Visible = True
            Screen.MousePointer = vbHourglass
            m_CurrKEY = MSHFlexGrid1.TextMatrix(tmpMouseRow, 0)
            UpdateCtrlData
            Me.SSTab1.Tab = 1
            Screen.MousePointer = vbDefault
       End If
   End If
End Sub

Private Sub MSHFlexGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
   
   getGrdColRow MSHFlexGrid1, x, y, nCol, nRow
   MSHFlexGrid1.col = nCol
   MSHFlexGrid1.row = nRow
End Sub

Private Sub MSHFlexGrid1_SelChange()
Dim tmpMouseRow
Dim strText As String
   
   MSHFlexGrid1.Visible = False
   tmpMouseRow = MSHFlexGrid1.row
   MSHFlexGrid1.Visible = True
   If tmpMouseRow <> 0 Then
       MSHFlexGrid1.row = tmpMouseRow
       MSHFlexGrid1.col = 0
       If MSHFlexGrid1.CellBackColor = QBColor(15) Then
            MSHFlexGrid1.Visible = False
            For j = 1 To MSHFlexGrid1.Rows - 1
                MSHFlexGrid1.row = j
                For i = 0 To MSHFlexGrid1.Cols - 1
                     MSHFlexGrid1.col = i
                     MSHFlexGrid1.CellBackColor = QBColor(15)
                Next i
            Next j
            MSHFlexGrid1.row = tmpMouseRow
            For i = 0 To MSHFlexGrid1.Cols - 1
                MSHFlexGrid1.col = i
                MSHFlexGrid1.CellBackColor = &HFFC0C0
            Next i
            MSHFlexGrid1.Visible = True
'            Screen.MousePointer = vbHourglass
'            m_CurrKEY = MSHFlexGrid1.TextMatrix(tmpMouseRow, 0)
'            UpdateCtrlData
'            Me.SSTab1.Tab = 2
'            Screen.MousePointer = vbDefault
       End If
   End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
Dim bolMustChk As Boolean
   
   cmdQuery(0).Default = False
   If SSTab1.Tab = 1 And (m_EditMode = 1 Or m_EditMode = 2) Then
      bolMustChk = False
      If lstUsers(1).ListCount <> lstUsers(2).ListCount Then
         bolMustChk = True
      Else
         For i = 0 To lstUsers(1).ListCount - 1
            If lstUsers(1).List(i) <> lstUsers(2).List(i) Then
               bolMustChk = True
               Exit For
            End If
         Next
      End If
      If bolMustChk = True Then
         ChkIsMuchCust '檢查是否為多筆客戶
         Call UpdateCtrlData2(strCustID) '查詢帳號資料
      End If
   ElseIf SSTab1.Tab = 2 Then
      cmdQuery(0).Default = True
   End If
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

' 更新 Create 及 Update 的人
Private Sub UpdateCUID(ByRef rsSrcTmp As ADODB.Recordset)
Dim strTemp As String
Dim strCName As String
Dim strCDate As String
Dim strCTime As String
Dim strUName As String
Dim strUDate As String
Dim strUTime As String
   
   If IsNull(rsSrcTmp.Fields("cw06")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("cw06")) = False Then
         strCName = GetStaffName(rsSrcTmp.Fields("cw06"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("cw07")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("cw07")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("cw07"))
         strCDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("cw08")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("cw08")) = False Then
         strTemp = rsSrcTmp.Fields("cw08")
         'Modified by Morgan 2023/6/16
         'strCTime = Format(strTemp, "00:00:00")
         strCTime = Format(strTemp, "##:##:##")
         'end 2023/6/16
      End If
   End If
   If IsNull(rsSrcTmp.Fields("cw09")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("cw09")) = False Then
         strUName = GetStaffName(rsSrcTmp.Fields("cw09"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("cw10")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("cw10")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("cw10"))
         strUDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("cw11")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("cw11")) = False Then
         strTemp = rsSrcTmp.Fields("cw11")
         'Modified by Morgan 2023/6/16
         'strUTime = Format(strTemp, "##:##:##")
         strUTime = Format(strTemp, "00:00:00")
         'end 2023/6/16
      End If
   End If
   
   ' 設定CUID中的文字
   Label23.Caption = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime & String(10, " ") & _
              "UPDATE : " & strUName & " " & _
              " " & strUDate & " " & _
              " " & strUTime
End Sub

Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

   TxtValidate = False
   
   If Me.textCW(2).Enabled = True Then
      Cancel = False
      textCW_Validate 2, Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   'Add by Sindy 2021/5/13 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me) = False Then
      Exit Function
   End If
   '2021/5/13 END
   
   TxtValidate = True
End Function

Private Function TxtValidate_CD() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

   TxtValidate_CD = False
   
   If Me.textCD(7).Enabled = True Then
      Cancel = False
      textCD_Validate 7, Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   'Add by Sindy 2021/5/13 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me) = False Then
      Exit Function
   End If
   '2021/5/13 END
   
   TxtValidate_CD = True
End Function

' 設定欄位的內容
Private Sub SetFieldNewData(ByVal strName As String, Optional ByVal strData As String = "#==#")
Dim nIndex As Integer
   
   For nIndex = 0 To tf_CW - 1
      If strName = m_FieldList(nIndex).fiName Then
         If strData = "#==#" Then
            m_FieldList(nIndex).fiNewData = m_FieldList(nIndex).fiOldData
         Else
            m_FieldList(nIndex).fiNewData = strData
         End If
         Exit For
      End If
   Next nIndex
End Sub

' 從記錄中更新欄位內容
Private Sub UpdateFieldOldData(ByRef rsTmp As ADODB.Recordset)
Dim nIndex As Integer
Dim strTmp As String
   
   For nIndex = 0 To tf_CW - 1
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
EXITSUB:
End Sub

' 新增記錄
Private Function AddRecord() As Boolean
Dim strTmp As String
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim nIndex As Integer
Dim bDifference As Boolean
Dim bFirst As Boolean
Dim strCW01 As String
   
   AddRecord = False
   
   strCW01 = textCW(1)

   ' 檢查記錄是否已存在
   If IsRecordExist(strCW01) = True Then
      strTit = "新增資料"
      strMsg = "該筆記錄已存在"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      UpdateCtrlData
      Exit Function
   End If
   
   bFirst = True
   bDifference = False
   strSql = "INSERT INTO custweb ("
   For nIndex = 0 To tf_CW - 1
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
   For nIndex = 0 To tf_CW - 1
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
   
   Call SaveCustWebFile
   
   cnnConnection.CommitTrans
   
   If ((strCW01) < (m_FirstKEY)) Or ((strCW01) > (m_LastKEY)) Then
      RefreshRange
   End If
   
   ShowCurrRecord strCW01
   AddRecord = True
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox " 新增失敗！" & vbCrLf & Err.Description
End Function

Private Sub SaveCustWebFile()
Dim stFilePath As String
Dim iFileNo As Integer
Dim bytes() As Byte
Dim lngSize As Long '檔案大小
Dim adoRst As New ADODB.Recordset
Const BlockSize = 500000
Dim Numblocks As Integer
Dim LeftOver As Long
Dim stReName As String, strFtpPath As String
   
   For ii = 0 To lstAtt.ListCount - 1
      'Modify By Sindy 2021/5/13
      'If lstAtt.ItemData(ii) = 0 Then
      If InStr(lstAtt.Tag, lstAtt.List(ii)) = 0 Then
      '2021/5/13 END
         stFilePath = lstAtt.List(ii)
         stFilePath = Left(stFilePath, InStrRev(stFilePath, " (") - 1)
         If iFileNo > 0 Then Close #iFileNo
         iFileNo = FreeFile
         Open stFilePath For Binary Access Read As #iFileNo
         lngSize = LOF(iFileNo)
         Close #iFileNo
         'Add By Sindy 2017/5/25
         '改上傳FTP File Server
         stReName = lngSize & "." & GetFileName(stFilePath)
         PUB_PutFtpFile stFilePath, textCW(1), stReName, strFtpPath, UCase("custwebfile")
         If strFtpPath <> "" Then
            'Modified by Lydia 2023/06/07 +流水號CF04
            'strSql = "insert into custwebfile(cf01,cf02,cf03,cf08) " & _
                     "values(" & CNULL(textCW(1)) & "," & CNULL(GetFileName(stFilePath)) & _
                     "," & lngSize & "," & CNULL(strFtpPath) & ")"
            strSql = "insert into custwebfile(cf01,cf02,cf03,cf08,cf04) " & _
                     "select " & CNULL(textCW(1)) & " as cf01," & CNULL(GetFileName(stFilePath)) & " as cf02 " & _
                     "," & lngSize & " as cf03, " & CNULL(strFtpPath) & " as cf08,max(cf04)+1 as newcf04 from custwebfile "
            cnnConnection.Execute strSql
         End If
      End If
   Next
End Sub

' 修改記錄
Private Function ModRecord() As Boolean
Dim strTmp As String
Dim nIndex As Integer
Dim bDifference As Boolean
Dim bFirst As Boolean
Dim strCW01 As String

   ModRecord = False
   
   strCW01 = m_CurrKEY
   
   strSql = "begin user_data.user_enabled:=1; UPDATE custweb SET "

   bFirst = True
   bDifference = False
   For nIndex = 0 To tf_CW - 1
      strTmp = Empty
      'If nIndex < 42 Or nIndex > 47 Then
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
      'End If
   Next nIndex

   strSql = strSql & " " & _
                  "WHERE CW01 = '" & strCW01 & "' ; end; "
'On Error GoTo ErrHand
   cnnConnection.BeginTrans
   If bDifference = True Then
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql
   End If
   
   '刪除附件
   For ii = 1 To UBound(m_FilesRemoved)
      strSql = "delete custwebfile where cf01=" & strCW01 & " and cf02='" & ChgSQL(m_FilesRemoved(ii)) & "'"
      cnnConnection.Execute strSql, intI
   Next
   
   Call SaveCustWebFile
   
   cnnConnection.CommitTrans
   
   ShowCurrRecord strCW01
   ModRecord = True
   Exit Function
   
ErrHand:
   cnnConnection.RollbackTrans
   MsgBox (Err.Description)
End Function

' 刪除記錄
Private Function DelRecord(strDelKey As String) As Boolean
Dim strCW01 As String
   
   DelRecord = False
   
On Error GoTo ErrHand

   'Add By Sindy 2025/11/10 平台資料刪除時，要先檢查是否已有往來記錄，若有，則不可刪除平台資料。
   strExc(0) = "select * from ContactRecord where cr03='" & strDelKey & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      MsgBox "此平台有建立往來記錄，不可刪除平台資料！", vbExclamation
      Exit Function
   End If
   '2025/11/10 END
   
   cnnConnection.BeginTrans
   
   strCW01 = strDelKey
   
   '客戶平台主檔
   strSql = "DELETE FROM custweb WHERE cw01 = '" & strDelKey & "'  "
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql
   
   '客戶平台帳號資料
   strSql = "DELETE FROM custwebid WHERE cd01 = '" & strDelKey & "'  "
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql
   
   '客戶平台附件檔
   PUB_DelFtpFile2 textCW(1), , UCase("custwebfile") 'Add By Sindy 2017/5/25 檔案改放 FTP,必須在DB資料刪除前執行
   strSql = "DELETE FROM custwebfile WHERE cf01 = '" & strDelKey & "'  "
   cnnConnection.Execute strSql
   
   If (strCW01 = m_LastKEY) Or (strCW01 = m_FirstKEY) Then
      RefreshRange
   End If
   ShowCurrRecord strCW01
   DelRecord = True
   cnnConnection.CommitTrans
   
   Exit Function
   
ErrHand:
   cnnConnection.RollbackTrans
   MsgBox "刪除失敗！" & vbCrLf & Err.Description
End Function

' 查詢記錄
Public Function QueryRecord(strCW01) As Boolean
      
   QueryRecord = False
   If IsRecordExist(strCW01) = True Then
      m_CurrKEY = strCW01
      QueryRecord = True
      UpdateCtrlData
   Else
      QueryRecord = False
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
         If DelRecord(m_CurrKEY) = True Then
            RefreshRange
            ClearField
            ShowCurrRecord m_CurrKEY
         Else
            Exit Function
         End If
      Case 4: '查詢
         If textCW(1) <> "" Then
            textCW(1) = Format(textCW(1), "0000")
            If QueryRecord(textCW(1)) = False Then
               strMsg = "無此資料"
               strTit = "查詢資料"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               UpdateCtrlData
            End If
         Else
            GoTo EXITSUB
         End If
   End Select
   m_EditMode = 0
   SetCtrlReadOnly True
   OnWork = True
EXITSUB:
End Function

' 開始輸入資料
Private Sub SetInputEntry()
   Select Case m_EditMode
      Case 1: If Me.Visible = True Then textCW12.SetFocus
      Case 2: If Me.Visible = True Then textCW12.SetFocus
      Case 4: If Me.Visible = True Then textCW(1).SetFocus
   End Select
End Sub

' 檢查平台記錄是否已經存在
Private Function IsRecordExist(ByVal strKEY01 As String) As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   IsRecordExist = False
      
   strSql = "SELECT cw01 FROM custweb " & _
            "WHERE cw01 = '" & strKEY01 & "' "
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

' 檢查網址是否已經存在
Private Function IsRecordExist_CW02(ByVal strKEY01 As String) As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String, strCon As String
   
   IsRecordExist_CW02 = False
   
   If m_EditMode = 2 Then '修改時
      strCon = "and CW01<>'" & textCW(1) & "' "
   End If
   
   strSql = "SELECT cw01,cw02,cw17,cw18 FROM custweb " & _
            "WHERE (upper(cw02)='" & UCase(strKEY01) & "' or upper(cw17)='" & UCase(strKEY01) & "' or upper(cw18)='" & UCase(strKEY01) & "') " & _
            strCon
   ' 讀取資料庫
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   ' 檢查讀取的資料筆數
   If rsTmp.RecordCount > 0 Then
      IsRecordExist_CW02 = True
   Else
      IsRecordExist_CW02 = False
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 檢查帳號記錄是否已經存在
Private Function IsRecordExist_CD(ByVal strKEY01 As String, ByVal strKEY02 As String, ByVal strKEY03 As String) As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   IsRecordExist_CD = False
   
   strSql = "SELECT cd01,cd02,cd03 FROM custwebid " & _
            "WHERE cd01 = '" & strKEY01 & "' " & _
              "and cd02 = '" & strKEY02 & "' " & _
              "and cd03 = '" & strKEY03 & "' "
   ' 讀取資料庫
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   ' 檢查讀取的資料筆數
   If rsTmp.RecordCount > 0 Then
      IsRecordExist_CD = True
   Else
      IsRecordExist_CD = False
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 檢查此客戶帳號資料是否已經存在
Private Function IsRecordExist_CD02(ByVal strKEY01 As String, ByVal strKEY02 As String) As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   IsRecordExist_CD02 = False
   
   strSql = "SELECT cd01,cd02,cd03 FROM custwebid " & _
            "WHERE cd01 = '" & strKEY01 & "' " & _
              "and cd02 = '" & strKEY02 & "' "
   ' 讀取資料庫
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   ' 檢查讀取的資料筆數
   If rsTmp.RecordCount > 0 Then
      IsRecordExist_CD02 = True
   Else
      IsRecordExist_CD02 = False
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 顯示資料
Private Sub ShowCurrRecord(ByVal strKEY01 As String)
Dim strSql As String
Dim rsTmp As New ADODB.Recordset

   If IsRecordExist(strKEY01) = True Then
      m_CurrKEY = strKEY01
   Else
      strSql = "SELECT cw01 FROM custweb " & _
               "WHERE cw01 = '" & m_CurrKEY & "' "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("cw01")) = False Then: m_CurrKEY = rsTmp.Fields("cw01")
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      End If
      rsTmp.Close
      
      strSql = "SELECT cw01 FROM custweb " & _
               "WHERE cw01 = (SELECT MIN(cw01) FROM custweb)"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("cw01")) = False Then: m_CurrKEY = rsTmp.Fields("cw01")
      Else
         ShowLastRecord
         GoTo EXITSUB
      End If
      rsTmp.Close
   End If
   
   UpdateCtrlData
EXITSUB:
End Sub

' 顯示第一筆資料
Private Sub ShowFirstRecord()
   m_CurrKEY = m_FirstKEY
  
   UpdateCtrlData
End Sub

' 顯示上一筆資料
Private Sub ShowPrevRecord()
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   If m_CurrKEY = m_FirstKEY Then
      ShowMsg MsgText(9008)
      GoTo EXITSUB
   End If
   
   strSql = "SELECT cw01 FROM custweb " & _
            "WHERE cw01 = (SELECT MAX(cw01) FROM custweb " & _
                          "WHERE cw01 < '" & m_CurrKEY & "' )"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("cw01")) = False Then: m_CurrKEY = rsTmp.Fields("cw01")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT cw01 FROM custweb " & _
            "WHERE cw01 = (SELECT Min(cw01) FROM custweb ) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("cw01")) = False Then: m_CurrKEY = rsTmp.Fields("cw01")
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
   
   If m_CurrKEY = m_LastKEY Then
      ShowMsg MsgText(9009)
      GoTo EXITSUB
   End If
   
   strSql = "SELECT cw01 FROM custweb " & _
            "WHERE cw01 = (SELECT MIN(cw01) FROM custweb " & _
                          "WHERE cw01  > '" & m_CurrKEY & "' )"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("cw01")) = False Then: m_CurrKEY = rsTmp.Fields("cw01")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT cw01 FROM custweb " & _
            "WHERE cw01 = (SELECT max(cw01) FROM custweb ) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("cw01")) = False Then: m_CurrKEY = rsTmp.Fields("cw01")
   End If
   rsTmp.Close
   
   UpdateCtrlData
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 顯示最後一筆資料
Private Sub ShowLastRecord()
   m_CurrKEY = m_LastKEY
   UpdateCtrlData
End Sub

' 執行指令
Private Sub OnAction(ByVal KeyCode As Integer)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
'   cmdOpenAtt(0).Enabled = True
'   cmdSelect(0).Enabled = True
'   cmdSaveAtt(0).Enabled = True
   Select Case KeyCode
      ' 新增
      Case vbKeyF2:
         m_EditMode = 1
         ClearField
         textCW(1) = GetNewCW01
         SetKeyReadOnly True
         SetCtrlReadOnly False
         UpdateToolbarState
         SetInputEntry
         Me.SSTab1.Tab = 0
         cmdAddAtt(0).Enabled = True
         cmdRemAtt(0).Enabled = True
         SSTab1.TabEnabled(2) = False
      ' 修改
      Case vbKeyF3:
         m_EditMode = 2
         SetCtrlReadOnly False
         SetKeyReadOnly True
         UpdateToolbarState
         SetInputEntry
         cmdAddAtt(0).Enabled = True
         cmdRemAtt(0).Enabled = True
         SSTab1.TabEnabled(2) = False
      ' 刪除
      Case vbKeyF5:
         strTit = "詢問"
         strMsg = "是否確定要刪除此筆資料?"
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
         SetCtrlReadOnly True
         SetKeyReadOnly False
         ClearField
         UpdateToolbarState
         SetInputEntry
         SSTab1.Tab = 0
      ' 第一筆
      Case vbKeyHome:
         ShowFirstRecord
         SSTab1.Tab = 0
      ' 前一筆
      Case vbKeyPageUp:
         ShowPrevRecord
         SSTab1.Tab = 0
      ' 後一筆
      Case vbKeyPageDown:
         ShowNextRecord
         SSTab1.Tab = 0
      ' 最後一筆
      Case vbKeyEnd:
         ShowLastRecord
         SSTab1.Tab = 0
      ' 確定
      Case vbKeyF9:
         If cmdOK(1).Enabled = True Then
            If MsgBox("帳號資料尚未作業完畢，確定放棄編輯嗎?", vbYesNo + vbDefaultButton2) = vbNo Then
               SSTab1.Tab = 1
               Exit Sub
            End If
         End If
         ' 將所有欄位的內容更新到欄位串列中的欄位內容項目
         UpdateFieldNewData
         If OnWork = True Then
            UpdateToolbarState
         Else
            Exit Sub
         End If
         SSTab1.TabEnabled(2) = True
      ' 取消
      Case vbKeyF10:
         Select Case m_EditMode
            Case 1, 2:
               strTit = "詢問"
               strMsg = "你並未存檔, 確定離開嗎?"
               nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit)
               If nResponse = vbYes Then
                  If m_EditMode = 1 Then '新增時又按取消鍵,預防已鍵入帳號資料,因此必須執行刪除動作
                     If DelRecord(textCW(1)) = False Then Exit Sub
                  End If
                  m_EditMode = 0
                  UpdateCtrlData
                  SetCtrlReadOnly True
                  UpdateToolbarState
               End If
            Case Else
               If m_EditMode <> 0 Then
                  m_EditMode = 0
                  UpdateCtrlData
                  SetCtrlReadOnly True
                  UpdateToolbarState
               End If
         End Select
         SSTab1.TabEnabled(2) = True
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
   
   strSql = "SELECT cw01 FROM custweb " & _
            "WHERE cw01 = (SELECT MIN(cw01) FROM custweb)"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("cw01")) = False Then: m_FirstKEY = rsTmp.Fields("cw01")
   End If
   rsTmp.Close
   
   strSql = "SELECT cw01 FROM custweb " & _
            "WHERE cw01 = (SELECT MAX(cw01) FROM custweb)"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("cw01")) = False Then: m_LastKEY = rsTmp.Fields("cw01")
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

'檢查是否為多筆客戶
Private Sub ChkIsMuchCust()
Dim rsTmp As New ADODB.Recordset
   
   If textCW(4) = "" Then Exit Sub
   strCustID = ""
   SetlstUsers 2, textCW(4) '取得客戶名稱
   '先預設為單筆客戶
   Check1.Value = 0
   Check1.Enabled = True
   Check1.Visible = False
   Text2.Visible = False
   lstUsers(2).Enabled = False
   bolMuchCust = False
   '檢查是否為多筆客戶
   If InStr(textCW(4), ",") > 0 Then
      Check1.Visible = True
   End If
   '檢查是否依不同客戶做帳號設定
   strSql = "SELECT distinct(cd02) FROM custwebid WHERE cd01 = '" & textCW(1) & "' and cd02<>cd01 order by cd02 asc"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      '多筆客戶
      Check1.Value = 1
      'Check1.Enabled = False
      bolMuchCust = True
      strCustID = Trim(rsTmp.Fields(0))
      For i = 0 To lstUsers(2).ListCount - 1
         MyArr = Split(lstUsers(2).List(i), "@")
         If MyArr(1) = strCustID Then
            lstUsers(2).Selected(i) = True
         End If
      Next i
   End If
   rsTmp.Close
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   SSTab1.Enabled = False
   strSql = "SELECT * FROM custweb " & _
            "WHERE cw01 = '" & m_CurrKEY & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      ClearField
      If IsNull(rsTmp.Fields("cw01")) = False Then: textCW(1) = rsTmp.Fields("cw01")
      If IsNull(rsTmp.Fields("cw02")) = False Then: textCW(2) = rsTmp.Fields("cw02")
      If IsNull(rsTmp.Fields("cw03")) = False Then: SelCombo cboCW03, rsTmp.Fields("cw03"), 1
      If IsNull(rsTmp.Fields("cw04")) = False Then: textCW(4) = rsTmp.Fields("cw04") '客戶編號
      If IsNull(rsTmp.Fields("cw05")) = False Then: textCW05 = rsTmp.Fields("cw05")
      If IsNull(rsTmp.Fields("cw12")) = False Then: textCW12 = rsTmp.Fields("cw12")
      If IsNull(rsTmp.Fields("cw13")) = False Then: textCW(13) = ChangeWStringToTString(rsTmp.Fields("cw13"))
      If IsNull(rsTmp.Fields("cw14")) = False Then
         MyArr = Split(rsTmp.Fields("cw14"), ",")
         For j = 0 To UBound(MyArr)
            If MyArr(j) = "1" Then Check2(0).Value = 1
            If MyArr(j) = "2" Then Check2(1).Value = 1
            If MyArr(j) = "3" Then Check2(2).Value = 1
         Next j
         textCW(14) = rsTmp.Fields("cw14")
      End If
      If IsNull(rsTmp.Fields("cw15")) = False Then
         If rsTmp.Fields("cw15") = "1" Then Option1(0).Value = True
         If rsTmp.Fields("cw15") = "2" Then Option1(1).Value = True
         textCW(15) = rsTmp.Fields("cw15")
      End If
      If IsNull(rsTmp.Fields("cw16")) = False Then: textCW(16) = rsTmp.Fields("cw16")
      If IsNull(rsTmp.Fields("cw17")) = False Then: textCW(17) = rsTmp.Fields("cw17")
      If IsNull(rsTmp.Fields("cw18")) = False Then: textCW(18) = rsTmp.Fields("cw18")
      'Add by Amy 2015/03/27 +性質
      cboCW19 = "" & rsTmp.Fields("cw19")
      SetlstUsers 1, textCW(4)
      
      ' 更新CUID
      UpdateCUID rsTmp
      
      ' 更新暫存區的資料
      UpdateFieldOldData rsTmp
      
      '查詢附件檔
'      m_upFileServer = True 'Add By Sindy 2017/5/25
      strExc(0) = "select cf02,cf03,cf08 from custwebfile where cf01=" & m_CurrKEY & " order by 1"
      intI = 1
      lstAtt.Tag = "" 'Add By Sindy 2021/5/13
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         With RsTemp
         .MoveFirst
         Do While Not .EOF
'            If "" & .Fields("cf08") = "" Then m_upFileServer = False
            lstAtt.AddItem .Fields("cf02") & " (" & Round(.Fields("cf03") / 1024, 2) & " KB)", 0
            'Modify By Sindy 2021/5/13
            'lstAtt.ItemData(0) = 1
            lstAtt.Tag = lstAtt.Tag & ";" & .Fields("cf02") & " (" & Round(.Fields("cf03") / 1024, 2) & " KB)"
            '2021/5/13 END
            .MoveNext
         Loop
         End With
         If lstAtt.Tag <> "" Then lstAtt.Tag = Mid(lstAtt.Tag, 2) 'Add By Sindy 2021/5/13
         cmdOpenAtt(0).Enabled = True
         cmdSelect(0).Enabled = True
         cmdSaveAtt(0).Enabled = True
      End If
      'If lstAtt.ListCount > 0 Then SetListScroll lstAtt
      
      ChkIsMuchCust '檢查是否為多筆客戶
      Call UpdateCtrlData2(strCustID) '查詢帳號資料
   End If
   rsTmp.Close
   Me.Enabled = True
   SSTab1.Enabled = True
   Screen.MousePointer = vbDefault
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

'查詢帳號資料
Private Sub UpdateCtrlData2(strCD02 As String)
Dim rsTmp2 As New ADODB.Recordset
Dim strSql As String
Dim strID As String, strText As String
Dim strCon As String
   
   GRD1.Clear
   GRD1.Rows = 2
   ClearCD
   If strCD02 > "" And strCD02 <> textCW(1) Then '依客戶分別設定帳號資料
      bolMuchCust = True
   Else 'If strCD02 = textCW(1) Then
      bolMuchCust = False
      m_CurrKeyCD02 = textCW(1)
      m_CurrKeyCD02_Name = textCW(1)
   End If
   Call SetGrd(bolMuchCust)
   '抓取帳號資料
   strCon = ""
   If strCD02 > "" Then
      strCon = " and cd02='" & strCD02 & "'"
   End If
   strSql = "select ' ' as V,cd02 as 客戶編號,NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) as 客戶名稱,cd03 as 帳號,cd04 as 密碼,decode(cd05,'1','1 管理者','2','2 使用者') as 身份別,cd06 as 使用者,sqldatet(cd09) as 建置日期,sqldatet(cd07) as 下次更新日期,cd08 as 註解,cd06 as CD06 from custwebid,customer where cd01='" & textCW(1) & "'" & strCon & " and substr(cd02,1,1)='X' and cu01>' ' and cd02=cu01||cu02 " & _
            "union " & _
            "select ' ' as V,cd02 as 客戶編號,NVL(FA04,NVL(FA05||FA63||FA64||FA65,FA06)) as 客戶名稱,cd03 as 帳號,cd04 as 密碼,decode(cd05,'1','1 管理者','2','2 使用者') as 身份別,cd06 as 使用者,sqldatet(cd09) as 建置日期,sqldatet(cd07) as 下次更新日期,cd08 as 註解,cd06 as CD06 from custwebid,fagent where cd01='" & textCW(1) & "'" & strCon & " and substr(cd02,1,1)='Y' and fa01>' ' and cd02=fa01||fa02 " & _
            "union " & _
            "select ' ' as V,cd02 as 客戶編號,cd02 as 客戶名稱,cd03 as 帳號,cd04 as 密碼,decode(cd05,'1','1 管理者','2','2 使用者') as 身份別,cd06 as 使用者,sqldatet(cd09) as 建置日期,sqldatet(cd07) as 下次更新日期,cd08 as 註解,cd06 as CD06 from custwebid where cd01='" & textCW(1) & "'" & strCon & " and substr(cd02,1,1)<>'X' and substr(cd02,1,1)<>'Y' " & _
            "order by 客戶編號,身份別,帳號 asc"
   rsTmp2.CursorLocation = adUseClient
   rsTmp2.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp2.RecordCount > 0 Then
      Set GRD1.Recordset = rsTmp2
      m_CurrKeyCD02 = rsTmp2.Fields(1)
      m_CurrKeyCD02_Name = rsTmp2.Fields(2)
      For i = 1 To GRD1.Rows - 1
         '轉換使用者姓名
         strID = GRD1.TextMatrix(i, 6)
         If ChangeCD06CN(strID, strText) = True Then
            GRD1.TextMatrix(i, 6) = strText
         Else
            GRD1.TextMatrix(i, 6) = strID
         End If
      Next i
      cmdOK(1).Enabled = False
      cmdOK(2).Enabled = False
   End If
   rsTmp2.Close
   
EXITSUB:
   Set rsTmp2 = Nothing
End Sub

'轉換使用者姓名
Private Function ChangeCD06CN(strID As String, ByRef strText As String) As Boolean
Dim strTempName As String
   
   ChangeCD06CN = False
   If strID <> "" Then
      MyArr = Split(strID, ",")
      strText = ""
      For j = 0 To UBound(MyArr)
         If ClsPDGetStaff(MyArr(j), strTempName) = True Then
            strText = strText & "," & strTempName
         Else
            strText = strText & "," & MyArr(j)
         End If
      Next j
      strText = Mid(strText, 2, Len(strText))
      ChangeCD06CN = True
   End If
End Function

' 更新toolbar上按紐的狀態
Private Sub UpdateToolbarState()
   Me.Enabled = False
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
   Me.Enabled = True
End Sub

Private Function CheckDataValid() As Boolean
Dim nResponse As Boolean
Dim strTit As String
Dim strMsg As String
Dim intIndex As Integer, strText As String
   
   CheckDataValid = False
   
   If IsEmptyText(textCW12) = True Then
      strTit = "檢核資料"
      strMsg = "平台名稱不可為空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      SSTab1.Tab = 0
      textCW12.SetFocus
      GoTo EXITSUB
   End If
   If IsEmptyText(cboCW03) = True Then
      strTit = "檢核資料"
      strMsg = "平台類別不可為空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      SSTab1.Tab = 0
      cboCW03.SetFocus
      GoTo EXITSUB
   End If
   If IsEmptyText(textCW(13)) = True Then
      strTit = "檢核資料"
      strMsg = "建置日期不可為空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      SSTab1.Tab = 0
      textCW(13).SetFocus
      GoTo EXITSUB
   End If
   If Left(Trim(cboCW03.Text), 1) <> "4" Then '憑證時,網址可空白
      If IsEmptyText(textCW(2)) = True Then
         strTit = "檢核資料"
         strMsg = "網址不可為空白"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         SSTab1.Tab = 0
         textCW(2).SetFocus
         GoTo EXITSUB
      End If
   End If
   If Left(Trim(cboCW03.Text), 1) <> "4" Then '憑證時,可輸入重覆的網址
      If Trim(textCW(2)) <> "" Then
         If IsRecordExist_CW02(textCW(2)) Then
            strTit = "檢核資料"
            strMsg = "此網址資料已存在!!"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            SSTab1.Tab = 0
            Call textCW_GotFocus(2)
            textCW(2).SetFocus
            Exit Function
         End If
      End If
      If Trim(textCW(17)) <> "" Then
         If IsRecordExist_CW02(textCW(17)) Then
            strTit = "檢核資料"
            strMsg = "此網址2資料已存在!!"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            SSTab1.Tab = 0
            Call textCW_GotFocus(17)
            textCW(17).SetFocus
            Exit Function
         End If
      End If
      If Trim(textCW(18)) <> "" Then
         If IsRecordExist_CW02(textCW(18)) Then
            strTit = "檢核資料"
            strMsg = "此網址3資料已存在!!"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            SSTab1.Tab = 0
            Call textCW_GotFocus(18)
            textCW(18).SetFocus
            Exit Function
         End If
      End If
   End If
   
   If Trim(textCW(2)) <> "" Or Trim(textCW(17)) <> "" Or Trim(textCW(18)) <> "" Then
      '網址必須依順序輸入
      If Trim(textCW(18)) <> "" And (Trim(textCW(2)) = "" Or Trim(textCW(17)) = "") Then
         strTit = "檢核資料"
         strMsg = "網址必須依順序輸入!!"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         SSTab1.Tab = 0
         Call textCW_GotFocus(18)
         textCW(18).SetFocus
         Exit Function
      End If
      If Trim(textCW(17)) <> "" And Trim(textCW(2)) = "" Then
         strTit = "檢核資料"
         strMsg = "網址必須依順序輸入!!"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         SSTab1.Tab = 0
         Call textCW_GotFocus(17)
         textCW(17).SetFocus
         Exit Function
      End If
      '網址不可重覆輸入
      For i = 1 To 3
         If i = 1 Then intIndex = 2
         If i = 2 Then intIndex = 17
         If i = 3 Then intIndex = 18
         strText = Trim(textCW(intIndex))
         If strText <> "" Then
            For j = i + 1 To 3
               If j = 2 Then intIndex = 17
               If j = 3 Then intIndex = 18
               If Trim(textCW(intIndex)) <> "" Then
                  If Trim(textCW(intIndex)) = strText Then
                     strTit = "檢核資料"
                     strMsg = "網址不可重覆輸入!!"
                     nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                     SSTab1.Tab = 0
                     Call textCW_GotFocus(intIndex)
                     textCW(intIndex).SetFocus
                     Exit Function
                  End If
               End If
            Next j
         End If
      Next i
   End If
   'Add by Amy 2015/03/27 +性質
   If Trim(cboCW19) = MsgText(601) Then
        strTit = "檢核資料"
        strMsg = "性質不可為空白"
        nResponse = MsgBox(strMsg, vbOKOnly, strTit)
        SSTab1.Tab = 0
        Call cboCW19_GotFocus
        cboCW19.SetFocus
        GoTo EXITSUB
   ElseIf CheckLengthIsOK(cboCW19, 10) = False Then
        SSTab1.Tab = 0
        Call cboCW19_GotFocus
        cboCW19.SetFocus
        GoTo EXITSUB
   End If
   'end 2015/03/27
   If IsEmptyText(textCW(4)) = True Then
      strTit = "檢核資料"
      strMsg = "客戶不可為空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      SSTab1.Tab = 0
      txtUserNo(1).SetFocus
      GoTo EXITSUB
   End If
   If Check2(0).Value = 0 And Check2(1).Value = 0 And Check2(2).Value = 0 Then
      strTit = "檢核資料"
      strMsg = "驗證方式至少勾選一項"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      SSTab1.Tab = 0
      GoTo EXITSUB
   End If
   If Option1(0).Value = False And Option1(1).Value = False Then
      strTit = "檢核資料"
      strMsg = "管理方式必須選擇一項"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      SSTab1.Tab = 0
      GoTo EXITSUB
   End If
   
   CheckDataValid = True
EXITSUB:
End Function

Private Function CheckDataValid_CD() As Boolean
Dim nResponse As Boolean
Dim strTit As String
Dim strMsg As String
   
   CheckDataValid_CD = False
   
   If lstUsers(2).ListCount = 0 Then
      strTit = "檢核資料"
      strMsg = "請輸入客戶!!"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      txtUserNo(1).SetFocus
      SSTab1.Tab = 0
      Exit Function
   End If
   If IsEmptyText(textCD(3)) = True Then
      strTit = "檢核資料"
      strMsg = "帳號不可為空白!!"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCD(3).SetFocus
      Exit Function
   End If
   If m_EditMode2 = 1 Then
      If IsRecordExist_CD(textCW(1), m_CurrKeyCD02, textCD(3)) Then
         strTit = "檢核資料"
         strMsg = "此帳號資料已存在!!"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         Call textCD_GotFocus(3)
         textCD(3).SetFocus
         Exit Function
      End If
   End If
   If IsEmptyText(textCD(4)) = True Then
      strTit = "檢核資料"
      strMsg = "密碼不可為空白!!"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCD(4).SetFocus
      Exit Function
   End If
   If IsEmptyText(cboCD05.Text) = True Then
      strTit = "檢核資料"
      strMsg = "身份別不可為空白!!"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      Exit Function
   End If
   If IsEmptyText(textCD(9)) = True Then
      strTit = "檢核資料"
      strMsg = "建置日期不可為空白!!"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCD(9).SetFocus
      Exit Function
   End If
   If IsEmptyText(textCD(7)) = True Then
      strTit = "檢核資料"
      strMsg = "下次更新日期不可為空白!!"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCD(7).SetFocus
      Exit Function
   End If
   
   CheckDataValid_CD = True
EXITSUB:
End Function

' 更新Key的狀態
Private Sub SetKeyReadOnly(ByVal bEnable As Boolean)
   textCW(1).Locked = bEnable
   If bEnable Then textCW(1).BackColor = &H8000000F Else textCW(1).BackColor = &H80000005
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
Dim i As Integer
   'textCW(1).Locked = bEnable
   textCW(2).Locked = bEnable
   cboCW03.Enabled = Not bEnable
   txtUserNo(1).Enabled = Not bEnable
   textCW05.Locked = bEnable
   textCW12.Locked = bEnable
   textCW(13).Locked = bEnable
   textCW(16).Locked = bEnable
   textCW(17).Locked = bEnable
   textCW(18).Locked = bEnable
   cboCW19.Enabled = Not bEnable 'Add by Amy 2015/03/27
   Frame4.Enabled = Not bEnable
   Frame5.Enabled = Not bEnable
   
   cmdOK(0).Enabled = Not bEnable
   EnableCD False
   cmdOK(1).Enabled = False
   cmdOK(2).Enabled = False
   cmdOK(3).Enabled = Not bEnable
   cmdOK(4).Enabled = Not bEnable
   GRD1.Enabled = True
End Sub

Private Sub ClearField()
Dim nIndex As Integer
   
   textCW(1) = Empty
   textCW(2) = Empty
   cboCW03.ListIndex = 0
   textCW05 = Empty
   textCW12 = Empty
   textCW(13) = strSrvDate(2)
   textCW(16) = Empty
   textCW(17) = Empty
   textCW(18) = Empty
   cboCW19 = Empty
   Label23.Caption = Empty
   
   '驗證方式
   Check2(0).Value = 0
   Check2(1).Value = 0
   Check2(2).Value = 0
   '管理方式
   Option1(0).Value = False
   Option1(1).Value = False
   
   '客戶編號
   textCW(4) = Empty
   lstUsers(1).Clear
   txtUserNo(1).Text = Empty
   lblName(1).Caption = Empty
   
   For nIndex = 0 To tf_CW - 1
      m_FieldList(nIndex).fiOldData = Empty
      m_FieldList(nIndex).fiNewData = Empty
   Next nIndex
   
   lstAtt.Clear '操作手冊
   Erase m_FilesRemoved
   ReDim m_FilesRemoved(0) As String
   cmdOpenAtt(0).Enabled = False
   cmdSelect(0).Enabled = False
   cmdSaveAtt(0).Enabled = False
   cmdAddAtt(0).Enabled = False
   cmdRemAtt(0).Enabled = False
   
   '帳號資料
   lstUsers(2).Clear
   Text2.Visible = False
   Check1.Visible = False
   GRD1.Clear
   GRD1.Rows = 2
   ClearCD
   Call SetGrd(False)
End Sub

'清除帳號資料
Sub ClearCD()
   textCD(3) = Empty
   textCD(4) = Empty
   cboCD05.ListIndex = 1
   textCD(7) = Empty
   textCD(8) = Empty
   textCD(9) = strSrvDate(2)
   '使用者
   textCD(6) = Empty
   lstUsers(0).Clear
   txtUserNo(0).Text = Empty
   lblName(0).Caption = Empty
End Sub

Private Sub UpdateFieldNewData()
   '若新增資料
   If m_EditMode = 1 Then
      SetFieldNewData "CW01", textCW(1)
   End If
   SetFieldNewData "CW02", textCW(2)
   If cboCW03.Text <> "" Then
        MyArr = Split(cboCW03, " ")
        SetFieldNewData "CW03", MyArr(0)
   Else
        SetFieldNewData "CW03", Empty
   End If
   SetFieldNewData "CW04", textCW(4)
   SetFieldNewData "CW05", textCW05
   SetFieldNewData "CW12", textCW12
   SetFieldNewData "CW13", DBDATE(textCW(13))
   '驗證方式
   textCW(14) = ""
   If Check2(0).Value = 1 Then textCW(14) = textCW(14) & ",1"
   If Check2(1).Value = 1 Then textCW(14) = textCW(14) & ",2"
   If Check2(2).Value = 1 Then textCW(14) = textCW(14) & ",3"
   If textCW(14) <> "" Then textCW(14) = Mid(textCW(14), 2, Len(textCW(14)))
   SetFieldNewData "CW14", textCW(14)
   '管理方式
   textCW(15) = ""
   If Option1(0).Value = True Then
      textCW(15) = "1"
   ElseIf Option1(1).Value = True Then
      textCW(15) = "2"
   End If
   SetFieldNewData "CW15", textCW(15)
   SetFieldNewData "CW16", textCW(16)
   SetFieldNewData "CW17", textCW(17)
   SetFieldNewData "CW18", textCW(18)
   'Add by Amy 2015/03/27
   SetFieldNewData "CW19", cboCW19
End Sub

' 初始化欄位陣列
Private Sub InitialField()
Dim nIndex As Integer
Dim strTmp As String
   
   ' 初始化欄位陣列
   For nIndex = 1 To tf_CW
      strTmp = Format(nIndex, "00")
      m_FieldList(nIndex - 1).fiName = "CW" & strTmp
      m_FieldList(nIndex - 1).fiOldData = Empty
      m_FieldList(nIndex - 1).fiNewData = Empty
      m_FieldList(nIndex - 1).fiType = 0 '文字型態
'      Select Case nIndex
'         Case 13, 23, 28, 29, 31, 32, 40, 41, 44, 45, 47, 48:
'            m_FieldList(nIndex - 1).fiType = 1 '數值型態
'      End Select
   Next nIndex
End Sub

'帶預設資料
Private Sub InitialData()
   'Removed by Morgan 2017/10/24
   'cboCW03.Clear
   'cboCW03.AddItem "1 IP管理"
   'cboCW03.AddItem "2 檔案存取"
   'cboCW03.AddItem "3 電子帳單"
   'cboCW03.AddItem "4 憑證"
   PUB_SetCW03Combo cboCW03
   'end 2017/10/24
   
   cboCD05.Clear
   cboCD05.AddItem "1 管理員"
   cboCD05.AddItem "2 使用者"
   'Add by Amy 2015/03/27 +性質
   cboCW19.Clear
   cboCW19.AddItem ""
   cboCW19.AddItem "案件"
   cboCW19.AddItem "財務"
   cboCW19.AddItem "投標"
   cboCW19.AddItem "憑證"
   'end 2015/03/27
   Call SetGrd(False)
End Sub

Private Sub SetGrd(bolMuchCust As Boolean)
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer
   
   arrGridHeadText = Array("V", "客戶編號", "客戶名稱", "帳號", "密碼", "身份別", "使用者", "建置日期", "下次更新日期", "註解", "CD06")
   If bolMuchCust = True Then
      arrGridHeadWidth = Array(0, 1000, 1000, 1000, 1000, 800, 1000, 1000, 1200, 2000, 0)
   Else
      arrGridHeadWidth = Array(0, 0, 0, 1000, 1000, 800, 1000, 1000, 1200, 2000, 0)
   End If
   GRD1.Visible = False
   GRD1.Cols = UBound(arrGridHeadText) + 1
   For iRow = 0 To GRD1.Cols - 1
      GRD1.row = 0
      GRD1.col = iRow
      GRD1.Text = arrGridHeadText(iRow)
      GRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      'grd1.CellAlignment = flexAlignCenterCenter
   Next
   GRD1.Visible = True
End Sub

Private Sub EnableCD(ByVal bEnable As Boolean)
'   If m_EditMode2 = 2 Then
'      textCD(3).Locked = bEnable
'      textCD(3).BackColor = &H8000000F
'   Else
      textCD(3).Locked = Not bEnable
      textCD(3).BackColor = &H80000005
'   End If
   textCD(4).Locked = Not bEnable
   cboCD05.Locked = Not bEnable
   txtUserNo(0).Locked = Not bEnable
   textCD(7).Locked = Not bEnable
   textCD(8).Locked = Not bEnable
End Sub

'選取選單
Private Sub SelCombo(ByRef pCBO As Object, ByVal pValue As String, Optional pLen As Integer = 2)
Dim idx As Integer
   
   If pValue = "" Then
      pCBO.ListIndex = 1
   Else
      For idx = 0 To pCBO.ListCount - 1
         If Left(pCBO.List(idx), pLen) = pValue Then
            pCBO.ListIndex = idx
            Exit For
         End If
      Next
   End If
End Sub

Private Sub Check1_Click()
   If Check1.Value = 0 Then
      Text2.Visible = False
      lstUsers(2).Enabled = False
      If lstUsers(2).ListIndex < 0 Then
         Call UpdateCtrlData2(textCW(1)) '查詢共同帳號資料
      Else
         lstUsers(2).Selected(lstUsers(2).ListIndex) = False
      End If
   Else '依客戶分別設定帳號資料
      Text2.Visible = True
      lstUsers(2).Enabled = True
      If lstUsers(2).ListIndex < 0 Then
         lstUsers(2).Selected(0) = True
      Else
         lstUsers(2).Selected(lstUsers(2).ListIndex) = True
      End If
   End If
End Sub

'進入網站
Private Sub OpenIE(strWebAddr As String)
   Dim myweb As Object
   Dim hLocalFile As Long
   
   'Modify By Sindy 2021/8/25 薛:平台資料0041,請修改為由chrome 開始。（網站已不支援ＩＥ）
   '  Tymetrix360 指定用Chrome開啟--經理
   If textCW(1) = "0026" Or textCW(1) = "0041" Then
      PUB_OpenURL strWebAddr, 1
   Else
      'Modify By Sindy 2020/12/29
      ShellExecute hLocalFile, "open", strWebAddr, vbNullString, vbNullString, 1
   End If
   Exit Sub
   '2020/12/29 END
   
   Set myweb = CreateObject("InternetExplorer.Application")
   Screen.MousePointer = vbHourglass
   With myweb
      .ToolBar = 0
      .Visible = True ' 顯示IE
      '.Navigate "http://" & textCW(2) ' 瀏覽網址 www.lativ.com.tw/Home/Login
      .Navigate strWebAddr ' 瀏覽網址 www.lativ.com.tw/Home/Login
'      ' 等待網頁載入完成
'      Do While .Busy
'         DoEvents
'      Loop
      '.Document.All("email").Value = "xxxx" '帳號 login
      '.Document.All("pw").Value = "xxxx"    '密碼 passwd
      '.Document.All("submit").Click         '登入 signIn
   End With
   Set myweb = Nothing ' 釋放IE 物件
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmdQuery_Click(Index As Integer)
Dim rsTmp As New ADODB.Recordset
Dim strID As String, strText As String
   
   Select Case Index
      Case 0
         If txt1(0) = "" And _
            txt1(1) = "" And _
            txtName = "" And _
            (txt1(3) = "" Or txt1(4) = "") Then
            MsgBox "請輸入查詢條件!!"
            txt1(0).SetFocus
            Exit Sub
         End If
         
         Screen.MousePointer = vbHourglass
         MSHFlexGrid1.Clear
         MSHFlexGrid1.Rows = 2
         GridHead1
         strSql = ""
         If txt1(0) <> "" Then '客戶編號
             strSql = strSql & " and instr(cw04,'" & txt1(0) & "')>0"
         End If
         If txt1(1) <> "" Then '使用者
             strSql = strSql & " and instr(cd06,'" & txt1(1) & "')>0"
         End If
         If txtName <> "" Then '平台名稱
             strSql = strSql & " and instr(upper(cw12),'" & UCase(txtName) & "')>0"
         End If
         If txt1(3) <> "" And txt1(4) <> "" Then '下次更新日期
             strSql = strSql & " and cd07>=" & DBDATE(txt1(3)) & " and cd07<=" & DBDATE(txt1(4))
         End If
         '抓取資料
         strSql = "SELECT distinct cw01 as 平台編號,cw04 as 客戶編號,' ' as 客戶名稱,cw12 as 平台名稱," & PUB_GetCW03SQL & " as 平台類別,cw02 as 網址" & _
                  " FROM custweb,custwebid" & _
                  " where cw01=cd01(+)" & strSql & _
                  " order by cw01 asc"
         If rsTmp.State = 1 Then rsTmp.Close
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            Set MSHFlexGrid1.Recordset = rsTmp
            For i = 1 To MSHFlexGrid1.Rows - 1
               '轉換客戶名稱
               strID = MSHFlexGrid1.TextMatrix(i, 1)
               If strID <> "" Then
                  MyArr = Split(strID, ",")
                  strText = ""
                  For j = 0 To UBound(MyArr)
                     If Left(MyArr(j), 1) = "X" Then
                        strText = strText & "," & GetPrjPeople1(CStr(MyArr(j)), "1")
                     ElseIf Left(MyArr(j), 1) = "Y" Then
                        strText = strText & "," & GetPrjName1(CStr(MyArr(j)))
                     End If
                  Next j
                  strText = Mid(strText, 2, Len(strText))
                  MSHFlexGrid1.TextMatrix(i, 2) = strText
               End If
               rsTmp.MoveNext
            Next i
         Else
            MsgBox "無查詢資料!!"
         End If
   End Select
   Screen.MousePointer = vbDefault
End Sub

Private Sub GridHead1()
   With MSHFlexGrid1
      .Visible = False
      .Cols = 6
      .row = 0
      .col = 0: .ColWidth(0) = 600: .Text = "平台編號"
      .CellAlignment = flexAlignCenterCenter
      .ColAlignment(0) = flexAlignCenterCenter
      
      .col = 1: .ColWidth(1) = 1000: .Text = "客戶編號"
      .CellAlignment = flexAlignCenterCenter
      .ColAlignment(1) = flexAlignCenterCenter
      
      .col = 2: .ColWidth(2) = 2500: .Text = "客戶名稱"
      .CellAlignment = flexAlignCenterCenter
      .ColAlignment(2) = flexAlignCenterCenter
      
      .col = 3: .ColWidth(3) = 2000: .Text = "平台名稱"
      .CellAlignment = flexAlignCenterCenter
      .ColAlignment(3) = flexAlignCenterCenter
      
      .col = 4: .ColWidth(4) = 800: .Text = "平台類別"
      .CellAlignment = flexAlignCenterCenter
      .ColAlignment(4) = flexAlignCenterCenter
      
      .col = 5: .ColWidth(5) = 3500: .Text = "網址1"
      .CellAlignment = flexAlignCenterCenter
      .ColAlignment(5) = flexAlignCenterCenter
      
      .Visible = True
   End With
End Sub

'Add By Sindy 2021/5/13
Private Sub textCW05_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = 2 Then Forms(0).PopupMenu2 textCW05
End Sub
Private Sub textCW12_GotFocus()
   InverseTextBox textCW12
   CloseIme
End Sub
Private Sub textCW05_GotFocus()
   InverseTextBox textCW05
   OpenIme
End Sub
Private Sub textCW05_KeyPress(KeyAscii As MSForms.ReturnInteger)
   '擋掉備註欄位的Enter動作
   If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub textCW_DblClick(Index As Integer)
   If Trim(textCW(Index)) <> "" Then
      'Added by Morgan 2021/4/15 Tymetrix360 指定用Chrome開啟--經理
      If textCW(1) = "0026" Then
         PUB_OpenURL Trim(textCW(Index)), 1
      Else
      'end 2021/4/15
         Call OpenIE(Trim(textCW(Index)))
      End If
   End If
End Sub

Private Sub textCW_GotFocus(Index As Integer)
   InverseTextBox textCW(Index)
   If Index = 5 Then
      OpenIme
   Else
      CloseIme
   End If
End Sub

Private Sub textCW_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
'      Case 0, 1
'         KeyAscii = UpperCase(KeyAscii)
      Case 13, 16
         KeyAscii = Pub_NumAscii(KeyAscii)
'      Case 5
'         '擋掉備註欄位的Enter動作
'         If KeyAscii = 13 Then KeyAscii = 0
   End Select
End Sub

Private Sub textCW_Validate(Index As Integer, Cancel As Boolean)
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   Select Case Index
      Case 2, 17, 18 '網址
         If Trim(textCW(Index)) <> "" Then
            If InStr(UCase(textCW(Index)), "HTTP://") = 0 And InStr(UCase(textCW(Index)), "HTTPS://") = 0 Then
               Call textCW_GotFocus(Index)
               Cancel = True
               MsgBox "請輸入完整的網址含HTTP:..."
               SSTab1.Tab = 0
               textCW(Index).SetFocus
               Exit Sub
            End If
         End If
      Case 13 '日期欄位
         If Trim(textCW(Index)) <> "" Then
            If CheckIsTaiwanDate(textCW(Index), False) = False Then
               Call textCW_GotFocus(Index)
               Cancel = True
               MsgBox "請輸入民國日期不含/！", vbInformation, "輸入日期錯誤"
               textCW(Index).SetFocus
               Exit Sub
            End If
            If Index = 13 Then
               If DBDATE(textCW(Index)) > strSrvDate(1) Then
                  Call textCW_GotFocus(Index)
                  Cancel = True
                  MsgBox "建置日期不可大於系統日！"
                  textCW(Index).SetFocus
                  Exit Sub
               End If
            End If
         End If
   End Select
End Sub

Private Sub textCD_GotFocus(Index As Integer)
   InverseTextBox textCD(Index)
   If Index = 8 Then
      OpenIme
   Else
      CloseIme
   End If
End Sub

Private Sub textCD_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   Select Case Index
'      Case 0, 1
'         KeyAscii = UpperCase(KeyAscii)
      Case 7, 9
         KeyAscii = Pub_NumAscii(KeyAscii)
   End Select
End Sub

Private Sub textCD_Validate(Index As Integer, Cancel As Boolean)
   If m_EditMode2 <> 1 And m_EditMode2 <> 2 Then Exit Sub
   If Trim(textCD(Index)) = "" Then Exit Sub
   
   Select Case Index
      Case 3 '帳號
         If m_EditMode2 = 2 Then '修改時,檢查帳號不可重覆輸入
            For i = 1 To GRD1.Rows - 1
               If iCurrSelRow <> i Then
                  If Trim(textCD(Index)) = GRD1.TextMatrix(i, 3) Then
                     Call textCD_GotFocus(Index)
                     Cancel = True
                     MsgBox "此帳號重覆，請重新輸入！"
                     textCD(Index).SetFocus
                     Exit Sub
                  End If
               End If
            Next i
         End If
      Case 7, 9 '日期欄位
         If CheckIsTaiwanDate(textCD(Index), False) = False Then
            Call textCD_GotFocus(Index)
            Cancel = True
            MsgBox "請輸入民國日期不含/！", vbInformation, "輸入日期錯誤"
            textCD(Index).SetFocus
            Exit Sub
         End If
         If Index = 7 Then
            'Modify By Sindy 2015/12/21
            'If DBDATE(textCD(Index)) <= strSrvDate(1) Then
            If DBDATE(textCD(Index)) <= strSrvDate(1) And textCD(Index) <> "111111" Then
            '2015/12/21 END
               Call textCD_GotFocus(Index)
               Cancel = True
               MsgBox "下次更新日期必須大於系統日！"
               textCD(Index).SetFocus
               Exit Sub
            End If
         End If
         If Index = 9 Then
            If DBDATE(textCD(Index)) > strSrvDate(1) Then
               Call textCD_GotFocus(Index)
               Cancel = True
               MsgBox "建置日期不可大於系統日！"
               textCD(Index).SetFocus
               Exit Sub
            End If
         End If
   End Select
End Sub

Private Sub textCW12_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = 2 Then Forms(0).PopupMenu2 textCW12
End Sub

Private Sub txtName_GotFocus()
   InverseTextBox txtName
   CloseIme
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   InverseTextBox txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0, 1
         KeyAscii = UpperCase(KeyAscii)
      Case 3, 4
         KeyAscii = Pub_NumAscii(KeyAscii)
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
Dim strTempName As String
   
   If Trim(txt1(Index)) = "" Then Exit Sub
   
   Select Case Index
      Case 0
         If Left(txt1(Index), 1) <> "X" And Left(txt1(Index), 1) <> "Y" Then
            MsgBox "客戶編號第1碼只能輸入X或Y！", vbExclamation
            Call txt1_GotFocus(Index)
            txt1(Index).SetFocus
            Cancel = True
            Exit Sub
         End If
         If Trim(txt1(Index)) > "" Then txt1(Index) = Left(Trim(txt1(Index)) & "000000000", 9)
         If Len(txt1(Index)) = 9 Then
            If Left(txt1(Index), 1) = "X" Then
               Lbl1(Index) = GetPrjPeople1(txt1(Index), "1")
            ElseIf Left(txt1(Index), 1) = "Y" Then
               Lbl1(Index) = GetPrjName1(txt1(Index))
            End If
         Else
            Lbl1(Index) = ""
         End If
         If Lbl1(Index) = "" Then
            MsgBox "客戶編號輸入錯誤！", vbExclamation
            Call txt1_GotFocus(Index)
            txt1(Index).SetFocus
            Cancel = True
            Exit Sub
         End If
      Case 1
         If Len(txt1(Index)) = 5 Then
            If ClsPDGetStaff(txt1(Index), strTempName) = True Then
               Lbl1(Index) = strTempName
            End If
         Else
            Lbl1(Index) = ""
         End If
         If Lbl1(Index) = "" Then
            MsgBox "使用者輸入錯誤！", vbExclamation
            Call txt1_GotFocus(Index)
            txt1(Index).SetFocus
            Cancel = True
            Exit Sub
         End If
      Case 3, 4
         If CheckIsTaiwanDate(txt1(Index), False) = False And Trim(txt1(Index)) <> "" Then
            Call txt1_GotFocus(Index)
            Cancel = True
            MsgBox "請輸入民國日期不含/！", vbInformation, "輸入日期錯誤"
            Exit Sub
         End If
         If Index = 3 Then
            If txt1(Index) <> "" And txt1(Index + 1) = "" Then
               txt1(Index + 1) = txt1(Index)
            End If
         ElseIf Index = 4 Then
            If RunNick2(txt1(Index - 1), txt1(Index)) Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
   End Select
End Sub

Private Sub txtName_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = 2 Then Forms(0).PopupMenu2 txtName
End Sub

Private Sub txtUserNo_GotFocus(Index As Integer)
   TextInverse txtUserNo(Index)
End Sub

Private Sub txtUserNo_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtUserNo_Validate(Index As Integer, Cancel As Boolean)
   Dim strTempName As String
   If txtUserNo(Index).Visible = True Then
      If txtUserNo(Index) <> "" Then 'And lblName(Index) = ""
         Select Case Index
            Case 0 '員工編號
               If Len(txtUserNo(Index)) = 5 Then
                  If ClsPDGetStaff(txtUserNo(Index), strTempName) = True Then
                     lblName(Index) = strTempName
                  End If
               Else
                  lblName(Index) = ""
               End If
               If lblName(Index) = "" Then
                  MsgBox "員工編號輸入錯誤！", vbExclamation
                  Call txtUserNo_GotFocus(Index)
                  txtUserNo(Index).SetFocus
                  Cancel = True
                  Exit Sub
               End If
            Case 1 '客戶編號
               If Left(txtUserNo(Index), 1) <> "X" And Left(txtUserNo(Index), 1) <> "Y" Then
                  MsgBox "客戶編號第1碼只能輸入X或Y！", vbExclamation
                  Call txtUserNo_GotFocus(Index)
                  txtUserNo(Index).SetFocus
                  Cancel = True
                  Exit Sub
               End If
               If Trim(txtUserNo(Index)) > "" Then txtUserNo(Index) = Left(Trim(txtUserNo(Index)) & "000000000", 9)
               If Len(txtUserNo(Index)) = 9 Then
                  If Left(txtUserNo(Index), 1) = "X" Then
                     lblName(Index) = GetPrjPeople1(txtUserNo(Index), "1")
                  ElseIf Left(txtUserNo(Index), 1) = "Y" Then
                     lblName(Index) = GetPrjName1(txtUserNo(Index))
                  End If
               Else
                  lblName(Index) = ""
               End If
               If lblName(Index) = "" Then
                  MsgBox "客戶編號輸入錯誤！", vbExclamation
                  Call txtUserNo_GotFocus(Index)
                  txtUserNo(Index).SetFocus
                  Cancel = True
                  Exit Sub
               End If
         End Select
      End If
   End If
End Sub

'加入員工編號或客戶編號
Private Sub cmdAdd_Click(Index As Integer)
   If Index = 1 Then
      If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   Else
      If m_EditMode2 <> 1 And m_EditMode2 <> 2 Then Exit Sub
   End If
   AddlstUsers Index
End Sub

'移除員工編號或客戶編號
Private Sub cmdRemove_Click(Index As Integer)
   If Index = 1 Then
      If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   Else
      If m_EditMode2 <> 1 And m_EditMode2 <> 2 Then Exit Sub
   End If
   RemovelstUsers Index
End Sub

Private Sub AddlstUsers(p_idx As Integer)
   Dim idx As Integer, bFound As Boolean
   If txtUserNo(p_idx) <> "" And lblName(p_idx) <> "" Then
      '非數字需做轉換
      'For idx = 0 To lstUsers(p_idx).ListCount - 1
         'If lstUsers(p_idx).ItemData(idx) = PUB_Id2Num(txtUserNo(p_idx), CStr(p_idx)) Then
         If InStr(lstUsers(p_idx).Tag, txtUserNo(p_idx)) > 0 Then
            Select Case p_idx
               Case 0 '員工編號
                  MsgBox "員工已存在於使用者清單中！"
               Case 1 '客戶編號
                  MsgBox "客戶已存在於客戶清單中！"
            End Select
            txtUserNo(p_idx).SetFocus
            txtUserNo_GotFocus p_idx
            bFound = True
            'Exit For
         End If
      'Next
      If bFound = False Then
         'lstUsers(p_idx).AddItem lblName(p_idx), 0
         lstUsers(p_idx).AddItem lblName(p_idx) & "                                                            @" & txtUserNo(p_idx), 0
         'Modify By Sindy 2021/5/12
         'lstUsers(p_idx).ItemData(0) = PUB_Id2Num(txtUserNo(p_idx), CStr(p_idx))
         lstUsers(p_idx).Tag = txtUserNo(p_idx) & IIf(lstUsers(p_idx).Tag <> "", "," & lstUsers(p_idx).Tag, "")
         Select Case p_idx
            Case 0 '員工編號
               If Trim(textCD(6)) <> "" Then textCD(6) = Trim(textCD(6)) & ","
               textCD(6) = Trim(textCD(6)) & txtUserNo(p_idx) 'PUB_Num2Id(lstUsers(p_idx).ItemData(0), 0)
            Case 1 '客戶編號
               If Trim(textCW(4)) <> "" Then textCW(4) = Trim(textCW(4)) & ","
               textCW(4) = Trim(textCW(4)) & txtUserNo(p_idx) 'PUB_Num2Id(lstUsers(p_idx).ItemData(0), 1)
         End Select
         '2021/5/12 END
         txtUserNo(p_idx) = ""
         lblName(p_idx) = ""
         txtUserNo(p_idx).SetFocus
      End If
   End If
End Sub

Private Function ComposeListX(p_index As Integer) As String
Dim varTmp As Variant
Dim bolFind As Boolean

   strExc(1) = ""
   'Modify By Sindy 2021/5/12
'   If lstUsers(p_index).ListCount > 0 Then
'      strExc(1) = PUB_Num2Id(lstUsers(p_index).ItemData(0), CStr(p_index))
'      For intI = 1 To lstUsers(p_index).ListCount - 1
'         strExc(1) = strExc(1) & "," & PUB_Num2Id(lstUsers(p_index).ItemData(intI), CStr(p_index))
'      Next
'   End If
   varTmp = Split(lstUsers(p_index).Tag, ",")
   For intI = LBound(varTmp) To UBound(varTmp)
      bolFind = False
      For ii = 0 To lstUsers(p_index).ListCount - 1
         If InStr(lstUsers(p_index).List(ii), varTmp(intI)) > 0 Then
            bolFind = True
            Exit For
         End If
      Next ii
      If bolFind = True Then
         strExc(1) = strExc(1) & "," & varTmp(intI)
      End If
   Next intI
   If strExc(1) <> "" Then
      If Left(strExc(1), 1) = "," Then strExc(1) = Mid(strExc(1), 2)
      lstUsers(p_index).Tag = strExc(1)
   End If
   '2021/5/12 END
   ComposeListX = strExc(1)
End Function

Private Sub SetlstUsers(p_idx As Integer, p_stNums As String)
   Dim arrID
   
   lstUsers(p_idx).Clear
   lstUsers(p_idx).Tag = "" 'Add By Sindy 2021/5/12 原本放在ItemData,改放在Tag
   If p_stNums <> "" Then
      Select Case p_idx
         Case 0 '員工編號
            strExc(0) = "select st01,st02 from staff where instr('" & p_stNums & "',st01)>0"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               arrID = Split(p_stNums, ",")
               With RsTemp
               '照原順序排
               For intI = UBound(arrID) To LBound(arrID) Step -1
                  .MoveFirst
                  Do While Not .EOF
                     If .Fields("st01") = arrID(intI) Then
                        'lstUsers(p_idx).AddItem "" & .Fields(1), 0
                        lstUsers(p_idx).AddItem "" & .Fields(1) & "                                                            @" & .Fields(0), 0
                        'Modify By Sindy 2021/5/12
                        'lstUsers(p_idx).ItemData(0) = PUB_Id2Num(.Fields(0), CStr(p_idx))
                        lstUsers(p_idx).Tag = .Fields(0) & IIf(lstUsers(p_idx).Tag <> "", "," & lstUsers(p_idx).Tag, "")
                        '2021/5/12 END
                        .MoveLast
                     End If
                     .MoveNext
                  Loop
               Next
               End With
            End If
         Case 1, 2 '客戶編號
            strExc(0) = "select cu01||cu02,NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) from customer where cu01>' ' and instr('" & p_stNums & "',cu01||cu02)>0" & _
                        " union" & _
                        " select fa01||fa02,NVL(FA04,NVL(FA05||FA63||FA64||FA65,FA06)) from fagent where fa01>' ' and instr('" & p_stNums & "',fa01||fa02)>0"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               arrID = Split(p_stNums, ",")
               With RsTemp
               '照原順序排
               For intI = UBound(arrID) To LBound(arrID) Step -1
                  .MoveFirst
                  Do While Not .EOF
                     If .Fields(0) = arrID(intI) Then
                        lstUsers(p_idx).AddItem "" & .Fields(1) & "                                                            @" & .Fields(0), 0
                        'Modify By Sindy 2021/5/12
                        'lstUsers(p_idx).ItemData(0) = PUB_Id2Num(.Fields(0), CStr(p_idx))
                        lstUsers(p_idx).Tag = .Fields(0) & IIf(lstUsers(p_idx).Tag <> "", "," & lstUsers(p_idx).Tag, "")
                        '2021/5/12 END
                        .MoveLast
                     End If
                     .MoveNext
                  Loop
               Next
               End With
            End If
      End Select
   End If
End Sub

Private Sub RemovelstUsers(p_idx As Integer)
   Dim idx As Integer
   
   If lstUsers(p_idx).ListCount > 0 Then
      For idx = lstUsers(p_idx).ListCount - 1 To 0 Step -1
         If lstUsers(p_idx).Selected(idx) = True Then
            If p_idx = 1 Then
               '若要移除客戶時,須檢查是否有該客戶的帳號資料存在
               If p_idx = 1 Then
                  If IsRecordExist_CD02(textCW(1), txtUserNo(p_idx).Text) = True Then
                     MsgBox "此客戶尚有帳號資料，不可移除！（若要移除請先刪除該客戶的帳號資料）"
                     Exit Sub
                  End If
               End If
            End If
            
            lstUsers(p_idx).RemoveItem idx
            Exit For 'Add By Sindy 2023/9/6 僅供一筆一筆移除,若點在最後一筆又是Run迴圈會全部刪光
         End If
      Next
   End If
   Select Case p_idx
      Case 0 '員工編號
         textCD(6) = ComposeListX(p_idx)
      Case 1 '客戶編號
         textCW(4) = ComposeListX(p_idx)
   End Select
   txtUserNo(p_idx).Text = ""
   txtUserNo(p_idx).SetFocus
End Sub

'Removed by Morgan 2016/9/8 客戶/代理人末三碼會有B,C,非數字會有錯,改呼叫公用函數
''員工或客戶編號轉數字
'Public Function PUB_Id2Num(pID As String, strType As String) As Long
'   Select Case strType
'      Case 0 '員工編號
'         PUB_Id2Num = "&H" & pID
'      Case 1 '客戶編號
'         If Left(Trim(pID), 1) = "X" Then
'            PUB_Id2Num = "1" & Mid(Trim(pID), 2, Len(Trim(pID)) - 1)
'         ElseIf Left(Trim(pID), 1) = "Y" Then
'            PUB_Id2Num = "2" & Mid(Trim(pID), 2, Len(Trim(pID)) - 1)
'         End If
'   End Select
'End Function
'
''數字轉員工或客戶編號
'Public Function PUB_Num2Id(pNum As Long, strType As String) As String
'   Select Case strType
'      Case 0 '員工編號
'         PUB_Num2Id = Hex(pNum)
'      Case 1 '客戶編號
'         If Left(Trim(pNum), 1) = "1" Then
'            PUB_Num2Id = "X" & Mid(Trim(pNum), 2, Len(Trim(pNum)) - 1)
'         ElseIf Left(Trim(pNum), 1) = "2" Then
'            PUB_Num2Id = "Y" & Mid(Trim(pNum), 2, Len(Trim(pNum)) - 1)
'         End If
'         PUB_Num2Id = Left(PUB_Num2Id & "000000000", 9)
'   End Select
'End Function
'end 2016/9/8

'自動給號
Private Function GetNewCW01() As String
Dim rsTmp As New ADODB.Recordset
   
   GetNewCW01 = ""
   Screen.MousePointer = vbHourglass
   '檢查是否資料
   strSql = "SELECT count(cw01) FROM custweb"
   If rsTmp.State = 1 Then rsTmp.Close
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.Fields(0) > 0 Then
      rsTmp.Clone
      '讀取最大編號數
      strSql = "SELECT max(cw01) FROM custweb"
      If rsTmp.State = 1 Then rsTmp.Close
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount >= 0 Then
         GetNewCW01 = Val(rsTmp.Fields(0)) + 1
      End If
   Else
      GetNewCW01 = 1
   End If
   rsTmp.Clone
   Set rsTmp = Nothing
   GetNewCW01 = Format(GetNewCW01, "0000")
   Screen.MousePointer = vbDefault
End Function

Private Function GetAttachFile(ByRef pFileName As String, Optional pSavePath As String) As Boolean
   Dim stAttPath As String
   Dim lngSize As Long
   Dim iFileNo As Integer
   Dim bytes() As Byte
   
On Error GoTo ErrHnd
   
   If pSavePath = "" Then
      If Dir(m_AttachPath, vbDirectory) = "" Then
         MkDir m_AttachPath
      End If
      stAttPath = m_AttachPath & "\" & pFileName
      '檔案已存在時不必重新下載
      If Dir(stAttPath) <> "" Then
         'Kill stAttPath
         pFileName = stAttPath
         GetAttachFile = True
         Exit Function
      End If
   Else
      stAttPath = pSavePath
   End If
      
   strExc(0) = "select * from custwebfile b where cf01=" & textCW(1) & " and cf02='" & ChgSQL(pFileName) & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If Dir(stAttPath) <> "" Then Kill stAttPath
      
      'Add By Sindy 2017/5/25
      If "" & RsTemp.Fields("cf08") <> "" Then
         GetAttachFile = PUB_GetFtpFile(RsTemp.Fields("cf08"), stAttPath, UCase("custwebfile"))
      Else
      '2017/5/25 END
         With RsTemp
         lngSize = Val(.Fields("cf03").Value)
         ReDim bytes(lngSize)
         If lngSize > 0 Then bytes() = .Fields("cf04").GetChunk(lngSize)
         End With
         iFileNo = FreeFile
         Open stAttPath For Binary Access Write As #iFileNo
         If lngSize > 0 Then Put #iFileNo, , bytes()
         Close #iFileNo
      End If
      
      pFileName = stAttPath
      GetAttachFile = True
   End If
   
   Exit Function
   
ErrHnd:
   MsgBox Err.Description, vbCritical
   If iFileNo > 0 Then Close #iFileNo
End Function

'開啟附件
Private Sub cmdOpenAtt_Click(Index As Integer)
   Dim hLocalFile As Long
   Dim stFileName As String
   Dim strAtt As String, strType As String
   
   Screen.MousePointer = vbHourglass
   
   If Index = 0 Then
      strAtt = lstAtt.Text
   End If
   
   If strAtt = "" Then
      MsgBox "請選擇欲開啟的附件！"
   Else
      For ii = 0 To lstAtt.ListCount - 1
         If lstAtt.Selected(ii) Then
            stFileName = lstAtt.List(ii)
            'stFileName = strAtt
            If InStrRev(stFileName, " (") > 0 Then
               stFileName = Left(stFileName, InStrRev(stFileName, " (") - 1)
            End If
            
            If InStr(stFileName, "\") = 0 Then
               If GetAttachFile(stFileName) = False Then
                  Exit Sub
               End If
            End If
            
            ShellExecute hLocalFile, "open", stFileName, vbNullString, vbNullString, 1
         End If
      Next ii
   End If
   
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSelect_Click(Index As Integer)
   Dim ii As Integer, oList As Object
   If Index = 0 Then
      Set oList = lstAtt
   End If
   
   For ii = 0 To oList.ListCount - 1
      lstAtt.Selected(ii) = True
   Next
End Sub

Private Sub cmdSaveAtt_Click(Index As Integer)
   
   Dim stFileName As String, stFolderPath As String, stFullName As String
   Dim bMultiFile As Boolean
   Dim ii As Integer, oList As Object
   
   Screen.MousePointer = vbHourglass
   
   If Index = 0 Then
      Set oList = lstAtt
   End If
   
   stFileName = ""
   bMultiFile = False
   For ii = 0 To oList.ListCount - 1
      If oList.Selected(ii) Then
         If stFileName <> "" Then
            bMultiFile = True
            Exit For
         Else
            stFileName = oList.Text
         End If
      End If
   Next
   
   If stFileName = "" Then
      MsgBox "請選擇欲存檔的附件！"
   Else
      '多選
      If bMultiFile Then
         stFolderPath = BrowseForFolder()
         If stFolderPath <> "" Then
            For ii = 0 To oList.ListCount - 1
               If oList.Selected(ii) Then
                  stFileName = oList.List(ii)
                  If InStrRev(stFileName, " (") > 0 Then
                     stFileName = Left(stFileName, InStrRev(stFileName, " (") - 1)
                  End If
                  stFullName = stFolderPath & stFileName
                  If stFullName <> "" Then
                     If Dir(stFullName) <> "" Then
                        If MsgBox("檔案[ " & stFileName & " ]已存在是否要覆蓋??", vbYesNo + vbDefaultButton2) = vbNo Then
                           stFullName = ""
                        End If
                     End If
                     If stFullName <> "" Then
                        If GetAttachFile(stFileName, stFullName) = False Then
                           MsgBox "無法儲存檔案[ " & stFileName & " ]！"
                        End If
                     End If
                  End If
               End If
            Next
         End If
      
      Else
         stFileName = Left(stFileName, InStrRev(stFileName, " (") - 1)
         stFullName = GetSaveName(stFileName)
         If stFullName <> "" Then
            If Dir(stFullName) <> "" Then
               If MsgBox("檔案[ " & stFileName & " ]已存在是否要覆蓋??", vbYesNo + vbDefaultButton2) = vbNo Then
                  stFullName = ""
               End If
            End If
            If stFullName <> "" Then
               If GetAttachFile(stFileName, stFullName) = False Then
                  MsgBox "無法儲存檔案[ " & stFileName & " ]！"
               End If
            End If
         End If
      End If
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmdAddAtt_Click(Index As Integer)
   Dim stFileName As String
   Dim sFile
   Dim ii As Integer
   Dim fs, f, s

On Error GoTo ErrHnd
   
   stFileName = "*.*"
   With CommonDialog1
      .CancelError = True
      .FileName = stFileName
      .Filter = "All Files (*.*)|*.*"
      'Modify by Amy 2014/07/14 第一次DB可能沒資料所以先抓Client
      If PUB_GetLastDate(Me.Name, "Dir") <> "" Then
         .InitDir = PUB_GetLastDate(Me.Name, "Dir")
      ElseIf GetSetting("TAIE", "FCP", Me.Name & "Dir", "") <> "" Then
         .InitDir = GetSetting("TAIE", "FCP", Me.Name & "Dir", "")
      Else
         .InitDir = PUB_Getdesktop
      End If
      .MaxFileSize = 3000
      .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
      .ShowOpen
      If .FileName <> "" Then
         If InStr(.FileName, ChrW$(0)) > 0 Then
            sFile = Split(.FileName, ChrW$(0))
            '記錄路徑
            'Modify by Amy 2014/07/14
            'SaveSetting "TAIE", "FCP", Me.Name & "Dir", sFile(0)
            PUB_SaveLastDate Me.Name, "Dir", CStr(sFile(0))
            For ii = 1 To UBound(sFile)
               If InStr(sFile(ii), "\") > 0 Then
                  stFileName = sFile(ii)
               Else
                  stFileName = sFile(0) & "\" & sFile(ii)
               End If
               Set fs = CreateObject("Scripting.FileSystemObject")
               Set f = fs.GetFile(stFileName)
               
               If Index = 0 Then
                  AddListX lstAtt, stFileName & " (" & Round(f.Size / 1024, 2) & " KB)", lstAtt
               End If
            Next
         Else
            '記錄路徑
            If InStr(.FileName, "\") > 0 Then
               For ii = Len(.FileName) To 1 Step -1
                  If Mid(Trim(.FileName), ii, 1) = "\" Then
                     'Modify by Amy 2014/07/14
                     'SaveSetting "TAIE", "FCP", Me.Name & "Dir", Mid(Trim(.FileName), 1, ii - 1)
                     PUB_SaveLastDate Me.Name, "Dir", Mid(Trim(.FileName), 1, ii - 1)
                     Exit For
                  End If
               Next ii
            End If
            stFileName = .FileName
            Set fs = CreateObject("Scripting.FileSystemObject")
            Set f = fs.GetFile(stFileName)
            If Index = 0 Then
               AddListX lstAtt, stFileName & " (" & Round(f.Size / 1024, 2) & " KB)", lstAtt
            End If
         End If
      End If
   End With
   Exit Sub
ErrHnd:
   If Err.Number <> 32755 Then
      MsgBox Err.Description
   End If
End Sub

Private Sub cmdRemAtt_Click(Index As Integer)
   If Index = 0 Then
      RemoveList lstAtt
   End If
End Sub

Private Function GetSaveName(ByVal pFileName As String) As String
   
On Error GoTo ErrHnd

   With CommonDialog1
      .CancelError = True
      .FileName = pFileName
      .Filter = "All Files (*.*)|*.*"
      .InitDir = PUB_Getdesktop
      .MaxFileSize = 3000
      .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
      .ShowSave
      If .FileName <> "" Then
         GetSaveName = .FileName
      End If
   End With
   
   Exit Function
   
ErrHnd:
   If Err.Number <> 32755 Then
      MsgBox Err.Description
   End If
End Function

Private Function RemoveList(oList As Object) As Boolean
   Dim ii As Integer
   If oList.ListCount > 0 Then
      ii = 0
      Do While ii < oList.ListCount
         If oList.Selected(ii) = True Then
         
            If oList.List(ii) <> "" Then
               'Add By Sindy 2017/5/25
'               If m_upFileServer = True Then
                  If MsgBox("確定要永久刪除" & oList.List(ii) & "電子檔？", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbNo Then
                     Screen.MousePointer = vbDefault
                     Exit Function
                  End If
                  '直接從資料庫刪除檔案
                  If PUB_DelFtpFile2(textCW(1), " and cf02='" & GetFileName(oList.List(ii)) & "'", UCase("custwebfile")) = True Then '檔案改放FTP,必須在DB資料刪除前執行
                     strSql = "delete from custwebfile where cf01='" & textCW(1) & "' and cf02='" & GetFileName(oList.List(ii)) & "'"
                     cnnConnection.Execute strSql
                  End If
'               Else
'               '2017/5/25 END
'                  intI = UBound(m_FilesRemoved) + 1
'                  ReDim Preserve m_FilesRemoved(intI) As String
'                  m_FilesRemoved(intI) = GetFileName(oList.List(ii))
'               End If
            End If
            
            oList.RemoveItem ii
            'SetListScroll oList
            RemoveList = True
            ii = ii - 1
         End If
         ii = ii + 1
      Loop
   End If
End Function

Private Function GetFileName(ByVal FullPath As String) As String
   Dim stItem As String, iPos As Integer
   
   stItem = FullPath
   iPos = InStr(stItem, "\")
   Do While iPos > 0
      stItem = Mid(stItem, iPos + 1)
      iPos = InStr(stItem, "\")
   Loop
   
   If InStrRev(stItem, " (") > 0 And Right(stItem, 1) = ")" Then
      stItem = Left(stItem, InStrRev(stItem, " (") - 1)
   End If
   
   GetFileName = stItem
End Function

Private Sub SetListScroll(oList As Object)
   Dim ii As Integer
   Dim lWnow As Long, lWmax As Long
   
   lWmax = 0
   For ii = 0 To oList.ListCount - 1
      lWnow = TextWidth(oList.List(ii) & " ")
      If lWnow > lWmax Then
         lWmax = lWnow
      End If
   Next
  
   If ScaleMode = vbTwips Then lWmax = lWmax / Screen.TwipsPerPixelX  ' if twips change to pixels
   SendMessageByNum oList.hWnd, LB_SETHORIZONTALEXTENT, lWmax, 0
End Sub

Private Function AddListX(oList As Object, stNewItem As String, oList1 As Object) As Boolean
   Dim idx As Integer, bFound As Boolean, stFileName As String
'   If InStr(stNewItem, ",") > 0 Then
'      MsgBox "逗號[,]為系統保留字，請重新命名！", vbExclamation
'      cmdAddAtt.SetFocus
'      Exit Function
'   End If
   If stNewItem <> "" Then
      For idx = 0 To oList.ListCount - 1
         stFileName = GetFileName(oList.List(idx))
         If GetFileName(stNewItem) = stFileName Then
            MsgBox "附件 " & stFileName & " 已存在！"
            AddListX = False
            bFound = True
            Exit For
         End If
      Next
      
      If bFound = False Then
         For idx = 0 To oList1.ListCount - 1
            stFileName = GetFileName(oList1.List(idx))
            If GetFileName(stNewItem) = stFileName Then
               MsgBox "附件 " & stFileName & " 已存在！"
               AddListX = False
               bFound = True
               Exit For
            End If
         Next
      End If
      
      If bFound = False Then
         oList.AddItem stNewItem, 0
         'SetListScroll oList
         AddListX = True
      End If
   End If
End Function
