VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm210153 
   BorderStyle     =   1  '單線固定
   Caption         =   "網頁提供國內專利公報資訊"
   ClientHeight    =   5460
   ClientLeft      =   6096
   ClientTop       =   1548
   ClientWidth     =   9132
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   9132
   Begin VB.CommandButton cmdNote 
      Caption         =   "注意事項"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   8.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7860
      TabIndex        =   39
      Top             =   630
      Width           =   1035
   End
   Begin VB.CheckBox Check1 
      Caption         =   "使用智慧局新版檢索系統"
      Height          =   225
      Left            =   5490
      TabIndex        =   38
      Top             =   660
      Width           =   2355
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7440
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
            Picture         =   "frm210153.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210153.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210153.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210153.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210153.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210153.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210153.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210153.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210153.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210153.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210153.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4545
      Left            =   120
      TabIndex        =   2
      Top             =   870
      Width           =   8865
      _ExtentX        =   15642
      _ExtentY        =   8022
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "基本資料"
      TabPicture(0)   =   "frm210153.frx":20F4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(2)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(3)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(5)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(6)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(7)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(8)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txt2PS(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txt2PS(2)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txt2PS(3)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txt2PS(11)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txt2Path"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txt2ImgFolder"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txt2MailTo"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lbl2Dates(1)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lbl2Dates(2)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "lbl2MailTo"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "MSHFlexGrid1"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "cmdPath"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Command1(0)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Command1(1)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Command2(1)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Command2(2)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Command2(3)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Frame1"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).ControlCount=   25
      TabCaption(1)   =   "多筆查詢"
      TabPicture(1)   =   "frm210153.frx":2110
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdQuery"
      Tab(1).Control(1)=   "MSHFlexGrid2"
      Tab(1).Control(2)=   "lbl2Find"
      Tab(1).Control(3)=   "txt2Find(1)"
      Tab(1).Control(4)=   "txt2Find(0)"
      Tab(1).Control(5)=   "txt2Find(2)"
      Tab(1).Control(6)=   "Label1(9)"
      Tab(1).Control(7)=   "Line2"
      Tab(1).Control(8)=   "Label1(4)"
      Tab(1).ControlCount=   9
      Begin VB.Frame Frame1 
         Height          =   495
         Left            =   4500
         TabIndex        =   34
         Top             =   -60
         Visible         =   0   'False
         Width           =   4305
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   255
            Left            =   1140
            TabIndex        =   36
            Top             =   150
            Width           =   3045
            _ExtentX        =   5376
            _ExtentY        =   445
            _Version        =   393216
            Appearance      =   0
         End
         Begin VB.Label lblProgress 
            AutoSize        =   -1  'True
            Caption         =   "簡圖下載中"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   120
            TabIndex        =   35
            Top             =   180
            Width           =   975
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "EMail連結"
         Height          =   315
         Index           =   3
         Left            =   7620
         TabIndex        =   16
         Top             =   2468
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "開啟連結"
         Height          =   315
         Index           =   2
         Left            =   7620
         TabIndex        =   13
         Top             =   2123
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "上傳網站"
         Height          =   315
         Index           =   1
         Left            =   7620
         TabIndex        =   12
         Top             =   1778
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "本機預覽"
         Height          =   315
         Index           =   1
         Left            =   7620
         TabIndex        =   11
         Top             =   1448
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "匯入(&I)"
         Height          =   315
         Index           =   0
         Left            =   7620
         TabIndex        =   7
         Top             =   1103
         Width           =   975
      End
      Begin VB.CommandButton cmdPath 
         Height          =   315
         Left            =   7230
         Picture         =   "frm210153.frx":212C
         Style           =   1  '圖片外觀
         TabIndex        =   5
         Top             =   1103
         Width           =   330
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "查詢(&Q)"
         Default         =   -1  'True
         Height          =   400
         Left            =   -68190
         TabIndex        =   0
         Top             =   390
         Width           =   912
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   3540
         Left            =   -74850
         TabIndex        =   18
         Top             =   930
         Width           =   8415
         _ExtentX        =   14838
         _ExtentY        =   6244
         _Version        =   393216
         Cols            =   5
         HighLight       =   0
         AllowUserResizing=   1
         FormatString    =   "流水號|檢索內容|建立日期|建立人員|網站連結"
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1620
         Left            =   150
         TabIndex        =   19
         Top             =   2850
         Width           =   8415
         _ExtentX        =   14838
         _ExtentY        =   2858
         _Version        =   393216
         Cols            =   8
         HighLight       =   0
         AllowUserResizing=   1
         FormatString    =   "#|專利編號|專利名稱|公告/公開日|申請號|申請人|簡圖檔名|FTP路徑"
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
         _Band(0).Cols   =   8
      End
      Begin MSForms.Label lbl2Find 
         Height          =   180
         Left            =   -69660
         TabIndex        =   33
         Top             =   540
         Width           =   1005
         Caption         =   "XXX"
         Size            =   "1773;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt2Find 
         Height          =   330
         Index           =   1
         Left            =   -72780
         TabIndex        =   32
         Top             =   465
         Width           =   975
         VariousPropertyBits=   679495707
         Size            =   "1720;582"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt2Find 
         Height          =   330
         Index           =   0
         Left            =   -73980
         TabIndex        =   31
         Top             =   465
         Width           =   975
         VariousPropertyBits=   679495707
         Size            =   "1720;582"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt2Find 
         Height          =   330
         Index           =   2
         Left            =   -70740
         TabIndex        =   30
         Top             =   465
         Width           =   975
         VariousPropertyBits=   679495707
         Size            =   "1720;582"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl2MailTo 
         Height          =   180
         Left            =   6480
         TabIndex        =   29
         Top             =   2550
         Width           =   1005
         Caption         =   "XXX"
         Size            =   "1773;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl2Dates 
         Height          =   180
         Index           =   2
         Left            =   1230
         TabIndex        =   28
         Top             =   2550
         Width           =   2865
         Caption         =   "上傳：XXX"
         Size            =   "5054;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl2Dates 
         Height          =   180
         Index           =   1
         Left            =   2370
         TabIndex        =   27
         Top             =   480
         Width           =   6195
         Caption         =   "CREATE：XXX   UPDATE：XXX"
         Size            =   "10927;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt2MailTo 
         Height          =   330
         Left            =   5370
         TabIndex        =   26
         Top             =   2490
         Width           =   975
         VariousPropertyBits=   679495707
         Size            =   "1720;582"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt2ImgFolder 
         Height          =   330
         Left            =   1380
         TabIndex        =   25
         Top             =   1440
         Width           =   6195
         VariousPropertyBits=   679495711
         Size            =   "10927;582"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt2Path 
         Height          =   330
         Left            =   1560
         TabIndex        =   24
         Top             =   1095
         Width           =   5685
         VariousPropertyBits=   679495711
         Size            =   "10028;582"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt2PS 
         Height          =   330
         Index           =   11
         Left            =   1200
         TabIndex        =   23
         Top             =   2130
         Width           =   6375
         VariousPropertyBits=   679495711
         Size            =   "11245;582"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt2PS 
         Height          =   330
         Index           =   3
         Left            =   1200
         TabIndex        =   22
         Top             =   1800
         Width           =   6375
         VariousPropertyBits=   679495711
         Size            =   "11245;582"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt2PS 
         Height          =   330
         Index           =   2
         Left            =   1200
         TabIndex        =   21
         Top             =   750
         Width           =   6375
         VariousPropertyBits=   679495707
         Size            =   "11245;582"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt2PS 
         Height          =   330
         Index           =   1
         Left            =   1200
         TabIndex        =   20
         Top             =   405
         Width           =   1065
         VariousPropertyBits=   679495707
         Size            =   "1879;582"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "收件者："
         BeginProperty Font 
            Name            =   "新細明體-ExtB"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   8
         Left            =   4590
         TabIndex        =   17
         Top             =   2535
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "簡圖資料夾："
         Height          =   180
         Index           =   7
         Left            =   210
         TabIndex        =   15
         Top             =   1515
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "網頁檔名："
         Height          =   180
         Index           =   6
         Left            =   210
         TabIndex        =   14
         Top             =   1845
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "網站連結："
         Height          =   180
         Index           =   5
         Left            =   210
         TabIndex        =   10
         Top             =   2190
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "檢索內容："
         Height          =   180
         Index           =   3
         Left            =   210
         TabIndex        =   9
         Top             =   810
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "流水號："
         Height          =   180
         Index           =   2
         Left            =   210
         TabIndex        =   8
         Top             =   480
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "待匯入資料夾："
         Height          =   180
         Index           =   1
         Left            =   210
         TabIndex        =   6
         Top             =   1170
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "建立人員:"
         Height          =   180
         Index           =   9
         Left            =   -71580
         TabIndex        =   4
         Top             =   540
         Width           =   765
      End
      Begin VB.Line Line2 
         X1              =   -73260
         X2              =   -72840
         Y1              =   623
         Y2              =   623
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "建立日期:"
         Height          =   180
         Index           =   4
         Left            =   -74850
         TabIndex        =   3
         Top             =   540
         Width           =   765
      End
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   576
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9132
      _ExtentX        =   16108
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "網頁公報連結至智慧局網站，僅於公告/公開日起半年內有效!!"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   120
      TabIndex        =   37
      Top             =   660
      Width           =   5235
   End
End
Attribute VB_Name = "frm210153"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/14 改成Form2.0 (原本就是)
'Created by Morgan 2019/5/13
'***考慮檢索內容可能含Unicode字元，本程式使用 OraOLEDB 獨立連線並搭配 Form2.0 元件***
Option Explicit

'前次紀錄KEY
Dim lst_KEY As String
'目前狀態 1=新增, 2=修改, 4=查詢, 0=瀏覽, 9=無資料
Dim iCurState As Integer
'使用者權限設定
Dim bolInsert As Boolean
Dim bolUpdate As Boolean
Dim bolDelete As Boolean
Dim bolSelect As Boolean
Dim m_strTempFolder As String, m_strPS04 As String
Dim m_bolImport As Boolean
Dim m_strKey As String
Dim iPrevRow1 As Integer, iPrevRow2 As Integer
Dim m_Connection As New ADODB.Connection

'連接資料庫
Private Function ConnectToServer(pConnection As ADODB.Connection) As Boolean
On Error GoTo ErrHand
   ConnectToServer = False
   If pConnection.State = adStateOpen Then pConnection.Close
   pConnection.ConnectionTimeout = 60
   pConnection.Provider = cOraProvider
   pConnection.Properties("Data Source").Value = IIf(strServerName <> "", strServerName, ServerName)
   pConnection.Properties("User ID").Value = UserName
   pConnection.Properties("Password").Value = Password
   pConnection.Open
   ConnectToServer = True
   Exit Function
   
ErrHand:
   MsgBox Err.Description
End Function

Private Sub ImportXLS()
   Dim strXLS As String, strImgFolder As String, strPageName As String
   
   If ChkPath() = True Then
      strXLS = Dir(txt2Path & "\*.xls")
      If strXLS = "" Then
         MsgBox "找不到 XLS 檔！", vbCritical
      Else
         strImgFolder = Dir(txt2Path & "\*_files", vbDirectory)
         If strImgFolder = "" Then
            MsgBox "找不到 簡圖檔 資料夾！", vbCritical
         Else
            txt2ImgFolder = txt2Path & "\" & strImgFolder
            If LoadXLS(strXLS, strImgFolder) = False Then
               MsgBox "XLS 匯入失敗！", vbCritical
            Else
               If txt2PS(3) = "" Then
                  strPageName = "TWG" & strSrvDate(2) & ".HTML"
               Else
                  strPageName = txt2PS(3)
               End If
               
               If MakeHTML(txt2ImgFolder & "\" & strPageName) = False Then
                  MsgBox "網頁建立失敗！", vbCritical
               Else
                  txt2PS(3) = strPageName
                  m_bolImport = True
                  
                  '刪除本機舊網頁
                  If iCurState = 2 Then
                     If Dir(m_strTempFolder & "\" & txt2PS(3)) <> "" Then
                        Kill m_strTempFolder & "\" & txt2PS(3)
                     End If
                  End If
                  
               End If
            End If
         End If
      End If
   End If
End Sub

Private Sub Check1_Click()
   If Check1.Value = vbChecked Then
      cmdNote.Enabled = True
   Else
      cmdNote.Enabled = False
   End If
End Sub

Private Sub cmdNote_Click()
   'Modify by Amy 2024/05/21 +公報網址
   MsgBox "1.若有多個網頁(圖檔)要存檔時，請用[ 不同檔名 ]，不可覆蓋！" & vbCrLf & _
                   "2.[ 申請案號 ]為對應簡圖的必要欄位，請務必保留！" & vbCrLf & vbCrLf & _
                   "[註]網路公報網址(PDF下載)：https://cloud.tipo.gov.tw/S220/gazette/patent", vbInformation, "新版檢所系統注意事項"
End Sub

Private Sub cmdPath_Click()
   Dim fName As String, strStartFolder As String
   
   If Dir(txt2Path & "\", vbDirectory) <> "" Then strStartFolder = txt2Path
   
   fName = PUB_GetFolder(Me.hWnd, strStartFolder, "請選取資料夾:")
   If fName <> "" Then 'they did not hit cancel
      txt2Path = fName
      SaveSetting "TAIE", "P", UCase(Me.Name) & "Dir", txt2Path
   End If
End Sub

Private Sub cmdQuery_Click()
   Dim stCon As String
   
   If SSTab1.Tab = 0 Then Exit Sub
   
   stCon = ""
   If txt2Find(0) <> "" Then
      stCon = stCon & " and to_char(PS06,'yyyymmdd')>='" & DBDATE(txt2Find(0)) & "'"
   End If
   
   If txt2Find(1) <> "" Then
      stCon = stCon & " and to_char(PS06,'yyyymmdd')<='" & DBDATE(txt2Find(1)) & "'"
   End If
   
   If txt2Find(2) <> "" Then
      stCon = stCon & " and PS05='" & txt2Find(2) & "'"
   End If
   
   SetGrid2 True
   strSql = "select PS01,PS02,sqldatet(to_char(PS06,'yyyymmdd')) Date1" & _
      ",st02,PS03" & _
      " from PatentSearch,Staff where 1=1" & stCon & _
      " and st01(+)=PS05 order by 1"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If RsTemp.RecordCount > 0 Then
      Set MSHFlexGrid2.Recordset = RsTemp.Clone
      SetGrid2
   End If
   
End Sub

Private Sub Command1_Click(Index As Integer)
   SetMouseBusy
   Select Case Index
      Case 0 '匯入
         ImportXLS
         
      Case 1 '本機預覽
         OpenLocal
   End Select
   SetMouseReady
End Sub

Private Sub Command2_Click(Index As Integer)
   SetMouseBusy
   Select Case Index
      Case 1 '上傳網站
         If Upload2WWW() = True Then
            doQuery 4
         End If
         
      Case 2 '開啟連結
         OpenURL
         
      Case 3 'EMail連結
         MailURL
   End Select
   SetMouseReady
End Sub

Private Sub Form_Activate()
   Static bolActivted As Boolean
   If bolActivted = False Then
      bolActivted = True
      If m_Connection.State <> adStateOpen Then
         Unload Me
      End If
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   
   '建立 OraOLEDB 連線
   If ConnectToServer(m_Connection) = True Then
      Me.Caption = Me.Caption & GetDbTerminal
   Else
      Exit Sub
   End If
   
   SetPath
   setAuthority
   FormReset
   SSTab1.Tab = 0
   '預設為瀏覽最後一筆
   If doQuery(9) = True Then
      iCurState = 0
   Else
      iCurState = 9
   End If
   Call SetToolBar(iCurState)
   Call SetInputs(iCurState)
   
   SetTempFolder
   KillTemp
   
   txt2MailTo = strUserNum
   lbl2MailTo = strUserName
   lbl2Find = ""
   
   'Added by Morgan 2021/8/24
   If strSrvDate(1) > "20210930" Then
      Check1.Value = vbChecked
      Check1.Enabled = False
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set m_Connection = Nothing
   Set frm210153 = Nothing
End Sub

Private Sub MSHFlexGrid1_Click()
   Dim nRow As Integer
   nRow = MSHFlexGrid1.MouseRow
   If nRow > 0 Then
      SelectRow nRow, MSHFlexGrid1, iPrevRow1
      iPrevRow1 = nRow
   End If
End Sub

Private Sub MSHFlexGrid2_Click()
   Dim nRow As Integer
   nRow = MSHFlexGrid2.MouseRow
   If nRow > 0 Then
      SelectRow nRow, MSHFlexGrid2, iPrevRow2
      iPrevRow2 = nRow
   End If
End Sub

Private Sub MSHFlexGrid2_DblClick()
   If iPrevRow2 > 0 And MSHFlexGrid2.row = iPrevRow2 Then
      SSTab1.Tab = 0
      doAction 4
      txt2PS(1) = MSHFlexGrid2.TextMatrix(iPrevRow2, 0)
      doAction 11
   End If
End Sub

Private Sub SelectRow(ByRef pRow As Integer, ByRef FlexGrid As MSHFlexGrid, ByRef pPrevRow As Integer)
   Dim nCol As Integer, iCol As Integer
   With FlexGrid
   nCol = .col
   If pPrevRow > 0 Then
      If pPrevRow <> pRow Then
         .row = pPrevRow
         If .FixedCols > 0 Then
            .col = .FixedCols - 1
            .CellBackColor = .BackColorFixed
            .CellForeColor = .ForeColor
         End If
         For iCol = .FixedCols To .Cols - 1
            .col = iCol
            .CellBackColor = .BackColor
         Next
      End If
   End If

   If pRow > 0 Then
      .row = pRow
      If .FixedCols > 0 Then
         .col = .FixedCols - 1
         .CellBackColor = .BackColorSel
         .CellForeColor = .ForeColorSel
      End If
      For iCol = .FixedCols To .Cols - 1
        .col = iCol
        .CellBackColor = &HFFC0C0
      Next
   End If
   .col = nCol
   .Refresh
   pPrevRow = pRow
   End With
End Sub

Private Sub doAction(pIndex As Integer)
   SSTab1.Tab = 0
   Select Case pIndex
      Case 1 '新增
         iCurState = 1
         FormReset
         SSTab1.TabEnabled(1) = False
         txt2PS(2).SetFocus
         
      Case 2 '修改
         iCurState = 2
         SSTab1.TabEnabled(1) = False
         txt2PS(2).SetFocus
         
      Case 3 '刪除
         SSTab1.Tab = 0
         If MsgBox("是否要刪除此筆資料?", vbCritical + vbYesNo + vbDefaultButton2, "詢問") = vbYes Then
            SetMouseBusy
            If DeleteData = True Then
               FormReset
               If doQuery(8, False) = True Then
                  iCurState = 0
               ElseIf doQuery(9) = True Then
                  iCurState = 0
               Else
                  lst_KEY = ""
                  iCurState = 9
               End If
            End If
            SetMouseReady
         End If
         
      Case 4 '查詢
         iCurState = 4
         SSTab1.TabEnabled(1) = False
         FormReset
         txt2PS(1).SetFocus
      Case 6 '第一筆
         Call doQuery(6)
         
      Case 7 '上一筆
         Call doQuery(7)
         
      Case 8 '下一筆
         Call doQuery(8)
         
      Case 9 '最後筆
         Call doQuery(9)
         
      Case 11 '確定
         If CheckConfirm() = False Then Exit Sub
         
         SetMouseBusy
         Select Case iCurState
            Case 1 '新增
               If txt2PS(3) <> "" Then
                  If txt2PS(3).Tag <> txt2PS(2) Then
                     If MakeHTML(txt2ImgFolder & "\" & txt2PS(3)) = False Then
                        MsgBox "網頁建立失敗！", vbCritical
                        SetMouseReady
                        Exit Sub
                     End If
                  End If
               End If
            
               If insertdata() = False Then
                  SetMouseReady
                  Exit Sub
               End If
               
            Case 2 '修改
               '若檢索內容有改時要檢查是否網頁也有更新
               If txt2PS(3) <> "" Then
                  If txt2PS(3).Tag <> txt2PS(2) Then
                     If MakeHTML(txt2ImgFolder & "\" & txt2PS(3)) = False Then
                        MsgBox "網頁建立失敗！", vbCritical
                        SetMouseReady
                        Exit Sub
                     End If
                  End If
               End If
   
               If UpdateData() = False Then
                  SetMouseReady
                  Exit Sub
                  
               '若檢索內容有改時要提醒重新上傳網站
               ElseIf txt2PS(11) <> "" Then
                  If m_bolImport = True Then
                     Call Upload2WWW
                  ElseIf txt2PS(11).Tag <> txt2PS(2) Then
                     Call Upload2WWW(True)
                  End If
               End If
               
            Case 4 '查詢
               
         End Select
         SetMouseReady
         
         '重新查詢
         If doQuery(4) = True Then
            iCurState = 0
            
         ElseIf iCurState = 4 Then
            Exit Sub
         End If
                  
      Case 12 '取消
         '新增/修改
         If iCurState = 1 Or iCurState = 2 Then
            If MsgBox("你並未存檔，確定要取消?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
               Exit Sub
            End If
         End If
         
         If lst_KEY <> "" Then
            txt2PS(1) = lst_KEY
            If doQuery(4) = True Then
               iCurState = 0
            ElseIf doQuery(9) = True Then
               iCurState = 0
            Else
               lst_KEY = ""
               iCurState = 9
            End If
         Else
            iCurState = 9
         End If
         
      Case 14 '結束
         '新增/修改
         If iCurState = 1 Or iCurState = 2 Then
            If MsgBox("你並未存檔，是否確定要離開?", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
               Unload Me
               Exit Sub
            End If
         Else
            Unload Me
            Exit Sub
         End If
         
   End Select
   
   If iCurState = 0 Or iCurState = 9 Then SSTab1.TabEnabled(1) = True
   
   Call SetToolBar(iCurState)
   Call SetInputs(iCurState)
End Sub

Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   doAction Button.Index
End Sub

'讀取主檔資料
Private Function doQuery(ByVal iAct As Integer, Optional ByVal bolMsg As Boolean = True) As Boolean

   Dim stMessage As String
   
   Select Case iAct
      Case 4
      '查詢
         strSql = "Select a.*,s1.st02 name1,sqldatet(to_char(ps06,'yyyymmdd'))||to_char(ps06,' hh24:mi') date1" & _
            ",s2.st02 name2,sqldatet(to_char(ps08,'yyyymmdd'))||to_char(ps08,' hh24:mi') date2" & _
            ",s3.st02 name3,sqldatet(to_char(ps10,'yyyymmdd'))||to_char(ps10,' hh24:mi') date3" & _
            " From PatentSearch a,staff s1,staff s2,staff s3" & _
            " where PS01=" & Val(txt2PS(1)) & _
            " and s1.st01(+)=ps05 and s2.st01(+)=ps07 and s3.st01(+)=ps09"
         stMessage = "查無資料！"
   
      Case 6
      '第一筆
         strSql = "Select a.*,s1.st02 name1,sqldatet(to_char(ps06,'yyyymmdd'))||to_char(ps06,' hh24:mi') date1" & _
            ",s2.st02 name2,sqldatet(to_char(ps08,'yyyymmdd'))||to_char(ps08,' hh24:mi') date2" & _
            ",s3.st02 name3,sqldatet(to_char(ps10,'yyyymmdd'))||to_char(ps10,' hh24:mi') date3" & _
            " From PatentSearch a,staff s1,staff s2,staff s3" & _
            " where ps01=(select min(b.ps01) from PatentSearch b)" & _
            " and s1.st01(+)=ps05 and s2.st01(+)=ps07 and s3.st01(+)=ps09"
         stMessage = "無資料！"
      Case 7
      '上一筆
         strSql = "Select a.*,s1.st02 name1,sqldatet(to_char(ps06,'yyyymmdd'))||to_char(ps06,' hh24:mi') date1" & _
            ",s2.st02 name2,sqldatet(to_char(ps08,'yyyymmdd'))||to_char(ps08,' hh24:mi') date2" & _
            ",s3.st02 name3,sqldatet(to_char(ps10,'yyyymmdd'))||to_char(ps10,' hh24:mi') date3" & _
            " From PatentSearch a,staff s1,staff s2,staff s3" & _
            " where ps01=(select max(b.ps01) from PatentSearch b where b.ps01<" & lst_KEY & ")" & _
            " and s1.st01(+)=ps05 and s2.st01(+)=ps07 and s3.st01(+)=ps09"
         stMessage = "已是第一筆了！"

      Case 8
      '下一筆
         strSql = "Select a.*,s1.st02 name1,sqldatet(to_char(ps06,'yyyymmdd'))||to_char(ps06,' hh24:mi') date1" & _
            ",s2.st02 name2,sqldatet(to_char(ps08,'yyyymmdd'))||to_char(ps08,' hh24:mi') date2" & _
            ",s3.st02 name3,sqldatet(to_char(ps10,'yyyymmdd'))||to_char(ps10,' hh24:mi') date3" & _
            " From PatentSearch a,staff s1,staff s2,staff s3" & _
            " where ps01=(select min(b.ps01) from PatentSearch b where b.ps01>" & lst_KEY & ")" & _
            " and s1.st01(+)=ps05 and s2.st01(+)=ps07 and s3.st01(+)=ps09"
         stMessage = "已是最後一筆了！"

      Case 9
      '最後筆
         strSql = "Select a.*,s1.st02 name1,sqldatet(to_char(ps06,'yyyymmdd'))||to_char(ps06,' hh24:mi') date1" & _
            ",s2.st02 name2,sqldatet(to_char(ps08,'yyyymmdd'))||to_char(ps08,' hh24:mi') date2" & _
            ",s3.st02 name3,sqldatet(to_char(ps10,'yyyymmdd'))||to_char(ps10,' hh24:mi') date3" & _
            " From PatentSearch a,staff s1,staff s2,staff s3" & _
            " where ps01=(select max(b.ps01) from PatentSearch b)" & _
            " and s1.st01(+)=ps05 and s2.st01(+)=ps07 and s3.st01(+)=ps09"
         stMessage = "無資料！"
        
   End Select
   
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      
      txt2PS(1) = RsTemp("PS01")
      txt2PS(2) = "" & RsTemp("PS02")
      txt2PS(3) = "" & RsTemp("PS03")
      m_strPS04 = "" & RsTemp("PS04")
      txt2PS(11) = "" & RsTemp("PS11")
      '記錄原檢索內容(判斷網頁及連結是否要更新)
      txt2PS(2).Tag = txt2PS(2)
      txt2PS(3).Tag = txt2PS(2)
      txt2PS(11).Tag = txt2PS(2)
      
      lbl2Dates(1) = "CREATE：" & RsTemp("name1") & "  " & RsTemp("date1") & "   UPDATE："
      If Not IsNull(RsTemp("name2")) Then
         lbl2Dates(1) = lbl2Dates(1) & RsTemp("name2") & "  " & RsTemp("date2")
      End If
      
      lbl2Dates(2) = "上傳："
      If Not IsNull(RsTemp("name3")) Then
         lbl2Dates(2) = lbl2Dates(2) & RsTemp("name3") & "  " & RsTemp("date3")
      End If
      
      lst_KEY = txt2PS(1)
      
      If Pub_StrUserSt03 = "M51" Or strUserNum = RsTemp("PS05") Then
         bolUpdate = True
      Else
         bolUpdate = False
      End If
      
      If QueryDetail() = True Then doQuery = True
   ElseIf intI = 0 Then
      If bolMsg Then
         MsgBox stMessage, vbExclamation
      End If
   End If
   
End Function

'讀取明細資料
Private Function QueryDetail() As Boolean
   m_bolImport = False
   SetGrid True
   'Modify by Amy 2024/05/07 PSD05 顯示/
   strSql = ",SubStr(PSD05,1,4)||Decode(SubStr(PSD05,5,2),null,'','/'||SubStr(PSD05,5,2))||Decode(SubStr(PSD05,7,2),null,'','/'||SubStr(PSD05,7,2)) as PSD05"
   strSql = "select PSD02,PSD03,PSD04" & strSql & ",PSD06,PSD07,PSD08,PSD09 from PatentSearchDetail where PSD01=" & Val(txt2PS(1)) & " order by 1"
   'end 2025/05/07
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   '若無資料時不可設定，否則可能會造成 MouseRow 與實際點選位置不同
   If RsTemp.RecordCount > 0 Then
      Set MSHFlexGrid1.Recordset = RsTemp.Clone
      SetGrid
   End If
   QueryDetail = True
   
End Function

'工具列控制
Private Sub SetToolBar(Optional ByVal iStatus As Integer)

   Dim i As Integer
   For i = 1 To 13
      TBar1.Buttons(i).Enabled = False
   Next
   TBar1.Buttons(14).Enabled = True
   
   Select Case iStatus
      Case 0 '瀏覽
         If bolInsert Then
            TBar1.Buttons(1).Enabled = True
         End If
         If bolUpdate Then
            TBar1.Buttons(2).Enabled = True
         End If
         If bolDelete Then
            TBar1.Buttons(3).Enabled = True
         End If
         If bolSelect Then
            TBar1.Buttons(4).Enabled = True
         End If
         TBar1.Buttons(6).Enabled = True
         TBar1.Buttons(7).Enabled = True
         TBar1.Buttons(8).Enabled = True
         TBar1.Buttons(9).Enabled = True
         
      Case 1, 2, 4 '1:新增  '2:修改  '4查詢
         TBar1.Buttons(11).Enabled = True
         TBar1.Buttons(12).Enabled = True
               
      Case 9 '無資料
         If bolInsert Then
            TBar1.Buttons(1).Enabled = True
         End If
   End Select
   
End Sub

'使用者權限設定
Private Sub setAuthority()
   'Modified by Morgan 2019/6/17
   '新增/瀏覽不必限制,修改只能是建立人,刪除看權限
   'bolSelect = IsUserHasRightOfFunction(Me.Name, strFind, False)
   'bolInsert = IsUserHasRightOfFunction(Me.Name, strAdd, False)
   'bolUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
   bolSelect = True
   bolInsert = True
   bolUpdate = False
   'end 2019/6/17
   bolDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyF2
      '新增
         If SSTab1.Tab = 0 And TBar1.Buttons(1).Enabled = True Then
            Call Tbar1_ButtonClick(TBar1.Buttons(1))
         End If
      Case vbKeyF3
      '修改
         If SSTab1.Tab = 0 And TBar1.Buttons(2).Enabled = True Then
            Call Tbar1_ButtonClick(TBar1.Buttons(2))
         End If
      Case vbKeyF5
      '刪除
         If SSTab1.Tab = 0 And TBar1.Buttons(3).Enabled = True Then
            Call Tbar1_ButtonClick(TBar1.Buttons(3))
         End If
      Case vbKeyF4
      '查詢
         If SSTab1.Tab = 0 And TBar1.Buttons(4).Enabled = True Then
            Call Tbar1_ButtonClick(TBar1.Buttons(4))
         End If
      Case vbKeyHome
      '第一筆
         If SSTab1.Tab = 0 And TBar1.Buttons(6).Enabled = True Then
            Call Tbar1_ButtonClick(TBar1.Buttons(6))
         End If

      Case vbKeyPageUp
      '上一筆
         If SSTab1.Tab = 0 And TBar1.Buttons(7).Enabled = True Then
            Call Tbar1_ButtonClick(TBar1.Buttons(7))
         End If
      Case vbKeyPageDown
      '下一筆
         If SSTab1.Tab = 0 And TBar1.Buttons(8).Enabled = True Then
            Call Tbar1_ButtonClick(TBar1.Buttons(8))
         End If
      Case vbKeyEnd
      '最後筆
         If SSTab1.Tab = 0 And TBar1.Buttons(9).Enabled = True Then
            Call Tbar1_ButtonClick(TBar1.Buttons(9))
         End If
      Case vbKeyF9
      '存檔
         If SSTab1.Tab = 0 And TBar1.Buttons(11).Enabled = True Then
            Call Tbar1_ButtonClick(TBar1.Buttons(11))
         End If
      Case vbKeyReturn
      '確定
         If SSTab1.Tab = 1 Then
            'Call cmdQuery_Click(0)
         ElseIf SSTab1.Tab = 0 And TBar1.Buttons(11).Enabled = True Then
            Call Tbar1_ButtonClick(TBar1.Buttons(11))
         End If
      Case vbKeyF10
      '取消
         If SSTab1.Tab = 0 And TBar1.Buttons(12).Enabled = True Then
            Call Tbar1_ButtonClick(TBar1.Buttons(12))
         End If
      Case vbKeyEscape
      '結束
        If TBar1.Buttons(14).Enabled = True Then
            Call Tbar1_ButtonClick(TBar1.Buttons(14))
         End If
    End Select
End Sub

'設定輸入物件
Private Sub SetInputs(Optional ByVal iStatus As Integer = 0)
   Dim bolCommand1 As Boolean, bolCommand2 As Boolean
   Dim oCommand As CommandButton
   
   
   Select Case iStatus
      Case 0 '瀏覽
         txt2PS(1).Locked = True
         txt2PS(2).Locked = True
         bolCommand1 = False
         bolCommand2 = True
         Command1(1).Enabled = True
         
      Case 1 '新增
         txt2PS(1).Locked = True
         txt2PS(2).Locked = False
         bolCommand1 = True
         bolCommand2 = False
         Command1(1).Enabled = True
         
      Case 2 '修改
         txt2PS(1).Locked = True
         txt2PS(2).Locked = False
         bolCommand1 = True
         bolCommand2 = False
         Command1(1).Enabled = True
         
      Case 4 '查詢
         txt2PS(1).Locked = False
         txt2PS(2).Locked = True
         bolCommand1 = False
         bolCommand2 = False
         Command1(1).Enabled = False
         
      Case 9 '無資料
         txt2PS(1).Locked = True
         txt2PS(2).Locked = True
         bolCommand1 = False
         bolCommand2 = False
         Command1(1).Enabled = False
   End Select
   
   Command1(0).Enabled = bolCommand1
   For Each oCommand In Command2
      oCommand.Enabled = bolCommand2
   Next
   
   If bolUpdate = False Then Command2(1).Enabled = False '有修改權限才能上傳
   
End Sub

Private Sub SetPath()
   '讀取前次設定路徑
   txt2Path.Text = GetSetting("TAIE", "P", UCase(Me.Name) & "Dir", "")
   If txt2Path <> "" Then ChkPath
End Sub

Private Function ChkPath() As Boolean
   If txt2Path = "" Then
      MsgBox "[ 待匯入資料夾 ] 尚未設定！", vbExclamation
   Else
      If PUB_ChkDir(txt2Path) = True Then
         ChkPath = True
      Else
         MsgBox "待匯入資料夾 [ " & txt2Path & " ] 不存在，請重新設定！", vbCritical
         txt2Path = ""
      End If
   End If
End Function

Private Function CheckConfirm() As Boolean
   Dim ii As Integer, iMissCount As Integer
   
   If iCurState = "1" Or iCurState = "2" Then
      If txt2PS(2) = "" Then
         MsgBox "請輸入檢索內容！", vbExclamation
         txt2PS(2).SetFocus
         Exit Function
      End If
      
      If txt2PS(3) = "" Then
         MsgBox "尚未匯入完成！", vbCritical
         Exit Function
      End If
      
      If m_bolImport = True Then
         iMissCount = 0
         With MSHFlexGrid1
         For ii = 1 To .Rows - 1
            If .TextMatrix(ii, 6) = "" Then
               iMissCount = iMissCount + 1
            End If
         Next
         End With
         If iMissCount > 0 Then
             If MsgBox("有 " & iMissCount & " 筆資料無簡圖檔，是否確定要繼續？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
               Exit Function
             End If
         End If
      End If
      
   ElseIf iCurState = "4" Then
      If txt2PS(1) = "" Then
         MsgBox "請輸入欲查詢的流水號！", vbExclamation
         txt2PS(1).SetFocus
         SetMouseReady
         Exit Function
      End If
   End If
   
   CheckConfirm = True
End Function

Private Function insertdata() As Boolean
   Dim stNo As String, strFtpPath As String, bolInTrans As Boolean
   Dim strPSD03 As String
   Dim ii As Integer
   
On Error GoTo ErrHnd
   
   m_Connection.BeginTrans
   bolInTrans = True
   stNo = GetAutoNo("PSS")
   
   strFtpPath = ""
   If txt2PS(3) <> "" Then
      strPSD03 = "TWG" & strSrvDate(2) & Format(stNo, "000000") & ".HTML"
      PUB_PutFtpFile txt2ImgFolder & "\" & txt2PS(3), stNo, strPSD03, strFtpPath, UCase("PatentSearch")
   End If
   
   strSql = "insert into PatentSearch(PS01,PS02,PS03,PS04,PS05,PS06)" & _
      " values(" & stNo & ",'" & ChgSQL(txt2PS(2)) & "','" & strPSD03 & "','" & strFtpPath & "','" & strUserNum & "',sysdate)"
   m_Connection.Execute strSql
   
   With MSHFlexGrid1
   ProgressBar1.max = .Rows - 1
   ProgressBar1.Value = 0
   lblProgress.Caption = "簡圖上傳中"
   Frame1.Visible = True
   For ii = 1 To .Rows - 1
      strFtpPath = ""
      If .TextMatrix(ii, 6) <> "" Then
         PUB_PutFtpFile txt2ImgFolder & "\" & .TextMatrix(ii, 6), stNo, .TextMatrix(ii, 6), strFtpPath, UCase("PatentSearchDetail")
      End If
      'Modify by Amy 2024/05/07 PSD05 要去/存入
      strSql = "insert into PatentSearchDetail(PSD01,PSD02,PSD03,PSD04,PSD05,PSD06,PSD07,PSD08,PSD09)" & _
         " values(" & stNo & "," & .TextMatrix(ii, 0) & ",'" & .TextMatrix(ii, 1) & "','" & ChgSQL(.TextMatrix(ii, 2)) & "'," & Val(Replace(.TextMatrix(ii, 3), "/", "")) & ",'" & .TextMatrix(ii, 4) & "','" & ChgSQL(.TextMatrix(ii, 5)) & "','" & .TextMatrix(ii, 6) & "','" & strFtpPath & "')"
      m_Connection.Execute strSql
      ProgressBar1.Value = ii
      DoEvents
   Next
   Frame1.Visible = False
   End With
   
   If Dir(txt2ImgFolder & "\" & strPSD03) <> "" Then Kill txt2ImgFolder & "\" & strPSD03
   Name txt2ImgFolder & "\" & txt2PS(3) As txt2ImgFolder & "\" & strPSD03
   
   m_Connection.CommitTrans
   
   txt2PS(1) = stNo
   insertdata = True
   Exit Function
   
ErrHnd:
   If bolInTrans Then m_Connection.RollbackTrans
   MsgBox Err.Description, vbCritical
   Frame1.Visible = False
End Function

Private Function UpdateData() As Boolean
   Dim stNo As String, strFtpPath As String, bolInTrans As Boolean
   Dim strPSD03 As String
   Dim ii As Integer
         
On Error GoTo ErrHnd
   
   m_Connection.BeginTrans
   bolInTrans = True
   stNo = txt2PS(1)
   
   If txt2PS(3) <> "" Then
      
      If m_bolImport = True Then
         PUB_DelFtpFile2 txt2PS(1), , UCase("PatentSearchDetail")   '檔案放FTP,必須在DB資料刪除前執行
         strSql = "delete PatentSearchDetail where PSD01=" & txt2PS(1)
         m_Connection.Execute strSql
         
         PUB_DelFtpFile2 txt2PS(1), , UCase("PatentSearch")
         strPSD03 = "TWG" & strSrvDate(2) & Format(stNo, "000000") & ".HTML"
         PUB_PutFtpFile txt2ImgFolder & "\" & txt2PS(3), stNo, strPSD03, strFtpPath, UCase("PatentSearch")
         m_strPS04 = strFtpPath
         
         With MSHFlexGrid1
         For ii = 1 To .Rows - 1
            If .TextMatrix(ii, 6) <> "" Then
               PUB_PutFtpFile txt2ImgFolder & "\" & .TextMatrix(ii, 6), stNo, .TextMatrix(ii, 6), strFtpPath, UCase("PatentSearchDetail")
            End If
            'Modify by Amy 2024/05/07 PSD05 要去/存入
            strSql = "insert into PatentSearchDetail(PSD01,PSD02,PSD03,PSD04,PSD05,PSD06,PSD07,PSD08,PSD09)" & _
               " values(" & stNo & "," & .TextMatrix(ii, 0) & ",'" & .TextMatrix(ii, 1) & "','" & ChgSQL(.TextMatrix(ii, 2)) & "'," & Val(Replace(.TextMatrix(ii, 3), "/", "")) & ",'" & .TextMatrix(ii, 4) & "','" & ChgSQL(.TextMatrix(ii, 5)) & "','" & .TextMatrix(ii, 6) & "','" & strFtpPath & "')"
            m_Connection.Execute strSql
         Next
         End With
      Else
         strPSD03 = txt2PS(3)
         
         '檢索內容有變時
         If txt2PS(3).Tag <> txt2PS(2).Tag Then
            PUB_DelFtpFile2 txt2PS(1), , UCase("PatentSearch")
            PUB_PutFtpFile txt2ImgFolder & "\" & txt2PS(3), stNo, strPSD03, strFtpPath, UCase("PatentSearch")
            m_strPS04 = strFtpPath
         End If
      End If
   End If
   
   strSql = "Update PatentSearch set PS02='" & ChgSQL(txt2PS(2)) & "',PS03='" & strPSD03 & "'" & _
      ",PS04='" & m_strPS04 & "',PS07='" & strUserNum & "',PS08=sysdate" & _
      " where PS01=" & stNo
   m_Connection.Execute strSql
   
   m_Connection.CommitTrans
   
   UpdateData = True
   Exit Function
   
ErrHnd:
   If bolInTrans Then m_Connection.RollbackTrans
   MsgBox Err.Description, vbCritical
End Function

Private Function DeleteData() As Boolean
   Dim bolInTrans As Boolean
   
   '刪除網站檔案
   If txt2PS(11) <> "" Then
      If PUB_DeleteWWW(txt2PS(1)) = False Then Exit Function
   End If
   
On Error GoTo ErrHnd
   m_Connection.BeginTrans
   bolInTrans = True
   
   PUB_DelFtpFile2 txt2PS(1), , UCase("PatentSearchDetail")   '檔案放FTP,必須在DB資料刪除前執行
   strSql = "delete PatentSearchDetail where PSD01=" & txt2PS(1)
   m_Connection.Execute strSql
   
   PUB_DelFtpFile2 txt2PS(1), , UCase("PatentSearch")   '檔案放FTP,必須在DB資料刪除前執行
   strSql = "delete PatentSearch where PS01=" & txt2PS(1)
   Pub_SeekTbLog strSql
   m_Connection.Execute strSql
   
   m_Connection.CommitTrans
   DeleteData = True
   Exit Function
   
ErrHnd:
   If bolInTrans Then m_Connection.RollbackTrans
   MsgBox Err.Description, vbCritical
End Function

Private Sub SetGrid(Optional pReset As Boolean = False)
   Dim arrMSHFlexGrid1HeadWidth
   Dim iCol As Integer
   Dim iUbound As Integer

   arrMSHFlexGrid1HeadWidth = Array(350, 800, 1600, 1000, 1000, 1500, 1750)
   iUbound = UBound(arrMSHFlexGrid1HeadWidth)
   
   With MSHFlexGrid1
   .Redraw = False
   If pReset = True Then
      .Clear
      .Rows = 2
      .FixedCols = 0
      iPrevRow1 = 0
   End If
   
   .FormatString = "#|專利編號|專利名稱|公告/公開日|申請號|申請人|簡圖檔名|FTP路徑"
   For iCol = 0 To .Cols - 1
      If iCol <= iUbound Then
         .ColWidth(iCol) = arrMSHFlexGrid1HeadWidth(iCol)
         .ColAlignment(iCol) = flexAlignLeftCenter
      Else
         .ColWidth(iCol) = 0
      End If
   Next
   If pReset = False Then
      .FixedCols = 1
   End If
   .Redraw = True
   End With
End Sub

Private Sub SetGrid2(Optional pReset As Boolean = False)
   Dim arrMSHFlexGrid1HeadWidth
   Dim iCol As Integer
   Dim iUbound As Integer

   arrMSHFlexGrid1HeadWidth = Array(650, 3000, 1000, 1000, 2400)
   iUbound = UBound(arrMSHFlexGrid1HeadWidth)
   
   With MSHFlexGrid2
   .Redraw = False
   If pReset = True Then
      .Clear
      .Rows = 2
      .FixedCols = 0
      iPrevRow2 = 0
   End If
   
   .FormatString = "流水號|檢索內容|建立日期|建立人員|網頁檔名"
   For iCol = 0 To .Cols - 1
      If iCol <= iUbound Then
         .ColWidth(iCol) = arrMSHFlexGrid1HeadWidth(iCol)
         .ColAlignment(iCol) = flexAlignLeftCenter
      Else
         .ColWidth(iCol) = 0
      End If
   Next
   If pReset = False Then
      .FixedCols = 1
   End If
   .Redraw = True
   End With
End Sub

'清除畫面
Private Sub FormReset()
   txt2PS(1) = ""
   txt2PS(2) = ""
   txt2PS(3) = ""
   txt2PS(11) = ""
   txt2ImgFolder = ""
   txt2PS(11) = ""
   lbl2Dates(1) = ""
   lbl2Dates(2) = ""
   
   txt2PS(2).Tag = ""
   txt2PS(3).Tag = ""
   txt2PS(11).Tag = ""
   
   SetGrid True
End Sub

Private Sub txt2Find_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txt2MailTo_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txt2PS_GotFocus(Index As Integer)
   TextInverse txt2PS(Index)
End Sub

Private Function LoadXLS(pXLS As String, pImgFolder As String, Optional pPicError As Boolean = False) As Boolean
   Dim xlsReport As New Excel.Application
   Dim wksReport As New Worksheet
   Dim ii As Integer, iRow As Integer, iMissCount As Integer
   Dim stPSD08 As String, arrCol(5) As String, strCode As String, strName As String
   
   SetGrid True
   xlsReport.Workbooks.Open txt2Path & "\" & pXLS
   
On Error GoTo ErrHnd
   'xlsReport.Visible = True
   Set wksReport = xlsReport.Worksheets(1)
   
   '欄位檢查
   With wksReport
   ii = 1
   If .Range("B1") = "" Then ii = ii + 1 'Added by Morgan 2022/11/2 第1列增加輸出的日期及筆數
   
   strCode = "A"
   Do While (.Range(strCode & ii) <> "")
      Select Case .Range(strCode & ii)
         'Modified by Morgan 2021/8/19 配合智慧局檢索系統 8/16 改新版,舊版用到 9/30,增加新的欄位 "公開公告號","公開公告日"
         Case "專利編號", "公開公告號"
            arrCol(1) = strCode
         Case "專利名稱"
            arrCol(2) = strCode
         Case "公告/公開日", "公開公告日"
            arrCol(3) = strCode
         Case "申請號"
            arrCol(4) = strCode
         Case "申請人"
            arrCol(5) = strCode
      End Select
      
      If strCode = "ZZ" Then
         Exit Do
      ElseIf Right(strCode, 1) = "Z" Then
         strCode = Chr(Asc(Left(strCode, 1)) + 1) & "A"
      ElseIf Len(strCode) = 1 Then
         strCode = Chr(Asc(strCode) + 1)
      Else
         strCode = Left(strCode, 1) & Chr(Asc(Right(strCode, 1)) + 1)
      End If
   Loop
   End With
   
         
   If arrCol(1) = "" Then
      If strSrvDate(1) > "20210930" Then
         MsgBox "找不到欄位 [ 公開公告號 ]！", vbExclamation, "欄位檢查"
      Else
         MsgBox "找不到欄位 [ 專利編號 ]！", vbExclamation, "欄位檢查"
      End If
      GoTo ExitPoint
   End If
   If arrCol(2) = "" Then
      MsgBox "找不到欄位 [ 專利名稱 ]！", vbExclamation, "欄位檢查"
      GoTo ExitPoint
   End If
   If arrCol(3) = "" Then
      MsgBox "找不到欄位 [ 公告/公開日 ]！", vbExclamation, "欄位檢查"
      GoTo ExitPoint
   End If
   If arrCol(4) = "" Then
      MsgBox "找不到欄位 [ 申請號 ]！", vbExclamation, "欄位檢查"
      GoTo ExitPoint
   End If
   If arrCol(5) = "" Then
      MsgBox "找不到欄位 [ 申請人 ]！", vbExclamation, "欄位檢查"
      GoTo ExitPoint
   End If
   
   iMissCount = 0
   ii = ii + 1
   'Added by Morgan 2019/6/14 欄位名稱可能跨列
   If wksReport.Range(arrCol(1) & ii) = "" Then
      ii = ii + 1
   End If
   If wksReport.Range(arrCol(1) & ii) = "" Then
      MsgBox "格式錯誤或無資料，請確認 XLS 檔案內容是否正確！", vbExclamation, "XLS 檔案檢查"
      GoTo ExitPoint
   End If
   'end 2019/6/14
   
   iRow = 1
   With MSHFlexGrid1
   Do While wksReport.Range(arrCol(1) & ii) <> ""
      .Rows = iRow + 1
      .TextMatrix(iRow, 0) = iRow
      .TextMatrix(iRow, 1) = wksReport.Range(arrCol(1) & ii)
      .TextMatrix(iRow, 2) = wksReport.Range(arrCol(2) & ii)
      'Memo by Amy 2024/05/07 公告/公開日 要顯示有/的西元年月日
      .TextMatrix(iRow, 3) = wksReport.Range(arrCol(3) & ii)
      .TextMatrix(iRow, 4) = wksReport.Range(arrCol(4) & ii)
      .TextMatrix(iRow, 5) = wksReport.Range(arrCol(5) & ii)
      
      'Added by Morgan 2021/8/24
      If Check1.Value = vbChecked Then
         stPSD08 = SearhPic(.TextMatrix(iRow, 4))
      Else
         stPSD08 = Dir(txt2Path & "\" & pImgFolder & "\TWG1" & wksReport.Range(arrCol(4) & ii) & ".png")
      End If
      
      If stPSD08 = "" Then
         iMissCount = iMissCount + 1
      Else
         .TextMatrix(iRow, 6) = stPSD08
      End If
      
      ii = ii + 1
      iRow = iRow + 1
   Loop
   .Refresh
   End With
   
   
   '檢索結果不一定有附簡圖
   'If iMissCount > 0 Then
   '    MsgBox "匯入失敗，缺簡圖檔(" & iMissCount & ")！", vbExclamation
   '    GoTo ExitPoint
   'End If
   
   LoadXLS = True
   
   
ErrHnd:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
   
ExitPoint:
   xlsReport.ActiveWorkbook.Close savechanges:=False
   xlsReport.Quit
   Set wksReport = Nothing
   Set xlsReport = Nothing
End Function

Private Function MakeHTML(pFileName As String) As Boolean
   Dim strHTML As String
   Dim ii As Integer, jj As Integer, stGNo As String
   Dim objStream As Object
   
   'Modified by Morgan 2020/3/30 事務所名稱改用函數抓
   strHTML = "<!DOCTYPE html>" & vbCrLf & _
      "<html><head>" & vbCrLf & _
      "<link rel=""shortcut icon"" href=""https://www.taie.com.tw/favicon.ico"" type=""image/x-icon"" />" & vbCrLf & _
      "<meta charset=""utf-8"" />" & vbCrLf & _
      "<style type=""text/css"">" & vbCrLf & _
      ".row{margin-right:-15px;margin-left:-15px}" & vbCrLf & _
      ".clsColName{color:#125e05;letter-spacing:1px;background-color:#ade5a6;white-space: nowrap;font-size:13px;font-weight:bold;word-break:keep-all;word-wrap:normal;text-align:center;border-bottom:1px solid #fff;width:80px}" & vbCrLf & _
      ".clsColValue{padding-left:3px;border-bottom:1px dotted #ccc;}" & vbCrLf & _
      "</style>" & vbCrLf & _
      "<title>" & PUB_GetCompName2("1") & "</title>" & vbCrLf & _
      "</head>" & vbCrLf
   strHTML = strHTML & "<body>" & vbCrLf
   'Modified by Morgan 2024/8/26 配合公司網頁更新,改首頁及圖片的網址(兩個IMG要直接串在一起中間才不會多一個底線符號)
   strHTML = strHTML & "<div>" & vbCrLf & _
      "<a href=""https://www.taie.com.tw"">" & vbCrLf & _
      "<img src=""https://www.taie.com.tw/assets/images/common/logo_red.svg"" width=""80"" height=""55"" /><img src=""https://www.taie.com.tw/assets/images/common/logo_text.svg"" width=""240"" height="""" alt=""" & PUB_GetCompName2("1") & """ />" & vbCrLf & _
      "</a>" & vbCrLf & _
      "</div>" & vbCrLf
      
   strHTML = strHTML & "<table>" & vbCrLf
   strHTML = strHTML & "<tr><td width=100% height=40 style=""border:1px solid D9D9C3;padding-left:10pt""><font style=""font-size: 22px; color: blue; display: inline-block; text-align: left; font-weight:bold"">專利公報資訊</font></td></tr>" & vbCrLf
   strHTML = strHTML & "<tr><td width=100% height=20 style=""border:1px solid D9D9C3;padding-left:10pt"" bgcolor=lemonchiffon><font style=""font-size: 14px;color:brown;font-weight:bold"">檢索內容：" & txt2PS(2) & "</font></td></tr>" & vbCrLf
   strHTML = strHTML & "</table>"
   
   strHTML = strHTML & "<table border=1>" & vbCrLf
   With MSHFlexGrid1
   For ii = 1 To .Rows - 1
      If ii Mod 3 = 1 Then
         strHTML = strHTML & "<tr valign=""top"">" & vbCrLf
      End If
         
      strHTML = strHTML & "<td><table>" & vbCrLf & _
         "<tr>" & vbCrLf & _
         "<td class=""clsColName""  valign=""top"">#</td>" & vbCrLf & _
         "<td class=""clsColValue"">" & .TextMatrix(ii, 0) & "</a></td>" & vbCrLf & _
         "</tr>" & vbCrLf & _
         "<tr>" & vbCrLf & _
         "<td class=""clsColName""  valign=""top"">專利編號</td>" & vbCrLf
         'Memo by Amy 2024/05/07 公告日(.TextMatrix(ii, 3))有改要看UpdateData 和 insertdata PSD05 欄位是否也要改
         '公告號
         If .TextMatrix(ii, 1) > "A" Or Len(.TextMatrix(ii, 1)) < 8 Then
            'Modified by Morgan 2021/8/19 新版日期有"/"
            'stGNo = Format(Left(.TextMatrix(ii, 3), 4) - 1973, "000") & Format(Mid(.TextMatrix(ii, 3), 5, 2) * 3 - (2 - Mid(.TextMatrix(ii, 3), 7, 1)), "000")
            stGNo = Replace(.TextMatrix(ii, 3), "/", "")
            stGNo = Format(Left(stGNo, 4) - 1973, "000") & Format(Mid(stGNo, 5, 2) * 3 - (2 - Mid(stGNo, 7, 1)), "000")
            'end 2021/8/19
            'Modify by Amy 2024/05/07 改網址
            strHTML = strHTML & _
               "<td class=""clsColValue""><a title="""" href=""https://cloud.tipo.gov.tw/S220/downloads/patent/isu" & stGNo & "/pdfdata/" & .TextMatrix(ii, 4) & ".pdf"" target=""_blank"">" & .TextMatrix(ii, 1) & "</a></td>" & vbCrLf
         '公開號
         Else
            'Modified by Morgan 2021/8/19 新版日期有"/"
            'stGNo = Format(Left(.TextMatrix(ii, 3), 4) - 2002, "000") & Format(Mid(.TextMatrix(ii, 3), 5, 2) * 2 - (1 - Mid(.TextMatrix(ii, 3), 7, 1)), "000")
            stGNo = Replace(.TextMatrix(ii, 3), "/", "")
            stGNo = Format(Left(stGNo, 4) - 2002, "000") & Format(Mid(stGNo, 5, 2) * 2 - (1 - Mid(stGNo, 7, 1)), "000")
            'end 2021/8/19
            'Modify by Amy 2024/05/07 改網址
            strHTML = strHTML & _
               "<td class=""clsColValue""><a title="""" href=""https://cloud.tipo.gov.tw/S220/downloads/invention/pub" & stGNo & "/pdfdata/" & .TextMatrix(ii, 4) & ".pdf"" target=""_blank"">" & .TextMatrix(ii, 1) & "</a></td>" & vbCrLf
         End If
      'Memo by Amy 2024/05/07 公告/公開日 要顯示有/的西元年月日
      strHTML = strHTML & "</tr>" & vbCrLf & _
         "<tr>" & vbCrLf & _
         "<td class=""clsColName""  valign=""top"">公告/公開日</td>" & vbCrLf & _
         "<td class=""clsColValue"">" & .TextMatrix(ii, 3) & "</a></td>" & vbCrLf & _
         "</tr>" & vbCrLf & _
         "<tr>" & vbCrLf & _
         "<td class=""clsColName""  valign=""top"">申請號</td>" & vbCrLf & _
         "<td class=""clsColValue"">" & .TextMatrix(ii, 4) & "</a></td>" & vbCrLf & _
         "</tr>" & vbCrLf & _
         "<tr>" & vbCrLf & _
         "<td class=""clsColName""  valign=""top"">專利名稱</td>" & vbCrLf & _
         "<td class=""clsColValue"" height=""40"" >" & .TextMatrix(ii, 2) & "</a></td>" & vbCrLf & _
         "</tr>" & vbCrLf & _
         "<tr>" & vbCrLf & _
         "<td class=""clsColName""  valign=""top"">申請人</td>" & vbCrLf & _
         "<td class=""clsColValue"" height=""40"">" & .TextMatrix(ii, 5) & "</a></td>" & vbCrLf & _
         "</tr>" & vbCrLf & _
         "<tr>" & vbCrLf & _
         "<td class=""clsColName""  valign=""top"">簡圖</td>" & vbCrLf & _
         "<td valign=""middle"">" & IIf(.TextMatrix(ii, 6) = "", "缺", "<img src=""./" & .TextMatrix(ii, 6) & """ border=""0"" width=""300"" height=""300"" align=""right"">") & "</td>" & vbCrLf & _
         "</tr>" & vbCrLf & _
         "</table></td>"
         
      If ii Mod 3 = 0 Then
         strHTML = strHTML & "</tr>" & vbCrLf
      End If
   Next ii
   End With
   
   If ii Mod 3 <> 1 Then
      strHTML = strHTML & "</tr>" & vbCrLf
   End If
   strHTML = strHTML & "</table>" & vbCrLf
   
   '免責聲明
   strHTML = strHTML & "<div style=""font-size: 12px; color: blue; display: inline-block; text-align: left;margin-top: 10px;"">"
   strHTML = strHTML & "本網頁所列專利公報資料完全取自經濟部智慧財產局之專利資訊檢索系統及開放資料，其僅供一般性參考使用；另本網頁製作時已採取合理措施，力求資料內容的正確性，當資料出現歧異時，以經濟部智慧財產局公告之資料為準。"
   strHTML = strHTML & "</div>"
   
   strHTML = strHTML & "</body></html>"
   
   If Dir(pFileName) <> "" Then Kill pFileName
   
   Set objStream = CreateObject("ADODB.Stream")
   With objStream
      .Type = 2
      .Mode = 3
      .Open
      .Charset = "Unicode"
      .WriteText strHTML
      .SaveToFile pFileName
      .Close
   End With
   MakeHTML = True
   txt2PS(3).Tag = txt2PS(2)
End Function

Private Sub OpenLocal()
   Dim hLocalFile As Long
   
   If txt2PS(3) = "" Then
      MsgBox "尚未匯入完成無法預覽！", vbCritical
      Exit Sub
      
   ElseIf DownLoadFiles() = True Then
      
      If txt2PS(3).Tag <> txt2PS(2) Then
         If MakeHTML(txt2ImgFolder & "\" & txt2PS(3)) = False Then
            MsgBox "網頁建立失敗！", vbCritical
            Exit Sub
         End If
      End If
      
      ShellExecute hLocalFile, "open", txt2ImgFolder & "\" & txt2PS(3), vbNullString, vbNullString, 3
      
   End If
End Sub

Private Sub OpenURL()
   Dim hLocalFile As Long
   If txt2PS(11) = "" Then
      MsgBox "連結尚未建立，請先上傳網站！", vbCritical
      Exit Sub
   Else
      ShellExecute hLocalFile, "open", txt2PS(11).Text, vbNullString, vbNullString, 3
   End If
End Sub

Private Sub MailURL()
   If txt2PS(11) = "" Then
      MsgBox "連結尚未建立，請先上傳網站！", vbCritical
      Exit Sub
   ElseIf txt2MailTo = "" Then
      MsgBox "請輸入收件者！", vbCritical
      txt2MailTo.SetFocus
      Exit Sub
   Else
      PUB_SendMail strUserNum, txt2MailTo, "", "「" & txt2PS(2) & "」專利公報檢索連結", "<FONT size=3>&nbsp;<A href=""" & txt2PS(11).Text & """>" & txt2PS(11).Text & "</A></FONT>"
      If bolMailSendOk Then MsgBox "連結已寄出！", vbInformation
   End If
End Sub

Private Function DownLoadFiles() As Boolean
   Dim ii As Integer
      
   '網頁存在就不再下載
   If txt2ImgFolder <> "" And txt2PS(3) <> "" Then
      If Dir(txt2ImgFolder & "\" & txt2PS(3)) <> "" Then
         DownLoadFiles = True
         Exit Function
      End If
   End If
      
   If m_strTempFolder = "" Then
      SetTempFolder
   End If
   
   If m_strTempFolder <> "" Then
      If PUB_GetFtpFile(m_strPS04, m_strTempFolder & "\" & txt2PS(3), "PATENTSEARCH", True) = False Then
         Exit Function
      Else
         With MSHFlexGrid1
         ProgressBar1.max = .Rows - 1
         ProgressBar1.Value = 0
         lblProgress.Caption = "簡圖下載中"
         Frame1.Visible = True
         For ii = 1 To .Rows - 1
            If .TextMatrix(ii, 6) <> "" Then
               If PUB_GetFtpFile(.TextMatrix(ii, 7), m_strTempFolder & "\" & .TextMatrix(ii, 6), "PATENTSEARCHDETAIL", True) = False Then
                  Frame1.Visible = False
                  Exit Function
               End If
            End If
            ProgressBar1.Value = ii
            DoEvents
         Next
         Frame1.Visible = False
         End With
      End If
      txt2ImgFolder = m_strTempFolder
      DownLoadFiles = True
   End If
End Function

Private Function Upload2WWW(Optional pNoPic As Boolean = False) As Boolean
   Dim stDir As String, stFolder As String, stURL As String
   
   If UCase(pub_DbTerminalName) <> UCase(正式資料庫電腦名稱) Then
      If MsgBox("目前連線非正式資料庫是否確定要上傳？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
         Exit Function
      End If
   End If
   
On Error GoTo ErrHnd
   
   If txt2PS(3) = "" Then
      MsgBox "網頁尚未產生，無法上傳！", vbCritical
      Exit Function
   End If

   stDir = Pub_GetSpecMan("FTP_WWW_Path")
   If stDir = "" Then
      MsgBox "網站的 FTP目錄[ FTP_WWW_Path ]尚未設定，無法上傳！", vbExclamation
      Exit Function
   End If
   stFolder = Format(txt2PS(1), "000000")
   stDir = stDir & "/" & stFolder
   
   If pNoPic = False Then
      If DownLoadFiles() = False Then
         Exit Function
      End If
   End If
   
   If Ftp2WWW(stDir, pNoPic) = True Then
      If txt2PS(1) <> "" Then
         stURL = "https://www.taie.com.tw/sales/" & stFolder & "/" & txt2PS(3)
         strSql = "update patentsearch set ps09='" & strUserNum & "',ps10=sysdate,ps11='" & ChgSQL(stURL) & "' where ps01='" & txt2PS(1) & "'"
         m_Connection.Execute strSql, intI
         txt2PS(11) = stURL
         Upload2WWW = True
      End If
   End If
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Function

Private Function Ftp2WWW(pDir As String, Optional pNoPic As Boolean = False) As Boolean
   Dim hConnection As Long, hFind As Long
   Dim stFtpIP As String, stLocalPath As String, stFtpFolder As String, stFileName As String
   Dim ii As Integer
   Dim dwInternetFlags As Integer
   Dim pData As WIN32_FIND_DATA
   
On Error GoTo ErrHnd
   
   hConnection = PUB_GetFtpConnect(, , "FTP_WWW_IP")
   If hConnection <> 0 Then
      '切換目錄來檢查是否存在
      If PUB_SetFtpDirectory(hConnection, pDir) = True Then
         If pNoPic = False Then
         
            '上傳圖檔
            With MSHFlexGrid1
            ProgressBar1.max = .Rows - 1
            ProgressBar1.Value = 0
            lblProgress.Caption = "簡圖上傳中"
            Frame1.Visible = True
            For ii = 1 To .Rows - 1
               If .TextMatrix(ii, 6) <> "" Then
                  stLocalPath = txt2ImgFolder & "\" & .TextMatrix(ii, 6)
                  stFileName = .TextMatrix(ii, 6)
                  hFind = FtpFindFirstFile(hConnection, stFileName, pData, 0, 0)
                  If hFind <> 0 Then
                     InternetCloseHandle hFind
                     hFind = 0
                     If InStr(pData.cFileName, stFileName & Chr(0)) = 1 Then
                        If FtpDeleteFile(hConnection, stFileName) = 0 Then
                           MsgBox stFileName & "的 FTP 舊檔無法刪除！", vbCritical
                           GoTo ErrHnd
                        End If
                     End If
                  End If
                  '上傳檔案
                  dwInternetFlags = 2 'INTERNET_FLAG_TRANSFER_BINARY
                  If FtpPutFile(hConnection, stLocalPath, stFileName, dwInternetFlags, 0) <> 1 Then
                     MsgBox stLocalPath & " 檔案上傳失敗！", vbCritical
                     GoTo ErrHnd
                  Else
                     '檢查檔案是否確實存在
                     hFind = FtpFindFirstFile(hConnection, stFileName, pData, 0, 0)
                     If hFind = 0 Then
                        MsgBox stLocalPath & " 檔案上傳失敗！", vbCritical
                        GoTo ErrHnd
                     Else
                        InternetCloseHandle hFind
                        hFind = 0
                     End If
                  End If
               End If
               ProgressBar1.Value = ii
               DoEvents
            Next
            Frame1.Visible = False
            End With
         End If
         
         '上傳網頁
         stLocalPath = txt2ImgFolder & "\" & txt2PS(3)
         stFileName = txt2PS(3)
         hFind = FtpFindFirstFile(hConnection, stFileName, pData, 0, 0)
         If hFind <> 0 Then
            InternetCloseHandle hFind
            hFind = 0
            If InStr(pData.cFileName, stFileName & Chr(0)) = 1 Then
               If FtpDeleteFile(hConnection, stFileName) = 0 Then
                  MsgBox stFileName & "的 FTP 舊檔無法刪除！", vbCritical
                  GoTo ErrHnd
               End If
            End If
         End If
            
         dwInternetFlags = 2 'INTERNET_FLAG_TRANSFER_BINARY
         '上傳檔案
         If FtpPutFile(hConnection, stLocalPath, stFileName, dwInternetFlags, 0) <> 1 Then
            MsgBox stLocalPath & " 檔案上傳失敗！", vbCritical
            GoTo ErrHnd
         Else
            '檢查檔案是否確實存在
            hFind = FtpFindFirstFile(hConnection, stFileName, pData, 0, 0)
            If hFind = 0 Then
               MsgBox stLocalPath & " 檔案上傳失敗！", vbCritical
               GoTo ErrHnd
            Else
               InternetCloseHandle hFind
               hFind = 0
               Ftp2WWW = True
            End If
         End If
      End If
   End If
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   If hConnection <> 0 Then InternetCloseHandle hConnection
   Frame1.Visible = False
End Function

Private Sub SetTempFolder()
   Dim strTempFolder As String
   
   m_strTempFolder = ""
On Error GoTo ErrHnd
   strTempFolder = App.path & "\" & strUserNum
   If Dir(strTempFolder, vbDirectory) = "" Then
      MkDir strTempFolder
   End If
   strTempFolder = strTempFolder & "\PatenSearch"
   If Dir(strTempFolder, vbDirectory) = "" Then
      MkDir strTempFolder
   End If
   m_strTempFolder = strTempFolder
   Exit Sub
ErrHnd:
   MsgBox Err.Description, vbCritical, "暫存資料夾設定失敗"
End Sub

Private Sub KillTemp()
   Dim iTimes As Integer
   If m_strTempFolder = "" Then Exit Sub
On Error GoTo ErrHnd
   If Dir(m_strTempFolder & "\.") <> "" Then
      Kill m_strTempFolder & "\*.*"
   End If
   Exit Sub
   
ErrHnd:
   If iTimes < 2 Then
      iTimes = iTimes + 1
      Sleep 1000
      Resume
   Else
      'MsgBox "暫存檔無法清除！" & vbCrLf & vbCrLf & "請重新執行本作業，否則有可能載入的不是最新的定稿！", vbExclamation
   End If
   Err.Clear
End Sub

Private Sub SetMouseBusy()
   Screen.MousePointer = vbHourglass
   MSHFlexGrid1.MousePointer = vbHourglass
End Sub

Private Sub SetMouseReady()
   Screen.MousePointer = vbDefault
   MSHFlexGrid1.MousePointer = vbDefault
End Sub

Private Sub SetMouseBusy2()
   Screen.MousePointer = vbHourglass
   MSHFlexGrid2.MousePointer = vbHourglass
End Sub

Private Sub SetMouseReady2()
   Screen.MousePointer = vbDefault
   MSHFlexGrid2.MousePointer = vbDefault
End Sub


Private Sub txt2Find_Change(Index As Integer)
   If Index = 2 Then
      lbl2Find = ""
   End If
End Sub

Private Sub txt2Find_GotFocus(Index As Integer)
   TextInverse txt2Find(Index)
End Sub

Private Sub txt2Find_Validate(Index As Integer, Cancel As Boolean)
   If Index = 2 Then
      txt2Find(2) = Trim(txt2Find(2))
      If txt2Find(2) <> "" And txt2Find(2).Tag <> txt2Find(2) Then
         If txt2Find(2) > "6" And txt2Find(2) < "F" Then
            If ClsPDGetStaff(txt2Find(2), strExc(1)) Then
               lbl2Find = strExc(1)
            Else
               Cancel = True
               txt2Find_GotFocus 2
            End If
         Else
            If GetIdFromName(txt2Find(2), strExc(1)) Then
               strExc(0) = txt2Find(2)
               txt2Find(2) = strExc(1)
               lbl2Find = strExc(0)
            Else
               Cancel = True
               txt2Find_GotFocus 2
            End If
         End If
         txt2Find(2).Tag = txt2Find(2)
      End If
   End If
End Sub

Private Sub txt2ImgFolder_GotFocus()
   TextInverse txt2ImgFolder
End Sub

Private Sub txt2MailTo_Change()
   lbl2MailTo = ""
End Sub

Private Sub txt2MailTo_GotFocus()
   TextInverse txt2MailTo
End Sub

Private Sub txt2MailTo_Validate(Cancel As Boolean)
   txt2MailTo = Trim(txt2MailTo)
   If txt2MailTo <> "" And txt2MailTo.Tag <> txt2MailTo Then
      If txt2MailTo > "6" And txt2MailTo < "F" Then
         If ClsPDGetStaff(txt2MailTo, strExc(1)) Then
            lbl2MailTo = strExc(1)
         Else
            Cancel = True
            txt2MailTo_GotFocus
         End If
      Else
         If GetIdFromName(txt2MailTo, strExc(1)) Then
            strExc(0) = txt2MailTo
            txt2MailTo = strExc(1)
            lbl2MailTo = strExc(0)
         Else
            Cancel = True
            txt2MailTo_GotFocus
         End If
      End If
      txt2MailTo.Tag = txt2MailTo
   End If
End Sub

Private Sub txt2Path_GotFocus()
   TextInverse txt2Path
End Sub


Private Function GetIdFromName(ByVal pName As String, ByRef pID As String) As Boolean
   strExc(0) = "select st01,st02 from staff where st02='" & ChgSQL(pName) & "' and st04='1' and st01>'6' and st01<'F'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If RsTemp.RecordCount = 1 Then
         pID = RsTemp.Fields("st01")
         GetIdFromName = True
      Else
         MsgBox "員工名稱重複，請直接輸入員工編號！"
      End If
   Else
      MsgBox "該員工名稱不存在！"
   End If
End Function

Private Function ClsLawReadRstMsg(ByRef i As Integer, ByRef strSql As String, Optional Server As Boolean = False, _
      Optional ErrShow As Boolean = False) As ADODB.Recordset
 
 Dim Rc As New ADODB.Recordset
On Error GoTo ErrHand
   If Server = False Then
      Rc.CursorLocation = adUseClient
      Rc.CursorType = adOpenStatic
      Rc.LockType = adLockReadOnly
      Rc.Open strSql, m_Connection
      If Rc.RecordCount > 0 Then
         Set ClsLawReadRstMsg = Rc
         i = 1
      Else
         If i = 0 Then MsgBox "資料庫無資料 !", vbInformation
         Set ClsLawReadRstMsg = Rc
         i = 0
      End If
   Else
      Rc.CursorLocation = adUseServer
      Rc.CursorType = adOpenStatic
      Rc.LockType = adLockReadOnly
      
      Rc.Open strSql, m_Connection
      If Rc.EOF And Rc.BOF Then
         If i = 0 Then MsgBox "資料庫無資料 !", vbInformation
         Set ClsLawReadRstMsg = Rc
         i = 0
      Else
         Set ClsLawReadRstMsg = Rc
         i = 1
      End If
   End If
   Exit Function
ErrHand:
   i = 2
   If ErrShow = False Then MsgBox "錯誤 : " & Err.Description, vbCritical
End Function


'讀資料庫電腦名稱
Private Function GetDbTerminal() As String
   Dim stSQL As String, intR As Integer
   Dim rsQuery As ADODB.Recordset
   
On Error GoTo ErrHnd
   
   stSQL = "select TERMINAL FROM V$SESSION where SID=1"
   intR = 1
   Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      pub_DbTerminalName = rsQuery.Fields(0)
      GetDbTerminal = "(" & rsQuery.Fields(0) & ")"
   End If
   
ErrHnd:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
   Set rsQuery = Nothing
End Function
'Added by Morgan 2021/8/24
'pAppNo:申請號,pPatNo:公開/公告號
Private Function SearhPic(pAppNo As String) As String
   Dim strHTMLFile As String, strHTMLText As String, strPicPath As String, strPicName As String
   Dim adoStream As ADODB.Stream
   Dim lngP1 As Long, lngP2 As Long
   
   strHTMLFile = Dir(txt2Path & "\*.html")
   If strHTMLFile <> "" Then
      Set adoStream = New ADODB.Stream
      adoStream.Charset = "UTF-8"
      adoStream.Open
      Do While strHTMLFile <> ""
         strHTMLFile = txt2Path & "\" & strHTMLFile
         adoStream.LoadFromFile strHTMLFile
         strHTMLText = adoStream.ReadText
         lngP1 = InStr(strHTMLText, pAppNo) '申請號
         If lngP1 > 0 Then
            lngP1 = InStr(lngP1, strHTMLText, "簡圖") '簡圖
            If lngP1 > 0 Then
               lngP1 = InStr(lngP1, strHTMLText, "<td ")
               If lngP1 > 0 Then
                  lngP2 = InStr(lngP1, strHTMLText, "</td>")
                  If lngP2 > 0 Then
                     lngP1 = InStr(lngP1, strHTMLText, "<img src=")
                     If lngP1 > 0 And lngP1 < lngP2 Then
                        lngP1 = lngP1 + Len("<img src=") + 1
                        lngP2 = InStr(lngP1, strHTMLText, """")
                        strPicPath = Mid(strHTMLText, lngP1, lngP2 - lngP1)
                        strPicPath = Replace(strPicPath, "/", "\")
                        strPicPath = Replace(strPicPath, ".\", txt2Path & "\")
                        strPicName = Mid(strPicPath, InStrRev(strPicPath, "\") + 1)
                        '若非簡圖資料夾時要複製
                        If InStr(strPicPath, txt2ImgFolder & "\") <> 1 Then
                           FileCopy strPicPath, txt2ImgFolder & "\" & strPicName
                        End If
                     End If
                  End If
               End If
            End If
         End If
         If strPicName <> "" Then Exit Do
         strHTMLFile = Dir()
      Loop
      adoStream.Close
      SearhPic = strPicName
   End If
   Set adoStream = Nothing
End Function
