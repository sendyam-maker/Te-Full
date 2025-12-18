VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060504 
   BorderStyle     =   1  '單線固定
   Caption         =   "案件命名追蹤"
   ClientHeight    =   6108
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7572
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6108
   ScaleWidth      =   7572
   Begin VB.CommandButton cmdDelFTP 
      BackColor       =   &H00C0FFFF&
      Caption         =   "刪除原始檔"
      Height          =   300
      Left            =   2808
      Style           =   1  '圖片外觀
      TabIndex        =   53
      Top             =   2688
      Width           =   1284
   End
   Begin VB.CommandButton cmdAddFile 
      Caption         =   "上傳檔案"
      Height          =   380
      Left            =   936
      TabIndex        =   52
      Top             =   1128
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "新案認領"
      Height          =   1275
      Left            =   120
      TabIndex        =   35
      Top             =   4752
      Width           =   7395
      Begin VB.Frame FraTCN13 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  '沒有框線
         Caption         =   "Frame3"
         Height          =   315
         Left            =   60
         TabIndex        =   49
         Top             =   900
         Width           =   4368
         Begin VB.CheckBox Chk1 
            Caption         =   "待確定"
            Height          =   195
            Index           =   2
            Left            =   3300
            TabIndex        =   44
            Top             =   60
            Width           =   945
         End
         Begin VB.CheckBox Chk1 
            Caption         =   "有"
            Height          =   195
            Index           =   1
            Left            =   2640
            TabIndex        =   43
            Top             =   60
            Width           =   615
         End
         Begin VB.CheckBox Chk1 
            Caption         =   "無"
            Height          =   195
            Index           =   0
            Left            =   1980
            TabIndex        =   42
            Top             =   60
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "外文本的對應英/中說："
            Height          =   195
            Index           =   17
            Left            =   90
            TabIndex        =   50
            Top             =   60
            Width           =   1845
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   276
         Left            =   4170
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   570
         Width           =   2445
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   11
         Left            =   4440
         TabIndex        =   51
         Top             =   888
         Visible         =   0   'False
         Width           =   360
         VariousPropertyBits=   679495707
         MaxLength       =   1
         Size            =   "635;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   10
         Left            =   6180
         TabIndex        =   39
         Top             =   210
         Width           =   360
         VariousPropertyBits=   679495707
         MaxLength       =   1
         Size            =   "635;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "客戶有提供彩圖：            (Y:是)"
         Height          =   255
         Index           =   15
         Left            =   4740
         TabIndex        =   48
         Top             =   233
         Width           =   2535
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   9
         Left            =   3420
         TabIndex        =   38
         Top             =   210
         Width           =   360
         VariousPropertyBits=   679495707
         MaxLength       =   1
         Size            =   "635;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "英文組認領：            (Y:是)"
         Height          =   255
         Index           =   12
         Left            =   2280
         TabIndex        =   47
         Top             =   233
         Width           =   2235
      End
      Begin VB.Label Label1 
         Caption         =   "指定組別："
         Height          =   195
         Index           =   16
         Left            =   3240
         TabIndex        =   46
         Top             =   630
         Width           =   975
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   8
         Left            =   1410
         TabIndex        =   40
         Top             =   570
         Width           =   1500
         VariousPropertyBits=   679495707
         MaxLength       =   12
         Size            =   "2646;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "相似舊案案號："
         Height          =   255
         Index           =   13
         Left            =   150
         TabIndex        =   45
         Top             =   600
         Width           =   1335
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   7
         Left            =   1050
         TabIndex        =   37
         Top             =   210
         Width           =   360
         VariousPropertyBits=   679495707
         MaxLength       =   1
         Size            =   "635;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "暫不認領：            (Y:是)"
         Height          =   252
         Index           =   14
         Left            =   156
         TabIndex        =   36
         Top             =   228
         Width           =   2052
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "急件翻譯"
      Height          =   1335
      Left            =   120
      TabIndex        =   20
      Top             =   3300
      Width           =   7335
      Begin VB.CommandButton cmdMail 
         Caption         =   "E-Mail"
         Height          =   380
         Left            =   6360
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   6
         Left            =   1560
         TabIndex        =   6
         Top             =   937
         Width           =   960
         VariousPropertyBits=   679495707
         MaxLength       =   7
         Size            =   "1693;529"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "只交Claims期限:："
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   34
         Top             =   960
         Width           =   1695
      End
      Begin MSForms.ComboBox cboTarget 
         Height          =   300
         Left            =   4800
         TabIndex        =   8
         Top             =   600
         Width           =   1395
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "2461;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox cboSource 
         Height          =   300
         Left            =   4800
         TabIndex        =   7
         Top             =   240
         Width           =   1395
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "2469;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   255
         Index           =   4
         Left            =   4800
         TabIndex        =   33
         Top             =   960
         Width           =   1440
         BackColor       =   -2147483643
         Size            =   "2540;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox cboTCN15 
         Height          =   300
         Left            =   1080
         TabIndex        =   4
         Top             =   240
         Width           =   1815
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "3201;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   5
         Left            =   1080
         TabIndex        =   5
         Top             =   600
         Width           =   960
         VariousPropertyBits=   679495707
         MaxLength       =   7
         Size            =   "1693;529"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   270
         Index           =   4
         Left            =   2760
         TabIndex        =   29
         Top             =   480
         Visible         =   0   'False
         Width           =   615
         VariousPropertyBits=   746604571
         Size            =   "1085;476"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label31 
         Caption         =   "對外翻用-"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3000
         TabIndex        =   27
         Top             =   285
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "原文語種："
         Height          =   255
         Index           =   10
         Left            =   3840
         TabIndex        =   26
         Top             =   285
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "翻譯語種："
         Height          =   255
         Index           =   7
         Left            =   3840
         TabIndex        =   24
         Top             =   645
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "翻譯人員："
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   23
         Top             =   288
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "交稿期限："
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   22
         Top             =   608
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "翻譯收文號："
         Height          =   255
         Index           =   8
         Left            =   3675
         TabIndex        =   21
         Top             =   975
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdOpenDir 
      Caption         =   "上傳檔案查詢"
      Height          =   380
      Left            =   2520
      TabIndex        =   11
      Top             =   720
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtPath 
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "txtPath"
      Top             =   960
      Width           =   4815
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "上傳檔案"
      Height          =   380
      Left            =   960
      TabIndex        =   10
      Top             =   720
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6840
      Top             =   600
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
            Picture         =   "frm060504.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060504.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060504.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060504.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060504.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060504.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060504.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060504.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060504.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060504.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060504.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   576
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   7572
      _ExtentX        =   13356
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
   Begin MSForms.Label Label2 
      Height          =   255
      Index           =   2
      Left            =   1080
      TabIndex        =   32
      Top             =   2670
      Width           =   1440
      BackColor       =   -2147483643
      Size            =   "2540;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   300
      Index           =   1
      Left            =   2100
      TabIndex        =   31
      Top             =   2040
      Width           =   960
      Size            =   "1693;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   30
      Top             =   2364
      Width           =   1440
      BackColor       =   -2147483643
      Size            =   "2540;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   2
      Left            =   1080
      TabIndex        =   2
      Top             =   2016
      Width           =   960
      VariousPropertyBits=   679495707
      MaxLength       =   6
      Size            =   "1693;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   1
      Left            =   1080
      TabIndex        =   1
      Top             =   1668
      Width           =   960
      VariousPropertyBits=   679495707
      MaxLength       =   7
      Size            =   "1693;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   0
      Left            =   1080
      TabIndex        =   0
      Top             =   1320
      Width           =   960
      VariousPropertyBits=   679495707
      MaxLength       =   5
      Size            =   "1693;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label23 
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   3000
      Width           =   7365
      Caption         =   "Create ID:           Date         Time             Update ID:                Date  "
      Size            =   "12991;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   1335
      Index           =   3
      Left            =   3240
      TabIndex        =   3
      Top             =   1320
      Width           =   4215
      VariousPropertyBits=   -1466939365
      MaxLength       =   200
      ScrollBars      =   2
      Size            =   "7435;2355"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   25
      Top             =   2700
      Width           =   945
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "預設上傳後檔案路徑："
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   2520
      TabIndex        =   19
      Top             =   720
      Width           =   1800
   End
   Begin VB.Label Label1 
      Caption         =   "收文號："
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   18
      Top             =   2355
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "備註："
      Height          =   255
      Index           =   5
      Left            =   2640
      TabIndex        =   17
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "追蹤號："
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   15
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "管制人："
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   14
      Top             =   2010
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "    期限："
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   1665
      Width           =   855
   End
End
Attribute VB_Name = "frm060504"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Modified by Lydia 2023/08/18 改用FTP(原始檔區)存放檔案
'Modified by Lydia 2023/05/23 外專新案認領：原本”107.03.09 TCN12上傳msg檔數量+TCN13收文讀取的msg檔數量”在2020/06/30 取消通知(因為使用者還是有人工複製檔案到TrackingNo, 造成5/4~6/30有19個TrackingNo資料夾沒有刪除)；112.5.22 改放"客戶有提供彩圖、外文本的對應英/中說"
'Modified by Lydia 2023/02/13 外專新案認領：增加TCN16~TCN23
'Memo by Lydia 2018/11/13 改成Form2.0 (Label2, Label23和Textbox)
'2013/06/24 Create by Amy -- 案件命名追蹤
Option Explicit

Private Const cArr As String = "&" 'Added by Lydia 2018/06/07 檔案的分隔符號
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
Dim ManageYN As Boolean '登入者是否為主管
Dim ActionEdit As Integer '0:add 1:update 2:query 3:cancel
Dim i As Integer

' 第一筆資料
Dim m_FirstKEY(1) As String
' 最後一筆資料
Dim m_LastKEY(1) As String
' 目前正在顯示
Dim m_CurrKEY(1) As String

'Modified by Lydia 2018/11/13  改成Form2.0
'Dim oText As TextBox, idx As Integer
Dim oText, txt, Lbl
Dim idx As Integer
'end 2018/11/13
Dim bolMsgRight As Boolean 'Added by Lydia 2018/11/21 Form 2.0表單是否彈過提示滑鼠右鍵無效
Dim SyxMsg As String 'Added by Lydia 2018/11/21 Form 2.0表單是否彈過提示滑鼠右鍵無效(記錄前一位置)

'Added by Lydia 2017/12/04 FCP案件命名電子化
Public m_strSaveFiles As String '上傳檔案路徑
Dim tmpArr As Variant
'Added by Lydia 2018/06/07
Dim m_PrevForm As Form '前一畫面
Dim m_PFkey As String '前一畫面-翻譯收文號
Dim staTF As String '急件翻譯狀態(1=新增,2=修改,4=刪除)
'Add By Sindy 2022/5/20
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Public m_RDate As String
Dim m_Done As Boolean
'2022/5/20 END
Dim m_UserSt16 As String 'Added by Lydia 2023/02/13 承辦人之組別

'Added by Lydia 2018/06/07 傳入前一畫面
Public Sub SetParent(ByRef pForm As Form, Optional ByVal pTNo As String = "")
     Set m_PrevForm = pForm
     m_PFkey = pTNo
End Sub

'Add By Sindy 2022/5/27
Private Sub Form_Activate()
   If m_strIR01 <> "" And m_Done = False Then
      '新增
      Text1(0).SetFocus
      RbEdit 0
      m_Done = True
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF2 'add
            If m_bInsert Then
                Text1(1).SetFocus
                Text1(0).TabStop = False
                RbEdit 0
            End If
        Case vbKeyF3 'update
            If m_bUpdate Then
                Text1(1).SetFocus
                Text1(0).TabStop = False
                RbEdit 1
            End If
        Case vbKeyF5  'delete
            If m_bDelete Then
                RbEdit 2
            End If
        
        Case vbKeyF4 'query
            If m_bQuery Then
                Text1(0).SetFocus
                RbEdit 5
            End If
        
        Case vbKeyHome
             If Not (ActionEdit = 0 Or ActionEdit = 1) Then
                ActionRb 0
             End If
        Case vbKeyPageUp
             If Not (ActionEdit = 0 Or ActionEdit = 1) Then
                 ActionRb 1
             End If
        Case vbKeyPageDown
             If Not (ActionEdit = 0 Or ActionEdit = 1) Then
                 ActionRb 2
             End If
        Case vbKeyEnd
             If Not (ActionEdit = 0 Or ActionEdit = 1) Then
                  ActionRb 3
             End If
        Case vbKeyF9 'ok
        '欄位驗證
        RbEdit 3
        '抓資料
        Case vbKeyF10 'cancel
         RbEdit 4
        Case vbKeyEscape
        Unload Me
        Set frm060504 = Nothing
End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'Enter 事件，等於存檔，做完取消，不然 form 內其他物件有寫 keycode 或是 keyascii 事件的話，也會做到Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
      Case vbKeyReturn:
         If ActionEdit <> 3 Then
            KeyAscii = 0
            Form_KeyDown vbKeyF9, 0
         End If
    End Select
End Sub

Private Sub Form_Load()

 '取得使用者執行各項功能的權限
  m_bInsert = IsUserHasRightOfFunction("frm060504", strAdd, False)
  m_bUpdate = IsUserHasRightOfFunction("frm060504", strEdit, False)
  m_bDelete = IsUserHasRightOfFunction("frm060504", strDel, False)
  m_bQuery = IsUserHasRightOfFunction("frm060504", strFind, False)
  
   ManageYN = CheckManage()
   MoveFormToCenter Me
   RefreshRange
   GetFirstRecordVal     '設定第一筆key值
   ToolBarSet 1             '設定ToolBar按鈕顯示
   ActionEdit = 3           'cancel/第一次進入
   
   'Added by Lydia 2018/06/07
   If cboSource.ListCount = 0 Then
       Call SetCombList
   End If
   
   'Added by Lydia 2023/02/13
   If strSrvDate(1) < 外專新案認領啟用日 Then
       Frame2.Visible = False
       Me.Height = 5565
   Else
      Frame2.Visible = True
   End If
   m_UserSt16 = PUB_GetStaffST16(strUserNum)
   If m_UserSt16 = "" Then
      m_UserSt16 = "1" '預設英文組
   End If
   If m_UserSt16 = "1" Then
       Label1(12).Visible = False
       Text1(9).Visible = False
       'Added by Lydia 2023/05/23
       Label1(15).Visible = False
       Text1(10).Visible = False
   Else
       Label1(12).Visible = True
       Text1(9).Visible = True
       'Added by Lydia 2023/05/23 日文組承辦增加「客戶有提供彩圖TCN12」欄位輸入，英文組則隱藏此欄位
       Label1(15).Visible = True
       Text1(10).Visible = True
   End If
   'end 2023/02/13
   FraTCN13.BackColor = &H8000000F   'Added by Lydia 2023/06/14
   
   'Added by Lydia 2023/08/24
   If Pub_StrUserSt03 <> "M51" Then
      cmdDelFTP.Visible = False
   End If
   
End Sub

Private Sub ActionRb(ByVal stA As Integer)
   TxtLock 2
      Select Case stA
         Case 0 'MoveFirst
           GetFirstRecordVal
         Case 1 'MovePrv
            GetPreRecordVal
         Case 2 'MoveNext
            GetNextRecordVal
         Case 3 'MoveLast
            GetLastRecordVal
      End Select
End Sub

Private Sub SetTxtValue()
Dim strTmp As String, m_ibf01 As String, m_ibf02 As String
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   'Added by Lydia 2018/06/07
   If cboSource.ListCount = 0 Then
       Call SetCombList
   End If
   
   'Modified by Lydia 2018/06/07 +急件翻譯
   'strSql = "SELECT TCN01, TCN02,TCN03,TCN04,NVL(TCN05,'NULL') TCN05,TCN06,TCN07,TCN08,TCN09,TCN10,TCN11,NVL(ST02,'NULL') As ST02 FROM TrackingCaseName ,STAFF " & _
                "WHERE TCN03=ST01(+) And TCN01='" & m_CurrKEY(0) & "' ORDER BY TCN01"
   'Modified by Lydia 2018/12/04 +tf32
   'Added by Lydia 2023/02/13 +tcn16~20,cp10,cp27
   If strSrvDate(1) >= 外專新案認領啟用日 Then
      'Modified by Lydia 2023/05/23 +TCN12,TCN13
      strSql = "select tcn01, tcn02,tcn03,tcn04,nvl(tcn05,'NULL') tcn05,tcn06,tcn07,tcn08,tcn09,tcn10,tcn11,TCN12,TCN13," & _
                   "nvl(s1.st02,'NULL') as st02,cp01,cp02,cp03,cp04,tcn14,tcn15,s2.st02 tcn15n,tf26,tf27,tf28,tf32 " & _
                   ",tcn16,tcn17,tcn18,tcn19,tcn20,cp10,cp27 " & _
                   "from TrackingCaseName,CaseProgress ,staff s1,staff s2,TransFee " & _
                   "where tcn01='" & m_CurrKEY(0) & "'  and tcn05=cp09(+) " & _
                   "and tcn03=s1.st01(+) and tcn15=s2.st01(+) and tcn14=tf01(+) " & _
                   "order by tcn01"
   Else
      strSql = "select tcn01, tcn02,tcn03,tcn04,nvl(tcn05,'NULL') tcn05,tcn06,tcn07,tcn08,tcn09,tcn10,tcn11," & _
                   "nvl(s1.st02,'NULL') as st02,cp01,cp02,cp03,cp04,tcn14,tcn15,s2.st02 tcn15n,tf26,tf27,tf28,tf32 " & _
                   "from TrackingCaseName,CaseProgress ,staff s1,staff s2,TransFee " & _
                   "where tcn01='" & m_CurrKEY(0) & "'  and tcn05=cp09(+) " & _
                   "and tcn03=s1.st01(+) and tcn15=s2.st01(+) and tcn14=tf01(+) " & _
                   "order by tcn01"
   End If
   TxtClear 'Added by Lydia 2017/12/27 查詢前先清資料
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
        'Modified by Lydia 2018/06/07
'        For Each oText In Text1
'          idx = oText.Index
'          If IsNull(rsTmp(idx)) Then
'             oText = ""
'          Else
'            Select Case idx
'               Case 1
'                oText = ChangeWStringToTString(rsTmp(idx))
'               Case 0, 2, 3
'                oText = rsTmp(idx)
'            End Select
'          End If
'        Next
        Text1(0).Text = "" & rsTmp.Fields("tcn01") '追蹤號
        Text1(1).Text = ChangeWStringToTString("" & rsTmp.Fields("tcn02")) '期限
        Text1(2).Text = "" & rsTmp.Fields("tcn03") '管制人
        Text1(3).Text = "" & rsTmp.Fields("tcn04") '備註
        Text1(4).Text = "" & rsTmp.Fields("tcn15") '翻譯人員
        If Text1(4).Text <> "" Then
            cboTCN15.Text = rsTmp.Fields("tcn15") & " " & rsTmp.Fields("tcn15n")
        End If
        Text1(5).Text = ChangeWStringToTString("" & rsTmp.Fields("tf26")) '交稿期限
        Text1(6).Text = ChangeWStringToTString("" & rsTmp.Fields("tf32")) 'Added by Lydia 2018/12/04 只交Claims期限
        '本所案號
        Label2(2).Caption = IIf("" & rsTmp.Fields("cp01") <> "", rsTmp.Fields("cp01") & "-" & rsTmp.Fields("cp02") & "-" & rsTmp.Fields("cp03") & "-" & rsTmp.Fields("cp04"), "")
        Label2(4).Caption = "" & rsTmp.Fields("tcn14") '翻譯收文號

        '原文語種
        cboSource.Tag = "" & rsTmp.Fields("tf27")
        If "" & rsTmp.Fields("tf27") <> "" Then
             cboSource.ListIndex = Val("" & rsTmp.Fields("tf27")) - 1
        End If
        '翻譯語種
        cboTarget.Tag = "" & rsTmp.Fields("tf28")
        If "" & rsTmp.Fields("tf28") <> "" Then
             cboTarget.ListIndex = Val("" & rsTmp.Fields("tf28")) - 1
        End If
        
        'Added by Lydia 2023/02/13
        If strSrvDate(1) >= 外專新案認領啟用日 Then
            Text1(7) = "" & rsTmp.Fields("tcn16") '暫不認領
            Text1(8) = "" & rsTmp.Fields("tcn17")  '有相似案
            '指定組別(相似舊案)
            If "" & rsTmp.Fields("tcn18") <> "" Then
                Combo1.Text = rsTmp.Fields("tcn18") & "." & PUB_GetFCPGrpName(rsTmp.Fields("tcn18"))
            Else
                Combo1.Text = ""
            End If
            Combo1.Tag = Combo1.Text
            Text1(9) = "" & rsTmp.Fields("tcn19") '英文組認領
            Text1(10) = "" & rsTmp.Fields("tcn12") 'Added by Lydia 2023/05/23 客戶有提供彩圖
            Text1(11) = "" & rsTmp.Fields("tcn13") 'Added by Lydia 2023/06/14 外文本的對應英/中說
        End If
        'end 2023/02/13
        
        For Each oText In Text1
            oText.Tag = oText.Text
        Next
        'end 2018/06/07
        'Added by Lydia 2023/06/14 外文本的對應英/中說
        If Trim(Text1(11)) <> "" Then
           Select Case Text1(11)
              Case "0", "4" '0=無,4=(客戶提供文件)確定無文件
                  chk1(0).Value = 1
              Case "1", "3" '1=有,3=(客戶提供文件)確定已收文件
                  chk1(1).Value = 1
              Case "2"
                  chk1(2).Value = 1
           End Select
        End If
        'end 2023/06/14
        
        If "" & rsTmp.Fields("ST02") = "" Then  '管制人
            MsgBox ("管制人有誤")
            Label2(1).Caption = ""
        Else
            Label2(1).Caption = rsTmp.Fields("ST02")
        End If
        
        TBar1.Buttons(3).Enabled = False 'Added by Lydia 2018/06/07
        If rsTmp.Fields("TCN05") = "NULL" Then  '收文號
            TBar1.Buttons(2).Enabled = True
            'Modified by Lydia 2018/06/07 權限控制
            'TBar1.Buttons(3).Enabled = True
            If m_bDelete = True Then TBar1.Buttons(3).Enabled = True
            Label2(0).Caption = ""
        Else
            Label2(0).Caption = rsTmp.Fields("TCN05")
            TBar1.Buttons(2).Enabled = False
            TBar1.Buttons(3).Enabled = False
        End If

       'Added by Lydia 2017/12/04 FCP案件命名電子化
        CmdFile.Visible = False
        'Add By Sindy 2024/4/24 系統收件區會傳電子檔名進來
        If m_strIR01 = "" Then
        '2024/4/24 END
            m_strSaveFiles = ""
        End If
        'end 2017/12/04
        cmdOpenDir.Visible = False 'Added by Lydia 2018/02/23
        
       'Added by Lydia 2023/08/18  改用FTP(原始檔區)存放檔案
       cmdAddFile.Visible = False
       cmdAddFile.Top = CmdFile.Top
       cmdAddFile.Left = CmdFile.Left
      'Added by Lydia 2017/12/27 有資料才顯示，上傳檔案的資料夾--->Typing2\English_Vers
      'Mark by Lydia 2023/10/12 改用原始檔區存放
      'Memo by Lydia 2024/12/13 刪除舊Code
          If Trim(Label2(0).Caption) = "" Then
             cmdAddFile.Visible = True
             If Trim(Text1(0)) <> "" Then
                If PUB_ChkTCNfileExist(Text1(0)) = False Then
                   MsgBox "請上傳相關檔案！", vbExclamation, "上傳檔案稽核"
                End If
             End If
          Else
             cmdAddFile.Visible = False
          End If
      
   'Mark by Lydia 2017/12/27 查詢前先清資料
   'Else
  '      TxtClear
   End If
   '更新CUID
   UpdateCUID rsTmp
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Added by Lydia 2018/06/07 跳回前一畫面
   If UCase(TypeName(m_PrevForm)) <> "NOTHING" Then
      If TypeName(m_PrevForm) = "frm060122" Then
            m_PrevForm.cmdState = 0
            Call m_PrevForm.PubShowNextData
      'Add By Sindy 2022/5/20
      ElseIf UCase(TypeName(m_PrevForm)) = UCase("frm06010616") Then
         Set m_PrevForm = Nothing
         GoTo gotoExit
         '2022/5/20 END
      End If
      m_PrevForm.Show
   End If
   'end 2018/06/07
   
gotoExit:
   Set frm060504 = Nothing
End Sub

Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
      Case 1 'Add
         Text1(0).SetFocus
         RbEdit 0
      Case 2 'Update
         Text1(1).SetFocus
         Text1(0).TabStop = False
         RbEdit 1
      Case 3 'Del
         RbEdit 2
      Case 4 'query
        TxtClear
        RbEdit 5
      Case 6 'MoveFirst
         ActionRb 0
      Case 7 'MovePrv
         ActionRb 1
      Case 8 'MoveNext
         ActionRb 2
      Case 9 'MoveLast
         ActionRb 3
      Case 11 'OK
        RbEdit 3
        Text1(0).TabStop = True
      Case 12 'Cancel
        RbEdit 4
      Case 14 'Exit
        Unload Me
        Set frm060504 = Nothing
   End Select
End Sub

Private Sub RbEdit(stA As Integer)
Dim StrSQLa, Str01 As String, nTime As String
Dim SeqNo As Integer
Dim NowTime As Long
'Added by Lydia 2017/12/04
Dim bolSaveMsg As Boolean '是否上傳msg檔
Dim strFilePath As String '上傳檔案存放路徑
Dim strSFName As String
Dim Str02 As String 'Added by Lydia 2018/03/06
Dim fs, f 'Added by Lydia 2023/08/18

nTime = ServerTime
NowTime = IIf(Len(nTime) = 6, Left(nTime, 4), Left(nTime, 3))
    Select Case stA
      Case 0 'add
         Screen.MousePointer = vbHourglass 'Added by Lydia 2023/08/18
         TxtClear
         ToolBarSet 0
         ActionEdit = 0
         'Added by Lydia 2017/12/04 FCP案件命名電子化
         If strSrvDate(1) >= FCP案件命名啟用日 Then
            '命名追蹤期限改預設3個工作天(含當日)
            Text1(1).Text = TransDate(CompWorkDay(3, strSrvDate(1)), 1)
         End If
         'end 2017/12/04
         'Added by Lydia 2023/02/13
         If Frame2.Visible = True Then
             If m_UserSt16 = "1" Then '英文組預設欄位隱藏
                 Text1(9) = "Y"
             End If
         End If
         'end 2023/02/13
         Screen.MousePointer = vbDefault 'Added by Lydia 2023/08/18
         Text1(1).SetFocus
         TextInverse Text1(1)
         'Text1(0).Appearance = 0 'Remove by Lydia 2018/11/13 Form2.0不支援
         Text1(0).BorderStyle = 0
         If Pub_StrUserSt03 <> "M51" Then Text1(2).Text = strUserNum: Label2(1).Caption = strUserName
         cmdAddFile.Visible = True
         
      Case 1 'update
         Screen.MousePointer = vbHourglass 'Added by Lydia 2023/08/18
         ToolBarSet 0
         ActionEdit = 1

         Screen.MousePointer = vbDefault 'Added by Lydia 2023/08/18
         'end 2018/03/29
         Text1(0).Locked = True
         Text1(1).SetFocus
         TextInverse Text1(1)
         'Text1(0).Appearance = 0 'Remove by Lydia 2018/11/13 Form2.0不支援
         Text1(0).BorderStyle = 0
         
      Case 2 'delete
         'Added by Lydia 2018/06/07
         If Label2(0).Caption <> "" Then
              MsgBox "追蹤號已收文，不可刪除!", vbCritical
              Exit Sub
         End If
         'end 2018/06/07
         If MsgBox("是否要刪除此筆資料?", vbCritical + vbYesNo + vbDefaultButton2, "詢問") = vbYes Then
            Screen.MousePointer = vbHourglass 'Added by Lydia 2023/08/18
            If DelRecord(NowTime) = True Then
                RefreshRange
            Else
                Exit Sub
            End If
         End If
         Screen.MousePointer = vbDefault 'Added by Lydia 2023/08/18
      Case 3 'ok
        If ActionEdit = 0 Then  '在新增狀態按Enter鍵
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Sub
            
            Screen.MousePointer = vbHourglass 'Added by Lydia 2023/08/18
            SeqNo = GetSerialNo(Me.Text1(0).Text)
            'Added by Lydia 2017/12/27
            On Error GoTo ErrorHandleAdd
            cnnConnection.BeginTrans
            'end 2017/12/27
            'Modified by Lydia 2018/06/07 +翻譯收文號TCN14、翻譯人員TCN15
            'Modified by Lydia 2023/05/23 +客戶有提供彩圖TCN12
            'Modified by Lydia 2033/05/29 外文本的對應英/中說TCN13
            'StrSQLa = "Insert Into TrackingCaseName (TCN01, TCN02, TCN03, TCN04, TCN06, TCN07, TCN08,TCN14,TCN15) " & _
                            "Values('" & SeqNo & "','" & ChangeTStringToWString(Me.Text1(1)) & "','" & Me.Text1(2) & "','" & ChgSQL(Me.Text1(3)) & "','" & strUserNum & "','" & strSrvDate(1) & "','" & NowTime & "' " & _
                            IIf(staTF = "", ",null,null", ",'" & SeqNo & "','" & Text1(4).Text & "' ") & ")"
            StrSQLa = "Insert Into TrackingCaseName (TCN01, TCN02, TCN03, TCN04, TCN06, TCN07, TCN08,TCN14,TCN15,TCN12,TCN13) " & _
                            "Values('" & SeqNo & "','" & ChangeTStringToWString(Me.Text1(1)) & "','" & Me.Text1(2) & "','" & ChgSQL(Me.Text1(3)) & "','" & strUserNum & "','" & strSrvDate(1) & "','" & NowTime & "' " & _
                            IIf(staTF = "", ",null,null", ",'" & SeqNo & "','" & Text1(4).Text & "' ") & ", '" & Me.Text1(10) & "','" & Me.Text1(11) & "' )"
            cnnConnection.Execute StrSQLa
            'Added by Lydia 2018/06/07 新增-翻譯費用檔
            If staTF = "1" Then
               'Modified by Lydia 2018/12/04 + 只交Claims期限TF32
               StrSQLa = "insert into TransFee (TF01,TF26,TF27,TF28,TF32) values ('" & SeqNo & "', " & CNULL(DBDATE(Text1(5).Text), True) & ", " & CNULL(Left(cboSource.Text, 1)) & ", " & CNULL(Left(cboTarget.Text, 1)) & ", " & CNULL(DBDATE(Text1(6).Text), True) & ") "
               cnnConnection.Execute StrSQLa
            End If
            'end 2018/06/07
            
            Text1(0).Text = Format(Val(SeqNo), "00000")
            'Added by Lydai 2023/08/18 上傳檔案：改用FTP(原始檔區)存放檔案
            'Memo by Lydia 2024/12/13 刪除舊Code
            If cmdAddFile.Visible = True And ActionEdit = 0 And m_strSaveFiles <> "" Then
               tmpArr = Empty
               tmpArr = Split(m_strSaveFiles, cArr)
               For intI = 0 To UBound(tmpArr)
                  If Trim(tmpArr(intI)) <> "" Then
                     strSFName = Trim(tmpArr(intI))
                     If InStrRev(Trim(tmpArr(intI)), " (") > 0 Then
                        strSFName = Left(Trim(tmpArr(intI)), InStrRev(Trim(tmpArr(intI)), " (") - 1)
                     End If
                     Str01 = Dir(strSFName)
                     If Str01 <> "" Then
                        Str02 = PUB_GetSimpleName(Str01, , True)
                        Sleep 1000
                        Set fs = CreateObject("Scripting.FileSystemObject")
                        Set f = fs.GetFile(strSFName)
                        If SaveAttFile_Org(Text1(0), strSFName, Str02, Format(f.DateLastModified, "YYYYMMDD"), Format(f.DateLastModified, "HHMMSS"), "A") = False Then
                           GoTo ErrorHandleAdd
                        End If
                     Else
                        MsgBox "查無此檔案: " & strSFName
                        DoEvents
                     End If
                  End If
               Next intI
               Set fs = Nothing
               Set f = Nothing
            End If 'Added by Lydia 2023/08/18
            
            'Add by Sindy 2022/5/20
            If m_strIR01 <> "" Then
               PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm060504", , , "追蹤號=" & SeqNo
            End If
            '2022/5/20 END
            
            'Added by Lydia 2023/02/13
            If Frame2.Visible = True And Frame2.Enabled = True Then
                StrSQLa = "update TrackingCaseName set TCN16=" & CNULL(Text1(7)) & ",TCN17=" & CNULL(Text1(8)) & _
                    " , TCN18=" & CNULL(Left(Combo1, 1)) & " , TCN19=" & CNULL(Text1(9)) & " where TCN01='" & SeqNo & "' "
                cnnConnection.Execute StrSQLa
            End If
            'end 2023/02/13
            
             cnnConnection.CommitTrans 'Added by Lydia 2017/12/27
             
             'Added by Lydia 2018/06/05 新增或刪除急件翻譯,自動發mail
             If m_PFkey = "" And staTF = "1" Then
                 If ProcEmail(staTF, Text1(0)) = False Then
                     MsgBox "取消送信，急件翻譯欄位清空 !", vbCritical
                     'Added by Lydia 2018/08/10 與承辦開會後,決定若取消送信就清空欄位
                     cnnConnection.BeginTrans
                          StrSQLa = "delete from transfee where tf01='" & Text1(0) & "' "
                          cnnConnection.Execute StrSQLa, intI
                          StrSQLa = "update TrackingCaseName set TCN14=null , TCN15=null where TCN01='" & Text1(0) & "' "
                          cnnConnection.Execute StrSQLa
                     cnnConnection.CommitTrans
                     'end 2018/08/10
                 End If
             End If
             'end 2018/06/05
             
             GetCurrRecordVal SeqNo
             
            'Add By Sindy 2022/5/20
            If Me.m_strIR01 <> "" Then
               If Not m_PrevForm Is Nothing Then
                  Call m_PrevForm.GoNext
               End If
               'Modify By Sindy 2022/8/29 Mark
               'Unload Me
               'Exit Sub
               '2022/8/29 END
            End If
            '2022/5/20 END
             
         ElseIf ActionEdit = 1 Then '在修改狀態按Enter鍵
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Sub
            Screen.MousePointer = vbHourglass 'Added by Lydia 2023/08/18
             'Added by Lydia 2018/03/29
             On Error GoTo ErrorHandleAdd
             cnnConnection.BeginTrans
             'end 2018/03/29
             'Modified by Lydia 2018/06/07 +翻譯人員
             'StrSQLa = "Update TrackingCaseName set TCN02='" & ChangeTStringToWString(Text1(1)) & "',TCN03= '" & Me.Text1(2).Text & "',TCN04='" & ChgSQL(Text1(3)) & "', TCN09='" & strUserNum & "', TCN10='" & strSrvDate(1) & "', TCN11='" & NowTime & "' " & _
                              "Where TCN01=" & Val(Text1(0))
             strExc(1) = ""
             If Text1(1).Text <> Text1(1).Tag Then
                 strExc(1) = strExc(1) & ", TCN02='" & ChangeTStringToWString(Text1(1).Text) & "'"
             End If
             If Text1(2).Text <> Text1(2).Tag Then
                 strExc(1) = strExc(1) & ", TCN03='" & Text1(2).Text & "'"
             End If
             If Text1(3).Text <> Text1(3).Tag Then
                 strExc(1) = strExc(1) & ", TCN04='" & ChgSQL(Text1(3).Text) & "'"
             End If
             If staTF = "1" Then
                  strExc(1) = strExc(1) & ", TCN14=" & CNULL(Text1(0).Text)
             ElseIf staTF = "4" Then
                  strExc(1) = strExc(1) & ", TCN14=NULL, TCN15=NULL"
             End If
             If Text1(4).Text <> Text1(4).Tag And staTF <> "4" Then
                 strExc(1) = strExc(1) & ", TCN15=" & CNULL(Text1(4).Text)
             End If
             'Added by Lydia 2023/05/23
             If Text1(10).Text <> Text1(10).Tag Then '客戶有提供彩圖TCN12
                 strExc(1) = strExc(1) & ", TCN12='" & ChgSQL(Text1(10).Text) & "'"
             End If
             'Added by Lydia 2023/06/14
             If Text1(11).Text <> Text1(11).Tag Then '外文本的對應英/中說TCN13
                 strExc(1) = strExc(1) & ", TCN13='" & ChgSQL(Text1(11).Text) & "'"
             End If
             
             StrSQLa = "Update TrackingCaseName set TCN09='" & strUserNum & "', TCN10='" & strSrvDate(1) & "', TCN11='" & NowTime & "' " & strExc(1) & " Where TCN01=" & Val(Text1(0))
             If m_PFkey <> "" Then Pub_SeekTbLog StrSQLa  'Sharon修改翻譯人員，留記錄
             'end 2018/06/07
             cnnConnection.Execute StrSQLa
             
             'Added by Lydia 2018/06/07
             If staTF = "4" Then '刪除-翻譯費用檔
                  StrSQLa = "Delete from TransFee where TF01=" & CNULL(Label2(4).Caption)
                  cnnConnection.Execute StrSQLa
             ElseIf staTF = "1" Then '新增-翻譯費用檔
                  'Modified by Lydia 2018/12/04 只交Claims期限TF32
                  StrSQLa = "insert into TransFee (TF01,TF26,TF27,TF28,TF32) values ('" & Val(Text1(0)) & "', " & CNULL(DBDATE(Text1(5).Text), True) & ", " & CNULL(Left(cboSource.Text, 1)) & ", " & CNULL(Left(cboTarget.Text, 1)) & ", " & CNULL(DBDATE(Text1(6).Text), True) & " ) "
                  cnnConnection.Execute StrSQLa
             Else '修改-翻譯費用檔
                  StrSQLa = ""
                  If Text1(5).Text <> Text1(5).Tag Then
                       StrSQLa = StrSQLa & ", TF26=" & CNULL(DBDATE(Text1(5).Text))
                  End If
                  'Added by Lydia 2018/12/04 只交Claims期限TF32
                  If Text1(6).Text <> Text1(6).Tag Then
                       StrSQLa = StrSQLa & ", TF32=" & CNULL(DBDATE(Text1(6).Text))
                  End If
                  'end 2018/12/04
                  If cboSource.Tag <> Left(cboSource.Text, 1) Then
                       StrSQLa = StrSQLa & ", TF27=" & CNULL(Left(cboSource.Text, 1))
                  End If
                  If cboTarget.Tag <> Left(cboTarget.Text, 1) Then
                       StrSQLa = StrSQLa & ", TF28=" & CNULL(Left(cboTarget.Text, 1))
                  End If
                  If StrSQLa <> "" Then
                       StrSQLa = "Update TransFee set " & Mid(StrSQLa, 2) & " where TF01=" & CNULL(Label2(4).Caption)
                       cnnConnection.Execute StrSQLa, intI
                  End If
             End If
             'end 2018/06/07
           
            'Added by Lydia 2023/02/13
            If Frame2.Visible = True And Frame2.Enabled = True Then
                StrSQLa = ""
                '暫不認領
                If Text1(7).Text <> Text1(7).Tag Then
                   StrSQLa = StrSQLa & ", TCN16=" & CNULL(Text1(7))
                   If Text1(7).Text = "Y" And Label2(0).Caption <> "" Then
                       StrSQLa = StrSQLa & ", TCN21=99999999 "
                   End If
                End If
                '相似舊案 & 指定組別
                If Text1(8).Text <> Text1(8).Tag Or Combo1.Text <> Combo1.Tag Then
                   StrSQLa = StrSQLa & ", TCN17=" & CNULL(Text1(8))
                   StrSQLa = StrSQLa & ", TCN18=" & CNULL(Left(Combo1, 1)) & " "
                End If
                '英文組認領
                If Text1(9).Text <> Text1(9).Tag Then
                   StrSQLa = StrSQLa & ", TCN19=" & CNULL(Text1(9))
                End If
                If StrSQLa <> "" Then
                   StrSQLa = "update TrackingCaseName set " & Mid(StrSQLa, 2) & " where TCN01='" & Val(Text1(0)) & "' "
                    cnnConnection.Execute StrSQLa
                End If
            End If
            'end 2023/02/13
             cnnConnection.CommitTrans
             'end 2018/03/29
             
             'Added by Lydia 2018/06/05 新增或刪除急件翻譯,自動發mail
             If m_PFkey = "" And (staTF = "1" Or staTF = "4") Then
                 If ProcEmail(staTF, Text1(0)) = False Then
                     MsgBox "取消送信，急件翻譯欄位清空 !", vbCritical
                     'Added by Lydia 2018/08/10 與承辦開會後,決定若取消送信就清空欄位
                     If staTF = "1" Then
                            cnnConnection.BeginTrans
                                 StrSQLa = "delete from transfee where tf01='" & Text1(0) & "' "
                                 cnnConnection.Execute StrSQLa, intI
                                 StrSQLa = "update TrackingCaseName set TCN14=null , TCN15=null where TCN01='" & Text1(0) & "' "
                                 cnnConnection.Execute StrSQLa, intI
                            cnnConnection.CommitTrans
                     End If
                     'end 2018/08/10
                 End If
             End If
             'end 2018/06/05
             RefreshRange
             SetTxtValue
             
         ElseIf ActionEdit = 2 Then '在查詢狀態按Enter鍵
            If Len(Trim(Text1(0))) > 0 Then
                QueryRecord Text1(0)
            ElseIf Len(Trim(Text1(0))) = 0 Then
                MsgBox ("請輸入追蹤號")
                Text1(0).SetFocus
                Exit Sub
            End If
            
         End If
         Screen.MousePointer = vbDefault 'Added by Lydia 2023/08/18
         ToolBarSet 1
         ActionEdit = 3
         'Text1(0).Appearance = 1 'Remove by Lydia 2018/11/13 Form2.0不支援
         Text1(0).BorderStyle = 1
      Case 4 'cancel
         If ActionEdit <> 2 Then
            If MsgBox("並未存檔,確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2) = vbYes Then
               If ActionEdit = 0 Then
                  ActionRb 3
               ElseIf ActionEdit = 1 Then
                  SetTxtValue
               End If
               ToolBarSet 1
               ActionEdit = 3 'cancel
               'SetTxtValue 'Mark by Lydia 2023/08/18 重複抓值
               'Text1(0).Appearance = 1 'Remove by Lydia 2018/11/13 Form2.0不支援
               Text1(0).BorderStyle = 1
            Else
               Exit Sub
            End If
            Text1(0).SetFocus
            
            'Add By Sindy 2022/6/28
            If Me.m_strIR01 <> "" Then
               Unload Me
               Exit Sub
            End If
            '2022/6/28 END
         Else
            ToolBarSet 1
            ActionEdit = 3 'cancel
            SetTxtValue
            'Text1(0).Appearance = 1 'Remove by Lydia 2018/11/13 Form2.0不支援
            Text1(0).BorderStyle = 1
         End If
         
       Case 5 'query
         ToolBarSet 0
         TxtLock 3
         ActionEdit = 2
         Text1(0).Locked = False
         Text1(0).SetFocus
         'Text1(0).Appearance = 1 'Remove by Lydia 2018/11/13 Form2.0不支援
         Text1(0).BorderStyle = 1
   End Select

'Added by Lydia 2017/12/27
   Exit Sub
   
ErrorHandleAdd:
   If Err.Number <> "" Then
       Screen.MousePointer = vbDefault 'Added by Lydia 2023/08/18
       MsgBox Err.Description
       'Modified by Lydia 2018/03/29
       'If sta = 3 And ActionEdit = 0 Then
       If stA = 3 And (ActionEdit = 0 Or ActionEdit = 1) Then
           cnnConnection.RollbackTrans
       End If
   End If
End Sub

Public Function CheckManage() As Boolean
 CheckManage = False
  strExc(0) = "Select count(*) Manage From Staff " & _
                      "WHERE ST52= '" & strUserNum & "' OR ST53='" & strUserNum & "' " & _
                           "OR ST54 = '" & strUserNum & "' OR ST55='" & strUserNum & "' "
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If Val("" & RsTemp.Fields("Manage")) > 0 Then
        CheckManage = True
      End If
   End If
End Function

Private Function TxtValidate() As Boolean
Dim Cancel As Boolean
'Dim txt As TextBox  'Remove by Lydia 2018/11/13 改成Form2.0
Dim strFileName As String 'Added by Lydia 2017/12/04
'Added by Lydia 2018/06/07
Dim bolAttPdf As Boolean '是否有外文本
Dim strFTlist As String
'end 2018/06/07

  TxtValidate = False
  
  For Each txt In Text1
    Text1_Validate txt.Index, Cancel
    If Cancel = True Then
        Exit Function
     End If
  Next
        
  'Added by Lydai 2023/08/18 上傳檔案：改用FTP(原始檔區)存放檔案
  If cmdAddFile.Visible = True And (ActionEdit = 0 Or ActionEdit = 1) Then
     If m_strSaveFiles = "" And ActionEdit = 0 Then
        MsgBox "請上傳相關檔案 !", vbExclamation
        Cancel = True
        Exit Function
     ElseIf m_strSaveFiles <> "" And m_strSaveFiles <> "Y" Then
         tmpArr = Empty
         tmpArr = Split(m_strSaveFiles, cArr)
         Cancel = True
         For intI = 0 To UBound(tmpArr)
            If Trim(tmpArr(intI)) <> "" Then
               strFileName = Trim(tmpArr(intI))
               '檢查檔名是否為英數字
               strExc(1) = Mid(strFileName, InStrRev(strFileName, "\") + 1)
               If InStr(strExc(1), " (") > 0 Then strExc(1) = Mid(strExc(1), 1, InStr(strExc(1), " (") - 1)
               strExc(1) = PUB_GetSimpleName(strExc(1), , True)
               If Trim(strExc(1)) = "" Or Mid(strExc(1), 1, 1) = "." Then
                    MsgBox "檔名請以英數字命名!"
                    Exit Function
               End If
               If InStrRev(Trim(tmpArr(intI)), " (") > 0 Then
                  strFileName = Left(Trim(tmpArr(intI)), InStrRev(Trim(tmpArr(intI)), " (") - 1)
               End If
               If Right(UCase(strFileName), 4) = UCase(FcpTcnFKey01) Then
                  Cancel = False
               End If
            End If
         Next intI
         If Cancel = True Then
            If MsgBox("未上傳郵件檔(*" & FcpTcnFKey01 & ")，是否繼續存檔？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
                Exit Function
            End If
         End If
     End If
     If ActionEdit = 1 Then
        If PUB_ChkTCNfileExist(Text1(0)) = False Then
           MsgBox "請上傳相關檔案！", vbExclamation, "上傳檔案稽核"
           Exit Function
        End If
     End If
  Else
  'end 2023/08/18
      'Added by Lydia 2017/12/04 FCP案件命名電子化:檢查MSG檔
      'Modified by Lydia 2018/03/29 修改也可上傳檔案
      'If cmdFile.Visible = True And ActionEdit = 0 Then '新增
      'Mark by Lydia 2023/10/12 改用原始檔區存放
'      If cmdFile.Visible = True And (ActionEdit = 0 Or ActionEdit = 1) Then  '新增
'         'Added by Lydia 2018/03/29 檢查資料夾內的檔名
'         'Modified by Lydia 2018/06/07 另外抓外文本
'         'If ChkFileName(Text1(0).Text, strFileName) = False Then
'         If ChkFileName(Text1(0).Text, strFileName, FcpTcnFKey02, strFTlist) = False Then
'             Exit Function
'         End If
'         'end 2018/03/29
'         'Modified by Lydia 2018/03/29 + 現有資料夾內的檔名
'         'If m_strSaveFiles = "" Then
'         If m_strSaveFiles = "" And strFileName = "" Then
'            'Modified by Lydia 2018/03/01
'            'MsgBox "請上傳相關檔案 !", vbExclamation
'            'Modified by Lydia 2018/05/09 +FMP案
'            'If MsgBox("FCP案必須上傳相關檔案，請問是否為FMP案？", vbYesNo + vbInformation + vbDefaultButton2) = vbNo Then
'            MsgBox "請上傳相關檔案 !", vbExclamation
'                Cancel = True
'                Exit Function
'            'End If 'Mark by by Lydia 2018/05/09
'         'Modified by Lydia 2018/03/29 限新增時，彈訊息
'         'Else
'         ElseIf ActionEdit = 0 Then
'            tmpArr = Empty
'            tmpArr = Split(m_strSaveFiles, cArr)
'            Cancel = True
'            For intI = 0 To UBound(tmpArr)
'               If Trim(tmpArr(intI)) <> "" Then
'                  strFileName = Trim(tmpArr(intI))
'                  'Added by Lydia 2018/03/06 檢查檔名是否為英數字
'                  strExc(1) = Mid(strFileName, InStrRev(strFileName, "\") + 1)
'                  If InStr(strExc(1), " (") > 0 Then strExc(1) = Mid(strExc(1), 1, InStr(strExc(1), " (") - 1)
'                  'Modified by Lydia 2018/06/6 保留檔名中間空白
'                  'strExc(1) = PUB_GetSimpleName(strExc(1))
'                  strExc(1) = PUB_GetSimpleName(strExc(1), , True)
'                  If Trim(strExc(1)) = "" Or Mid(strExc(1), 1, 1) = "." Then
'                       MsgBox "檔名請以英數字命名!"
'                       Exit Function
'                  End If
'                  'end 2018/03/06
'                  If InStrRev(Trim(tmpArr(intI)), " (") > 0 Then
'                     strFileName = Left(Trim(tmpArr(intI)), InStrRev(Trim(tmpArr(intI)), " (") - 1)
'                  End If
'                  If Right(UCase(strFileName), 4) = UCase(FcpTcnFKey01) Then
'                     Cancel = False
'                  End If
'                  'Added by Lydia 2018/06/07 是否有外文本
'                  If Right(UCase(strFileName), Len(FcpTcnFKey02)) = UCase(FcpTcnFKey02) Then
'                      bolAttPdf = True
'                  End If
'                  'end 2018/06/07
'               End If
'            Next intI
'            If Cancel = True Then
'               'Modified by Lydia 2017/12/27 不限msg檔案
'               'MsgBox "郵件檔(*" & FcpTcnFKey01 & ")為必要檔案，請上傳檔案 !", vbExclamation
'               If MsgBox("未上傳郵件檔(*" & FcpTcnFKey01 & ")，是否繼續存檔？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
'                   Exit Function
'               End If
'               'end 2017/12/27
'            End If
'         'Added by Lydia 2018/06/07 修改時檢查是否有外文本
'         ElseIf ActionEdit = 1 Then
'            If strFTlist <> "" Then  '已上傳
'                    bolAttPdf = True
'            ElseIf m_strSaveFiles <> "" Then '預定上傳檔案
'                    tmpArr = Empty
'                    tmpArr = Split(m_strSaveFiles, cArr)
'                    For intI = 0 To UBound(tmpArr)
'                       If Trim(tmpArr(intI)) <> "" Then
'                           strExc(1) = tmpArr(intI)
'                           If InStr(strExc(1), " (") > 0 Then strExc(1) = Mid(strExc(1), 1, InStr(strExc(1), " (") - 1)
'                           If UCase(Right(strExc(1), Len(FcpTcnFKey02))) = UCase(FcpTcnFKey02) Then
'                               bolAttPdf = True
'                               strFTlist = strFTlist & cArr & strExc(1)
'                           End If
'                       End If
'                    Next
'            End If
'         'end 2018/06/07
'         End If
'      End If
      'end 2017/12/04
      'end 2023/10/12
  End If 'Added by Lydia 2023/08/18
  
  'Added by Lydia 2018/06/07 若追蹤號已經收文立案，不可輸入急件翻譯。
  staTF = ""
  'Modified by Lydia 2018/12/04
  'If Trim(Text1(4).Text & Text1(5).Text & cboSource.Text & cboTarget.Text) <> "" Then
       'If Label2(0).Caption <> "" And Trim(Text1(4).Tag & Text1(5).Tag & cboTarget.Tag) = "" Then
  If Trim(Text1(4).Text & Text1(5).Text & Text1(6).Text & cboSource.Text & cboTarget.Text) <> "" Then
       If Label2(0).Caption <> "" And Trim(Text1(4).Tag & Text1(5).Tag & Text1(6).Tag & cboTarget.Tag) = "" Then
            MsgBox "追蹤號已經收文立案，不可輸入急件翻譯 !", vbCritical
            Exit Function
       End If
       'Modified by Lydia 2018/12/04
       'If Trim(Text1(4).Text & Text1(5).Text) = "" Then
       '     MsgBox "急件翻譯請輸入翻譯人員或交稿期限 !", vbCritical
       If Trim(Text1(4).Text & Text1(5).Text & Text1(6).Text) = "" Then
            MsgBox "急件翻譯請輸入翻譯人員、交稿期限或只交Claims期限 !", vbCritical
            Exit Function
       End If
       'Added by Lydia 2018/12/04
       If Text1(5) <> "" And Text1(6) <> "" And Text1(6) > Text1(5) Then
            MsgBox "只交Claims期限不可大於交稿期限 !", vbCritical
            Exit Function
       End If
       'end 2018/12/04
       'Modified by Lydia 2024/02/21 +4.韓文
       If Left(cboSource.Text, 1) <> "1" And Left(cboSource.Text, 1) <> "2" And Left(cboSource.Text, 1) <> "3" And Left(cboSource.Text, 1) <> "4" Then
            MsgBox "原文語種請輸入1-4的選項 !", vbCritical
            Exit Function
       End If
       If Left(cboTarget.Text, 1) <> "1" And Left(cboTarget.Text, 1) <> "2" Then
            MsgBox "翻譯語種請輸入1-2的選項 !", vbCritical
            Exit Function
       End If
       If Label2(4).Caption <> "" Then
            staTF = "2" '修改
       Else
            staTF = "1" '新增
       End If
  End If
  
  'Added by Lydia 2023/08/18 改用FTP(原始檔區)存放檔案
  If staTF <> "" Then
      If Val(Text1(0)) = 0 Then
         If InStr(UCase(m_strSaveFiles) & ",", FcpTcnFKey02) = 0 Then
            MsgBox "請確認是否已上傳" & FcpTcnFKey02, vbCritical
            Exit Function
         End If
      Else
         If PUB_ChkTCNfileExist(Text1(0), FcpTcnFKey02) = False Then
            MsgBox "請確認是否已上傳" & FcpTcnFKey02, vbCritical
            Exit Function
         End If
      End If
  Else
  'end 2023/08/18
      If staTF <> "" And bolAttPdf = False Then
         MsgBox txtPath & "資料夾不存在" & FcpTcnFKey02 & "，請確認是否已上傳檔案 ! ", vbCritical
         Exit Function
      End If
  End If 'Added by Lydia 2023/08/18
  '取消急件翻譯
  'Modified by Lydia 2018/12/04 +Text1(6)
  If Trim(Text1(4).Text & Text1(5).Text & Text1(6).Text & cboSource.Text & cboTarget.Text) = "" And Label2(4).Caption <> "" Then
        If Val(Left(Label2(4).Caption, 1)) > 0 Then
             If MsgBox("是否取消急件翻譯？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
                  Exit Function
             End If
        Else
            MsgBox "追蹤號已經收文立案，不可取消急件翻譯 !", vbCritical
            Exit Function
        End If
        staTF = "4" '刪除
  End If
  'end 2018/06/07
   
  'Added by Lydia 2023/02/13
  If Frame2.Visible = True And Frame2.Enabled = True Then
     If Text1(8) <> "" Or Combo1.Text <> "" Then
         If Text1(8) = "" Then
             MsgBox "請輸入相似舊案案號 !", vbCritical
             Exit Function
         End If
         If Trim(Combo1.Text) = "" Then
             MsgBox "請輸入指定組別 !", vbCritical
             Exit Function
         End If
     End If
     'Added by Lydia 2023/06/14
     If FraTCN13.Visible = True And Text1(11).Text <> "" And Text1(9) <> "Y" Then
        MsgBox "請輸入英文組認領 !", vbCritical
        Exit Function
     End If
  End If
  'end 2023/02/13
  
    'Added by Lydia 2021/04/14 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
    If PUB_ChkUniText(Me, , True) = False Then
        Exit Function
    End If
    'end 2021/04/14
    
  TxtValidate = True
End Function

Private Sub TxtClear()
   'Dim txt As TextBox, Lbl As LABEL, chk As CheckBox 'Remove by Lydia 2018/11/13 改成Form2.0
   For Each txt In frm060504.Text1
      txt.Text = ""
      txt.Tag = "" 'Added by Lydia 2018/06/07
   Next
   Label23 = Empty
   'Modified by Lydia 2018/06/07 增加急件翻譯
   'Label2(0) = Empty
   'Label2(1) = Empty
   For Each Lbl In Label2
       Lbl.Caption = ""
   Next
   cboSource.Text = ""
   cboSource.Tag = ""
   cboTarget.Text = ""
   cboTarget.Tag = ""
   cboTCN15.Text = ""
   cboTCN15.Tag = ""
   'end 2018/06/07
   
    'Added by Lydia 2017/12/27
    Label3.Visible = False
    txtPath.Visible = False
    txtPath.Text = ""
    'end 2017/12/27
    cmdOpenDir.Visible = False 'Added by Lydia 2018/02/23
    
    'Added by Lydia 2023/02/13
    Combo1.Text = ""
    Combo1.Tag = ""
    'Added by Lydia 2023/06/14
    For Each oText In chk1
       oText.Value = 0
    Next
    'end 2023/06/14
    
    'Added by Lydia 2023/08/18
    cmdAddFile.Visible = False
    'Add By Sindy 2024/4/24 系統收件區會傳電子檔名進來
    If m_strIR01 = "" Then
    '2024/4/24 END
      m_strSaveFiles = ""
    End If
End Sub
Private Sub ToolBarSet(ByVal stA As Integer)
'Modified by Lydia 2018/11/13 改成Form2.0
 'Dim i As Integer, txt As TextBox
 Dim i As Integer
 
 Select Case stA
    Case 0
        TxtLock 1
      For i = 1 To 4
         TBar1.Buttons(i).Enabled = False
         TBar1.Buttons(i + 5).Enabled = False
      Next
      TBar1.Buttons(11).Enabled = True
      TBar1.Buttons(12).Enabled = True
      TBar1.Buttons(14).Enabled = False
    Case 1
        TxtLock 0
        If m_bInsert Or Pub_StrUserSt03 = "M51" Then
            TBar1.Buttons(1).Enabled = True
        Else
            TBar1.Buttons(1).Enabled = False
        End If
        'Modified by Lydia 2020/10/28 +新增後：判斷操作人=管制人，可以直接維護
        'If (m_bUpdate And ManageYN) Or Pub_StrUserSt03 = "M51" Then
        If (m_bUpdate And ManageYN) Or Pub_StrUserSt03 = "M51" Or (Label2(0).Caption = "" And Text1(2).Text = strUserNum) Then
            TBar1.Buttons(2).Enabled = True
        Else
            TBar1.Buttons(2).Enabled = False
        End If
        'Modified by Lydia 2020/10/28 +新增後：判斷操作人=管制人，可以直接維護
        'If (m_bDelete And ManageYN) Or Pub_StrUserSt03 = "M51" Then
        If (m_bDelete And ManageYN) Or Pub_StrUserSt03 = "M51" Or (Label2(0).Caption = "" And Text1(2).Text = strUserNum) Then
            TBar1.Buttons(3).Enabled = True
        Else
            TBar1.Buttons(3).Enabled = False
        End If
        'Added by Lydia 2023/08/18 已收文不可修改/刪除
        If Label2(0).Caption <> "" Then
           TBar1.Buttons(2).Enabled = False
           TBar1.Buttons(3).Enabled = False
        End If
        'end 2023/08/18
        
        If m_bQuery Or Pub_StrUserSt03 = "M51" Then
            TBar1.Buttons(4).Enabled = True
        Else
            TBar1.Buttons(4).Enabled = False
        End If
      'Added by Lydia 2018/06/07 從翻譯分案作業過來
      If m_PFkey <> "" Then
        For i = 6 To 10
           TBar1.Buttons(i).Enabled = False
        Next
      Else
      'end 2018/06/07
        For i = 6 To 10
           TBar1.Buttons(i).Enabled = True
        Next
      End If
      
      TBar1.Buttons(11).Enabled = False
      TBar1.Buttons(12).Enabled = False
      TBar1.Buttons(14).Enabled = True
    End Select
   
End Sub
Private Sub TxtLock(ByVal Lt As Integer)
 'Dim txt As TextBox, idx As Integer 'Remove by Lydia 2018/11/13 改成Form2.0
   Select Case Lt
      Case 0 'cancel/OK
         For Each txt In frm060504.Text1
            txt.Locked = True
            txt.Enabled = True
         Next
          'Added by Lydia 2018/06/07 急件翻譯
          cboTCN15.Locked = True
          cboSource.Locked = True
          cboTarget.Locked = True
          cmdMail.Enabled = True
          'end 2018/06/07
          Combo1.Locked = True   'Added by Lydia 2023/02/13
          FraTCN13.Enabled = False 'Added by Lydia 2023/06/14
      Case 1 'add/upd
         For Each txt In frm060504.Text1
            idx = txt.Index
            If idx <> 0 Then
                txt.Locked = False
            End If
         Next
          'Added by Lydia 2018/06/07
          cboTCN15.Locked = False
          cboSource.Locked = False
          cboTarget.Locked = False
          cmdMail.Enabled = False
          Combo1.Locked = False   'Added by Lydia 2023/02/13
          'Added by Lydia 2023/06/14
          '已確收文件不可再變更
          If Text1(11) = "3" Or Text1(11) = 4 Then
             FraTCN13.Enabled = False
          Else
             FraTCN13.Enabled = True
          End If
          'end 2023/06/14
          '程序只可輸入急件翻譯部份
          If m_PFkey <> "" Then
              Text1(1).Locked = True
              Text1(2).Locked = True
              Text1(3).Locked = True
          End If
          'end 2018/06/07
          
      Case 2 '上一筆下一筆…'第一次進入
         For Each txt In frm060504.Text1
            txt.Locked = True
         Next
          'Added by Lydia 2018/06/07 急件翻譯
          cboTCN15.Locked = True
          cboSource.Locked = True
          cboTarget.Locked = True
          cmdMail.Enabled = True
          'end 2018/06/07
          'Combo1.Locked = True   'Added by Lydia 2023/02/13 'Mark by Lydia 2023/05/03
          FraTCN13.Enabled = False 'Added by Lydia 2023/06/14
      Case 3 'query
        For Each txt In frm060504.Text1
            idx = txt.Index
             If idx = 0 Then
                txt.Locked = False
             Else
                txt.Locked = True
             End If
       Next
        'Added by Lydia 2018/06/07 急件翻譯
        cboTCN15.Locked = True
        cboSource.Locked = True
        cboTarget.Locked = True
        cmdMail.Enabled = True
        'end 2018/06/07
        'Combo1.Locked = True   'Added by Lydia 2023/02/13 'Mark by Lydia 2023/05/03
        FraTCN13.Enabled = False 'Added by Lydia 2023/06/14
      End Select
End Sub

'取得序號
Private Function GetSerialNo(strBU01 As String) As String
  Dim StrSQLa As String
  Dim rsA As New ADODB.Recordset

  '抓TrackingCaseName(案件命名追蹤)流水號 累加
  StrSQLa = "SELECT NVL(MAX(TCN01),0) as TCN01 From TrackingCaseName "
  rsA.CursorLocation = adUseClient
  rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly

  If Not rsA.EOF And Not rsA.BOF Then
      GetSerialNo = Format(Val(rsA("TCN01").Value) + 1, "00000")
  Else
      GetSerialNo = "00001"
  End If
  If rsA.State <> adStateClosed Then rsA.Close
  Set rsA = Nothing
End Function

Private Sub RefreshRange()

Dim strSql As String
Dim rsTmp As New ADODB.Recordset
 'Added by Lydia 2018/06/07 從翻譯分案作業過來
 If m_PFkey <> "" Then
       strSql = "Select TCN01 From TrackingCaseName where TCN01=" & CNULL(m_PFkey)
 Else
 'end 2018/06/07
    If ManageYN Then '主管可看部屬建的資料
         strSql = "Select TCN01 From TrackingCaseName " & _
                     "Where TCN01 = (Select MIN(TCN01) From TrackingCaseName,Staff Where TCN03=ST01(+)  And (TCN05 <>'111111' or TCN05 is null) And (ST52='" & strUserNum & "' " & _
                                            "OR ST53='" & strUserNum & "' OR ST54='" & strUserNum & "' OR ST55='" & strUserNum & "' OR TCN03='" & strUserNum & "' " & _
                                            "OR TCN06='" & strUserNum & "' OR TCN03='" & strUserNum & "')) "
    Else
      If Pub_StrUserSt03 = "M51" Then '電腦中心可看全部資料
          strSql = "Select TCN01 From TrackingCaseName " & _
                      "Where TCN01 = (Select MIN(TCN01) From TrackingCaseName Where TCN05 <>'111111' OR TCN05 is null)"
      Else
          strSql = "Select TCN01 From TrackingCaseName " & _
                      "Where TCN01 = (Select MIN(TCN01) From TrackingCaseName Where (TCN05 <>'111111' OR TCN05 is null) And (TCN06='" & strUserNum & "' OR TCN03='" & strUserNum & "')) "
      End If
    End If
 End If
 
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("TCN01")) = False Then: m_FirstKEY(0) = rsTmp.Fields("TCN01")
   Else
        m_FirstKEY(0) = 0
   End If
   rsTmp.Close

 'Added by Lydia 2018/06/07 從翻譯分案作業過來
 If m_PFkey <> "" Then
       strSql = "Select TCN01 From TrackingCaseName where TCN01=" & CNULL(m_PFkey)
 Else
 'end 2018/06/07
     If ManageYN Then
       strSql = "Select TCN01 From TrackingCaseName " & _
                   "Where TCN01 = (Select MAX(TCN01) From TrackingCaseName,Staff Where TCN03=ST01(+)  And (TCN05 <>'111111' or TCN05 is null) " & _
                                             "And (ST52='" & strUserNum & "' OR ST53='" & strUserNum & "' OR ST54='" & strUserNum & "' OR ST55='" & strUserNum & "' " & _
                                             " OR TCN03='" & strUserNum & "' OR TCN06='" & strUserNum & "' ))  "
    
     Else
       If Pub_StrUserSt03 = "M51" Then '電腦中心可看全部資料
           strSql = "Select TCN01 From TrackingCaseName " & _
                       "Where TCN01 = (Select MAX(TCN01) From TrackingCaseName Where TCN05 <>'111111' OR TCN05 is null)"
    
       Else
            strSql = "Select TCN01 From TrackingCaseName " & _
                        "Where TCN01 = (Select MAX(TCN01) From TrackingCaseName Where (TCN05 <>'111111' OR TCN05 is null) And (TCN06='" & strUserNum & "' OR TCN03='" & strUserNum & "')) "
       End If
     End If
 End If
 
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("TCN01")) = False Then: m_LastKEY(0) = rsTmp.Fields("TCN01")
   Else
    m_LastKEY(0) = 0
   End If
   rsTmp.Close
     
   Set rsTmp = Nothing
End Sub

' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal strKEY01 As String) As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim strDutyAgent As String 'Added by Lydia 2023/08/18

   IsRecordExist = False
   
   If ManageYN Then
     strSql = "Select * From TrackingCaseName,Staff " & _
                "Where TCN01 = " & strKEY01 & " And TCN03=ST01(+) And (TCN05 <>'111111' or TCN05 is null)  And (ST52='" & strUserNum & "' " & _
                                         "OR ST53='" & strUserNum & "' OR ST54='" & strUserNum & "' OR ST55='" & strUserNum & "' OR TCN03='" & strUserNum & "' OR TCN06='" & strUserNum & "' )"
   Else
     If Pub_StrUserSt03 = "M51" Then '電腦中心可看全部資料
       strSql = "Select TCN01 From TrackingCaseName " & _
                   "Where TCN01 = " & strKEY01 & " And (TCN05 <>'111111' or TCN05 is null) "

     Else
        strSql = "Select * From TrackingCaseName " & _
                "Where TCN01 = " & strKEY01 & " And (TCN05 <>'111111' or TCN05 is null) And (TCN06='" & strUserNum & "' OR TCN03='" & strUserNum & "')"
     End If
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
   
   'Added by Lydia 2023/08/18 啟用日控制: 開放休假人員（含主管）之職務代理人查詢及修改資料
   If IsRecordExist = False Then
      'Modified by Lydia 2023/09/11 +st16
      strSql = "select tcn01, tcn03, st02, st16, st52 as tcn03st52, st53 as tcn03st53, st54 as tcn03st54, st55 as tcn03st55 " & _
               "from trackingcasename ,staff where TCN01 = " & strKEY01 & " And tcn03=st01(+) "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         strExc(1) = "" & rsTmp.Fields("tcn03")
         strExc(2) = "" & rsTmp.Fields("tcn03st52")
         strExc(3) = "" & rsTmp.Fields("tcn03st53")
         strExc(4) = "" & rsTmp.Fields("tcn03st54")
         strExc(5) = "" & rsTmp.Fields("tcn03st55")
         For idx = 1 To 5
            If Trim(strExc(idx)) <> "" Then
               'Added by Lydia 2023/09/11 日文組因為有開會或接待客戶，所以非請假也要能互相管理(by Elaine)
               If "" & rsTmp.Fields("st16") = "2" Then
                   Call GetABS001_1(strExc(idx), strExc(7), strExc(8), strExc(9))
                   strDutyAgent = strExc(7) & "," & strExc(8) & "," & strExc(9)
               Else
               'end 2023/09/11
                   strDutyAgent = GetCaseDutyAgent(strExc(idx), "", False, "1")
               End If  'Added by Lydia 2023/09/11
               If strDutyAgent <> "" And InStr(strDutyAgent & ",", strUserNum) > 0 Then
                  If MsgBox("追蹤號：" & strKEY01 & vbCrLf & "管制人：" & rsTmp.Fields("st02") & vbCrLf & "是否繼續作業？", vbExclamation + vbYesNo + vbDefaultButton2) = vbYes Then
                     IsRecordExist = True
                     Exit For
                  End If
               End If
            End If
         Next idx
      End If
      rsTmp.Close
   End If
   'end 2023/08/18
   
   Set rsTmp = Nothing
End Function
' 顯示資料
Private Sub GetCurrRecordVal(ByVal strKEY01 As String)
Dim strSql As String
Dim rsTmp As New ADODB.Recordset

   If IsRecordExist(strKEY01) = True Then
      m_CurrKEY(0) = strKEY01
   Else
        If ManageYN Then
             strSql = "Select TCN01 From TrackingCaseName,Staff " & _
                            "Where TCN01 = " & m_CurrKEY(0) & " And TCN03=ST01(+) And (TCN05 <>'111111' or TCN05 is null) And (ST52='" & strUserNum & "' OR ST53='" & strUserNum & "' OR ST54=" & _
                            "'" & strUserNum & "' OR ST55='" & strUserNum & "'  OR TCN03='" & strUserNum & "' OR TCN06='" & strUserNum & "' ) "
        Else
            If Pub_StrUserSt03 = "M51" Then '電腦中心可看全部資料
                 strSql = "Select TCN01 From TrackingCaseName " & _
                             "Where TCN01 = '" & m_CurrKEY(0) & "' And (TCN05 <>'111111' OR TCN05 is null)"
            Else
                 strSql = "Select TCN01 From TrackingCaseName " & _
                             "Where TCN01 ='" & m_CurrKEY(0) & "' And (TCN05 <>'111111' OR TCN05 is null) And (TCN06='" & strUserNum & "' OR TCN03='" & strUserNum & "')"
            End If
        End If
   
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("TCN01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("TCN01")
         rsTmp.Close
         RefreshRange
         SetTxtValue
         GoTo EXITSUB
      End If
      rsTmp.Close
      
   If ManageYN Then
        strSql = "Select TCN01 From TrackingCaseName,Staff Where TCN01 >" & m_CurrKEY(0) & " And TCN03=ST01(+) And (TCN05 <>'111111' or TCN05 is null) " & _
                       "And (ST52='" & strUserNum & "' OR ST53='" & strUserNum & "' OR ST54='" & strUserNum & "' OR ST55='" & strUserNum & "'  OR TCN03='" & strUserNum & "' OR TCN06='" & strUserNum & "' ) Order by TCN01"
   Else
        If Pub_StrUserSt03 = "M51" Then '電腦中心可看全部資料
              strSql = "Select TCN01 From TrackingCaseName " & _
                          "Where TCN01 >" & m_CurrKEY(0) & " And (TCN05 <>'111111' or TCN05 is null) Order by TCN01"
        Else
            strSql = "Select TCN01 From TrackingCaseName " & _
                          "Where TCN01 > " & m_CurrKEY(0) & "  And (TCN05 <>'111111' or TCN05 is null) And (TCN06='" & strUserNum & "' OR TCN03='" & strUserNum & "') Order by TCN01 "
      End If
   End If
   
   
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("TCN01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("TCN01")
      Else
         GetLastRecordVal
         GoTo EXITSUB
      End If
      rsTmp.Close
   End If
   'RefreshRange  'Mark by Lydia 2022/08/25 保持在指定編號
   SetTxtValue
EXITSUB:
End Sub
' 第一筆資料
Private Sub GetFirstRecordVal()
   m_CurrKEY(0) = m_FirstKEY(0)
   
   SetTxtValue
End Sub
'上一筆資料
Private Sub GetPreRecordVal()
Dim strSql As String
Dim rsTmp As New ADODB.Recordset

   If m_CurrKEY(0) = m_FirstKEY(0) Then
      ShowMsg MsgText(9008)
      GoTo EXITSUB
   End If
   
   If ManageYN Then
        strSql = "Select MAX(TCN01) TCN01 From TrackingCaseName,Staff " & _
                       "Where TCN01<" & m_CurrKEY(0) & " And TCN03=ST01(+) And (TCN05 <>'111111' or TCN05 is null) And (ST52='" & strUserNum & "' OR ST53='" & strUserNum & "' OR ST54=" & _
                       "'" & strUserNum & "' OR ST55='" & strUserNum & "' OR TCN03='" & strUserNum & "' OR TCN06='" & strUserNum & "' ) "
   Else
        If Pub_StrUserSt03 = "M51" Then '電腦中心可看全部資料
              strSql = "Select MAX(TCN01) TCN01 From TrackingCaseName " & _
                          "Where TCN01<" & m_CurrKEY(0) & " And (TCN05 <>'111111' or TCN05 is null) "
       Else
              strSql = "Select MAX(TCN01) TCN01 From TrackingCaseName " & _
               "Where TCN01  <" & m_CurrKEY(0) & " And (TCN05 <>'111111' or TCN05 is null) And (TCN06='" & strUserNum & "' OR TCN03='" & strUserNum & "')"
        End If
  End If
  
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("TCN01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("TCN01")
      rsTmp.Close
      SetTxtValue
      GoTo EXITSUB
   End If
   rsTmp.Close
     
EXITSUB:
   Set rsTmp = Nothing
End Sub
'下一筆資料
Private Sub GetNextRecordVal()
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   If m_CurrKEY(0) = m_LastKEY(0) Then
      ShowMsg MsgText(9009)
      GoTo EXITSUB
   End If
   If ManageYN Then
        strSql = "Select MIN(TCN01) TCN01 From TrackingCaseName,Staff Where TCN01 > " & m_CurrKEY(0) & " And TCN03=ST01(+) And (TCN05 <>'111111' or TCN05 is null) " & _
                       "And (ST52='" & strUserNum & "' OR ST53='" & strUserNum & "' OR ST54='" & strUserNum & "' OR ST55='" & strUserNum & "' OR TCN03='" & strUserNum & "' OR TCN06='" & strUserNum & "' ) "
   Else
        If Pub_StrUserSt03 = "M51" Then '電腦中心可看全部資料
              strSql = "Select MIN(TCN01) TCN01  From TrackingCaseName " & _
                    "Where TCN01 >" & m_CurrKEY(0) & " And (TCN05 <>'111111' or TCN05 is null) "
        Else
            strSql = "Select MIN(TCN01) TCN01 From TrackingCaseName " & _
                    "Where TCN01 >" & m_CurrKEY(0) & " And (TCN05 <>'111111' or TCN05 is null) And (TCN06='" & strUserNum & "' OR TCN03='" & strUserNum & "')"
        End If
   End If
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("TCN01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("TCN01")
      rsTmp.Close
      SetTxtValue
      GoTo EXITSUB
   End If
   rsTmp.Close
     
EXITSUB:
   Set rsTmp = Nothing
End Sub
' 最後一筆資料
Private Sub GetLastRecordVal()
   m_CurrKEY(0) = m_LastKEY(0)
   SetTxtValue
End Sub

' 查詢記錄
Private Function QueryRecord(ByVal Str01 As String) As Boolean

   QueryRecord = False
   Str01 = Val(Str01)
  
   If IsRecordExist(Str01) = True Then
      m_CurrKEY(0) = Str01
      
      QueryRecord = True
      
   Else
      QueryRecord = False
      MsgBox ("無此資料")
   End If
    SetTxtValue
   ToolBarSet 1
End Function
'Added by Lydia 2023/05/03
Private Sub Text1_Change(Index As Integer)
   If Index = 3 Then
      PUB_RefreshText Text1(3)
   End If
End Sub

'Added by Lydia 2018/11/21
'Modified by Lydia 2023/05/03 改成可用滑鼠右鍵
'Private Sub Text1_KeyDown(Index As Integer, KeyCode As MSForms.ReturnInteger, Shift As Integer)
'    If Index = 3 Then
'         Call PUB_HandleForm2TextBox(Me.Text1(3), Me.Text1(1), KeyCode, Shift)  '模組化-統一控制
'    End If
'End Sub
'
''Added by Lydia 2018/11/21
'Private Sub Text1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'
'    If SyxMsg <> "Text1_" & Format(Index, "00") Then '避免連續產生訊息
'        bolMsgRight = False
'        SyxMsg = "Text1_" & Format(Index, "00")
'    End If
'    Call PUB_HandleForm2TextBoxR(Button, Shift, bolMsgRight) '模組化-統一控制
'
'End Sub
Private Sub Text1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   If Index = 3 Or Index = 8 Then
      If Button = 2 Then Forms(0).PopupMenu2 Text1(Index)
   End If
End Sub
'end 2023/05/03

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
  If ActionEdit = 0 Or ActionEdit = 1 Then
    If Text1(Index).Locked = True Then Exit Sub 'Added by Lydia 2018/06/07
    Select Case Index
        Case 1
            If Trim(Text1(1)) = "" Then
                MsgBox "期限不可為空值"
                'Modified by Lydia 2018/12/04
                'Cancel = True
                'Text1_GotFocus 1
                'Exit Sub
                GoTo EXITSUB
            ElseIf CheckIsTaiwanDate(Me.Text1(1)) = False Then
                'Modified by Lydia 2018/12/04
                'Cancel = True
                'Text1_GotFocus 1
                'Exit Sub
                GoTo EXITSUB
            ElseIf Not ChkWorkDay(DBDATE(Me.Text1(1))) Then
                MsgBox "期限必須是工作天 !"
                'Modified by Lydia 2018/12/04
                'Cancel = True
                'Text1_GotFocus 1
                'Exit Sub
                GoTo EXITSUB
            ElseIf DBDATE(Me.Text1(1)) < DBDATE(strSrvDate(2)) Then
                MsgBox "期限需大於系統日!"
                'Modified by Lydia 2018/12/04
                'Cancel = True
                'Text1_GotFocus 1
                'Exit Sub
                GoTo EXITSUB
            End If
        Case 2  '輸入員編帶姓名
        If Len(Trim(Text1(2))) > 0 Then
            If Len(Text1(2).Text) = 5 Then
               'Modified by Lydia 2023/12/26
               'strExc(1) = "SELECT ST02,A0902 FROM STAFF,ACC090 WHERE ST01='" & Text1(2) & "' And ST03=A0901 And ST04<>2 And ST03='F23' "
               strExc(1) = "SELECT ST02 FROM STAFF WHERE ST01='" & Text1(2) & "' And ST04<>2 And ST03='F23' "
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(1))
               If intI = 0 Then
                    Label2(1).Caption = ""
                    MsgBox ("管制人輸入錯誤,請輸入外專承辦組在職同仁！")
                    'Modified by Lydia 2018/12/04
                    'Cancel = True
                    'Text1_GotFocus 2
                    'Exit Sub
                    GoTo EXITSUB
               Else
                    If IsNull(RsTemp.Fields(0).Value) Then
                        Label2(1).Caption = ""
                   Else
                       Label2(1).Caption = RsTemp.Fields(0).Value
                   End If
               End If
            Else
                MsgBox ("管制人輸入錯誤,請輸入外專承辦組在職同仁！")
                'Modified by Lydia 2018/12/04
                'Cancel = True
                'Text1_GotFocus 2
                'Exit Sub
                GoTo EXITSUB
            End If
        Else
            MsgBox ("管制人不可為空！")
            'Modified by Lydia 2018/12/04
            'Cancel = True
            'Text1_GotFocus 2
            'Exit Sub
            GoTo EXITSUB
        End If
        Case 3
          If CheckLengthIsOK(Text1(3), 200) = False Then Text1_GotFocus 3: Cancel = True: Exit Sub
        'Added by Lydia 2018/06/07 交稿期限
        'Modified by Lydia 2018/12/04 +只交Claims期限Txt1(6)
        Case 5, 6
            If Text1(Index).Text <> "" Then
                If CheckIsTaiwanDate(Me.Text1(Index)) = False Then
                    GoTo EXITSUB
                ElseIf Not ChkWorkDay(DBDATE(Me.Text1(Index))) Then
                    'Modified by Lydia 2018/12/04
                    'MsgBox "交稿期限必須是工作天 !"
                    MsgBox IIf(Index = 5, "交稿期限", "只交Claims期限") & "必須是工作天 !"
                    GoTo EXITSUB
                ElseIf Me.Text1(Index).Tag = "" And DBDATE(Me.Text1(Index)) < DBDATE(strSrvDate(2)) Then
                    'Modified by Lydia 2018/12/04
                    'MsgBox "交稿期限需大於系統日!"
                    MsgBox IIf(Index = 5, "交稿期限", "只交Claims期限") & "需大於系統日!"
                    GoTo EXITSUB
                End If
            End If
        'end 2018/06/07
        'Added by Lydia 2023/05/03 相似舊案案號：依客戶指示有相似案／舊案
        Case 8
            If Text1(Index).Text <> Text1(Index).Tag Then Combo1.Text = ""
            If Text1(Index) <> "" Then
               Call ChgCaseNo(Text1(Index), strExc)
               If (strExc(1) = "FCP" Or strExc(1) = "P") And Len(strExc(2)) = 6 Then
                   strExc(0) = "select pa150 from patent where pa01='" & strExc(1) & "' and pa02='" & strExc(2) & "' and pa03='" & strExc(3) & "' and pa04='" & strExc(4) & "' "
                   intI = 1
                   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                   If intI = 1 Then
                      If Val("" & RsTemp.Fields("pa150")) > 0 Then
                          Combo1.Text = RsTemp.Fields("pa150") & "." & PUB_GetFCPGrpName(RsTemp.Fields("pa150"))
                      End If
                      Text1(Index) = strExc(1) & strExc(2) & strExc(3) & strExc(4)
                   Else
                      MsgBox "查無此案號！", vbCritical, "檢核資料"
                      GoTo EXITSUB
                   End If
               Else
                    MsgBox "請輸入FCP案或P案案號！", vbCritical, "檢核資料"
                    GoTo EXITSUB
               End If
            End If
        'end 2023/05/03
    End Select
  End If
  'Added by Lydia 2023/08/17 限制不可輸入全形字; ex.20245的暫不認領=Ｙ，造成直接通知工程師主管開始命名
  If Trim(Text1(Index)) <> "" And (Index = 7 Or Index = 9 Or Index = 10) Then
     Text1(Index) = Chr(89)
  End If
  
'Added by Lydia 2018/12/04
  Exit Sub
  
EXITSUB:
  Cancel = True
  Text1(Index).SetFocus
  Text1_GotFocus Index
'end 2018/12/04
End Sub

'反白
'Remove by Lydia 2018/06/05
'Public Sub TextInverse(ByRef txtTemp As TextBox)
'txtTemp.SelStart = 0
'txtTemp.SelLength = Len(txtTemp.Text)
'End Sub

Private Sub Text1_GotFocus(Index As Integer)
   If ActionEdit <> 3 Then
       'Modified by Lydia 2018/11/21  取消備註反白
       'TextInverse Text1(Index)
       If Index <> 3 Then TextInverse Text1(Index)
   End If
End Sub

'Modified by Lydia 2018/11/13 改成Form2.0
'Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
    Select Case Index
    'Modified by Lydia 2023/05/03 相似舊案案號 +8
    Case 2, 8 '管制人
        KeyAscii = UpperCase(KeyAscii)
    'Added by Lydia 2023/02/13
    'Modified by Lydia 2023/06/14 +10=TCN13
    Case 7, 9, 10
        KeyAscii = UpperCase(KeyAscii)
        If KeyAscii <> 89 And KeyAscii <> 8 Then
           KeyAscii = 0
           Beep
        End If
    'end 2023/02/13
    End Select
End Sub

Private Function DelRecord(ByVal nTime As Long) As Boolean
    Dim strSQLD As String, m_st01 As String
    DelRecord = False
On Error GoTo ErrHand
      
    '刪除只是將收文號設為6個1
    m_st01 = Val(Me.Text1(0))
    
    cnnConnection.BeginTrans 'Added by Lydia 2018/06/07
        'Modified by Lydia 2018/06/07 +清空急件翻譯,避免造成翻譯分案作業抓到
        'strSQLD = "Update TrackingCaseName  set TCN05='111111',TCN09='" & strUserNum & "',TCN10=" & strSrvDate(1) & ",TCN11=" & nTime & "  Where TCN01='" & m_st01 & "' "
        strSQLD = "Update TrackingCaseName  set TCN05='111111',TCN09='" & strUserNum & "',TCN10=" & strSrvDate(1) & ",TCN11=" & nTime & _
                         IIf(Label2(4).Caption <> "", ",TCN14=null, TCN15=null ", "") & " Where TCN01='" & m_st01 & "' "
        cnnConnection.Execute strSQLD
     'Added by Lydia 2018/06/07 刪除急件翻譯
        If Label2(4).Caption <> "" Then
            strSQLD = "delete from TransFee where TF01='" & Label2(4).Caption & "' "
            cnnConnection.Execute strSQLD, intI
        End If
    cnnConnection.CommitTrans
    'end 2018/06/07
    
    
       'Added by Lydia 2023/08/18-- 2023/08/21改用原始檔區存放
       'Modified by Lydia 2023/08/24 改成模組共用
       Call ProcDelFTP(Text1(0), "1")
   
   If m_st01 = m_LastKEY(0) And m_st01 = m_FirstKEY(0) Then
        m_CurrKEY(0) = 0
        SetTxtValue
   Else
        '只有刪除的是最後一筆才須重新取第一筆及最後一筆
        If m_st01 = m_LastKEY(0) Or m_st01 = m_FirstKEY(0) Then
             RefreshRange
        End If
            GetCurrRecordVal m_st01
   End If
    DelRecord = True
    
   Exit Function
ErrHand:
    
    MsgBox "刪除失敗！" & vbCrLf & Err.Description
End Function

'更新 Create 及 Update 的人
Private Sub UpdateCUID(ByRef rsSrcTmp As ADODB.Recordset)
Dim strTemp As String
Dim strCName As String
Dim strCDate As String
Dim strCTime As String
Dim strUName As String
Dim strUDate As String
Dim strUTime As String
   
If rsSrcTmp.RecordCount > 0 Then
   If IsNull(rsSrcTmp.Fields("TCN06")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("TCN06")) = False Then
         strCName = GetStaffName(rsSrcTmp.Fields("TCN06"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("TCN07")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("TCN07")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("TCN07"))
         strCDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("TCN08")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("TCN08")) = False Then
         strTemp = rsSrcTmp.Fields("TCN08")
         strCTime = Format(strTemp, "##:##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("TCN09")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("TCN09")) = False Then
         strUName = GetStaffName(rsSrcTmp.Fields("TCN09"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("TCN10")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("TCN10")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("TCN10"))
         strUDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("TCN11")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("TCN11")) = False Then
         strTemp = rsSrcTmp.Fields("TCN11")
         strUTime = Format(strTemp, "##:##")
      End If
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

'Added by Lydia 2017/12/04 FCP案件命名電子化:上傳檔案
Private Sub cmdFile_Click()
   'Modified by Lydia 2020/10/27 限制檔名80字元
   Call frm090801_8.SetParent(Me, 80, , "上傳檔案") 'Modified by Lydia 2018/03/02
   frm090801_8.m_strSaveFiles = Me.m_strSaveFiles
   frm090801_8.lblCaseNo = ""
   frm090801_8.lblCaseNo.Visible = False
   frm090801_8.Label1.Visible = False
   frm090801_8.Label4.Visible = False
   frm090801_8.bolNotPDF = True
   frm090801_8.Show vbModal
End Sub

Private Sub txtPath_GotFocus()
   TextInverse txtPath
End Sub

'Added by Lydia 2018/02/23 開啟Tracking_no暫存資料夾
Private Sub cmdOpenDir_Click()
Dim hLocalFile As Long 'Added by Lydia 2018/06/21

'Modified by Lydia 2018/03/23 無權限的錯誤訊息要改
'On Error Resume Next
On Error GoTo ErrHand01

    If txtPath = "" Then
        MsgBox "無檔案路徑!!"
    Else
        'Remove by Lydia 2018/03/23
        'If Pub_StrUserSt03 <> "M51" And Left(Pub_StrUserSt03, 1) <> "F" Then
        '      If MsgBox("非國外部人員無權限進入Tracking_No資料夾，是否繼續開啟？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
        '           Exit Sub
        '      End If
        'End If
        'end 2018/03/23
        If Dir(txtPath.Text & "\*.*") <> "" Then
             'Modified by Lydia 2018/06/21 用檔案總管開啟放置1~2分鐘後,檔案總管會出錯(ex. A2037, A4041)
             'SHELL "Explorer.exe " & txtPath.Text, vbNormalFocus  '開啟案件資料夾
             ShellExecute hLocalFile, "explore", txtPath.Text, vbNullString, vbNullString, 1
        Else
             MsgBox txtPath.Text & "的資料夾不存在或無檔案!", vbInformation
        End If
    End If
    
'Added by Lydia 2018/03/23
    Exit Sub
    
ErrHand01:
    If Err.Number <> 0 Then
         '全部錯誤訊息統一
         MsgBox "無法讀取" & txtPath.Text & "，請通知電腦中心！", vbCritical
         Resume Next
    End If
'end 2018/03/23
End Sub

'Added by Lydia 2018/06/07
Private Sub cboTCN15_Validate(Cancel As Boolean)
   If Trim(cboTCN15.Text) <> "" Then
      strExc(1) = Trim(Left(cboTCN15.Text, 6))
      If ClsPDGetStaff(strExc(1), strExc(2)) = True Then
           cboTCN15.Text = strExc(1) & " " & strExc(2)
      Else
           cboTCN15.Text = ""
      End If
   End If
   Text1(4).Text = Trim(Left(cboTCN15.Text, 6))
End Sub

Private Sub cboTarget_Validate(Cancel As Boolean)
Dim iR As Integer
   If cboTarget.Text <> "" Then
      iR = Val(cboTarget.Text)
      If iR = 0 Or iR > 2 Then
           MsgBox "請輸入1-2的選項！", vbCritical
           Cancel = True
           cboTarget.SetFocus
           Exit Sub
      Else
           If cboTarget.ListIndex <> iR - 1 Then
                cboTarget.ListIndex = iR - 1
           End If
      End If
   End If
End Sub

Private Sub cboSource_Validate(Cancel As Boolean)
Dim iR As Integer
   If cboSource.Text <> "" Then
        iR = Val(cboSource.Text)
        'Modified by Lydia 2024/02/21 +4.韓文
        If iR = 0 Or iR > 4 Then
             MsgBox "請輸入1-4的選項！", vbCritical
             Cancel = True
             cboSource.SetFocus
             Exit Sub
        Else
             If cboSource.ListIndex <> iR - 1 Then
                  cboSource.ListIndex = iR - 1
             End If
        End If
   End If
End Sub
'end 2018/06/07

'Added by Lydia 2018/06/05
Private Sub cmdMail_Click()
   'Modified by Lydia 2018/12/04 +只交Claims期限
   'If Text1(5).Text = "" And Trim(cboTCN15.Text) = "" Then
   '      MsgBox "急件翻譯Email必需輸入翻譯人員或交稿期限! ", vbCritical
   If Trim(Text1(5).Text & Text1(6)) = "" And Trim(cboTCN15.Text) = "" Then
      MsgBox "急件翻譯Email必需輸入翻譯人員、交稿期限或只交Claims期限! ", vbCritical
      Exit Sub
   End If
   If ProcEmail("0", Text1(0)) = False Then
   End If
End Sub

Private Function ProcEmail(ByVal iType As String, ByVal iNo As String) As Boolean
Dim stFtpIP As String
Dim stSub As String, stUsers As String, stConT As String
Dim stAttFile As String
Dim stLocalPath As String
Dim pbolDone As Boolean
Dim strDate As String, strTime As String, strRestKind As String, strExcept As String 'Added by Lydia 2018/09/19

   ProcEmail = False
    '收件人
    stUsers = Pub_GetSpecMan("M")
    'Added by Lydia 2018/09/19 判斷人員請假,指定代理人
    '目前系統日期時間
    strDate = strSrvDate(1)
    strTime = Left(Right("000000" & ServerTime, 6), 2) & ":" & Mid(Right("000000" & ServerTime, 6), 3, 2)
    '若過下班時間,檢查日期改為下一個工作天,時間預設為07:30
    If Val(Format(strTime, "hhmm")) > 1800 Then
       strDate = CompWorkDay(2, strSrvDate(1), 0)
       strTime = "07:30"
    End If
    If CheckIsPersonRest(stUsers, strDate, strTime, strRestKind) = True Then
         stUsers = "86024"
         strExcept = "因翻譯分案人請假,請代為辦理！"
    End If
    'end 2018/09/19
    
    '主旨
    stSub = "急件翻譯(追蹤號：" & iNo & ")"
    Select Case iType
       Case "4" '取消急件翻譯
          stSub = "取消->" & stSub
          'Modified by Lydia 2018/09/19
          'stConT = vbCrLf & vbCrLf & "同主旨"
          stConT = vbCrLf & vbCrLf & "同主旨" & IIf(strExcept <> "", vbCrLf & strExcept, "")
          Call PUB_SendMail(strUserNum, stUsers, "", stSub, stConT)
          ProcEmail = True
             
       Case Else '其他
          'Added by Lydia 2023/08/18 改用FTP(原始檔區)存放檔案
          'Modified by Lydia 2024/12/13 刪除舊Code
            stLocalPath = App.path & "\" & strUserNum
            Call Pub_ChkExcelPath(stLocalPath)
            stLocalPath = stLocalPath & "\FCmail1"
            Call Pub_ChkExcelPath(stLocalPath)
            PUB_KillAnyFile stLocalPath
            
            strSql = "select cpf01,cpf02,cpf13 from casepaperfile where cpf01='" & Val(iNo) & "' AND (UPPER(CPF02) LIKE '%" & FcpTcnFKey02 & "') and cpf10<>'D' and substr(upper(cpf02),-4)<>upper('.del') "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               RsTemp.MoveFirst
               Do While Not RsTemp.EOF
                  If "" & RsTemp.Fields("CPF01") <> "" And "" & RsTemp.Fields("CPF02") <> "" And "" & RsTemp.Fields("CPF13") <> "" Then
                     strExc(1) = stLocalPath & "\" & RsTemp.Fields("CPF02")  '下載檔案名稱+路徑
                     If PUB_GetFtpFile("" & RsTemp.Fields("CPF13"), strExc(1), "CASEPAPERFILE") = True Then
                        stAttFile = stAttFile & ";" & strExc(1)
                     End If
                  End If
                  RsTemp.MoveNext
               Loop
               stAttFile = Mid(stAttFile, 2)
            Else
               MsgBox "請確認是否已上傳" & FcpTcnFKey02, vbCritical
               Exit Function
            End If
         '-------------------------
          
          stConT = "追蹤號：" & iNo & vbCrLf
          stConT = stConT & "翻譯人員：" & cboTCN15.Text & vbCrLf
          stConT = stConT & "交稿期限：" & ChangeTStringToTDateString(Text1(5)) & vbCrLf
          If Text1(6) <> "" Then stConT = stConT & "只交Claims期限：" & ChangeTStringToTDateString(Text1(6)) & vbCrLf  'Added by Lydia 2018/12/04
          stConT = stConT & "原文語種：" & Mid(cboSource.Text, 3) & vbCrLf
          stConT = stConT & "翻譯語種：" & Mid(cboTarget.Text, 3) & vbCrLf
          stConT = stConT & String(30, "-") & vbCrLf
          stConT = stConT & "管制人：" & Label2(1).Caption & vbCrLf
          stConT = stConT & "備　註：" & Trim(Text1(3).Text) & vbCrLf
          'Added by Lydia 2018/09/19
          If strExcept <> "" Then
              stConT = strExcept & vbCrLf & stConT
          End If
          'end 2018/09/19
          
          '開啟Email畫面
          frm880019.txtSubject = stSub
          frm880019.txtContent = stConT
          frm880019.txtReceiver = stUsers
          frm880019.SetAttach stAttFile
          frm880019.cmdAttach.Visible = False
          frm880019.SetParent Me
          frm880019.Show vbModal
          pbolDone = frm880019.m_bolDone '是否傳送成功
          Unload frm880019
          If pbolDone = True Then
              ProcEmail = True
          End If
    End Select
      
End Function
'end 2018/06/05

'Added by Lydia 2018/06/07 設原文、翻譯語種和外翻編號的下拉清單
Private Sub SetCombList()
   cboSource.Clear
   cboSource.AddItem "1." & Pub_GetTransFeeL("1", "1")
   cboSource.AddItem "2." & Pub_GetTransFeeL("1", "2")
   cboSource.AddItem "3." & Pub_GetTransFeeL("1", "3")
   cboSource.AddItem "4." & Pub_GetTransFeeL("1", "4")  'Added by Lydia 2024/02/21
   cboTarget.Clear
   cboTarget.AddItem "1." & Pub_GetTransFeeL("2", "1")
   cboTarget.AddItem "2." & Pub_GetTransFeeL("2", "2")
   
   'Modified by Lydia 2025/03/13 改用模組取得
   'strSql = "select st01,st02 from staff where st01 in('" & 外翻_舜禹 & "','" & 外翻_捷恩凱 & "','" & 外翻_迅達 & "') order by st01"
   strSql = "select st01,st02 from staff where st01 in(" & GetAddStr(Pub_SetF51Order("F", "")) & ") order by st01"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
        RsTemp.MoveFirst
        Do While Not RsTemp.EOF
              cboTCN15.AddItem RsTemp.Fields("st01") & " " & RsTemp.Fields("st02")
              RsTemp.MoveNext
        Loop
   End If
End Sub

'Added by Lydia 2023/02/13
Private Sub Combo1_Validate(Cancel As Boolean)
   If Combo1 <> "" Then
      Combo1 = Left(Combo1, 1) + "." + PUB_GetFCPGrpName(Left(Combo1, 1))
      If Combo1 = Left(Combo1, 1) + "." Then
         Combo1 = Left(Combo1, 1)
         Cancel = True
         Combo1.SetFocus
      End If
   End If
End Sub

'Added by Lydia 2023/06/14
Private Sub Chk1_Click(Index As Integer)
   If chk1(Index).Value = vbChecked Then
       Text1(11).Text = Index
       For Each oText In chk1
          If oText.Index <> Val(Text1(11)) Then
             oText.Value = 0
          End If
       Next
   Else
       If Text1(11).Text <> "" And Val(Text1(11).Text) = Index Then
           Text1(11).Text = ""
       End If
   End If
End Sub

'Added by Lydia 2023/08/18 上傳檔案：改用FTP(原始檔區)存放檔案
Private Sub cmdAddFile_Click()
   Call frm060504_1.SetParent(Me, 80, True)
   frm060504_1.m_strTCN01 = Text1(0)
   frm060504_1.m_strSaveFiles = Me.m_strSaveFiles
   Call frm060504_1.QueryData(True)
   frm060504_1.Show vbModal
End Sub

'Added by Lydia 2023/08/24
Private Sub cmdDelFTP_Click()
   If MsgBox("追蹤號：" & Trim(Text1(0)) & "的原始檔，確定刪除？", vbInformation + vbYesNo + vbDefaultButton2, "刪除原始檔") = vbYes Then
      Call ProcDelFTP(Text1(0), "2")
   End If
End Sub

'Added by Lydia 2023/08/24 模組化
Private Sub ProcDelFTP(ByVal pNo As String, ByVal pKind As String)
Dim intQ As Integer, strQ1 As String
Dim rsQD As New ADODB.Recordset

   strQ1 = "select * from casepaperfile where cpf01='" & Trim(pNo) & "' order by 2 "
   intQ = 1
   Set rsQD = ClsLawReadRstMsg(intQ, strQ1)
   If intQ = 1 Then
      rsQD.MoveFirst
      Do While Not rsQD.EOF
         '直接從資料庫刪除檔案-->只保留最新檔案
         If DelAttFile_File("", Trim(pNo), "" & rsQD.Fields("cpf02"), , , , True) = False Then
            Exit Sub
         End If
         rsQD.MoveNext
      Loop
      If pKind = "2" Then MsgBox "刪除完畢!", vbInformation
   Else
      If pKind = "2" Then
          MsgBox "無資料可供刪除!", vbInformation
      End If
   End If
   Set rsQD = Nothing
End Sub
