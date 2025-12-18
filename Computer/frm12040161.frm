VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm12040161 
   BorderStyle     =   1  '單線固定
   Caption         =   "核判表設定作業"
   ClientHeight    =   6050
   ClientLeft      =   50
   ClientTop       =   340
   ClientWidth     =   8950
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6050
   ScaleWidth      =   8950
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7500
      Top             =   0
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
            Picture         =   "frm12040161.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040161.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040161.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040161.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040161.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040161.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040161.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040161.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040161.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040161.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040161.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   520
      Left            =   0
      TabIndex        =   32
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   5460
      Left            =   30
      TabIndex        =   33
      Top             =   570
      Width           =   8900
      _ExtentX        =   15699
      _ExtentY        =   9631
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "單筆資料"
      TabPicture(0)   =   "frm12040161.frx":20F4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label10"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(3)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "textPP02_2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cboPP04"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cboPP05"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "LblNoteNA01"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "GRD2"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "textPP02"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cboType"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "TextDept_2"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "TextDept"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Command1"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Frame_P"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Frame_T"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Frame1"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "FramePP0306"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "FrameSys"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Frame_CF"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).ControlCount=   21
      TabCaption(1)   =   "多筆瀏覽"
      TabPicture(1)   =   "frm12040161.frx":2110
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label16"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label15"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Line4"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Line5"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "GRD1"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdok"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "txt1(3)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "txt1(2)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "txt1(1)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "txt1(0)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).ControlCount=   10
      Begin VB.Frame Frame_CF 
         Appearance      =   0  '平面
         Caption         =   "外商CF核判分類"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   3075
         Left            =   2310
         TabIndex        =   61
         Top             =   2970
         Visible         =   0   'False
         Width           =   3465
         Begin VB.CommandButton Command4 
            BackColor       =   &H00C0C0FF&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Left            =   1950
            Style           =   1  '圖片外觀
            TabIndex        =   69
            Top             =   150
            Width           =   370
         End
         Begin VB.CheckBox Chk_CF 
            Caption         =   "E 提申"
            Height          =   345
            Index           =   4
            Left            =   510
            TabIndex        =   66
            Top             =   1612
            Width           =   2085
         End
         Begin VB.CheckBox Chk_CF 
            Caption         =   "D 其他來函"
            Height          =   345
            Index           =   3
            Left            =   510
            TabIndex        =   65
            Top             =   1290
            Width           =   2085
         End
         Begin VB.CheckBox Chk_CF 
            Caption         =   "C 審定爭議來函"
            Height          =   345
            Index           =   2
            Left            =   510
            TabIndex        =   64
            Top             =   926
            Width           =   2085
         End
         Begin VB.CheckBox Chk_CF 
            Caption         =   "B 非申請"
            Height          =   345
            Index           =   1
            Left            =   510
            TabIndex        =   63
            Top             =   583
            Width           =   1305
         End
         Begin VB.CheckBox Chk_CF 
            Caption         =   "A 申請"
            Height          =   345
            Index           =   0
            Left            =   510
            TabIndex        =   62
            Top             =   240
            Width           =   1305
         End
      End
      Begin VB.Frame FrameSys 
         Height          =   310
         Left            =   3120
         TabIndex        =   57
         Top             =   690
         Width           =   1270
         Begin VB.TextBox textPP01 
            Height          =   270
            Left            =   600
            MaxLength       =   3
            TabIndex        =   2
            Top             =   0
            Width           =   530
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "系統別"
            ForeColor       =   &H00004000&
            Height          =   180
            Index           =   1
            Left            =   0
            TabIndex        =   58
            Top             =   60
            Width           =   540
         End
      End
      Begin VB.Frame FramePP0306 
         Height          =   700
         Left            =   5970
         TabIndex        =   54
         Top             =   660
         Width           =   2860
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "案件性質"
            ForeColor       =   &H00004000&
            Height          =   180
            Left            =   30
            TabIndex        =   56
            Top             =   90
            Width           =   720
         End
         Begin MSForms.ComboBox cboPP03 
            Height          =   300
            Left            =   780
            TabIndex        =   3
            Top             =   30
            Width           =   2000
            VariousPropertyBits=   679495707
            DisplayStyle    =   3
            Size            =   "3519;529"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "申請國家"
            ForeColor       =   &H00004000&
            Height          =   180
            Left            =   30
            TabIndex        =   55
            Top             =   450
            Width           =   720
         End
         Begin MSForms.ComboBox cboPP06 
            Height          =   300
            Left            =   780
            TabIndex        =   6
            Top             =   390
            Width           =   2000
            VariousPropertyBits=   679495707
            DisplayStyle    =   3
            Size            =   "3519;529"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Enabled         =   0   'False
         Height          =   255
         Left            =   6120
         TabIndex        =   48
         Top             =   360
         Width           =   2720
         Begin VB.OptionButton OptDept 
            Caption         =   "外商CF"
            Height          =   225
            Index           =   2
            Left            =   1590
            TabIndex        =   60
            Top             =   30
            Width           =   860
         End
         Begin VB.OptionButton OptDept 
            Caption         =   "商標"
            Height          =   225
            Index           =   1
            Left            =   840
            TabIndex        =   50
            Top             =   30
            Width           =   735
         End
         Begin VB.OptionButton OptDept 
            Caption         =   "專利"
            Height          =   225
            Index           =   0
            Left            =   60
            TabIndex        =   49
            Top             =   30
            Width           =   735
         End
      End
      Begin VB.Frame Frame_T 
         Appearance      =   0  '平面
         Caption         =   "商標處核判分類"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   3075
         Left            =   4410
         TabIndex        =   47
         Top             =   1320
         Visible         =   0   'False
         Width           =   3465
         Begin VB.CommandButton Command3 
            BackColor       =   &H00C0C0FF&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Left            =   2130
            Style           =   1  '圖片外觀
            TabIndex        =   68
            Top             =   150
            Width           =   370
         End
         Begin VB.CheckBox Chk_T 
            Caption         =   "F 商申分析"
            Height          =   345
            Index           =   5
            Left            =   510
            TabIndex        =   53
            Top             =   1955
            Width           =   2085
         End
         Begin VB.CheckBox Chk_T 
            Caption         =   "G 商爭分析"
            Height          =   345
            Index           =   6
            Left            =   510
            TabIndex        =   52
            Top             =   2298
            Width           =   2085
         End
         Begin VB.CheckBox Chk_T 
            Caption         =   "H CMT申請"
            Height          =   345
            Index           =   7
            Left            =   510
            TabIndex        =   51
            Top             =   2640
            Width           =   2085
         End
         Begin VB.CheckBox Chk_T 
            Caption         =   "A 商申申請"
            Height          =   345
            Index           =   0
            Left            =   510
            TabIndex        =   22
            Top             =   240
            Width           =   1305
         End
         Begin VB.CheckBox Chk_T 
            Caption         =   "B 商申延期"
            Height          =   345
            Index           =   1
            Left            =   510
            TabIndex        =   23
            Top             =   583
            Width           =   1305
         End
         Begin VB.CheckBox Chk_T 
            Caption         =   "C 商爭智慧財產局"
            Height          =   345
            Index           =   2
            Left            =   510
            TabIndex        =   24
            Top             =   926
            Width           =   2085
         End
         Begin VB.CheckBox Chk_T 
            Caption         =   "D 商爭經濟部"
            Height          =   345
            Index           =   3
            Left            =   510
            TabIndex        =   25
            Top             =   1269
            Width           =   2085
         End
         Begin VB.CheckBox Chk_T 
            Caption         =   "E CMT變更"
            Height          =   345
            Index           =   4
            Left            =   510
            TabIndex        =   26
            Top             =   1612
            Width           =   2085
         End
      End
      Begin VB.Frame Frame_P 
         Appearance      =   0  '平面
         Caption         =   "專利處核判分類"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   2685
         Left            =   1320
         TabIndex        =   46
         Top             =   1320
         Visible         =   0   'False
         Width           =   3465
         Begin VB.CommandButton Command2 
            BackColor       =   &H00C0C0FF&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Left            =   3000
            Style           =   1  '圖片外觀
            TabIndex        =   67
            Top             =   150
            Width           =   370
         End
         Begin VB.CheckBox Chk_P 
            Caption         =   "M 對內信函"
            Height          =   345
            Index           =   13
            Left            =   1860
            TabIndex        =   21
            Top             =   2220
            Width           =   1335
         End
         Begin VB.CheckBox Chk_P 
            Caption         =   "L 對外信函"
            Height          =   345
            Index           =   12
            Left            =   1860
            TabIndex        =   20
            Top             =   1890
            Width           =   1335
         End
         Begin VB.CheckBox Chk_P 
            Caption         =   "K 比對分析報告"
            Height          =   345
            Index           =   11
            Left            =   1860
            TabIndex        =   19
            Top             =   1575
            Width           =   1550
         End
         Begin VB.CheckBox Chk_P 
            Caption         =   "J 爭議救濟"
            Height          =   345
            Index           =   10
            Left            =   1860
            TabIndex        =   18
            Top             =   1245
            Width           =   1335
         End
         Begin VB.CheckBox Chk_P 
            Caption         =   "I 答辯"
            Height          =   345
            Index           =   9
            Left            =   1860
            TabIndex        =   17
            Top             =   915
            Width           =   1335
         End
         Begin VB.CheckBox Chk_P 
            Caption         =   "H 設計案"
            Height          =   345
            Index           =   8
            Left            =   1860
            TabIndex        =   16
            Top             =   600
            Width           =   1335
         End
         Begin VB.CheckBox Chk_P 
            Caption         =   "G 申請案"
            Height          =   345
            Index           =   7
            Left            =   1860
            TabIndex        =   15
            Top             =   270
            Width           =   1335
         End
         Begin VB.CheckBox Chk_P 
            Caption         =   "F 比對分析報告"
            Height          =   345
            Index           =   6
            Left            =   300
            TabIndex        =   14
            Top             =   2220
            Width           =   1520
         End
         Begin VB.CheckBox Chk_P 
            Caption         =   "E 爭議案"
            Height          =   345
            Index           =   5
            Left            =   300
            TabIndex        =   13
            Top             =   1895
            Width           =   1305
         End
         Begin VB.CheckBox Chk_P 
            Caption         =   "D 救濟案"
            Height          =   345
            Index           =   4
            Left            =   300
            TabIndex        =   12
            Top             =   1570
            Width           =   1305
         End
         Begin VB.CheckBox Chk_P 
            Caption         =   "C 再審查"
            Height          =   345
            Index           =   3
            Left            =   300
            TabIndex        =   11
            Top             =   1245
            Width           =   1305
         End
         Begin VB.CheckBox Chk_P 
            Caption         =   "B 設計"
            Height          =   345
            Index           =   2
            Left            =   300
            TabIndex        =   10
            Top             =   920
            Width           =   1305
         End
         Begin VB.CheckBox Chk_P 
            Caption         =   "P 新型"
            Height          =   345
            Index           =   1
            Left            =   300
            TabIndex        =   9
            Top             =   595
            Width           =   1305
         End
         Begin VB.CheckBox Chk_P 
            Caption         =   "A 發明"
            Height          =   345
            Index           =   0
            Left            =   300
            TabIndex        =   8
            Top             =   270
            Width           =   1305
         End
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H000080FF&
         Caption         =   "全部的核判分類"
         Height          =   315
         Left            =   4410
         Style           =   1  '圖片外觀
         TabIndex        =   7
         Top             =   660
         Width           =   1575
      End
      Begin VB.TextBox TextDept 
         BackColor       =   &H8000000F&
         Height          =   270
         Left            =   3720
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   41
         Top             =   390
         Width           =   465
      End
      Begin VB.TextBox TextDept_2 
         Appearance      =   0  '平面
         BackColor       =   &H8000000A&
         BorderStyle     =   0  '沒有框線
         Height          =   270
         Left            =   4230
         Locked          =   -1  'True
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   390
         Width           =   1575
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   0
         Left            =   -73980
         MaxLength       =   6
         TabIndex        =   27
         Top             =   390
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   1
         Left            =   -72930
         MaxLength       =   6
         TabIndex        =   28
         Top             =   390
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   2
         Left            =   -70830
         MaxLength       =   3
         TabIndex        =   29
         Top             =   390
         Width           =   555
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   3
         Left            =   -70140
         MaxLength       =   3
         TabIndex        =   30
         Top             =   390
         Width           =   555
      End
      Begin VB.CommandButton cmdok 
         Caption         =   "查詢"
         Height          =   285
         Left            =   -68670
         TabIndex        =   31
         Top             =   390
         Width           =   915
      End
      Begin VB.ComboBox cboType 
         Height          =   260
         ItemData        =   "frm12040161.frx":212C
         Left            =   960
         List            =   "frm12040161.frx":212E
         TabIndex        =   1
         Top             =   690
         Width           =   1725
      End
      Begin VB.TextBox textPP02 
         Height          =   270
         Left            =   960
         MaxLength       =   6
         TabIndex        =   0
         Top             =   390
         Width           =   735
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Bindings        =   "frm12040161.frx":2130
         Height          =   4670
         Left            =   -74970
         TabIndex        =   34
         Top             =   750
         Width           =   8790
         _ExtentX        =   15505
         _ExtentY        =   8237
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         FormatString    =   "部門代碼|部門名稱|所別|員工編號|姓名"
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD2 
         Bindings        =   "frm12040161.frx":2145
         Height          =   3860
         Left            =   60
         TabIndex        =   43
         Top             =   1560
         Width           =   8730
         _ExtentX        =   15399
         _ExtentY        =   6809
         _Version        =   393216
         Cols            =   11
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         FormatString    =   "所別|員編|姓名|核判分類|核稿人1|離職|核稿人2|離職|判發人|離職|筆數"
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
         _Band(0).Cols   =   11
      End
      Begin VB.Label LblNoteNA01 
         Caption         =   "* 代表國家沒指定"
         ForeColor       =   &H00800000&
         Height          =   160
         Left            =   7350
         TabIndex        =   59
         Top             =   1380
         Width           =   1450
      End
      Begin MSForms.ComboBox cboPP05 
         Height          =   300
         Left            =   3720
         TabIndex        =   5
         Top             =   1020
         Width           =   2000
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "3528;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox cboPP04 
         Height          =   300
         Left            =   960
         TabIndex        =   4
         Top             =   1020
         Width           =   1995
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "3519;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textPP02_2 
         Height          =   270
         Left            =   1710
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   390
         Width           =   1190
         VariousPropertyBits=   671105055
         BackColor       =   -2147483633
         BorderStyle     =   1
         Size            =   "2099;476"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "判發人"
         ForeColor       =   &H000000C0&
         Height          =   180
         Left            =   3120
         TabIndex        =   45
         Top             =   1080
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "核稿人"
         ForeColor       =   &H000000C0&
         Height          =   180
         Left            =   360
         TabIndex        =   44
         Top             =   1080
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "部門代碼"
         Height          =   180
         Index           =   3
         Left            =   2970
         TabIndex        =   42
         Top             =   420
         Width           =   720
      End
      Begin VB.Line Line5 
         X1              =   -70440
         X2              =   -69840
         Y1              =   510
         Y2              =   510
      End
      Begin VB.Line Line4 
         X1              =   -73320
         X2              =   -72630
         Y1              =   510
         Y2              =   510
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "員工編號："
         Height          =   180
         Left            =   -74910
         TabIndex        =   38
         Top             =   420
         Width           =   900
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "部門代號："
         Height          =   180
         Left            =   -71760
         TabIndex        =   37
         Top             =   420
         Width           =   900
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "核判分類"
         Height          =   180
         Left            =   210
         TabIndex        =   36
         Top             =   750
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "員工編號"
         Height          =   180
         Index           =   0
         Left            =   210
         TabIndex        =   35
         Top             =   435
         Width           =   720
      End
   End
End
Attribute VB_Name = "frm12040161"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2021/12/13 Form2.0已修改(textPP02_2,cboPP04,cboPP05,cboPP06,GRD1及GRD2改Fonts)
'Create by Sindy 2018/3/23
Option Explicit

' 變數宣告區
Dim m_EditMode As Integer
'(執行各項功能的權限)
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
' 第一筆資料的本所案號
Dim m_FirstKEY(1) As String
' 最後一筆資料的本所案號
Dim m_LastKEY(1) As String
' 目前正在顯示的本所案號
Dim m_CurrKEY(1) As String
'Modify By Sindy 2024/7/8  鑑定報告 => 比對分析報告
Const P_Sort As String = "decode(cpm23,'P','A1(P)',cpm23)"
Const T_Sort As String = "decode(cpm23,'H','E1(H)',cpm23)"
Dim oChk As CheckBox
Dim m_CPM01List As String '核判的系統別
Dim m_ByCPMSet As Boolean, strPorType As String
Dim tmpArr As Variant


Private Sub cboPP04_GotFocus()
   InverseTextBox cboPP04
End Sub

'modify by sonia 2022/1/10
'Private Sub cboPP04_KeyPress(KeyAscii As Integer)
Private Sub cboPP04_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub cboPP04_Validate(Cancel As Boolean)
   If cboPP04.Text <> "" Then
      '檢查人員是否存在或離職
      If ChkStaffST04(Left(cboPP04, 5)) = True Then
         Call cboPP04_GotFocus
         Cancel = True
         Exit Sub
      End If
   End If
End Sub

Private Sub cboPP05_GotFocus()
   InverseTextBox cboPP05
End Sub

'modify by sonia 2022/1/10
'Private Sub cboPP05_KeyPress(KeyAscii As Integer)
Private Sub cboPP05_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub cboPP05_Validate(Cancel As Boolean)
   If cboPP05.Text <> "" Then
      '檢查人員是否存在或離職
      If ChkStaffST04(Left(cboPP05, 5)) = True Then
         Call cboPP05_GotFocus
         Cancel = True
         Exit Sub
      End If
   End If
End Sub

Private Sub cboPP06_GotFocus()
   InverseTextBox cboPP06
End Sub

'modify by sonia 2022/1/10
'Private Sub cboPP06_KeyPress(KeyAscii As Integer)
Private Sub cboPP06_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'核判分類
Private Sub cboType_Click()
   If OptDept(0).Value = True Then  '專利處
      'If Left(cboType.Text, 1) >= "A" And Left(cboType.Text, 1) <= "F" Then
      If Left(cboType.Text, 1) = "A" Or _
         Left(cboType.Text, 1) = "P" Or _
         Left(cboType.Text, 1) = "B" Or _
         Left(cboType.Text, 1) = "C" Or _
         Left(cboType.Text, 1) = "D" Or _
         Left(cboType.Text, 1) = "E" Or _
         Left(cboType.Text, 1) = "F" Then
         m_CPM01List = "'P','PS'"
      ElseIf Left(cboType.Text, 1) = "G" Or _
         Left(cboType.Text, 1) = "H" Or _
         Left(cboType.Text, 1) = "I" Or _
         Left(cboType.Text, 1) = "J" Or _
         Left(cboType.Text, 1) = "K" Or _
         Left(cboType.Text, 1) = "L" Or _
         Left(cboType.Text, 1) = "M" Then
         m_CPM01List = "'CFP','CPS'"
      Else
         MsgBox "系統別有誤!", vbExclamation
      End If
   ElseIf OptDept(1).Value = True Then
      m_CPM01List = "'T'"
   'Add Sindy 2024/9/16
   ElseIf OptDept(2).Value = True Then
      m_CPM01List = "'CFT','CFC','S'"
   '2024/9/16 END
   End If
End Sub

Private Sub cboType_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub cmdok_Click()
If txt1(0) & txt1(1) & txt1(2) & txt1(3) <> "" Then
    If RunNick(txt1(0), txt1(1)) Then
        txt1(0).SetFocus
        Exit Sub
    End If
    If RunNick2(txt1(2), txt1(3)) Then
        txt1(2).SetFocus
        Exit Sub
    End If
    GetData
Else
    If m_EditMode <> 1 Then
      MsgBox "查詢條件不可以空白！", vbExclamation, "操作錯誤！"
    End If
End If
End Sub

'全部的核判分類
Private Sub Command1_Click()
   If OptDept(0).Value = True Then '專利處
      Frame_P.Visible = True
      Frame_T.Visible = False
      Frame_CF.Visible = False
      cboType.Enabled = False
      Command2.Visible = False
   ElseIf OptDept(1).Value = True Then
      Frame_P.Visible = False
      Frame_T.Visible = True
      Frame_CF.Visible = False
      cboType.Enabled = False
      Command3.Visible = False
   'Add Sindy 2024/9/16
   ElseIf OptDept(2).Value = True Then
      Frame_P.Visible = False
      Frame_T.Visible = False
      Frame_CF.Visible = True
      cboType.Enabled = False
      Command4.Visible = False
   '2024/9/16 END
   End If
End Sub

Private Sub Command2_Click()
   Frame_P.Visible = False
End Sub

Private Sub Command3_Click()
   Frame_T.Visible = False
End Sub

Private Sub Command4_Click()
   Frame_CF.Visible = False
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

   m_bInsert = IsUserHasRightOfFunction(Me.Name, strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
   m_bDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   m_bQuery = IsUserHasRightOfFunction(Me.Name, strFind, False)
   
   textPP02.BackColor = &H8000000F
   MoveFormToCenter Me
   
   Command1.Enabled = False
   RefreshRange
   ShowFirstRecord
   UpdateToolbarState
   SetCtrlReadOnly True
   Me.SSTab1.Tab = 0
   
   Frame1.BorderStyle = 0
   FrameSys.BorderStyle = 0
   FramePP0306.BorderStyle = 0
   textPP02_2.BorderStyle = 0
   TextDept_2.BackColor = &H8000000F
   
   'Add By Sindy 2024/9/16
   Frame_T.Top = Frame_P.Top
   Frame_T.Left = Frame_P.Left
   Frame_CF.Top = Frame_P.Top
   Frame_CF.Left = Frame_P.Left
   'Sindy 2024/9/16 END
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm12040161 = Nothing
End Sub

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long

   getGrdColRow GRD1, x, y, nCol, nRow
   GRD1.col = nCol
   GRD1.row = nRow
End Sub

Private Sub grd1_SelChange()
Dim tmpMouseRow
Dim i, j

   GRD1.Visible = False
   tmpMouseRow = GRD1.row
   GRD1.Visible = True
   If tmpMouseRow <> 0 Then
      GRD1.row = tmpMouseRow
      GRD1.col = 0
      'If grd1.CellBackColor <> &HFFC0C0 Then
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
         textPP02.Text = GRD1.TextMatrix(tmpMouseRow, 3)
         Call QueryRecord
         GRD1.Visible = True
      'End If
   End If
End Sub

Private Sub grd2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long

   getGrdColRow grd2, x, y, nCol, nRow
   grd2.col = nCol
   If nRow < 0 Then Exit Sub
   grd2.row = nRow
End Sub

Private Sub GRD2_SelChange()
Dim tmpMouseRow As Integer
Dim i, j
   
   grd2.Visible = False
   tmpMouseRow = grd2.row
   grd2.Visible = True
   If tmpMouseRow <> 0 Then
      grd2.row = tmpMouseRow
      grd2.col = 0
      If grd2.CellBackColor <> &HFFC0C0 Then
         grd2.Visible = False
         For j = 1 To grd2.Rows - 1
            grd2.row = j
            For i = 0 To grd2.Cols - 1
                 grd2.col = i
                 grd2.CellBackColor = QBColor(15)
            Next i
         Next j
         grd2.row = tmpMouseRow
         For i = 0 To grd2.Cols - 1
             grd2.col = i
             grd2.CellBackColor = &HFFC0C0
         Next i
         '讀取資料
         Call GetRowData(tmpMouseRow)
         grd2.Visible = True
      End If
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

Private Function TxtValidate() As Boolean
Dim Cancel As Boolean
Dim bolChk As Boolean
   
   TxtValidate = False
   
   If Me.textPP02.Enabled = True Then
      Cancel = False
      textPP02_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If textPP02.Text = "" Then
      MsgBox "員工編號不可以空白！", vbExclamation
      textPP02.SetFocus
      Exit Function
   End If
   
   If Frame1.Visible = False Then
      If textPP01 = "" Then
         MsgBox "系統別不可以空白！", vbExclamation
         Exit Function
      End If
      
      If cboPP03 = "" Then
         MsgBox "案件性質不可以空白！", vbExclamation
         Exit Function
      End If
      
      If cboPP06 = "" Then
         MsgBox "申請國家不可以空白！", vbExclamation
         Exit Function
      End If
   Else
      'Modify By Sindy 2024/9/16 + And Frame_CF.Visible = False
      If cboType.Enabled = True And cboType.Text = "" _
         And Frame_P.Visible = False And Frame_T.Visible = False And Frame_CF.Visible = False Then
         MsgBox "核判分類不可以空白！", vbExclamation
         Exit Function
      Else
         If Frame_P.Visible = True Then
            bolChk = False
            For Each oChk In Chk_P
               If oChk.Value = 1 Then
                  bolChk = True
                  Exit For
               End If
            Next
         ElseIf Frame_T.Visible = True Then
            bolChk = False
            For Each oChk In Chk_T
               If oChk.Value = 1 Then
                  bolChk = True
                  Exit For
               End If
            Next
         'Add By Sindy 2024/9/16
         ElseIf Frame_CF.Visible = True Then
            bolChk = False
            For Each oChk In Chk_CF
               If oChk.Value = 1 Then
                  bolChk = True
                  Exit For
               End If
            Next
         '2024/9/16 END
         End If
         If Frame_P.Visible = True Or Frame_T.Visible = True Or Frame_CF.Visible = True Then
            If bolChk = False Then
               MsgBox "核判分類至少勾選一項！", vbExclamation
               Exit Function
            End If
         End If
      End If
   End If
   
   If cboPP04.Text = "" And cboPP05.Text = "" Then
      MsgBox "核判人員不可以空白！", vbExclamation
      Exit Function
   End If
   
   If Me.cboPP04.Enabled = True Then
      Cancel = False
      cboPP04_Validate Cancel
      If Cancel = True Then
         cboPP04.SetFocus
         Exit Function
      End If
   End If
   If Me.cboPP05.Enabled = True Then
      Cancel = False
      cboPP05_Validate Cancel
      If Cancel = True Then
         cboPP05.SetFocus
         Exit Function
      End If
   End If
   
   If m_EditMode = 1 Then '新增時
      'Add By Sindy 2024/6/24
      If m_ByCPMSet = True Then
         tmpArr = Split(cboPP03.Text, " ")
         If IsRecordExist(textPP02, tmpArr(0)) = True Then
            MsgBox "此筆核判資料已存在，不可新增！", vbExclamation
            Exit Function
         End If
      '2024/6/24 END
      Else
         If cboType.Enabled = True And cboType.Text <> "" Then
            If IsRecordExist(textPP02, Left(cboType.Text, 1)) = True Then '資料已存在
               MsgBox "此筆核判資料已存在，不可新增！", vbExclamation
               Exit Function
            End If
         End If
      End If
   End If
   
   TxtValidate = True
End Function

' 新增記錄
Private Function AddRecord() As Boolean
Dim strPP02 As String
Dim strPP03 As String
   
   AddRecord = False
   
   strPP02 = textPP02
   
On Error GoTo ErrHand
   
   cnnConnection.BeginTrans
   
   'Add By Sindy 2024/6/24
   If m_ByCPMSet = True Then
      tmpArr = Split(cboPP03.Text, " ")
      strPP03 = tmpArr(0)
      Call DelDataSql(strPP02, strPP03)
      Call InsDataSql(strPP02, strPP03)
   '2024/6/24 END
   'Modify By Sindy 2024/9/16 + And Frame_CF.Visible = False
   ElseIf cboType.Enabled = True And Frame_P.Visible = False And Frame_T.Visible = False And Frame_CF.Visible = False Then
      strPP03 = Left(cboType.Text, 1)
      Call DelDataSql(strPP02, strPP03)
      Call InsDataSql(strPP02, strPP03)
   Else
      If Frame_P.Visible = True Then
         For Each oChk In Chk_P
            If oChk.Value = 1 And oChk.Enabled = True Then
               If Left(oChk.Caption, 1) = "A" Or _
                  Left(oChk.Caption, 1) = "P" Or _
                  Left(oChk.Caption, 1) = "B" Or _
                  Left(oChk.Caption, 1) = "C" Or _
                  Left(oChk.Caption, 1) = "D" Or _
                  Left(oChk.Caption, 1) = "E" Or _
                  Left(oChk.Caption, 1) = "F" Then
                  m_CPM01List = "'P','PS'"
               ElseIf Left(oChk.Caption, 1) = "G" Or _
                  Left(oChk.Caption, 1) = "H" Or _
                  Left(oChk.Caption, 1) = "I" Or _
                  Left(oChk.Caption, 1) = "J" Or _
                  Left(oChk.Caption, 1) = "K" Or _
                  Left(oChk.Caption, 1) = "L" Or _
                  Left(oChk.Caption, 1) = "M" Then
                  m_CPM01List = "'CFP','CPS'"
               Else
                  MsgBox "系統別有誤!", vbExclamation
               End If
               
               strPP03 = Left(oChk.Caption, 1)
               Call DelDataSql(strPP02, strPP03)
               Call InsDataSql(strPP02, strPP03)
            End If
         Next
         Frame_P.Visible = False
      ElseIf Frame_T.Visible = True Then
         For Each oChk In Chk_T
            If oChk.Value = 1 And oChk.Enabled = True Then
               m_CPM01List = "'T'"
               strPP03 = Left(oChk.Caption, 1)
               Call DelDataSql(strPP02, strPP03)
               Call InsDataSql(strPP02, strPP03)
            End If
         Next
         Frame_T.Visible = False
      'Add By Sindy 2024/9/16
      ElseIf Frame_CF.Visible = True Then
         For Each oChk In Chk_CF
            If oChk.Value = 1 And oChk.Enabled = True Then
               m_CPM01List = "'CFT','CFC','S'"
               strPP03 = Left(oChk.Caption, 1)
               Call DelDataSql(strPP02, strPP03)
               Call InsDataSql(strPP02, strPP03)
            End If
         Next
         Frame_CF.Visible = False
      '2024/9/16 END
      End If
   End If
   
   cnnConnection.CommitTrans
   
   If strPP02 < m_FirstKEY(0) Or strPP02 > m_LastKEY(0) Then
      RefreshRange
   End If
   m_EditMode = 0 'Add By Sindy 2024/9/18
   ShowCurrRecord strPP02
   
   AddRecord = True
   Exit Function
   
ErrHand:
   cnnConnection.RollbackTrans
   MsgBox " 新增失敗！" & vbCrLf & Err.Description
End Function

' 刪除語法
Private Function DelDataSql(strPP02 As String, strPP03 As String) As Boolean
   'Add By Sindy 2024/6/21
   If m_ByCPMSet = True Then
      strSql = "delete From promoterproofreader where PP01='" & textPP01 & "' AND PP02='" & strPP02 & "' AND PP03='" & strPP03 & "' AND PP06='" & Trim(Left(cboPP06, 3)) & "'"
      cnnConnection.Execute strSql, intI
   Else
   '2024/6/21 END
      strSql = "delete From promoterproofreader" & _
               " where PP02='" & strPP02 & "' AND (PP01,PP03) IN (SELECT CPM01,CPM02 FROM CASEPROPERTYMAP WHERE CPM23='" & strPP03 & "' and cpm01 in(" & m_CPM01List & "))"
      cnnConnection.Execute strSql, intI
      strSql = "delete From promoterproofreader where PP02='" & strPP02 & "' AND PP03='" & strPP03 & "'"
      cnnConnection.Execute strSql, intI
   End If
End Function

' 新增語法
Private Function InsDataSql(strPP02 As String, strPP03 As String) As Boolean
   
   'Add By Sindy 2024/6/21
   If m_ByCPMSet = True Then
      tmpArr = Split(cboPP03.Text, " ")
      strSql = "INSERT INTO PROMOTERPROOFREADER values('" & textPP01 & "','" & strPP02 & "','" & Trim(tmpArr(0)) & "'," & CNULL(Trim(Left(cboPP04, 6))) & "," & CNULL(Trim(Left(cboPP05, 6))) & "," & CNULL(Trim(Left(cboPP06, 3))) & ")"
      cnnConnection.Execute strSql, intI
   Else
   '2024/6/21 END
      '自行核判
      If strPP02 = Trim(Left(cboPP04, 6)) And strPP02 = Trim(Left(cboPP05, 6)) Then
         If OptDept(0).Value = True Then '專利處
            If (Left(cboType.Text, 1) >= "A" And Left(cboType.Text, 1) <= "F") Or Left(cboType.Text, 1) = "P" Or _
               m_CPM01List = "'P','PS'" Then
               strSql = "INSERT INTO PROMOTERPROOFREADER values('P','" & strPP02 & "','" & strPP03 & "'," & CNULL(Trim(Left(cboPP04, 6))) & "," & CNULL(Trim(Left(cboPP05, 6))) & ",'*')"
               cnnConnection.Execute strSql, intI
               strSql = "INSERT INTO PROMOTERPROOFREADER values('PS','" & strPP02 & "','" & strPP03 & "'," & CNULL(Trim(Left(cboPP04, 6))) & "," & CNULL(Trim(Left(cboPP05, 6))) & ",'*')"
               cnnConnection.Execute strSql, intI
            Else
               strSql = "INSERT INTO PROMOTERPROOFREADER values('CFP','" & strPP02 & "','" & strPP03 & "'," & CNULL(Trim(Left(cboPP04, 6))) & "," & CNULL(Trim(Left(cboPP05, 6))) & ",'*')"
               cnnConnection.Execute strSql, intI
               strSql = "INSERT INTO PROMOTERPROOFREADER values('CPS','" & strPP02 & "','" & strPP03 & "'," & CNULL(Trim(Left(cboPP04, 6))) & "," & CNULL(Trim(Left(cboPP05, 6))) & ",'*')"
               cnnConnection.Execute strSql, intI
            End If
         ElseIf OptDept(1).Value = True Then '商標處
            strSql = "INSERT INTO PROMOTERPROOFREADER values('T','" & strPP02 & "','" & strPP03 & "'," & CNULL(Trim(Left(cboPP04, 6))) & "," & CNULL(Trim(Left(cboPP05, 6))) & ",'*')"
            cnnConnection.Execute strSql, intI
         'Add By Sindy 2024/9/16
         ElseIf OptDept(2).Value = True Then '外商CF
            strSql = "INSERT INTO PROMOTERPROOFREADER values('CFT','" & strPP02 & "','" & strPP03 & "'," & CNULL(Trim(Left(cboPP04, 6))) & "," & CNULL(Trim(Left(cboPP05, 6))) & ",'*')"
            cnnConnection.Execute strSql, intI
            strSql = "INSERT INTO PROMOTERPROOFREADER values('CFC','" & strPP02 & "','" & strPP03 & "'," & CNULL(Trim(Left(cboPP04, 6))) & "," & CNULL(Trim(Left(cboPP05, 6))) & ",'*')"
            cnnConnection.Execute strSql, intI
            strSql = "INSERT INTO PROMOTERPROOFREADER values('S','" & strPP02 & "','" & strPP03 & "'," & CNULL(Trim(Left(cboPP04, 6))) & "," & CNULL(Trim(Left(cboPP05, 6))) & ",'*')"
            cnnConnection.Execute strSql, intI
         '2024/9/16 END
         End If
      
      Else
         If OptDept(0).Value = True Then '專利處
            strSql = "INSERT INTO PROMOTERPROOFREADER (SELECT CPM01,'" & strPP02 & "',CPM02," & CNULL(Trim(Left(cboPP04, 6))) & "," & CNULL(Trim(Left(cboPP05, 6))) & ",'*' FROM CASEPROPERTYMAP WHERE CPM23='" & strPP03 & "' AND CPM01 in(" & m_CPM01List & "))"
            cnnConnection.Execute strSql, intI
         ElseIf OptDept(1).Value = True Then '商標處
            strSql = "SELECT CPM01 FROM CASEPROPERTYMAP WHERE CPM23='" & strPP03 & "' AND CPM01 in(" & m_CPM01List & ")"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               strSql = "INSERT INTO PROMOTERPROOFREADER (SELECT CPM01,'" & strPP02 & "',CPM02," & CNULL(Trim(Left(cboPP04, 6))) & "," & CNULL(Trim(Left(cboPP05, 6))) & ",'*' FROM CASEPROPERTYMAP WHERE CPM23='" & strPP03 & "' AND CPM01 in(" & m_CPM01List & "))"
               cnnConnection.Execute strSql, intI
               'Add By Sindy 2025/5/7 有些案件性質直接判斷為F=商申分析 或 G=商爭分析
               If strPP03 = "F" Or strPP03 = "G" Then
                  strSql = "INSERT INTO PROMOTERPROOFREADER values('T','" & strPP02 & "','" & strPP03 & "'," & CNULL(Trim(Left(cboPP04, 6))) & "," & CNULL(Trim(Left(cboPP05, 6))) & ",'*')"
                  cnnConnection.Execute strSql, intI
               End If
               '2025/5/7 END
            Else
               strSql = "INSERT INTO PROMOTERPROOFREADER values('T','" & strPP02 & "','" & strPP03 & "'," & CNULL(Trim(Left(cboPP04, 6))) & "," & CNULL(Trim(Left(cboPP05, 6))) & ",'*')"
               cnnConnection.Execute strSql, intI
            End If
         'Add By Sindy 2024/9/16
         ElseIf OptDept(2).Value = True Then '外商CF
            strSql = "INSERT INTO PROMOTERPROOFREADER (SELECT CPM01,'" & strPP02 & "',CPM02," & CNULL(Trim(Left(cboPP04, 6))) & "," & CNULL(Trim(Left(cboPP05, 6))) & ",'*' FROM CASEPROPERTYMAP WHERE CPM23='" & strPP03 & "' AND CPM01 in(" & m_CPM01List & "))"
            cnnConnection.Execute strSql, intI
            If intI = 0 Then
               strSql = "INSERT INTO PROMOTERPROOFREADER values('CFT','" & strPP02 & "','" & strPP03 & "'," & CNULL(Trim(Left(cboPP04, 6))) & "," & CNULL(Trim(Left(cboPP05, 6))) & ",'*')"
               cnnConnection.Execute strSql, intI
               strSql = "INSERT INTO PROMOTERPROOFREADER values('CFC','" & strPP02 & "','" & strPP03 & "'," & CNULL(Trim(Left(cboPP04, 6))) & "," & CNULL(Trim(Left(cboPP05, 6))) & ",'*')"
               cnnConnection.Execute strSql, intI
               strSql = "INSERT INTO PROMOTERPROOFREADER values('S','" & strPP02 & "','" & strPP03 & "'," & CNULL(Trim(Left(cboPP04, 6))) & "," & CNULL(Trim(Left(cboPP05, 6))) & ",'*')"
               cnnConnection.Execute strSql, intI
            End If
         '2024/9/16 END
         End If
      End If
   End If
End Function

' 修改記錄
Private Function ModRecord() As Boolean
Dim strPP02 As String
Dim strPP03 As String
   
   ModRecord = False
   
   strPP02 = textPP02
   'Add By Sindy 2024/6/24
   If m_ByCPMSet = True Then
      tmpArr = Split(cboPP03.Text, " ")
      strPP03 = tmpArr(0)
   Else
   '2024/6/24 END
      strPP03 = Left(cboType.Text, 1)
   End If
   
On Error GoTo ErrHand
   
   cnnConnection.BeginTrans
   
   Call DelDataSql(strPP02, strPP03)
   Call InsDataSql(strPP02, strPP03)
   
   cnnConnection.CommitTrans
   
   ShowCurrRecord strPP02
   ModRecord = True
   Exit Function
   
ErrHand:
   cnnConnection.RollbackTrans
   MsgBox (Err.Description)
End Function

' 刪除記錄
Private Function DelRecord() As Boolean
Dim strPP02 As String
Dim strPP03 As String
   
   DelRecord = False
   
   strPP02 = textPP02
   'Add By Sindy 2024/6/24
   If m_ByCPMSet = True Then
      tmpArr = Split(cboPP03.Text, " ")
      strPP03 = tmpArr(0)
   Else
   '2024/6/24 END
      strPP03 = Left(cboType.Text, 1)
   End If
   
On Error GoTo ErrHand
   
   cnnConnection.BeginTrans
   
   Call DelDataSql(strPP02, strPP03)
   
   cnnConnection.CommitTrans
   
   If strPP02 = m_LastKEY(0) Or strPP02 = m_FirstKEY(0) Then
      RefreshRange
   End If
   ShowCurrRecord strPP02
   DelRecord = True
   Exit Function
   
ErrHand:
   cnnConnection.RollbackTrans
   MsgBox "刪除失敗！" & vbCrLf & Err.Description
End Function

' 查詢記錄
Private Function QueryRecord() As Boolean
   Dim strPP02 As String
   
   strPP02 = textPP02
   
   m_CurrKEY(0) = strPP02
   
   QueryRecord = UpdateCtrlData

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
            If ModRecord = False Then Exit Function
         Else
            GoTo EXITSUB
         End If
      Case 3: '刪除
         If DelRecord = True Then
            'RefreshRange
            'ClearField
            'ShowCurrRecord m_CurrKEY(0)
            Me.SSTab1.TabEnabled(1) = True
            UpdateCtrlData
            SetCtrlReadOnly True
            UpdateToolbarState
         Else
            Exit Function
         End If
      Case 4: '查詢
         If textPP02 <> "" Then
            QueryRecord
'            If QueryRecord = False Then
'               strMsg = "無此資料"
'               strTit = "查詢資料"
'               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'               UpdateCtrlData
'            End If
         Else
            MsgBox "必須輸入員工編號才可進行查詢動作！", vbInformation
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
      Case 1: If Me.Visible = True Then textPP02.SetFocus
      Case 2: If Me.Visible = True Then cboPP04.SetFocus
      Case 4: If Me.Visible = True Then textPP02.SetFocus
   End Select
End Sub

' 檢查記錄是否已經存在 - 人 & 核判分類
Private Function IsRecordExist(ByVal strKEY01 As String, ByVal strKEY02 As String) As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   IsRecordExist = False
   'Add By Sindy 2024/6/24
   If m_ByCPMSet = True Then
      strSql = "select pp02,pp03" & _
               " From promoterproofreader" & _
               " where PP01='" & textPP01 & "' AND PP02='" & strKEY01 & "' AND PP03='" & strKEY02 & "' AND PP06='" & Trim(Left(cboPP06, 3)) & "'"
   Else
   '2024/6/24 END
      strSql = "select pp02,cpm23" & _
               " From promoterproofreader, casepropertymap" & _
               " where PP02='" & strKEY01 & "' AND (PP01,PP03) IN (SELECT CPM01,CPM02 FROM CASEPROPERTYMAP WHERE CPM23 IS NOT NULL AND CPM23<>'9')" & _
               " AND pp01=cpm01(+) and pp03=cpm02(+)" & _
               " and cpm23 is not null" & _
               " and cpm23='" & strKEY02 & "'" & _
               " Union " & _
               "select pp02,pp03" & _
               " From promoterproofreader" & _
               " where PP02='" & strKEY01 & "' AND PP03='" & strKEY02 & "'"
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

' 顯示資料
Private Sub ShowCurrRecord(ByVal strKEY01 As String)
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   If strKEY01 <> "" Then m_CurrKEY(0) = strKEY01 'Add By Sindy 2024/9/18
   strSql = "select pp02,a.st02,st03" & _
            " from promoterproofreader,staff a" & _
            " where pp02=a.st01(+)" & IIf(m_CurrKEY(0) <> "", " and pp02 = '" & m_CurrKEY(0) & "'", "") & _
            " group by st03,pp02,a.st02" & _
            " order by st03 asc,pp02 asc"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("pp02")) = False Then: m_CurrKEY(0) = rsTmp.Fields("pp02")
   Else
      ShowLastRecord
      GoTo EXITSUB
   End If
   rsTmp.Close
   UpdateCtrlData
EXITSUB:
End Sub

' 顯示第一筆資料
Private Sub ShowFirstRecord()
   m_CurrKEY(0) = m_FirstKEY(0)
   UpdateCtrlData
End Sub

' 顯示上一筆資料
Private Sub ShowPrevRecord()
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   If m_CurrKEY(0) = m_FirstKEY(0) Then
      ShowMsg MsgText(9008)
      GoTo EXITSUB
   End If
   
   strSql = "select pp02,a.st02,st03" & _
               " from promoterproofreader,staff a" & _
               " where pp02=a.st01(+) and a.st04='1'" & _
               IIf(m_CurrKEY(0) <> "", " and st03||pp02 < '" & PUB_GetST03(m_CurrKEY(0)) & m_CurrKEY(0) & "'", "") & _
               " group by st03,pp02,a.st02" & _
               " order by st03 asc,pp02 asc"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveLast
      If IsNull(rsTmp.Fields("pp02")) = False Then: m_CurrKEY(0) = rsTmp.Fields("pp02")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
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
   
   If m_CurrKEY(0) = m_LastKEY(0) Then
      ShowMsg MsgText(9009)
      GoTo EXITSUB
   End If
   
   strSql = "select pp02,a.st02,st03" & _
               " from promoterproofreader,staff a" & _
               " where pp02=a.st01(+) and a.st04='1'" & IIf(m_CurrKEY(0) <> "", " and st03||pp02 > '" & PUB_GetST03(m_CurrKEY(0)) & m_CurrKEY(0) & "'", "") & _
               " group by st03,pp02,a.st02" & _
               " order by st03 asc,pp02 asc"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      If IsNull(rsTmp.Fields("pp02")) = False Then: m_CurrKEY(0) = rsTmp.Fields("pp02")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   UpdateCtrlData
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 顯示最後一筆資料
Private Sub ShowLastRecord()
   m_CurrKEY(0) = m_LastKEY(0)
   UpdateCtrlData
End Sub

' 執行指令
Private Sub OnAction(ByVal KeyCode As Integer)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   If m_EditMode <> "1" Then Command1.Enabled = False: cboType.Enabled = False
   Select Case KeyCode
      ' 新增
      Case vbKeyF2:
         m_EditMode = 1
         GRD2_SelChange
         ClearField
         If Frame1.Visible = True Then Command1.Enabled = True
         Me.SSTab1.TabEnabled(1) = False
         SSTab1.Tab = 0
         SetCtrlReadOnly False
         UpdateToolbarState
         SetInputEntry
      ' 修改
      Case vbKeyF3:
         m_EditMode = 2
         Me.SSTab1.TabEnabled(1) = False
         SSTab1.Tab = 0
         SetCtrlReadOnly False
         SetKeyReadOnly True
         UpdateToolbarState
         SetInputEntry
      ' 刪除
      Case vbKeyF5:
         Me.SSTab1.TabEnabled(1) = False
         SSTab1.Tab = 0
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
         Else
            Me.SSTab1.TabEnabled(1) = True
            UpdateCtrlData
            SetCtrlReadOnly True
            UpdateToolbarState
         End If
      ' 查詢
      Case vbKeyF4:
         m_EditMode = 4
         cboType.Enabled = True
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
         If OnWork = True Then
            Me.SSTab1.TabEnabled(1) = True
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
                  Me.SSTab1.TabEnabled(1) = True
                  UpdateCtrlData
                  SetCtrlReadOnly True
                  UpdateToolbarState
               End If
            Case Else
               m_EditMode = 0
               Me.SSTab1.TabEnabled(1) = True
               UpdateCtrlData
               SetCtrlReadOnly True
               UpdateToolbarState
         End Select
      ' 離開
      Case vbKeyEscape:
         Unload Me
   End Select
End Sub

Private Sub RefreshRange()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   strSql = "select pp02,a.st02,st03" & _
            " from promoterproofreader,staff a" & _
            " where pp02=a.st01(+) and a.st04='1'" & _
            " group by st03,pp02,a.st02" & _
            " order by st03 asc,pp02 asc"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      If IsNull(rsTmp.Fields("pp02")) = False Then: m_FirstKEY(0) = rsTmp.Fields("pp02")
   End If
   rsTmp.Close
   
   strSql = "select pp02,a.st02,st03" & _
            " from promoterproofreader,staff a" & _
            " where pp02=a.st01(+) and a.st04='1'" & _
            " group by st03,pp02,a.st02" & _
            " order by st03 desc,pp02 desc"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      If IsNull(rsTmp.Fields("pp02")) = False Then: m_LastKEY(0) = rsTmp.Fields("pp02")
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Function UpdateCtrlData(Optional strKey As String = "") As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim i As Integer
Dim bolShowCommand1 As Boolean
   
   Call ClearField(False)
   If strKey = "" Then
      strKey = m_CurrKEY(0)
      Me.textPP02 = strKey
      Call textPP02_Validate(False)
   End If
   
   UpdateCtrlData = True
   Call SetComboList
   If m_ByCPMSet = False Then
      strSql = "select * from (select DECODE(A.ST06,'1','北所','2','中所','3','南所','4','高所','其他') as T1,pp02,a.st02 as T2,'',decode(cpm23," & IIf(strPorType = "P", P核判分類, IIf(strPorType = "T", T核判分類, CF核判分類)) & ",cpm23) as cpm23Nm,cpm23,pp06,pp04,b.st02 as T3,decode(b.st04,'2','Y','') as T4,pp05,c.st02 as T7,decode(c.st04,'2','Y','') as T8,COUNT(*)" & _
               " from promoterproofreader,staff a,staff b,staff c,casepropertymap" & _
               " where PP02='" & strKey & "' AND (PP01,PP03) IN (SELECT CPM01,CPM02 FROM CASEPROPERTYMAP WHERE CPM23 IS NOT NULL AND CPM23<>'9')" & _
               " and pp01=cpm01(+) and pp03=cpm02(+)" & _
               " and pp02=a.st01(+) and pp04=b.st01(+) and pp05=c.st01(+)" & _
               " and cpm23 is not null" & _
               " GROUP BY A.ST06,pp02,a.st02,decode(cpm23," & IIf(strPorType = "P", P核判分類, IIf(strPorType = "T", T核判分類, CF核判分類)) & ",cpm23),cpm23,pp04,pp06,b.st02,b.st04,pp05,c.st02,c.st04,c.st03" & _
               " Union " & _
               "select DECODE(A.ST06,'1','北所','2','中所','3','南所','4','高所','其他') as T1,pp02,a.st02 as T2,'',decode(pp03," & IIf(strPorType = "P", P核判分類, IIf(strPorType = "T", T核判分類, CF核判分類)) & ",pp03) as cpm23Nm,pp03 as cpm23,pp06,pp04,b.st02 as T3,decode(b.st04,'2','Y','') as T4,pp05,c.st02 as T7,decode(c.st04,'2','Y','') as T8,COUNT(*)" & _
               " from promoterproofreader,staff a,staff b,staff c" & _
               " where PP02='" & strKey & "' AND length(PP03)=1" & _
               " and pp02=a.st01(+) and pp04=b.st01(+) and pp05=c.st01(+)" & _
               " GROUP BY A.ST06,pp02,a.st02,decode(pp03," & IIf(strPorType = "P", P核判分類, IIf(strPorType = "T", T核判分類, CF核判分類)) & ",pp03),pp03,pp04,pp06,b.st02,b.st04,pp05,c.st02,c.st04,c.st03)" & _
               " order by " & IIf(strPorType = "P", P_Sort, IIf(strPorType = "T", T_Sort, "cpm23")) & " asc"
   Else
      strSql = "select DECODE(A.ST06,'1','北所','2','中所','3','南所','4','高所','其他') as T1,pp02,a.st02 as T2,pp01,pp03||' '||decode(pp06,'000',cpm03,decode(pp06,'*',cpm03,cpm04)),pp03,decode(pp06,'*',pp06,na01||' '||na03) pp06,pp04,b.st02 as T3,decode(b.st04,'2','Y','') as T4,pp05,c.st02 as T7,decode(c.st04,'2','Y','') as T8,1" & _
               " from promoterproofreader,staff a,staff b,staff c,casepropertymap,nation" & _
               " where PP02='" & strKey & "'" & _
               " and pp01=cpm01(+) and pp03=cpm02(+)" & _
               " and pp02=a.st01(+) and pp04=b.st01(+) and pp05=c.st01(+)" & _
               " and pp06=na01(+)" & _
               " order by pp06 asc,pp01 asc,pp03 asc"
   End If
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   Set grd2.Recordset = rsTmp
   Call SetGrd2(grd2)
   If rsTmp.RecordCount = 0 Then
      UpdateCtrlData = False
      grd2.col = 0
      grd2.row = 0
      If m_EditMode = 1 Then
         If Frame1.Visible = True Then Command1.Enabled = True
      Else
         MsgBox "查無資料！", vbInformation
      End If
   Else
      Me.SSTab1.Tab = 0
      'If OptDept(0).Value = True Then '專利處
      If strPorType = "P" Then
         For Each oChk In Chk_P
            If IsRecordExist(textPP02, Left(oChk.Caption, 1)) = True Then '資料已存在
               oChk.Enabled = False
            Else
               oChk.Enabled = True
               bolShowCommand1 = True
            End If
         Next
      ElseIf strPorType = "T" Then 'If OptDept(1).Value = True Then
         For Each oChk In Chk_T
            If IsRecordExist(textPP02, Left(oChk.Caption, 1)) = True Then '資料已存在
               oChk.Enabled = False
            Else
               oChk.Enabled = True
               bolShowCommand1 = True
            End If
         Next
      'Add By Sindy 2024/9/16
      ElseIf strPorType = "CF" Then
         For Each oChk In Chk_CF
            If IsRecordExist(textPP02, Left(oChk.Caption, 1)) = True Then '資料已存在
               oChk.Enabled = False
            Else
               oChk.Enabled = True
               bolShowCommand1 = True
            End If
         Next
      '2024/9/16 END
      End If
      'If bolShowCommand1 = True And Frame1.Visible = True Then Command1.Enabled = True
      
      '若有資料游標停在第一筆
      grd2.Visible = False
      grd2.col = 0
      grd2.row = 1
      'dblPrevRow = GRD2.row
      If rsTmp.RecordCount > 0 Then
         For i = 0 To grd2.Cols - 1
            grd2.col = i
            grd2.CellBackColor = &HFFC0C0
         Next i
         '讀取資料
         If m_EditMode <> 1 Then
            Call GetRowData(1)
         End If
      End If
      grd2.Visible = True
   End If
   rsTmp.Close
   
EXITSUB:
   Set rsTmp = Nothing
End Function

Private Sub SetGrd()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   '                        0           1           2       3           4
   arrGridHeadText = Array("部門代碼", "部門名稱", "所別", "員工編號", "姓名")
   arrGridHeadWidth = Array(1000, 1000, 1000, 1000, 1000)
   GRD1.Visible = False
   GRD1.Cols = UBound(arrGridHeadText) + 1
   For iRow = 0 To GRD1.Cols - 1
      GRD1.row = 0
      GRD1.col = iRow
      GRD1.Text = arrGridHeadText(iRow)
      GRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      GRD1.CellAlignment = flexAlignCenterCenter
   Next
   GRD1.Visible = True
End Sub

Private Sub SetGrd2(oGrd As Object)
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   '                        0       1           2       3         4           5        6           7       8          9       10      11        12      13
   arrGridHeadText = Array("所別", "員工編號", "姓名", "系統別", "核判分類", "cpm23", "申請國家", "pp04", "核稿人1", "離職", "pp05", "判發人", "離職", "筆數")
   If m_ByCPMSet = True Then
      arrGridHeadWidth = Array(500, 800, 1000, 1000, 1000, 0, 1000, 0, 800, 500, 0, 800, 500, 0)
   Else
      arrGridHeadWidth = Array(500, 800, 1000, 0, 1500, 0, 0, 0, 800, 500, 0, 800, 500, 500)
   End If
   oGrd.Visible = False
   oGrd.Cols = UBound(arrGridHeadText) + 1
   For iRow = 0 To oGrd.Cols - 1
      oGrd.row = 0
      oGrd.col = iRow
      oGrd.Text = arrGridHeadText(iRow)
      oGrd.ColWidth(iRow) = arrGridHeadWidth(iRow)
      oGrd.CellAlignment = flexAlignCenterCenter
   Next
   If m_ByCPMSet = True Then oGrd.TextMatrix(0, 4) = "案件性質"
   oGrd.Visible = True
End Sub

'讀取資料
Sub GetRowData(intRow As Integer)
Dim strST03 As String
Dim strST03Nm As String
   
   textPP02.Text = grd2.TextMatrix(intRow, 1)
   textPP02_2 = GetStaffName(textPP02, True, strST03Nm, strST03)
   TextDept = strST03
   TextDept_2 = strST03Nm
   If m_ByCPMSet = True Then
      textPP01.Text = grd2.TextMatrix(intRow, 3)
      cboPP03.Text = grd2.TextMatrix(intRow, 4)
      cboPP06.Text = grd2.TextMatrix(intRow, 6)
   Else
      cboType = grd2.TextMatrix(intRow, 4)
      Call cboType_Click
   End If
   cboPP04.Text = grd2.TextMatrix(intRow, 7)
   If cboPP04.Text <> "" Then
      cboPP04.Text = cboPP04.Text & " " & GetStaffName(cboPP04)
   End If
   cboPP05.Text = grd2.TextMatrix(intRow, 10)
   If cboPP05.Text <> "" Then
      cboPP05.Text = cboPP05.Text & " " & GetStaffName(cboPP05)
   End If
End Sub

Sub GetData()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String, strCon As String
Dim i As Integer

   strSql = "": strCon = ""
   If txt1(0) <> "" Then
       strCon = strCon & " and ST01>='" & txt1(0) & "' "
   End If
   If txt1(1) <> "" Then
       strCon = strCon & " and ST01<='" & txt1(1) & "' "
   End If
   If txt1(2) <> "" Then
       strCon = strCon & " and ST03>='" & DBDATE(txt1(2)) & "' "
   End If
   If txt1(3) <> "" Then
       strCon = strCon & " and ST03<='" & DBDATE(txt1(3)) & "' "
   End If
   
   strSql = "select st03,a0902,DECODE(A.ST06,'1','北所','2','中所','3','南所','4','高所','其他'),pp02,a.st02" & _
            " from promoterproofreader,staff a,acc090" & _
            " where pp02=a.st01(+) and st04='1'" & _
            " and st03=a0901(+)" & strCon & _
            " group by st06,st03,a0902,pp02,st02" & _
            " order by st03,st06,pp02"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   Set GRD1.Recordset = rsTmp
   Call SetGrd
   If rsTmp.RecordCount = 0 Then
      GRD1.col = 0
      GRD1.row = 0
      If m_EditMode <> 1 Then
         MsgBox "查無資料！", vbInformation
      End If
   Else
      GRD1.col = 0
      GRD1.row = 1
   End If
   rsTmp.Close
   
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

Private Function CheckDataValid() As Boolean
   Dim nResponse As Boolean
   Dim strTmp  As String
   CheckDataValid = False
   
   nResponse = False
   textPP02_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   
   CheckDataValid = True
   
EXITSUB:
End Function

' 更新Key的狀態
Private Sub SetKeyReadOnly(ByVal bEnable As Boolean)
   textPP02.Locked = bEnable
   If bEnable Then textPP02.BackColor = &H8000000F Else textPP02.BackColor = &H80000005
   If m_EditMode <> "2" Then
      cboType.Enabled = Not bEnable
      If bEnable Then cboType.BackColor = &H8000000F Else cboType.BackColor = &H80000005
   End If
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
   textPP02.Locked = bEnable
   If bEnable Then textPP02.BackColor = &H8000000F Else textPP02.BackColor = &H80000005
   If Frame1.Visible = True Then
      If m_EditMode = 1 Then
         cboType.Enabled = Not bEnable
         If bEnable Then cboType.BackColor = &H8000000F Else cboType.BackColor = &H80000005
      Else
         cboType.Enabled = False
         cboType.BackColor = &H8000000F
      End If
   Else
      cboType.Enabled = False
      cboType.BackColor = &H8000000F
      If m_EditMode = 1 Then
         textPP01.Enabled = Not bEnable
         If bEnable Then textPP01.BackColor = &H8000000F Else textPP01.BackColor = &H80000005
         cboPP03.Enabled = Not bEnable
         If bEnable Then cboPP03.BackColor = &H8000000F Else cboPP03.BackColor = &H80000005
         cboPP06.Enabled = Not bEnable
         If bEnable Then cboPP06.BackColor = &H8000000F Else cboPP06.BackColor = &H80000005
      Else
         textPP01.Enabled = False
         textPP01.BackColor = &H8000000F
         cboPP03.Enabled = False
         cboPP03.BackColor = &H8000000F
         cboPP06.Enabled = False
         cboPP06.BackColor = &H8000000F
      End If
   End If
   cboPP04.Enabled = Not bEnable
   If bEnable Then cboPP04.BackColor = &H8000000F Else cboPP04.BackColor = &H80000005
   cboPP05.Enabled = Not bEnable
   If bEnable Then cboPP05.BackColor = &H8000000F Else cboPP05.BackColor = &H80000005
   OptDept(0).Enabled = False 'Not bEnable
   If bEnable Then OptDept(0).BackColor = &H8000000F Else OptDept(0).BackColor = &H80000005
   OptDept(1).Enabled = False 'Not bEnable
   If bEnable Then OptDept(1).BackColor = &H8000000F Else OptDept(1).BackColor = &H80000005
   'Add By Sindy 2024/9/16
   OptDept(2).Enabled = False 'Not bEnable
   If bEnable Then OptDept(2).BackColor = &H8000000F Else OptDept(2).BackColor = &H80000005
   '2024/9/16 END
End Sub

Private Sub ClearField(Optional bolClearKey As Boolean = True)
   If bolClearKey = True Then
      textPP02 = Empty: textPP02_2 = Empty
      TextDept = Empty: TextDept_2 = Empty
      OptDept(0).Value = False
      OptDept(1).Value = False
      OptDept(2).Value = False 'Add By Sindy 2024/9/16
   End If
   Frame_P.Visible = False: Command2.Visible = True
   Frame_T.Visible = False: Command3.Visible = True
   Frame_CF.Visible = False: Command4.Visible = True 'Add By Sindy 2024/9/16
   textPP01 = Empty
   cboType.Text = ""
   cboPP04.Text = ""
   cboPP03.ListIndex = -1
   cboPP06.ListIndex = -1
   cboPP05.Text = ""
   
   'Add By Sindy 2024/9/18
   For Each oChk In Chk_P
      oChk.Enabled = True
      oChk.Value = 0
   Next
   For Each oChk In Chk_T
      oChk.Enabled = True
      oChk.Value = 0
   Next
   For Each oChk In Chk_CF
      oChk.Enabled = True
      oChk.Value = 0
   Next
   '2024/9/18 END
End Sub

Private Sub SetComboList()
   If TextDept.Text = "" Then Exit Sub
   
   If m_ByCPMSet = False Then
      'If TextDept.Text = TextDept.Tag And (OptDept(0).Value = True Or OptDept(1).Value = True) Then Exit Sub
      If Left(TextDept, 2) <> "P1" And Left(TextDept, 2) <> "P2" And Left(TextDept, 2) <> "F1" And m_EditMode = 1 Then
         strUserDept = InputBox("請輸入欲增修的核判表是屬於那一個部門？" & vbCrLf & "(P:專利處 T:商標處 CF:外商CF)")
         strUserDept = UCase(strUserDept) 'Add By Sindy 2023/8/18
         If strUserDept <> "P" And strUserDept <> "T" And strUserDept <> "CF" Then
            Me.textPP02 = ""
            Me.textPP02.SetFocus
            Exit Sub
         Else
            If strUserDept = "P" Then
               OptDept(0).Value = True
            ElseIf strUserDept = "T" Then
               OptDept(1).Value = True
            'Add By Sindy 2024/9/16
            ElseIf strUserDept = "CF" Then
               OptDept(2).Value = True
            '2024/9/16 END
            End If
         End If
      End If
      
      '核判分類下拉選單
      cboType.Clear
      'If Left(TextDept, 2) = "P1" Or strUserDept = "P" Then '專利處
      If OptDept(0).Value = True Then
         For Each oChk In Chk_P
            cboType.AddItem oChk.Caption
         Next
      End If
      'If Left(TextDept, 2) = "P2" Or strUserDept = "T" Then '商標處
      If OptDept(1).Value = True Then
         For Each oChk In Chk_T
            cboType.AddItem oChk.Caption
         Next
      End If
      'Add By Sindy 2024/9/16
      If OptDept(2).Value = True Then
         For Each oChk In Chk_CF
            cboType.AddItem oChk.Caption
         Next
      End If
      '2024/9/16 END
   End If
   TextDept.Tag = TextDept.Text
   
   '核判人員下拉選單
   cboPP04.Clear
   cboPP04.AddItem ""
   cboPP05.Clear
   cboPP05.AddItem ""
   strSql = "SELECT a0902,st01,st02 FROM staff,acc090 WHERE st04='1' and st01>'63' and st01<'F' and st03=a0901(+) and substr(st01,4,1)<>'9'"
   If m_ByCPMSet = False Then
      If OptDept(0).Value = True Then '專利處
         strSql = strSql & " and st03 in('P10','P11')"
      ElseIf OptDept(1).Value = True Then '商標處
         strSql = strSql & " and st03 in('P20','P21')"
      'Add By Sindy 2024/9/16
      ElseIf OptDept(2).Value = True Then '外商CF
         strSql = strSql & " and st03 in('F11','F12')"
      '2024/9/16 END
      End If
   Else
      strSql = strSql & " and st03='" & PUB_GetST03(textPP02) & "'"
   End If
   strSql = strSql & " order by st03 asc,st01 asc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      With RsTemp
         RsTemp.MoveFirst
         Do While RsTemp.EOF = False
            cboPP04.AddItem Trim(RsTemp.Fields("st01")) & " " & Trim(RsTemp.Fields("st02"))
            cboPP05.AddItem Trim(RsTemp.Fields("st01")) & " " & Trim(RsTemp.Fields("st02"))
            RsTemp.MoveNext
         Loop
      End With
   End If
      
   'Add By Sindy 2024/6/18 申請國家下拉選單
   cboPP06.Clear
   cboPP06.AddItem ""
   cboPP06.AddItem "*   沒指定"
   cboPP06.AddItem "020 中國大陸"
   '案件性質下拉選單
   cboPP03.Clear
   cboPP03.AddItem ""
   cboPP03.AddItem "103  外觀設計申請"
   '2024/6/18 END
End Sub

Private Sub textPP01_GotFocus()
   InverseTextBox textPP01
   CloseIme
End Sub

Private Sub textPP01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textPP01_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textPP01) = False Then
      Select Case m_EditMode
         Case 1, 4:
            If IsAlphabetic(textPP01) = False Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "系統類別不正確"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textPP01_GotFocus
               GoTo EXITSUB
            End If
            If IsUserHasRightOfSystem(strUserNum, textPP01) = False Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "您沒有使用該系統類別的權限"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textPP01_GotFocus
            End If
            If IsCorrectSysKind(textPP01) = False Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "系統類別不正確"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textPP01_GotFocus
            End If
      End Select
   End If
EXITSUB:
End Sub

Private Sub textPP02_GotFocus()
   InverseTextBox textPP02
End Sub

Private Sub textPP02_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textPP02_LostFocus()
   '新增狀態將游標停在員工代號的欄位
   If m_EditMode = 1 And textPP02 = "" Then textPP02.SetFocus
End Sub

Private Sub textPP02_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim strST03 As String, strST03Nm As String
   
   If textPP02.Text = "" Then
      textPP02_2.Text = ""
      TextDept.Text = ""
      TextDept_2.Text = ""
   End If
   
   If textPP02 <> "" And (textPP02.Text <> textPP02.Tag Or textPP02_2 = "") Then
      textPP02_2 = GetStaffName(textPP02, True, strST03Nm, strST03)
      TextDept = strST03
      TextDept_2 = strST03Nm
      textPP02.Tag = textPP02.Text
      Call SetEmpCase(textPP02.Text)
      'If m_EditMode = 1 Then  '新增時
         SetCtrlReadOnly False
         Call UpdateCtrlData(textPP02)
      'End If
   End If
   'If m_EditMode <> 0 And textPP02 <> "" Then
   '   '檢查員工編號規則
   '   If ChkStaffID(textPP02) Then
   '      Call textPP02_GotFocus
   '      Cancel = True
   '      Exit Sub
   '   End If
   '   If textPP02_2 = "" Then
   '       MsgBox "員工編號錯誤！查無此員工！", vbInformation
   '       Call textPP02_GotFocus
   '       Cancel = True
   '       Exit Sub
   '   End If
   'End If
End Sub

Private Sub SetEmpCase(strEmp As String)
Dim strST03 As String
   
   Me.Frame_P.Visible = False
   Me.Frame_T.Visible = False
   Me.Frame_CF.Visible = False 'Add By Sindy 2024/9/16
   strST03 = PUB_GetST03(strEmp)
   
   'P10 專利處主管
   'P11 專利工程師
   'P12 專利處程序
   'P13 專利處繪圖
   'P14 專利處英文顧問
   'P20 商標處主管
   'P21 商標處承辦
   'P22 商標處程序
   'F11 外商承辦
   'F12 外商程序
   strPorType = ""
   '以 系統別+案件性質 或 系統別+案件性質+申請國家 做設定
   If strST03 = "P13" Or Left(strST03, 2) = "F2" Then
      m_ByCPMSet = True
   '以 核判分類做設定
   Else
      'A5011=郭仁建
      'A8002=魏裕仁
      m_ByCPMSet = False
      If Left(strST03, 2) = "P1" Or Left(strST03, 2) = "F5" Or strEmp = "A5011" Or strEmp = "A8002" Then
         OptDept(0).Value = True
         strPorType = "P"
      ElseIf Left(strST03, 2) = "P2" Then
         OptDept(1).Value = True
         strPorType = "T"
      'Add By Sindy 2024/9/16
      ElseIf Left(strST03, 2) = "F1" Then
         OptDept(2).Value = True
         strPorType = "CF"
      '2024/9/16 END
      Else
         m_ByCPMSet = True
      End If
   End If
   
   If m_ByCPMSet = True Then
      Frame1.Visible = False
      Command1.Visible = False
      FrameSys.Visible = True
      FramePP0306.Visible = True
      Me.cboType.Enabled = False
      cboType.BackColor = &H8000000F
      LblNoteNA01.Visible = True
   Else
      Frame1.Visible = True
      Command1.Visible = True
      FrameSys.Visible = False
      FramePP0306.Visible = False
      LblNoteNA01.Visible = False
   End If
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   InverseTextBox txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0, 1, 2, 3
         KeyAscii = UpperCase(KeyAscii)
'      Case 2, 3
'         KeyAscii = Pub_NumAscii(KeyAscii)
      Case Else
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 0, 1
         ' 檢查員工編號規則
         If txt1(Index).Text <> "" Then
            If ChkStaffID(txt1(Index)) Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
         If Index = 0 Then
            If txt1(Index) <> "" And txt1(Index + 1) = "" Then
               txt1(Index + 1) = txt1(Index)
            End If
         ElseIf Index = 1 Then
            If txt1(Index) <> "" Then
               If RunNick(txt1(Index - 1), txt1(Index)) Then
                  Call txt1_GotFocus(Index)
                  Cancel = True
                  Exit Sub
               End If
            End If
         End If
         
      Case 2, 3
         If Index = 2 Then
            If txt1(Index) <> "" And txt1(Index + 1) = "" Then
               txt1(Index + 1) = txt1(Index)
            End If
         ElseIf Index = 3 Then
            If txt1(Index) <> "" Then
               If RunNick(txt1(Index - 1), txt1(Index)) Then
                  Call txt1_GotFocus(Index)
                  Cancel = True
                  Exit Sub
               End If
            End If
         End If
         
      Case Else
   End Select
End Sub
