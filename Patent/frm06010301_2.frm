VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm06010301_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "各式申請書-電子送件-補件, 翻譯, 檢視中說, 製作中說, 核對中說, 實體審查"
   ClientHeight    =   6940
   ClientLeft      =   410
   ClientTop       =   1500
   ClientWidth     =   8530
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6940
   ScaleWidth      =   8530
   Begin VB.TextBox txtDecreasePageFee 
      Height          =   288
      Left            =   5490
      TabIndex        =   104
      Top             =   30
      Width           =   1110
   End
   Begin VB.TextBox txtDecreaseItemFee 
      Height          =   288
      Left            =   5490
      TabIndex        =   103
      Top             =   300
      Width           =   1110
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定"
      Default         =   -1  'True
      Height          =   350
      Index           =   0
      Left            =   6630
      TabIndex        =   29
      Top             =   60
      Width           =   600
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "回前畫面"
      CausesValidation=   0   'False
      Height          =   350
      Index           =   2
      Left            =   7350
      TabIndex        =   30
      Top             =   60
      Width           =   900
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1110
      MaxLength       =   3
      TabIndex        =   34
      Top             =   12
      Width           =   550
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1650
      MaxLength       =   6
      TabIndex        =   33
      Top             =   12
      Width           =   855
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2490
      MaxLength       =   1
      TabIndex        =   32
      Top             =   12
      Width           =   255
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2760
      MaxLength       =   2
      TabIndex        =   31
      Top             =   12
      Width           =   375
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6375
      Left            =   45
      TabIndex        =   39
      Top             =   525
      Width           =   8445
      _ExtentX        =   14905
      _ExtentY        =   11236
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   8
      TabHeight       =   520
      ForeColor       =   -2147483630
      TabCaption(0)   =   "案件名稱"
      TabPicture(0)   =   "frm06010301_2.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label5(1)"
      Tab(0).Control(1)=   "Label14(1)"
      Tab(0).Control(2)=   "Label5(0)"
      Tab(0).Control(3)=   "lblNameAgent"
      Tab(0).Control(4)=   "Label27"
      Tab(0).Control(5)=   "Label2"
      Tab(0).Control(6)=   "Label3"
      Tab(0).Control(7)=   "Label4"
      Tab(0).Control(8)=   "Label6"
      Tab(0).Control(9)=   "Label10"
      Tab(0).Control(10)=   "lstNameAgent"
      Tab(0).Control(11)=   "Text6(1)"
      Tab(0).Control(12)=   "Text6(0)"
      Tab(0).Control(13)=   "txtCP84"
      Tab(0).Control(14)=   "cboResonCode"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Text5"
      Tab(0).Control(16)=   "Text7"
      Tab(0).Control(17)=   "Text8"
      Tab(0).Control(18)=   "Text9"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "FraPA174"
      Tab(0).ControlCount=   20
      TabCaption(1)   =   "補件內容"
      TabPicture(1)   =   "frm06010301_2.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label15"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label16(0)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label19"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label20"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label21"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label22"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Shape2"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label17"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Shape1"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label8"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label16(1)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label18"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Label23"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Label1(1)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "LblTotItem"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Label12"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "LblTotPage"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "chkDoc(0)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "txtDocCh(0)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "txtDocCh(1)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "txtDocCh(3)"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "txtCP135"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "txtCP136"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "txtDocCh(6)"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "txtDocCh(2)"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "chkDoc(1)"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "chkAtt(23)"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "chkAtt(24)"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "chkAtt(26)"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "chkAtt(27)"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "chkAtt(28)"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "chkAtt(29)"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "chkAtt(30)"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "chkDoc(3)"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "Frame1"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "Frame2"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "chkDoc(4)"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).Control(37)=   "txtDocCh(4)"
      Tab(1).Control(37).Enabled=   0   'False
      Tab(1).Control(38)=   "chkDoc(5)"
      Tab(1).Control(38).Enabled=   0   'False
      Tab(1).Control(39)=   "txtTotItem"
      Tab(1).Control(39).Enabled=   0   'False
      Tab(1).Control(40)=   "chkDoc(6)"
      Tab(1).Control(40).Enabled=   0   'False
      Tab(1).Control(41)=   "txtTotPage"
      Tab(1).Control(41).Enabled=   0   'False
      Tab(1).ControlCount=   42
      Begin VB.TextBox txtTotPage 
         Appearance      =   0  '平面
         BackColor       =   &H8000000F&
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   1230
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   101
         Top             =   1860
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.CheckBox chkDoc 
         Caption         =   "已於提申時減免規費800元整，補呈首頁及摘要英文資料"
         Height          =   405
         Index           =   6
         Left            =   1020
         TabIndex        =   20
         Top             =   3570
         Width           =   2760
      End
      Begin VB.Frame FraPA174 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  '沒有框線
         Caption         =   "Frame1"
         Height          =   525
         Left            =   -74820
         TabIndex        =   95
         Top             =   690
         Visible         =   0   'False
         Width           =   825
         Begin VB.CommandButton CmdPA174 
            BackColor       =   &H00C0FFFF&
            Caption         =   "特殊字"
            Height          =   280
            Left            =   0
            Style           =   1  '圖片外觀
            TabIndex        =   96
            Top             =   210
            Width           =   800
         End
         Begin VB.Label lblPA174 
            Caption         =   "有特殊字"
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   35
            TabIndex        =   97
            Top             =   0
            Width           =   765
         End
      End
      Begin VB.TextBox txtTotItem 
         Appearance      =   0  '平面
         BackColor       =   &H8000000F&
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   1230
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   94
         Top             =   2160
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.CheckBox chkDoc 
         Caption         =   "聲明中文本與申請時外文本實質內容一致"
         Height          =   570
         Index           =   5
         Left            =   1710
         TabIndex        =   23
         Top             =   4140
         Width           =   2130
      End
      Begin VB.TextBox txtDocCh 
         Height          =   270
         Index           =   4
         Left            =   3375
         MaxLength       =   4
         TabIndex        =   11
         Top             =   1050
         Width           =   420
      End
      Begin VB.CheckBox chkDoc 
         Caption         =   "首頁及摘要均附英文資料，減免規費800元整"
         Height          =   525
         Index           =   4
         Left            =   1020
         TabIndex        =   19
         Top             =   3060
         Width           =   2520
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  '平面
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   3930
         TabIndex        =   84
         Top             =   5130
         Width           =   4425
         Begin VB.CheckBox chkAtt 
            Caption         =   "文件檔名"
            Enabled         =   0   'False
            Height          =   195
            Index           =   22
            Left            =   810
            TabIndex        =   87
            Top             =   360
            Width           =   2430
         End
         Begin VB.CheckBox chkAtt 
            Caption         =   "文件描述"
            Enabled         =   0   'False
            Height          =   195
            Index           =   21
            Left            =   810
            TabIndex        =   86
            Top             =   150
            Width           =   2430
         End
         Begin VB.CheckBox chkDoc 
            Caption         =   "其他"
            Height          =   195
            Index           =   2
            Left            =   90
            TabIndex        =   85
            Top             =   150
            Width           =   1230
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  '平面
         ForeColor       =   &H80000008&
         Height          =   4785
         Left            =   3930
         TabIndex        =   59
         Top             =   330
         Width           =   4425
         Begin VB.CheckBox Check3 
            Caption         =   "個案"
            Height          =   195
            Left            =   3540
            TabIndex        =   63
            Top             =   3030
            Width           =   690
         End
         Begin VB.CheckBox chkAtt 
            Caption         =   "委任書(附譯文)"
            Height          =   195
            Index           =   25
            Left            =   2010
            TabIndex        =   62
            Top             =   3030
            Width           =   1530
         End
         Begin VB.CheckBox chkAtt 
            Caption         =   "發明圖式"
            Height          =   195
            Index           =   4
            Left            =   810
            TabIndex        =   82
            Tag             =   ".INV_DRAWINGS.pdf"
            Top             =   1275
            Width           =   1230
         End
         Begin VB.CheckBox chkAtt 
            Caption         =   "發明申請專利範圍"
            Height          =   195
            Index           =   3
            Left            =   810
            TabIndex        =   81
            Tag             =   ".INV_CLAIMS.pdf"
            Top             =   1050
            Width           =   1860
         End
         Begin VB.CheckBox chkAtt 
            Caption         =   "發明說明書"
            Height          =   195
            Index           =   2
            Left            =   810
            TabIndex        =   80
            Tag             =   ".INV_DESCRIPTION.pdf"
            Top             =   615
            Width           =   1230
         End
         Begin VB.CheckBox chkAtt 
            Caption         =   "發明摘要"
            Height          =   195
            Index           =   1
            Left            =   810
            TabIndex        =   79
            Tag             =   ".INV_ABSTRACT.pdf"
            Top             =   405
            Width           =   1230
         End
         Begin VB.CheckBox chkAtt 
            Caption         =   "基本資料表"
            Height          =   195
            Index           =   0
            Left            =   810
            TabIndex        =   78
            Tag             =   ".CONTACT.pdf"
            Top             =   180
            Value           =   2  '灰色
            Width           =   1230
         End
         Begin VB.CheckBox chkAtt 
            Caption         =   "新型圖式"
            Height          =   195
            Index           =   8
            Left            =   810
            TabIndex        =   77
            Tag             =   ".UTL_DRAWINGS.pdf"
            Top             =   2145
            Width           =   1230
         End
         Begin VB.CheckBox chkAtt 
            Caption         =   "新型申請專利範圍"
            Height          =   195
            Index           =   7
            Left            =   810
            TabIndex        =   76
            Tag             =   ".UTL_CLAIMS.pdf"
            Top             =   1935
            Width           =   1860
         End
         Begin VB.CheckBox chkAtt 
            Caption         =   "新型說明書"
            Height          =   195
            Index           =   6
            Left            =   810
            TabIndex        =   75
            Tag             =   ".UTL_DESCRIPTION.pdf"
            Top             =   1710
            Width           =   1230
         End
         Begin VB.CheckBox chkAtt 
            Caption         =   "新型摘要"
            Height          =   195
            Index           =   5
            Left            =   810
            TabIndex        =   74
            Tag             =   ".UTL_ABSTRACT.pdf"
            Top             =   1500
            Width           =   1230
         End
         Begin VB.CheckBox chkAtt 
            Caption         =   "設計圖式"
            Height          =   195
            Index           =   10
            Left            =   810
            TabIndex        =   73
            Tag             =   ".DES_DRAWINGS.pdf"
            Top             =   2595
            Width           =   1230
         End
         Begin VB.CheckBox chkAtt 
            Caption         =   "設計說明書"
            Height          =   195
            Index           =   9
            Left            =   810
            TabIndex        =   72
            Tag             =   ".DES_DESCRIPTION.pdf"
            Top             =   2370
            Width           =   1230
         End
         Begin VB.CheckBox chkAtt 
            Caption         =   "序列表"
            Height          =   195
            Index           =   11
            Left            =   810
            TabIndex        =   71
            Tag             =   ".SEQ.pdf"
            Top             =   840
            Width           =   1860
         End
         Begin VB.CheckBox chkAtt 
            Caption         =   "申請書"
            Height          =   195
            Index           =   12
            Left            =   810
            TabIndex        =   70
            Tag             =   ".ATT.DATA.pdf"
            Top             =   2805
            Width           =   1860
         End
         Begin VB.CheckBox chkAtt 
            Caption         =   "申復書"
            Height          =   195
            Index           =   13
            Left            =   810
            TabIndex        =   69
            Tag             =   ".EX.pdf"
            Top             =   4335
            Width           =   1230
         End
         Begin VB.CheckBox chkAtt 
            Caption         =   "再審查理由書"
            Height          =   195
            Index           =   14
            Left            =   810
            TabIndex        =   68
            Tag             =   ".RE.pdf"
            Top             =   4575
            Width           =   2430
         End
         Begin VB.CheckBox chkAtt 
            Caption         =   "優惠期證明文件"
            Height          =   195
            Index           =   15
            Left            =   810
            TabIndex        =   67
            Tag             =   ".EXHIBITION.pdf"
            Top             =   3465
            Width           =   1935
         End
         Begin VB.CheckBox chkAtt 
            Caption         =   "國內生物材料寄存證明文件"
            Height          =   195
            Index           =   16
            Left            =   810
            TabIndex        =   66
            Tag             =   ".DOMESTICPROOF.pdf"
            Top             =   3690
            Width           =   2535
         End
         Begin VB.CheckBox chkAtt 
            Caption         =   "國外生物材料寄存證明文件"
            Height          =   195
            Index           =   17
            Left            =   810
            TabIndex        =   65
            Tag             =   ".FOREIGNPROOF.pdf"
            Top             =   3900
            Width           =   2700
         End
         Begin VB.CheckBox chkAtt 
            Caption         =   "生物材料為通常知識者易於獲得證明文件"
            Height          =   195
            Index           =   18
            Left            =   810
            TabIndex        =   64
            Tag             =   ".EASILYOBTAINED.pdf"
            Top             =   4125
            Width           =   3540
         End
         Begin VB.CheckBox chkAtt 
            Caption         =   "委任書"
            Height          =   195
            Index           =   19
            Left            =   810
            TabIndex        =   61
            Tag             =   ".POA.pdf"
            Top             =   3030
            Width           =   1080
         End
         Begin VB.CheckBox chkAtt 
            Caption         =   "國際優先權證明文件"
            Height          =   195
            Index           =   20
            Left            =   810
            TabIndex        =   60
            Tag             =   ".PRI.pdf"
            Top             =   3240
            Width           =   2295
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "附件文書"
            Height          =   180
            Left            =   90
            TabIndex        =   83
            Top             =   255
            Width           =   720
         End
      End
      Begin VB.CheckBox chkDoc 
         Caption         =   "一併修正中文專利名稱"
         Height          =   195
         Index           =   3
         Left            =   1020
         TabIndex        =   18
         Top             =   2880
         Width           =   2280
      End
      Begin VB.CheckBox chkAtt 
         Caption         =   "變更申請人之國籍"
         Height          =   210
         Index           =   30
         Left            =   570
         TabIndex        =   28
         Top             =   5970
         Width           =   2430
      End
      Begin VB.CheckBox chkAtt 
         Caption         =   "變更申請人之姓名或名稱"
         ForeColor       =   &H000000C0&
         Height          =   210
         Index           =   29
         Left            =   570
         TabIndex        =   27
         Top             =   5760
         Width           =   2430
      End
      Begin VB.CheckBox chkAtt 
         Caption         =   "變更申請人之代表人"
         Height          =   210
         Index           =   28
         Left            =   570
         TabIndex        =   26
         Top             =   5550
         Width           =   2430
      End
      Begin VB.CheckBox chkAtt 
         Caption         =   "變更申請人之代理人"
         ForeColor       =   &H000000C0&
         Height          =   210
         Index           =   27
         Left            =   570
         TabIndex        =   25
         Top             =   5340
         Width           =   2430
      End
      Begin VB.CheckBox chkAtt 
         Caption         =   "變更申請人之地址"
         Height          =   210
         Index           =   26
         Left            =   570
         TabIndex        =   24
         Top             =   4920
         Width           =   1770
      End
      Begin VB.CheckBox chkAtt 
         Caption         =   "其他申復事項"
         Height          =   210
         Index           =   24
         Left            =   330
         TabIndex        =   22
         Top             =   4320
         Width           =   1410
      End
      Begin VB.CheckBox chkAtt 
         Caption         =   "補正其他文件"
         Height          =   210
         Index           =   23
         Left            =   330
         TabIndex        =   21
         Top             =   4050
         Width           =   2430
      End
      Begin VB.CheckBox chkDoc 
         Caption         =   "備註"
         Height          =   195
         Index           =   1
         Left            =   330
         TabIndex        =   17
         Top             =   2880
         Width           =   870
      End
      Begin VB.TextBox txtDocCh 
         Height          =   270
         Index           =   2
         Left            =   3375
         MaxLength       =   4
         TabIndex        =   12
         Top             =   1320
         Width           =   420
      End
      Begin VB.TextBox Text9 
         Alignment       =   2  '置中對齊
         Appearance      =   0  '平面
         BackColor       =   &H8000000F&
         BorderStyle     =   0  '沒有框線
         Height          =   225
         Left            =   -73380
         Locked          =   -1  'True
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   2115
         Width           =   375
      End
      Begin VB.TextBox Text8 
         Height          =   270
         Left            =   -70350
         MaxLength       =   11
         TabIndex        =   5
         Top             =   2085
         Width           =   1260
      End
      Begin VB.TextBox Text7 
         Height          =   270
         Left            =   -72330
         MaxLength       =   15
         TabIndex        =   4
         Text            =   "一(二)"
         Top             =   2085
         Width           =   1485
      End
      Begin VB.TextBox Text5 
         Height          =   270
         Left            =   -73560
         MaxLength       =   7
         TabIndex        =   3
         Top             =   1785
         Width           =   1005
      End
      Begin VB.ComboBox cboResonCode 
         Enabled         =   0   'False
         Height          =   260
         ItemData        =   "frm06010301_2.frx":0038
         Left            =   -73560
         List            =   "frm06010301_2.frx":0057
         Style           =   2  '單純下拉式
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1770
      End
      Begin VB.TextBox txtDocCh 
         Height          =   270
         Index           =   6
         Left            =   3375
         TabIndex        =   16
         Top             =   2430
         Width           =   420
      End
      Begin VB.TextBox txtCP136 
         Height          =   270
         Left            =   3375
         MaxLength       =   4
         TabIndex        =   15
         Top             =   2160
         Width           =   420
      End
      Begin VB.TextBox txtCP135 
         Height          =   270
         Left            =   3375
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   14
         Top             =   1890
         Width           =   420
      End
      Begin VB.TextBox txtDocCh 
         Height          =   270
         Index           =   3
         Left            =   3375
         MaxLength       =   4
         TabIndex        =   13
         Top             =   1605
         Width           =   420
      End
      Begin VB.TextBox txtDocCh 
         Height          =   270
         Index           =   1
         Left            =   3375
         MaxLength       =   4
         TabIndex        =   10
         Top             =   780
         Width           =   420
      End
      Begin VB.TextBox txtDocCh 
         Height          =   270
         Index           =   0
         Left            =   3375
         MaxLength       =   4
         TabIndex        =   9
         Top             =   510
         Width           =   420
      End
      Begin VB.CheckBox chkDoc 
         Caption         =   "中文本資訊"
         Height          =   195
         Index           =   0
         Left            =   330
         TabIndex        =   8
         Top             =   540
         Width           =   1230
      End
      Begin VB.TextBox txtCP84 
         Height          =   270
         Left            =   -73740
         MaxLength       =   7
         TabIndex        =   6
         Top             =   2760
         Width           =   1140
      End
      Begin VB.Label LblTotPage 
         AutoSize        =   -1  'True
         Caption         =   "目前總頁數:"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   240
         TabIndex        =   102
         Top             =   1935
         Visible         =   0   'False
         Width           =   945
      End
      Begin MSForms.TextBox Text6 
         Height          =   285
         Index           =   0
         Left            =   -73560
         TabIndex        =   0
         Top             =   420
         Width           =   6225
         VariousPropertyBits=   679495707
         Size            =   "10980;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text6 
         Height          =   285
         Index           =   1
         Left            =   -73560
         TabIndex        =   1
         Top             =   720
         Width           =   6225
         VariousPropertyBits=   679495707
         Size            =   "10980;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label12 
         Caption         =   "+ 300"
         ForeColor       =   &H000000C0&
         Height          =   165
         Left            =   1530
         TabIndex        =   88
         Top             =   4710
         Width           =   885
      End
      Begin MSForms.ListBox lstNameAgent 
         Height          =   315
         Left            =   -73740
         TabIndex        =   7
         Top             =   3060
         Width           =   1500
         VariousPropertyBits=   746586139
         ScrollBars      =   2
         DisplayStyle    =   2
         Size            =   "2646;556"
         MatchEntry      =   0
         ListStyle       =   1
         MultiSelect     =   1
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label LblTotItem 
         AutoSize        =   -1  'True
         Caption         =   "目前總項數:"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   240
         TabIndex        =   93
         Top             =   2205
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "不算超頁費"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   8.5
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   1
         Left            =   2490
         TabIndex        =   92
         Top             =   1095
         Width           =   900
      End
      Begin VB.Label Label23 
         Caption         =   "(請先至"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   165
         Left            =   2430
         TabIndex        =   91
         Top             =   4860
         Width           =   735
      End
      Begin VB.Label Label18 
         Caption         =   "客戶資料維護修改地址再產生申請書)"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   165
         Left            =   240
         TabIndex        =   90
         Top             =   5130
         Width           =   3195
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "序列表:"
         Height          =   180
         Index           =   1
         Left            =   1860
         TabIndex        =   89
         Top             =   1095
         Width           =   585
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "同時辦理事項"
         Height          =   210
         Left            =   330
         TabIndex        =   58
         Top             =   4710
         Width           =   1080
      End
      Begin VB.Shape Shape1 
         Height          =   3495
         Left            =   195
         Top             =   2790
         Width           =   3690
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "申請專利範圍頁數:"
         Height          =   180
         Left            =   1845
         TabIndex        =   57
         Top             =   1365
         Width           =   1485
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "範例：（  102  ）智專    一(二)15172     字第    10241450220    號"
         Height          =   180
         Left            =   -74145
         TabIndex        =   55
         Top             =   2460
         Width           =   4935
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "發文字號：（          ）智專                                     字第                               號"
         Height          =   180
         Left            =   -74505
         TabIndex        =   54
         Top             =   2130
         Width           =   5670
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "發文日期："
         Height          =   180
         Left            =   -74505
         TabIndex        =   53
         Top             =   1830
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "辦理依據:"
         Height          =   180
         Left            =   -74760
         TabIndex        =   52
         Top             =   1560
         Width           =   765
      End
      Begin VB.Shape Shape2 
         Height          =   2325
         Left            =   195
         Top             =   420
         Width           =   3690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "案由:"
         Height          =   180
         Left            =   -74760
         TabIndex        =   51
         Top             =   1230
         Width           =   405
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "圖式圖數:"
         Height          =   180
         Left            =   1845
         TabIndex        =   50
         Top             =   2475
         Width           =   765
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "申請專利範圍項數:"
         ForeColor       =   &H00000080&
         Height          =   180
         Left            =   1845
         TabIndex        =   49
         Top             =   2205
         Width           =   1485
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "頁數總計:"
         ForeColor       =   &H00000080&
         Height          =   180
         Left            =   1845
         TabIndex        =   48
         Top             =   1935
         Width           =   765
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "圖式頁數:"
         Height          =   180
         Left            =   1845
         TabIndex        =   47
         Top             =   1650
         Width           =   765
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "說明書頁數:"
         Height          =   180
         Index           =   0
         Left            =   1845
         TabIndex        =   46
         Top             =   825
         Width           =   945
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "摘要頁數:"
         Height          =   180
         Left            =   1845
         TabIndex        =   45
         Top             =   555
         Width           =   765
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "繳費金額:"
         Height          =   180
         Left            =   -74760
         TabIndex        =   44
         Top             =   2790
         Width           =   765
      End
      Begin VB.Label lblNameAgent 
         AutoSize        =   -1  'True
         Caption         =   "出名代理人:"
         Height          =   180
         Left            =   -74760
         TabIndex        =   43
         Top             =   3090
         Width           =   945
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   0
         Left            =   -73980
         TabIndex        =   42
         Top             =   465
         Width           =   345
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "案件名稱"
         Height          =   180
         Index           =   1
         Left            =   -74760
         TabIndex        =   41
         Top             =   465
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   1
         Left            =   -73980
         TabIndex        =   40
         Top             =   735
         Width           =   345
      End
   End
   Begin MSForms.Label Label7 
      Height          =   195
      Index           =   10
      Left            =   4080
      TabIndex        =   100
      Top             =   30
      Width           =   1335
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2355;344"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label7 
      Height          =   195
      Index           =   12
      Left            =   4080
      TabIndex        =   99
      Top             =   300
      Width           =   1335
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2355;344"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label7 
      Height          =   195
      Index           =   11
      Left            =   1080
      TabIndex        =   98
      Top             =   330
      Width           =   1335
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2355;344"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      Caption         =   "案件性質:"
      Height          =   180
      Left            =   3270
      TabIndex        =   38
      Top             =   15
      Width           =   765
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "智權人員:"
      Height          =   180
      Left            =   3270
      TabIndex        =   37
      Top             =   315
      Width           =   765
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "承辦人　:"
      Height          =   180
      Left            =   270
      TabIndex        =   36
      Top             =   315
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Index           =   0
      Left            =   270
      TabIndex        =   35
      Top             =   15
      Width           =   765
   End
End
Attribute VB_Name = "frm06010301_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/10/13 Form2.0已修改
'Created by Morgan 2013/7/11
Option Explicit

Dim strReceiveNo As String, intWhere As Integer
Dim pa() As String, cp() As String
Dim pageD() As String 'Add By Sindy 2023/3/15
Dim CP10 As String, m_CaseNo As String
'************************************************************
Dim m_bol99Case As Boolean '是否99年後申請案件
Dim m_bolChkFee As Boolean '是否需檢查規費
Dim m_bolNotSavePageItem As Boolean '不回存頁數項數 Add By Sindy 2018/4/25
Dim strNewCaseCP118 As String 'Add By Sindy 2018/4/25 抓新案的電子送件欄位值
Dim m_bolChkPageItem As Boolean '是否要輸頁數與項數
Dim m_lngOverPageFee As Long, m_lngOverItemFee As Long '超頁費,超項費
Dim m_lngOverPageFeeDiff As Long, m_lngOverItemFeeDiff As Long '超頁費,超項費差額
Dim m_lngRecOverPageFee As Long, m_lngRecOverItemFee As Long '已收文超頁費,超項費 Add by Morgan 2011/6/29
Dim m_str938RecvNo As String '回傳退費的超頁費文號 Add By Sindy 2023/4/7
Dim m_str939RecvNo As String '回傳退費的超項費文號 Add By Sindy 2023/4/7
Dim m_FeeMemo As String '規費備註
''Dim m_lngOfficialFee As Long '原始規費
Dim m_bol107NewFee As Boolean '台灣再審是否用102年新規費計算
Dim m_bolFixNewFee As Boolean '台灣修正是否用102年新規費計算
Dim m_strReExamCP27 As String '台灣再審發文日(若再審延期發文日)
Dim bolDelay As Boolean 'Add by Morgan 2004/9/8 是否延期過
Dim m_strDelayCP09 As String 'Added by Morgan 2011/11/11 延期收文號
Dim m_Div416OfficialFee As Long  '分割案實審規費  2010/12/8 add by sonia
Dim m_bolChkItem As Boolean '是否要檢查增刪項數 Add by Morgan 2010/9/27
'************************************************************
Dim m_strPA31 As String, m_strPA32 As String, m_strPA33 As String, m_strPA34 As String, m_strPA35 As String 'Add By Sindy 2018/4/18
Dim m_strPA36 As String, m_strPA37 As String, m_strPA38 As String, m_strPA39 As String, m_strPA40 As String 'Add By Sindy 2018/4/18
Dim oText As TextBox 'Add By Sindy 2018/4/23
Dim m_WriteNote As String 'Add By Sindy 2018/5/8
Dim str203CP09 As String 'Add By Sindy 2023/3/16
Dim m_allPage As String, m_allItem As String '總頁數,總項數
Dim m_AgentName As String 'Add By Sindy 2021/5/10
Dim m_IsRun As Boolean 'Add By Sindy 2023/3/24


Private Sub Check3_Click()
   If Check3.Value = 1 Then
      chkAtt(25).Value = 1
   End If
End Sub

'Modify By Sindy 2018/1/29
Private Sub chkAtt_Click(Index As Integer)
   If chkAtt(27).Value = 1 Or _
      chkAtt(29).Value = 1 Then
'      If Val(chkAtt(27).Tag) = 0 Then
'         txtCP84 = Val(txtCP84) + 300
'         chkAtt(27).Tag = 300
'      End If
      Label12.Visible = True
   Else
      Label12.Visible = False
'      If Val(chkAtt(27).Tag) = 300 And _
'         chkAtt(27).Value = 0 And _
'         chkAtt(29).Value = 0 Then
'         txtCP84 = Val(txtCP84) - 300
'         chkAtt(27).Tag = ""
'      End If
   End If
End Sub

Private Sub chkDoc_Click(Index As Integer)
   If Index = 0 Then
      LblTotItem.Visible = False
      txtTotItem.Visible = False
      'Add By Sindy 2023/3/21
      LblTotPage.Visible = False
      txtTotPage.Visible = False
      '2023/3/21 END
      'Add By Sindy 2019/1/16
      If cp(10) <> 實體審查 Then
         If chkDoc(0).Value = 1 Then
            LblTotItem.Visible = True
            txtTotItem.Visible = True
            'Add By Sindy 2023/3/21
            LblTotPage.Visible = True
            txtTotPage.Visible = True
            '2023/3/21 END
         End If
      End If
   '備註
   ElseIf Index = 1 Then
'      If chkDoc(1).Value = 0 Then
'         chkAtt(23).Enabled = False
'         chkAtt(23).Value = 0
'         chkAtt(24).Enabled = False
'         chkAtt(24).Value = 0
'         chkAtt(26).Enabled = False
'         chkAtt(26).Value = 0
'         chkAtt(27).Enabled = False
'         chkAtt(27).Value = 0
'         chkAtt(28).Enabled = False
'         chkAtt(28).Value = 0
'         chkAtt(29).Enabled = False
'         chkAtt(29).Value = 0
'         chkAtt(30).Enabled = False
'         chkAtt(30).Value = 0
'      Else
'         chkAtt(23).Enabled = True
'         chkAtt(24).Enabled = True
'         chkAtt(26).Enabled = True
'         chkAtt(27).Enabled = True
'         chkAtt(28).Enabled = True
'         chkAtt(29).Enabled = True
'         chkAtt(30).Enabled = True
'      End If
   '其他
   ElseIf Index = 2 Then
      If chkDoc(Index).Value = 0 Then
         chkAtt(21).Enabled = False
         chkAtt(21).Value = 0
         chkAtt(22).Enabled = False
         chkAtt(22).Value = 0
      Else
         chkAtt(21).Enabled = True
         chkAtt(22).Enabled = True
      End If
   'Add By Sindy 2018/1/11
   '一併修正專利名稱
   ElseIf Index = 3 Then
      If chkDoc(Index).Value = 1 Then
         chkDoc(1).Value = chkDoc(Index).Value
      Else
         If chkDoc(Index).Value = 0 And chkDoc(4).Value = 0 Then
            chkDoc(1).Value = 0
         End If
      End If
   'Add By Sindy 2018/4/17
   '首頁及摘要均附英文資料，減免規費800元整
   ElseIf Index = 4 Then
      If chkDoc(Index).Value = 1 Then
         chkDoc(1).Value = chkDoc(Index).Value
      Else
         If chkDoc(3).Value = 0 And chkDoc(Index).Value = 0 And chkDoc(6).Value = 0 Then
            chkDoc(1).Value = 0
         End If
      End If
      'Modify By Sindy 2019/12/4 Mark
      'Modify By Sindy 2025/7/21 詠心說要自動勾選
      chkDoc(2).Value = chkDoc(Index).Value '其他
      chkAtt(21).Value = chkDoc(Index).Value '文件描述
      chkAtt(22).Value = chkDoc(Index).Value '文件檔名
      '2019/12/4 END
   'Add By Sindy 2018/7/25
   '聲明中文本與申請時外文本實質內容一致
   ElseIf Index = 5 Then
      If chkDoc(Index).Value = 1 Then
         chkAtt(24).Value = chkDoc(Index).Value
      Else
         chkAtt(24).Value = chkDoc(Index).Value
      End If
      '2018/7/25 END
   'Add By Sindy 2021/8/31
   '已於提申時減免規費800元整，補呈首頁及摘要英文資料
   ElseIf Index = 6 Then
      If chkDoc(Index).Value = 1 Then
         chkDoc(1).Value = chkDoc(Index).Value
         chkAtt(12) = chkDoc(Index).Value  '申請書
         chkDoc(2) = chkDoc(Index).Value   '其他
         chkAtt(21) = chkDoc(Index).Value  '文件描述
         chkAtt(22) = chkDoc(Index).Value  '文件檔名
      Else
         If chkDoc(3).Value = 0 And chkDoc(Index).Value = 0 And chkDoc(4).Value = 0 Then
            chkDoc(1).Value = 0
         End If
      End If
   End If
End Sub

Private Sub cmdok_Click(Index As Integer)
Dim strFolder As String, strFileName As String
Dim m_Representative As String
   
   Select Case Index
      Case 0
         '重新檢查欄位有效性
         If TxtValidate = False Then Exit Sub
        'Added by Lydia 2020/02/21 產生各式申請書時，若基本檔「名稱有特殊字」已勾選，彈訊息提醒，並一併開啟原始檔。
        If (pa(1) = "FCP" Or pa(1) = "P") And pa(174) = "Y" Then
            MsgBox MsgText(1111), vbInformation
            If Pub_GetPA174toFile("0", pa(1), pa(2), pa(3), pa(4), Me, frm100101_M_1) = False Then
                Exit Sub
            End If
        End If
        'end 2020/02/21
            
         If FormSave = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Sub
         
         'Modified by Morgan 2018/3/8 檔名改數字第1碼0也要
         'm_CaseNo = pa(1) & IIf(Left(pa(2), 1) = "0", Mid(pa(2), 2), pa(2)) & IIf(pa(3) & pa(4) <> "000", pa(3) & pa(4), "")
         m_CaseNo = PUB_FCPCaseNo2FileName(pa(1), pa(2), pa(3), pa(4))
         'Ex.FCP47842 --> \\Typing2\專利案件\478\47842\FCP47842 (主管看過後會更改資料夾名稱加上 "-案件性質代碼" 如 FCP47842-101)
         'If Pub_StrUserSt03 = "M51" Then
         If UCase(pub_DbTerminalName) <> UCase(正式資料庫電腦名稱) Or Pub_StrUserSt03 = "M51" Then
            strFolder = PUB_Getdesktop
         Else
            strFolder = FCP電子送件檔案存放路徑
         End If
         'Modify By Sindy 2017/10/26 敏莉:請將電子送件所產生的新案及補文件申請書改路徑為:\\Typing2\電子送件暫存區\FCPXXXXX
'         strFolder = strFolder & "\" & Mid(m_CaseNo, 4, Len(m_CaseNo) - 5)
'         If Dir(strFolder, vbDirectory) = "" Then
'            MkDir strFolder
'         End If
         
'Removed by Morgan 2017/9/21 --葉敏莉
'         strFolder = strFolder & "\" & Mid(m_CaseNo, 4)
'         If Dir(strFolder, vbDirectory) = "" Then
'            MkDir strFolder
'         End If
         
         strFolder = strFolder & "\" & m_CaseNo
         If Dir(strFolder, vbDirectory) = "" Then
            MkDir strFolder
         End If
         
         'Add By Sindy 2018/11/23 判斷是否有變更代表人,若有,要傳入其資料,讀資料用
         If chkAtt(28).Value = 1 Then
            m_Representative = pa(79) & "@" & pa(80) & "@" & pa(81) & "@" & _
                               pa(82) & "@" & pa(83) & "@" & pa(84) & "@" & _
                               pa(109) & "@" & pa(110) & "@" & pa(111) & "@" & _
                               pa(112) & "@" & pa(113) & "@" & pa(114) & "@" & _
                               pa(115) & "@" & pa(116) & "@" & pa(117) & "@" & _
                               pa(118) & "@" & pa(119) & "@" & pa(120) & "@" & _
                               pa(121) & "@" & pa(122) & "@" & pa(123) & "@" & _
                               pa(124) & "@" & pa(125) & "@" & pa(126) & "@" & _
                               pa(127) & "@" & pa(128) & "@" & pa(129) & "@" & _
                               pa(130) & "@" & pa(131) & "@" & pa(132) & "@"
         Else
            m_Representative = ""
         End If
         '2018/11/23 END
         '1.基本資料
         StartLetterPA_EData "01", "14", strReceiveNo, pa, cp, False, IIf(chkAtt(26).Value = 1, True, False), , m_Representative, True
         NowPrint strReceiveNo, "01", "14", False, strUserNum, , , True, strExc(9)
         strFileName = strFolder & "\" & m_CaseNo & ".contact"
         Call PUB_MakeDoc(strExc(9), strFileName)
         
         'Add By Sindy 2018/1/29
         If cp(10) = 實體審查 Then
            '2.申請書
            If StartLetter2("01", "03") = False Then Exit Sub
            NowPrint strReceiveNo, "01", "03", False, strUserNum, , , True, strExc(9)
            strFileName = strFolder & "\" & "發明專利實體審查申請書"
            Call PUB_MakeDoc(strExc(9), strFileName)
         Else
         '2018/1/29 END
'            '1.基本資料
'            'Modified by Morgan 2018/1/17 刪除舊定稿(13)移除判斷
'            'Modify By Sindy 2014/11/14
'            'If strSrvDate(1) >= 專利發明人檔啟用日 Then
'               StartLetterPA_EData "01", "14", strReceiveNo, pa, cp, False
'               NowPrint strReceiveNo, "01", "14", False, strUserNum, , , True, strExc(9)
'            'Else
'            '2014/11/14 END
'            '   StartLetterPA_EData "01", "13", strReceiveNo, pa, cp
'            '   NowPrint strReceiveNo, "01", "13", False, strUserNum, , , True, strExc(9)
'            'End If
'            'end 2018/1/17
'            strFileName = strFolder & "\" & m_CaseNo & ".contact"
'            If PUB_MakeDoc(strExc(9), strFileName) = True Then
'               '目前 Html2Pdf 有問題,暫取消
'   '            If ConvertHtml2Pdf("A9801", strFileName) = False Then
'   '               Exit Sub
'   '            End If
'            End If
            
            '2.申請書
            If StartLetter2("01", "02") = False Then Exit Sub
            NowPrint strReceiveNo, "01", "02", False, strUserNum, , , True, strExc(9)
            strFileName = strFolder & "\" & "專利補正文件申請書"
            If PUB_MakeDoc(strExc(9), strFileName) = True Then
               '目前 Html2Pdf 有問題,暫取消
   '            If ConvertHtml2Pdf("P1202", strFileName) = False Then
   '               Exit Sub
   '            End If
            End If
            'Add By Sindy 2019/1/30
            Dim strCaseData As String
            Dim strBookDate As String
            If chkAtt(25).Value = 1 Then '委任書(附譯文)
               If Check3.Value = 1 Then
                  '個案
                  strCaseData = "　　有關台灣專利申請案第" & pa(11) & "號『" & pa(5) & "』之" & IIf(pa(8) = "1", "發明", IIf(pa(8) = "2", "新型", "設計")) & "專利申請案，"
               End If
               '委任日期
               strBookDate = Left(strSrvDate(2), 3) & " 年 " & Mid(strSrvDate(2), 4, 2) & " 月 " & Right(strSrvDate(2), 2) & " 日"
               If pa(26) <> "" Then
                  EndLetter "01", strReceiveNo, "31", strUserNum
                  '個案
                  If Check3.Value = 1 Then
                     strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                 "('01','" & strReceiveNo & "','31','" & strUserNum & "','個案','" & strCaseData & "')"
                     cnnConnection.Execute strExc(0)
                  End If
                  '委任日期
                  strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              "('01','" & strReceiveNo & "','31','" & strUserNum & "','委任日期','" & strBookDate & "')"
                  cnnConnection.Execute strExc(0)
                  '申請人國籍
                  strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              "('01','" & strReceiveNo & "','31','" & strUserNum & "','申請人國籍','　" & Trim(GetNationName(GetPrjNationNumber1(ChangeCustomerL(pa(26))))) & "　')"
                  cnnConnection.Execute strExc(0)
                  NowPrint strReceiveNo, "01", "31", False, strUserNum, , , True, strExc(9)
                  strFileName = strFolder & "\" & "委任狀中譯文1"
                  Call PUB_MakeDoc(strExc(9), strFileName, , pa(1), "01", "31")
               End If
               If pa(27) <> "" Then
                  EndLetter "01", strReceiveNo, "32", strUserNum
                  '個案
                  If Check3.Value = 1 Then
                     strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                 "('01','" & strReceiveNo & "','32','" & strUserNum & "','個案','" & strCaseData & "')"
                     cnnConnection.Execute strExc(0)
                  End If
                  '委任日期
                  strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              "('01','" & strReceiveNo & "','32','" & strUserNum & "','委任日期','" & strBookDate & "')"
                  cnnConnection.Execute strExc(0)
                  '申請人國籍
                  strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              "('01','" & strReceiveNo & "','32','" & strUserNum & "','申請人國籍','　" & Trim(GetNationName(GetPrjNationNumber1(ChangeCustomerL(pa(27))))) & "　')"
                  cnnConnection.Execute strExc(0)
                  NowPrint strReceiveNo, "01", "32", False, strUserNum, , , True, strExc(9)
                  strFileName = strFolder & "\" & "委任狀中譯文2"
                  Call PUB_MakeDoc(strExc(9), strFileName, , pa(1), "01", "32")
               End If
               If pa(28) <> "" Then
                  EndLetter "01", strReceiveNo, "33", strUserNum
                  '個案
                  If Check3.Value = 1 Then
                     strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                 "('01','" & strReceiveNo & "','33','" & strUserNum & "','個案','" & strCaseData & "')"
                     cnnConnection.Execute strExc(0)
                  End If
                  '委任日期
                  strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              "('01','" & strReceiveNo & "','33','" & strUserNum & "','委任日期','" & strBookDate & "')"
                  cnnConnection.Execute strExc(0)
                  '申請人國籍
                  strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              "('01','" & strReceiveNo & "','33','" & strUserNum & "','申請人國籍','　" & Trim(GetNationName(GetPrjNationNumber1(ChangeCustomerL(pa(28))))) & "　')"
                  cnnConnection.Execute strExc(0)
                  NowPrint strReceiveNo, "01", "33", False, strUserNum, , , True, strExc(9)
                  strFileName = strFolder & "\" & "委任狀中譯文3"
                  Call PUB_MakeDoc(strExc(9), strFileName, , pa(1), "01", "33")
               End If
               If pa(29) <> "" Then
                  EndLetter "01", strReceiveNo, "34", strUserNum
                  '個案
                  If Check3.Value = 1 Then
                     strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                 "('01','" & strReceiveNo & "','34','" & strUserNum & "','個案','" & strCaseData & "')"
                     cnnConnection.Execute strExc(0)
                  End If
                  '委任日期
                  strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              "('01','" & strReceiveNo & "','34','" & strUserNum & "','委任日期','" & strBookDate & "')"
                  cnnConnection.Execute strExc(0)
                  '申請人國籍
                  strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              "('01','" & strReceiveNo & "','34','" & strUserNum & "','申請人國籍','　" & Trim(GetNationName(GetPrjNationNumber1(ChangeCustomerL(pa(29))))) & "　')"
                  cnnConnection.Execute strExc(0)
                  NowPrint strReceiveNo, "01", "34", False, strUserNum, , , True, strExc(9)
                  strFileName = strFolder & "\" & "委任狀中譯文4"
                  Call PUB_MakeDoc(strExc(9), strFileName, , pa(1), "01", "34")
               End If
               If pa(30) <> "" Then
                  EndLetter "01", strReceiveNo, "35", strUserNum
                  '個案
                  If Check3.Value = 1 Then
                     strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                 "('01','" & strReceiveNo & "','35','" & strUserNum & "','個案','" & strCaseData & "')"
                     cnnConnection.Execute strExc(0)
                  End If
                  '委任日期
                  strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              "('01','" & strReceiveNo & "','35','" & strUserNum & "','委任日期','" & strBookDate & "')"
                  cnnConnection.Execute strExc(0)
                  '申請人國籍
                  strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              "('01','" & strReceiveNo & "','35','" & strUserNum & "','申請人國籍','　" & Trim(GetNationName(GetPrjNationNumber1(ChangeCustomerL(pa(30))))) & "　')"
                  cnnConnection.Execute strExc(0)
                  NowPrint strReceiveNo, "01", "35", False, strUserNum, , , True, strExc(9)
                  strFileName = strFolder & "\" & "委任狀中譯文5"
                  Call PUB_MakeDoc(strExc(9), strFileName, , pa(1), "01", "35")
               End If
            End If
            '2019/1/30 END
         End If
         
         frm060103_1.Show
         frm060103_1.ClearForm
         Unload Me
      Case 2
         frm060103_1.Show
         frm060103_1.cmdok_Click 1
         Unload Me
   End Select
End Sub

Private Sub Form_Activate()
   If m_IsRun = False Then
      m_IsRun = True
      
      '計算規費:
      Call_PUB_SetOfficialFee_P
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   m_IsRun = False 'Add By Sindy 2023/3/24
   
   intWhere = 國外_FC
   With frm060103_1
      Text1 = .Text1
      Text2 = .Text2
      Text3 = .Text3
      Text4 = .Text4
      strReceiveNo = .Tag
   End With
   ReDim pa(1 To TF_PA) As String
   ReDim cp(TF_CP)
   ReadPatent
   
   Label12.Visible = False
   'Add By Sindy 2018/1/29
   If cp(10) = 實體審查 Then
      'Modify By Sindy 2018/10/9 一併送中說 ex:FCP-59049
'      Frame1.Enabled = False
      For intI = 5 To 20
         If intI <> 11 Then
            chkAtt(intI).Enabled = False
         End If
      Next intI
      '2018/10/9 END
      'Add By Sindy 2019/6/19 ex:FCP-061202
      'Frame2.Enabled = False
      Frame2.Enabled = True
      '2019/6/19 END
      chkDoc(1).Enabled = False
      chkDoc(3).Enabled = False
      chkDoc(4).Enabled = False
      chkAtt(23).Enabled = False
      chkAtt(24).Enabled = False
      chkDoc(0).Value = 1
   End If
   '2018/1/29 END
   
   lstNameAgent.Clear
   PUB_SetOurAgent lstNameAgent, pa(), cp(110), cp(10), True
   'Added by Sindy 2021/5/10 如果一開始將ListBox拉到需要的大小，字型會自動放大；所以畫面預設為一列高度，Form_Load才放大到需要的大小
   lstNameAgent.Height = 1100
   lstNameAgent.Width = 1300
   
   SSTab1.Tab = 0
   
   cboResonCode.Clear
   cboResonCode.AddItem "24000 補送申請書"
   cboResonCode.AddItem "24002 補送說明書或圖式"
   cboResonCode.AddItem "24036 補送委任書"
   cboResonCode.AddItem "24046 補送生物材料寄存文件"
   cboResonCode.AddItem "24048 補送優先權證明文件"
   cboResonCode.AddItem "24054 補送優惠期證明文件"
   cboResonCode.AddItem "26000 繳納申請費"
   cboResonCode.AddItem "30700 其他"
   cboResonCode.ListIndex = 0
   
   'Add By Sindy 2018/7/25 設計案
   If pa(8) <> "1" Then
      chkAtt(11).Enabled = False '序列表
      chkAtt(16).Enabled = False '國內生物材料寄存證明文件
      chkAtt(17).Enabled = False '國外生物材料寄存證明文件
      chkAtt(18).Enabled = False '生物材料為通常知識者易於獲得證明文件
      '設計
      If pa(8) = "3" Then
         '摘要頁數
         Label15.Visible = False
         txtDocCh(0).Visible = False
         txtDocCh(0).Enabled = False
         '序列表
         Label16(1).Visible = False
         txtDocCh(4).Visible = False: Label1(1).Visible = False
         txtDocCh(4).Enabled = False
         '申請專利範圍頁數
         Label17.Visible = False
         txtDocCh(2).Visible = False
         txtDocCh(2).Enabled = False
         '申請專利範圍項數
         Label21.Visible = False
         txtCP136.Visible = False
         txtCP136.Enabled = False
      End If
   End If
   '2018/7/25 END
   
   FraPA174.BackColor = &H8000000F 'Added by Lydia 2020/02/21
   
   'Add By Sindy 2023/4/7
   If Pub_StrUserSt03 <> "M51" Then
      txtDecreasePageFee.Visible = False
      txtDecreaseItemFee.Visible = False
   End If
   '2023/4/7 END
End Sub

'************************************************
' 取回專利基本資料及收文資料
'
'************************************************
Private Sub ReadPatent()
Dim i As Integer, j As Integer, Lbl As Object, strTempName As String
Dim strCE01 As String 'Add By Sindy 2023/3/1
Dim m_TmpallPage As String, m_TmpallItem As String '總頁數,總項數
Dim strChgPA64 As String, strChgPA65 As String, strChgPA67 As String, strChgPA68 As String 'Add By Sindy 2023/3/16
   
   ReDim pageD(1 To 21) As String 'Add By Sindy 2023/3/15
   m_bolChkPageItem = False 'Add By Sindy 2018/2/21
   m_bolChkItem = False 'Add By Sindy 2018/2/21
   pa(1) = Text1
   pa(2) = Text2
   pa(3) = Text3
   pa(4) = Text4
   For Each Lbl In Label7
      Lbl.Caption = ""
   Next
   
   If ClsPDReadPatentDatabase(pa(), intWhere) Then
      For i = 0 To 1 '案件名稱
         Text6(i) = pa(i + 5)
      Next
   End If
   
   cp(9) = strReceiveNo
   If PUB_ReadCaseProgressDatabase(cp(), intWhere) Then
      If ClsPDGetStaff(cp(13), strExc(0)) Then Label7(12) = strExc(0)
      If ClsPDGetStaff(cp(14), strExc(0)) Then Label7(11) = strExc(0)
      If ClsPDGetCaseProperty("FCP", cp(10), strExc(0)) Then Label7(10) = strExc(0)
   End If
   
   'Add By Sindy 2018/5/22 敏莉說取消預設
'   '發明
'   'Add By Sindy 2018/2/22
'   If cp(10) <> 實體審查 Then
'   '2018/2/22 END
'      If pa(8) = "1" Then
'         chkAtt(1).Enabled = True
'         chkAtt(1).Value = 1 'Add By Sindy 2018/1/11
'         chkAtt(2).Enabled = True
'         chkAtt(2).Value = 1 'Add By Sindy 2018/1/11
'         chkAtt(11).Enabled = True
'         chkAtt(11).Value = 1 'Add By Sindy 2018/1/11
'         chkAtt(3).Enabled = True
'         chkAtt(3).Value = 1 'Add By Sindy 2018/1/11
'         chkAtt(4).Enabled = True
'         chkAtt(4).Value = 1 'Add By Sindy 2018/1/11
'      '新型
'      ElseIf pa(8) = "2" Then
'         chkAtt(5).Enabled = True
'         chkAtt(5).Value = 1 'Add By Sindy 2018/1/11
'         chkAtt(6).Enabled = True
'         chkAtt(6).Value = 1 'Add By Sindy 2018/1/11
'         chkAtt(7).Enabled = True
'         chkAtt(7).Value = 1 'Add By Sindy 2018/1/11
'         chkAtt(8).Enabled = True
'         chkAtt(8).Value = 1 'Add By Sindy 2018/1/11
'      '設計
'      Else
'         txtDocCh(0).Enabled = False
'         txtDocCh(2).Enabled = False
'         txtCP136.Enabled = False
'         chkAtt(9).Enabled = True
'         chkAtt(9).Value = 1 'Add By Sindy 2018/1/11
'         chkAtt(10).Enabled = True
'         chkAtt(10).Value = 1 'Add By Sindy 2018/1/11
'      End If
'   End If
   'Add By Sindy 2023/3/1 增加檢查是否有變更資料
   strExc(0) = "select cp09 From caseprogress" & _
               " where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'" & _
               " and cp10='401' and cp27||cp57 is null" & _
               " order by cp66 desc,cp67 desc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      strCE01 = RsTemp.Fields("cp09")
      '預設,同時辦理事項
      If PUB_GetChangeEvent(strCE01, 1) = True Then
         chkAtt(26).Value = 1: chkAtt(26).Tag = strCE01
      End If
      If PUB_GetChangeEvent(strCE01, 2) = True Then
         chkAtt(27).Value = 1: chkAtt(27).Tag = strCE01
      End If
      If PUB_GetChangeEvent(strCE01, 3) = True Then
         chkAtt(28).Value = 1: chkAtt(28).Tag = strCE01
      End If
      If PUB_GetChangeEvent(strCE01, 4) = True Then
         chkAtt(29).Value = 1: chkAtt(29).Tag = strCE01
      End If
      If PUB_GetChangeEvent(strCE01, 5) = True Then
         chkAtt(30).Value = 1: chkAtt(30).Tag = strCE01
      End If
   End If
   '2023/3/1 END
   
   '來函文號:
   'Modify By Sindy 2019/11/15 敏莉說補收款未必掛對應的相關總收文號,所以不需要抓來函文號
   If cp(10) <> 補收款 Then
   '2019/11/15 END
      'Modified by Morgan 2018/3/20 CP05 Desc -> NVL(ED08,CP05) Desc 歸卷公文會有一筆以上，考慮可能有紙本公文保留CP05的判斷
      'Modified by Morgan 2018/11/5 +1003通知補文件--敏莉 Ex:FCP057627
      'Modified by Sindy 2019/7/30 +125衍生設計--敏莉 Ex:FCP061447
      'Modified by Morgan 2022/5/12 +307分割(435 續行母案再審要抓) --陳亭妙
      strExc(0) = "select cp08,ed08 from caseprogress,edocument where CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' AND CP10 IN ('1003','1201','1004','101','102','103','125','307') AND ed11(+)=cp09 ORDER BY NVL(ED08,CP05) DESC"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If Not IsNull(RsTemp("ED08")) Then
            Text5 = RsTemp("ED08") - 19110000
            If Not IsNull(RsTemp("cp08")) Then
               strExc(0) = Mid(RsTemp("cp08"), InStr(RsTemp("cp08"), "智專") + 2, Len(RsTemp("cp08")))
               Text7 = Mid(strExc(0), 1, InStr(strExc(0), "字") - 1)
               strExc(0) = Replace(strExc(0), Text7 & "字第", "")
               Text8 = Mid(strExc(0), 1, InStr(strExc(0), "號") - 1)
            End If
         End If
      End If
   End If
   
   '*************************************************************************
   'Added by Morgan 2013/1/8
   m_bol107NewFee = True
   bolDelay = False
   'end 2013/1/8
   'Modify by Morgan 2006/8/18 加判斷107(再審),803(舉發),301,302,303,305(改請)才要
   'Modified by Morgan 2013/8/26 +507 -- FCP032929
   If InStr("107,803,301,302,303,305,507", cp(10)) > 0 Then
      'Add by Morgan 2004/9/8 檢查是否有延期，若有則規費預設0
      bolDelay = PUB_ChkDelay(strReceiveNo, m_strDelayCP09, strExc(1))
      If bolDelay = True Then
         If strExc(1) < "20130101" Then m_bol107NewFee = False 'Added by Morgan 2013/1/9
         cp(17) = "0"
      End If
   End If

   'Add by Morgan 2004/8/12
   txtCP84.Tag = cp(17)
   txtCP84.Text = txtCP84.Tag
   
   'Add by Morgan 2010/1/6 增加頁數,項數欄位
   m_bolChkFee = False
'   txtItem.Enabled = False
'   txtItem.BackColor = Me.BackColor
'   txtCP135.Enabled = False
'   txtCP135.BackColor = Me.BackColor
'   txtCP136.Enabled = False
'   txtCP136.BackColor = Me.BackColor
'   txtCP137.Enabled = False
'   txtCP137.BackColor = Me.BackColor
'   txtCP138.Enabled = False
'   txtCP138.BackColor = Me.BackColor
'   txtCount.Enabled = False
'   txtCount.BackColor = Me.BackColor
'   txtAddFee.Enabled = False
'   txtAddFee.BackColor = Me.BackColor
'   txtDecreaseFee.Enabled = False
'   txtDecreaseFee.BackColor = Me.BackColor
   txtCP84.Enabled = True
   
   'Added by Lydia 2018/12/27 中文本資訊-各項頁數
   txtDocCh(0).Text = pa(64) '摘要頁數
   txtDocCh(1).Text = pa(65) '說明書頁數
   txtDocCh(4).Text = pa(66) '序列表頁數
   txtDocCh(2).Text = pa(67) '申請專利範圍頁數
   txtDocCh(3).Text = pa(68) '圖式頁數
   'Added by Lydia 2019/01/10
   txtCP136.Text = pa(172) '申請專利範圍項數(最初項數)
   txtTotItem.Text = pa(172) '申請專利範圍項數(最初項數)
   txtDocCh(6).Text = pa(173) '圖式圖數
   'end 2019/01/10
   If Val(pa(64)) + Val(pa(65)) + Val(pa(66)) + Val(pa(67)) + Val(pa(68)) > 0 Then
      'Modify By Sindy 2019/3/20 各式申請書-電子送件-補文件 若有中文本資訊請不要自動帶到申請書，請將預設的勾勾拿掉
      'chkDoc(0).Value = 1
      'Modify By Sindy 2019/8/8 + 911.補收款
      'Modify By Sindy 2019/8/8 + 435.續行母案再審
      If cp(10) <> 補文件 And cp(10) <> 補收款 And cp(10) <> "435" Then chkDoc(0).Value = 1
      '2019/3/20 END
      Call txtDocCh_Validate(0, False)
      Call txtDocCh_Validate(1, False)
      Call txtDocCh_Validate(4, False)
      Call txtDocCh_Validate(2, False)
      Call txtDocCh_Validate(3, False)
   End If
   For Each oText In txtDocCh
      oText.Tag = oText.Text
   Next
   'end 2018/12/27
   
   'Add By Sindy 2018/4/23
   '讀取總頁數和總項數(統計已發文)
   'Modified by Lydia 2018/12/27 預設基本檔的頁數總計
   'm_allPage = 0: m_allItem = 0
   m_allPage = Val(txtCP135)
   m_allItem = Val(txtCP136)
   'end 2018/12/27
   'Add By Sindy 2019/1/22 因為會影響是否要加基本檔項數問題,改移至此處詢問
   m_WriteNote = "" 'Add By Sindy 2018/8/7
   '中說:
   If cp(10) = "201" Or cp(10) = "209" Or cp(10) = "235" Or cp(10) = "210" Then
      '實審已發文
      'If PUB_ChkCPExist(cp, "416", 2) Then
         If PUB_ChkCPExist(cp, "203", 1, str203CP09) Then
            If MsgBox("是否一併主動修正？", vbInformation + vbYesNo + vbDefaultButton2) = vbYes Then
               m_WriteNote = "Y"
            End If
         End If
      'End If
   End If
   If m_WriteNote <> "" Then Me.Caption = Me.Caption & "(" & m_WriteNote & ")" 'Add By Sindy 2018/8/7
   '2019/1/22 END
   'Add By Sindy 2023/3/15 取得總頁數/總項數
   Call PUB_GetAllPageItem(cp(9), cp, pa, m_TmpallPage, m_TmpallItem)
   
   'Modify By Sindy 2023/3/16
   txtTotPage = Val(txtCP135)
   If m_WriteNote = "Y" Then '有一併送主動修正才要加基本檔頁數
      '讀取專利說明書頁數明細
      Call PUB_ReadPageDetail(str203CP09, pageD, , , , strChgPA64, strChgPA65, strChgPA67, strChgPA68)
      '有一併送主動修正才要加基本檔頁數
      txtTotPage = Val(m_TmpallPage) + Val(txtCP135)
'         If Val(strChgPA64) <> 0 Then
'            txtDocCh(0) = txtDocCh(0) + Val(strChgPA64)
'         End If
'         If Val(strChgPA65) <> 0 Then
'            txtDocCh(1) = txtDocCh(1) + Val(strChgPA65)
'         End If
'         If Val(strChgPA67) <> 0 Then
'            txtDocCh(2) = txtDocCh(2) + Val(strChgPA67)
'         End If
'         If Val(strChgPA68) <> 0 Then
'            txtDocCh(3) = txtDocCh(3) + Val(strChgPA68)
'         End If
      'Modify By Sindy 2023/8/18 +210製作中說排除鎖頁項數欄位,可修改(因不會經中打室)
      'Modify By Sindy 2025/9/23 再開放 235核對中說格式,209檢視中說
      'If cp(10) <> "210" Then
      If cp(10) <> "210" And _
         Not (pa(8) = "3" And (cp(10) = "235" Or cp(10) = "210")) Then
      '2023/8/18 END
         '欄位鎖住,不可調整
         For Each oText In txtDocCh
            oText.Locked = True
            oText.BackColor = &H8000000F
         Next
         txtCP136.Locked = True
         txtCP136.BackColor = &H8000000F
         '有增修才要加基本檔項數
      End If
      txtTotItem = Val(m_TmpallItem) + Val(txtTotItem)
   End If
   '2023/3/16 END
   
'   'Modify By Sindy 2018/5/21 取消 and cp158>0,改為 and cp159=0
'   If Val(cp(135)) > 0 Then
'      'Modify By Sindy 2022/6/10 Ex:中說 Our Ref: FCP-066532:中打室有修改明細,但總頁數用進度的反而錯了
'      'Modify By Sindy 2023/12/5 Ex:實體審查 FCP-064392 反應程式應該用進度算出來的數值
''      If Val(m_allPage) = 0 Then
''      '2022/6/10 END
'         m_allPage = Val(cp(135))
''      End If
'   Else
   'Modify By Sindy 2023/12/6 中說一定抓基本檔 ex:FCP-069989
   If Not (cp(10) = "201" Or cp(10) = "209" Or cp(10) = "235" Or cp(10) = "210") Then
      'Modify By Sindy 2023/12/5 Ex:實體審查 FCP-064392 反應程式應該用進度算出來的數值
      '總頁數:最近一筆進度的頁數
      'Ex:FCP-62225 新案翻譯
      If Val(m_TmpallPage) > 0 Then
         m_allPage = Val(m_TmpallPage)
      End If
'      If Val(m_allPage) = 0 Then
'         'Modify By Sindy 2023/3/15
'         m_allPage = Val(m_TmpallPage)
''         strExc(0) = "select cp09,cp10,nvl(cp135,0) from caseprogress" & _
''                     " where CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "'" & _
''                     " and cp159=0" & _
''                     " and nvl(cp135,0)>0" & _
''                     " ORDER BY CP69 DESC,CP70 DESC"
''         intI = 1
''         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
''         If intI = 1 Then
''            m_allPage = Val("" & RsTemp.Fields(2))
''         End If
'         '2023/3/15 END
'      End If
      '2023/12/5 END
   End If
   txtCP135.Text = m_allPage '總頁數
'   If Val(cp(136)) > 0 Then
'      'Modify By Sindy 2022/6/10 Ex:中說 Our Ref: FCP-066532:中打室有修改明細,但總頁數用進度的反而錯了
'      'Modify By Sindy 2023/12/5 Ex:實體審查 FCP-064392 反應程式應該用進度算出來的數值
''      If Val(m_allItem) = 0 Then
''      '2022/6/10 END
'         m_allItem = Val(cp(136))
''      End If
'   Else
   'Modify By Sindy 2023/12/6 中說一定抓基本檔 ex:FCP-069989
   If Not (cp(10) = "201" Or cp(10) = "209" Or cp(10) = "235" Or cp(10) = "210") Then
      'Modify By Sindy 2023/12/5 Ex:實體審查 FCP-064392 反應程式應該用進度算出來的數值
'      'Modify By Sindy 2023/3/15
'      If Val(m_allItem) = 0 Then
'         'Modify By Sindy 2023/3/15
'         m_allItem = Val(m_TmpallItem)
'      End If
      'Modify By Sindy 2023/3/15
      If Val(m_TmpallItem) > 0 Then
         'Modify By Sindy 2023/3/15
         m_allItem = Val(m_TmpallItem)
      End If
      '2023/12/5 END
'      '總項數:增加項數-刪除未審項數-刪除已審項數
'      strExc(0) = "select sum(nvl(cp136,0)),sum(nvl(cp137,0)),sum(nvl(cp138,0)) from caseprogress" & _
'                  " where CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "'" & _
'                  " and cp159=0" & _
'                  " and (nvl(cp136,0)>0 or nvl(cp137,0)>0 or nvl(cp138,0)>0)" & _
'                  " ORDER BY CP69 DESC,CP70 DESC"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         'Add By Sindy 2019/1/25
'         If Val("" & RsTemp.Fields(0)) - Val("" & RsTemp.Fields(1)) - Val("" & RsTemp.Fields(2)) <> 0 Then
'         '2019/1/25 END
'            m_allItem = Val("" & RsTemp.Fields(0)) - Val("" & RsTemp.Fields(1)) - Val("" & RsTemp.Fields(2))
'            'Modify By Sindy 2019/1/30 有增修才要加基本檔項數 ex:FCP-059712
'            If m_WriteNote = "Y" Then 'Add By Sindy 2019/1/22 有一併送主動修正才要加基本檔項數
'               m_allItem = Val(m_allItem) + Val(txtPA172) 'Add By Sindy 2019/1/16 ex:FCP-59599
'            End If
'         End If
'      End If
      '2023/3/15 END
   End If
   txtCP136.Text = m_allItem '總項數
   '2018/4/23 END
   
   m_strReExamCP27 = "" 'Added by Morgan 2013/1/10
   m_bolFixNewFee = False 'Added by Morgan 2013/1/10
   'Modified by Morgan 2013/11/6 +235核對中說格式
'   416.實體審查
'   201.新案翻譯
'   209.檢視中說
'   235.核對中說格式
'   210.製作中說
'   307.分割
'   203.主動修正
'   204.修正
'   205.申復
'   206.補充說明
   If (cp(10) = "416" Or cp(10) = "201" Or cp(10) = "209" Or cp(10) = "235" Or cp(10) = "210" Or cp(10) = "203" Or cp(10) = "204" Or cp(10) = "205" Or cp(10) = "206") Then
      m_Div416OfficialFee = 0 '2010/12/8 add by sonia
      
      'Added by Morgan 2013/1/10
      If (cp(10) = "210" Or cp(10) = "203" Or cp(10) = "204" Or cp(10) = "205" Or cp(10) = "206") Then
         m_strReExamCP27 = PUB_GetReExamDate(cp)
         If m_strReExamCP27 > "20130000" Then
            m_bolFixNewFee = True
         End If
      End If
      'end 2013/1/10
      
      If m_strReExamCP27 = "" Then 'Added by Morgan 2013/1/10
         m_bol99Case = Chk99NewCase(cp(1), cp(2), cp(3), cp(4))
         If m_bol99Case Then
            '實審:
            If cp(10) = "416" Then
               '新案翻譯已發文
               'Modify by Morgan 2010/4/28 +307
               'Modified by Morgan 2013/11/6 +235核對中說格式
               '中說已發文
               If PUB_ChkCPExist(cp, "201", 2) Or PUB_ChkCPExist(cp, "209", 2) Or PUB_ChkCPExist(cp, "235", 2) Or PUB_ChkCPExist(cp, "210", 2) Or PUB_ChkCPExist(cp, "307", 2) Then
                  m_bolChkFee = True
                  m_bolChkPageItem = True
                  txtCP84.Enabled = False
                  'Add By Sindy 2018/4/23
                  txtCP135.Text = m_allPage '總頁數
                  txtCP136.Text = m_allItem '總項數
                  '2018/4/23 END
                  
'                  txtCP135.Enabled = True
'                  txtCP135.BackColor = vbWhite
'                  lblCP136.Caption = "總項數:"
'                  txtCP136.Enabled = True
'                  txtCP136.BackColor = vbWhite
               'Add By Sindy 2018/4/23
               '中說未收文或未發文
               Else
                  'Modify By Sindy 2018/8/14 ex:FCP-58675
'                  m_bolChkPageItem = False
'                  m_bolChkFee = False
                  If MsgBox("是否一併送中說？", vbInformation + vbYesNo + vbDefaultButton2) = vbYes Then
                     m_bolChkPageItem = True
                     m_bolChkFee = True
                  Else
                     m_bolChkPageItem = False
                     m_bolChkFee = False
                  End If
                  'txtCP84.Text = "7000" 'Add By Sindy 2018/4/23 7000應該不用寫死在程式裡,因為收實審時規費會輸入7000
               End If
               'Add By Sindy 2018/4/25
               '不回寫頁/項數
               Call GetCP31isY_CP05(cp(1), cp(2), cp(3), cp(4), "cp118", strNewCaseCP118)
               If strNewCaseCP118 = "Y" Or strNewCaseCP118 = "A" Then '新案電子送件時,才不回寫走新規則
                  m_bolNotSavePageItem = True
               End If
               '2018/4/25 END
'               For Each oText In txtDocCh
'                  oText.Enabled = False
'               Next
'               txtCP135.Enabled = False
'               txtCP136.Enabled = False
               '2018/4/23 END
            '新案翻譯,檢視中說,製作中說
            'Modified by Morgan 2013/11/6 +235核對中說格式
            '中說:
            ElseIf cp(10) = "201" Or cp(10) = "209" Or cp(10) = "235" Or cp(10) = "210" Then
               '實審已發文
               If PUB_ChkCPExist(cp, "416", 2) Then
                  m_bolChkPageItem = True
                  m_bolChkFee = True
                  'Modify By Sindy 2019/1/22 Mark, 改移至上頭詢問
'                  m_WriteNote = "N" 'Add By Sindy 2018/8/7
'                  If PUB_ChkCPExist(cp, "203", 1) Then
'                     If MsgBox("是否一併主動修正？", vbInformation + vbYesNo + vbDefaultButton2) = vbYes Then
'                        'm_bolChkFee = False
'                        'txtCP84.Text = "0" 'Add By Sindy 2018/4/23 繳費金額固定為0
'                        m_WriteNote = "Y"
'                     End If
'                  End If
'                  Me.Caption = Me.Caption & "(" & m_WriteNote & ")" 'Add By Sindy 2018/8/7
                  '2019/1/22 END
                  
'                  txtCP135.Enabled = True
'                  txtCP135.BackColor = vbWhite
'                  lblCP136.Caption = "總項數:"
'                  txtCP136.Enabled = True
'                  txtCP136.BackColor = vbWhite
               'Add By Sindy 2018/4/23
               '實審未收文或未發文
               Else
                  m_bolChkPageItem = False
                  m_bolChkFee = False
                  txtCP84.Text = "0" 'Add By Sindy 2018/4/23 繳費金額固定為0
               End If
               txtCP84.Enabled = False
               'Modify By Sindy 2020/4/20 Ex:FCP-62225
'               '回寫頁/項數
''               For Each oText In txtDocCh
''                  oText.Enabled = True
''               Next
''               txtCP135.Enabled = True
'               If Val(cp(135)) > 0 Then txtCP135 = cp(135)
''               txtCP136.Enabled = True
'               If Val(cp(136)) > 0 Then txtCP136 = cp(136)
'               '2018/4/23 END
            End If
            If m_bolChkFee Then
               Call_PUB_SetOfficialFee_P
            End If
         '2010/12/8 ADD BY SONIA 分割案之實審發文,若母案有已收未取消的再審程序且申請日在2010/1/1以前者,分割案實審規費應為8000元
         '2011/3/24 MODIFY BY SONIA FCP-034512申復發文誤帶規費8000
         'ElseIf PUB_ChkCPExist(cp, "307") Then
         ElseIf cp(10) = "416" And PUB_ChkCPExist(cp, "307") Then
            txtCP84 = "8000"
            'Modify by Morgan 2011/7/26 會有超頁費 Ex.FCP-044051
            'txtCP84.Enabled = False
            MsgBox "本案請依舊法規則計算規費！"
            'end 2011/7/26
            cp(17) = txtCP84      '同時改收文規費
            m_Div416OfficialFee = txtCP84
         '2010/12/8 END
         End If
         
      'Added by Morgan 2013/1/10
      '有再審102年後發文,修正要收超項費
      ElseIf m_bolFixNewFee = True Then
         m_bolChkItem = True
         m_bolChkFee = True
         txtCP84.Enabled = False
'         lblCP136.Caption = "增加項數"
'         txtCP136.Enabled = True
'         txtCP136.BackColor = vbWhite
'         txtCP137.Enabled = True
'         txtCP137.BackColor = vbWhite
'         txtCP138.Enabled = True
'         txtCP138.BackColor = vbWhite
'         'txtItem.Enabled = True
'         txtItem.BackColor = vbWhite
'         txtCount.Enabled = True
'         txtCount.BackColor = vbWhite
'         txtAddFee.Enabled = True
'         txtAddFee.BackColor = vbWhite
'         txtDecreaseFee.Enabled = True
'         txtDecreaseFee.BackColor = vbWhite
      End If
      'end 2013/1/10
   'Add By Sindy 2019/7/22
   ElseIf cp(10) = "435" Then '續行母案再審
      m_bolChkItem = True
      m_bolChkFee = True
      txtCP84.Enabled = False
   '2019/7/22 END
   End If
   '*************************************************************************
   
   'Added by Lydia 2018/04/20  若工程師提申後從命名系統修改專利名稱則自動加註"一併修改專利名稱",待中說或補文件發文後自動將註記的欄位清空。
   strExc(0) = "select cp09,cp10,tct01,tct15 from caseprogress,transcasetitle where CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' AND cp31='Y' and cp09=tct01(+) "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If "" & RsTemp.Fields("tct01") <> "" And "" & RsTemp.Fields("tct15") = "Y" Then
         chkDoc(3).Value = vbChecked
      End If
   End If
   'end 2018/04/20
   
   'Added by Lydia 2020/02/21 預設「名稱有特殊字」
   FraPA174.Visible = False
   If pa(1) = "FCP" Or pa(1) = "P" Then
       If pa(174) = "Y" Then
          FraPA174.Visible = True
       End If
   End If
   'end 2020/02/21
   
   'Add By Sindy 2021/10/27
   If cp(10) = "201" Or cp(10) = "209" Or cp(10) = "235" Then
   '2021/10/27 END
      'Add By Sindy 2021/8/31
      strExc(0) = "SELECT CP09,CP10,CP14,CP27,CP60,TF30 FROM CaseProgress,TransFee" & _
                  " WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                  " AND CP10 in ('201','209','235') AND CP159=0 AND CP09=TF01(+)"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         '有英文本收文號預設為勾選
         If "" & RsTemp.Fields("TF30") <> "" Then
            chkDoc(6).Value = 1 '首頁及摘要均附英文資料，減免規費800元整
            Call chkDoc_Click(6)
         End If
      End If
   End If
End Sub

Private Sub lstNameAgent_Validate(Cancel As Boolean)
   Dim ii As Integer
   Cancel = True
   cp(110) = "": m_AgentName = ""
   For ii = 0 To lstNameAgent.ListCount - 1
      If lstNameAgent.Selected(ii) = True Then
         'modify by sonia 2016/10/7 員工編號已可非數字需做轉換
         'cp(110) = cp(110) & "," & lstNameAgent.ItemData(ii)
         'Modify By Sindy 2021/5/10
         'cp(110) = cp(110) & "," & PUB_Num2Id(lstNameAgent.ItemData(ii))
         cp(110) = cp(110) & "," & PUB_GetItemData(lstNameAgent.Tag, ii)
         m_AgentName = m_AgentName & "、" & lstNameAgent.List(ii)
         '2021/5/10 END
         Cancel = False
      End If
   Next
   If Cancel = True Then
      MsgBox "出名代理人不可空白！", vbExclamation
   Else
      If Left(cp(110), 1) = "," Then cp(110) = Mid(cp(110), 2)
      m_AgentName = Mid(m_AgentName, 2) 'Add By Sindy 2021/5/10
   End If
End Sub

Private Sub Text5_GotFocus()
   TextInverse Text5
   CloseIme
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text5_Validate(Cancel As Boolean)
   If Text5 <> "" Then
      If Not ChkDate(Text5) Then
         Cancel = True
      ElseIf Val(Text5) > Val(strSrvDate(2)) Then
         MsgBox "發文日期不可大於系統日！"
         Cancel = True
      Else
         Text9 = Val(Text5) \ 10000
      End If
   End If
End Sub

Private Sub Text7_GotFocus()
   Dim iPos As Integer
   
   iPos = InStr(Text7, "一(二)")
   If iPos > 0 Then
      Text7.SelStart = iPos + 3
      Text7.SelLength = Len(Text7) - 4
   Else
      TextInverse Text7
   End If
End Sub

Private Sub Text8_GotFocus()
   TextInverse Text8
   CloseIme
End Sub

Private Sub txtCP135_GotFocus()
   TextInverse txtCP135
   CloseIme
End Sub

Private Sub txtCP135_KeyPress(KeyAscii As Integer)
   '只能輸倒退及數字鍵
   If KeyAscii <> 8 And (KeyAscii < 48 Or KeyAscii > 57) Then
      Beep
      KeyAscii = 0
   End If
End Sub

Private Sub txtCP135_Validate(Cancel As Boolean)
   'Add By Sindy 2018/8/7 人員會漏勾
   If Val(txtCP135) > 0 Then
      chkDoc(0).Value = 1
   End If
   '2018/8/7 END
   If m_bolChkFee Then
      Call_PUB_SetOfficialFee_P
   End If
End Sub

Private Sub txtCP136_GotFocus()
   TextInverse txtCP136
   CloseIme
End Sub

Private Sub txtCP136_KeyPress(KeyAscii As Integer)
   '只能輸倒退及數字鍵
   If KeyAscii <> 8 And (KeyAscii < 48 Or KeyAscii > 57) Then
      Beep
      KeyAscii = 0
   End If
End Sub

Private Sub txtCP136_Validate(Cancel As Boolean)
   If m_bolChkFee Then
      Call_PUB_SetOfficialFee_P
   End If
End Sub

Private Sub txtCP84_GotFocus()
   TextInverse txtCP84
   CloseIme
End Sub

Private Sub txtCP84_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtCP84_Validate(Cancel As Boolean)
   '台灣
   If pa(9) = "000" Then
      If Val(txtCP84.Text) <> Val(cp(17)) And Val(txtCP84.Text) <> Val(txtCP84.Tag) Then
         If MsgBox("發文規費【" & txtCP84.Text & "】與收文規費【" & cp(17) & "】不同，確定要繼續！", vbInformation + vbYesNo + vbDefaultButton2) = vbYes Then
            txtCP84.Tag = txtCP84.Text
         Else
            txtCP84_GotFocus
            Cancel = True
         End If
      End If
   End If
End Sub

Private Sub txtDocCh_GotFocus(Index As Integer)
   TextInverse txtDocCh(Index)
   CloseIme
End Sub

Private Sub txtDocCh_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtDocCh_Validate(Index As Integer, Cancel As Boolean)
   Dim iChecked As Single
   
   'Add By Sindy 2019/8/14 + 435.續行母案再審,下列項目不需要預設勾選
   If cp(10) = "435" Then Exit Sub
   
   If Val(txtDocCh(Index)) > 0 Then
      iChecked = vbChecked
   Else
      iChecked = vbUnchecked
   End If
   
   Select Case Index
   '摘要
   Case 0:
      If chkAtt(1).Enabled = True And pa(8) = "1" And cp(10) <> "416" Then
         chkAtt(1).Value = iChecked
      ElseIf chkAtt(5).Enabled = True And pa(8) = "2" Then
         chkAtt(5).Value = iChecked
      End If
      
   '說明書
   Case 1:
      If chkAtt(2).Enabled = True And pa(8) = "1" And cp(10) <> "416" Then
         chkAtt(2).Value = iChecked
      ElseIf chkAtt(6).Enabled = True And pa(8) = "2" Then
         chkAtt(6).Value = iChecked
      ElseIf chkAtt(9).Enabled = True And pa(8) = "3" Then
         chkAtt(9).Value = iChecked
      End If
      
   '申請專利範圍
   Case 2
      If chkAtt(3).Enabled = True And pa(8) = "1" And cp(10) <> "416" Then
         chkAtt(3).Value = iChecked
      ElseIf chkAtt(7).Enabled = True And pa(8) = "2" Then
         chkAtt(7).Value = iChecked
      End If
      
   '圖式
   Case 3:
      If pa(8) = "1" And cp(10) <> "416" Then '發明
         If chkAtt(4).Enabled = True Then '發明圖式
            chkAtt(4).Value = iChecked
         End If
      ElseIf pa(8) = "2" Then '新型
         If chkAtt(8).Enabled = True Then '新型圖式
            chkAtt(8).Value = iChecked
         End If
      ElseIf pa(8) = "3" Then '設計
         If chkAtt(10).Enabled = True Then '設計圖式
            chkAtt(10).Value = iChecked
         End If
      End If
   
   'Add By Sindy 2018/5/22
   '序列表
   Case 4
      If chkAtt(11).Enabled = True And pa(8) = "1" And cp(10) <> "416" Then
         chkAtt(11).Value = iChecked
      End If
   '2018/5/22 END
   End Select
   
   'Memo by Lydia 2018/12/27 序列表(不算超頁費=不算錢),所以不加入進度檔的總頁數
   If Index <= 3 Then
      txtCP135 = Val(txtDocCh(0)) + Val(txtDocCh(1)) + Val(txtDocCh(2)) + Val(txtDocCh(3))
   End If
End Sub

'Modify By Sindy 2023/3/24 規費計算
Private Sub Call_PUB_SetOfficialFee_P()
'Dim bolInChange As Boolean
'
'   'Modify By Sindy 2023/4/18
'   bolInChange = True
'   If PUB_ChkCPExist(cp, "203", 1) Then '有主動修正未發文
'      If m_WriteNote <> "Y" Then
'         bolInChange = False
'      End If
'   End If
'   If bolInChange = False Then
'      Call PUB_SetOfficialFee_P(cp(), pa(), bolDelay, m_strDelayCP09, m_strReExamCP27, _
'                              IIf(txtTotPage.Visible = True, txtTotPage, txtCP135), _
'                              IIf(txtTotItem.Visible = True, txtTotItem, txtCP136), _
'                              , txtCP84)
'   Else
'   '2023/4/18 END
   If m_bolChkFee Then
      Call PUB_SetOfficialFee_P(cp(), pa(), bolDelay, m_strDelayCP09, m_strReExamCP27, _
                              IIf(txtTotPage.Visible = True, txtTotPage, txtCP135), _
                              IIf(txtTotItem.Visible = True, txtTotItem, txtCP136), _
                              , txtCP84, , txtDecreaseItemFee, _
                              m_lngOverPageFee, m_lngOverItemFee, , , , , , txtDecreasePageFee, _
                              , , m_str938RecvNo, m_str939RecvNo)
   End If
End Sub

Private Function TxtValidate() As Boolean
   Dim Cancel As Boolean
   
   m_lngOverPageFee = 0
   m_lngOverItemFee = 0
   m_FeeMemo = ""
   If m_bolChkFee Then
      Call_PUB_SetOfficialFee_P
      'Modify By Sindy 2023/3/24 + txtTotPage和txtTotItem
      If Not PUB_CheckOfficialFee_P(cp(), m_bolChkPageItem, m_bolChkItem, _
                                    IIf(txtTotPage.Visible = True, txtTotPage, txtCP135), _
                                    IIf(txtTotItem.Visible = True, txtTotItem, txtCP136), , , txtCP84, _
                                    m_lngRecOverPageFee, m_lngRecOverItemFee, m_FeeMemo, _
                                    m_lngOverPageFee, m_lngOverItemFee, _
                                    m_lngOverPageFeeDiff, m_lngOverItemFeeDiff) Then
         Exit Function
      End If
   End If
   
   If Text6(0) = "" Then MsgBox "中文案件名稱不可空白!", vbCritical: Exit Function
   
   Cancel = False
   lstNameAgent_Validate Cancel
   If Cancel = True Then
      SSTab1.Tab = 0
      lstNameAgent.SetFocus
      Exit Function
   End If
   
   If chkDoc(0).Value <> 0 Then
      If txtDocCh(0).Enabled = True And Val(txtDocCh(0)) = 0 Then
         SSTab1.Tab = 1
         MsgBox "請輸入摘要頁數!!", vbInformation
         txtDocCh(0).SetFocus
         Exit Function
      End If
      If txtDocCh(1).Enabled = True And Val(txtDocCh(1)) = 0 Then
         SSTab1.Tab = 1
         MsgBox "請輸入說明書頁數!!", vbInformation
         txtDocCh(1).SetFocus
         Exit Function
      End If
      
      If txtDocCh(2).Enabled = True And Val(txtDocCh(2)) = 0 Then
         MsgBox "請輸入申請專利範圍頁數！", vbInformation
         SSTab1.Tab = 1
         txtDocCh(2).SetFocus
         Exit Function
      End If
      
      If txtCP136.Enabled = True And Val(txtCP136) = 0 Then
         SSTab1.Tab = 1
         MsgBox "請輸入專利範圍項數!!", vbInformation
         txtCP136.SetFocus
         Exit Function
      End If
      
      If Val(txtDocCh(6)) > 0 And Val(txtDocCh(3)) = 0 Then
         SSTab1.Tab = 1
         MsgBox "請輸入圖式頁數!!", vbInformation
         txtDocCh(3).SetFocus
         Exit Function
      End If
      If Val(txtDocCh(3)) > 0 And Val(txtDocCh(6)) = 0 Then
         SSTab1.Tab = 1
         MsgBox "請輸入圖式圖數!!", vbInformation
         txtDocCh(6).SetFocus
         Exit Function
      End If
   End If
   
   '繳費金額
   'Modify By Sindy 2018/2/14
'   If Val(txtCP84) = 0 Then
'      SSTab1.Tab = 0
'      MsgBox "請輸入繳費金額!!無需繳費請輸 0 !!", vbInformation
'      Exit Function
'   Else
   If Label12.Visible = True Then
      txtCP84 = Val(txtCP84) + 300
      txtCP84.Tag = txtCP84.Text
   End If
   'Modified by Morgan 2018/3/8
   'If Val(txtCP84) > 0 Then
   If Val(txtCP84) > 0 And txtCP84.Enabled Then
   'end 2018/3/8
   '2018/2/14 END
      Cancel = False
      txtCP84_Validate Cancel
      If Cancel = True Then
         SSTab1.Tab = 0
         If Label12.Visible = True Then txtCP84 = Val(txtCP84) - 300
         txtCP84.SetFocus
         Exit Function
      End If
   End If
   
   TxtValidate = True
End Function
'************************************************
' 儲存專利案件資料
'
'************************************************
Private Function FormSave() As Boolean
Dim ii As Integer, m_CP110 As String
Dim strCon As String
Dim strTmp As String
Dim stCP12 As String, stCP13 As String '最新收文智權人員,業務區
Dim strCP137 As String, strCP167 As String
   
   'Add By Sindy 2019/4/24 ex:FCP-60163補文件不用儲存頁項數
   'Modify By Sindy 2019/7/23 + And cp(10) <> 補文件 ex:FCP-60184
   'Modify By Sindy 2019/8/8 + 911.補收款
   If chkDoc(0).Value = 1 And cp(10) <> 補文件 And cp(10) <> 補收款 Then
   '2019/4/24 END
      '總頁數
      'Modify By Sindy 2023/3/20
      If txtTotPage.Visible = True And Val(txtTotPage.Text) > 0 Then
         cp(135) = Val(txtTotPage.Text)
         'Modify By Sindy 2023/8/31 設計案在產生製作中說申請書時，若有一併主動修正按確定後，
         '   請將原始頁數和修正頁數加總後之頁數，回寫到製作中說進度檔之增加頁數之欄位 ex:FCP-69796
         If m_WriteNote = "Y" And cp(10) = "210" Then
            If Val(pa(64) + pa(65) + pa(66) + pa(67) + pa(68)) = 0 Then '基本檔尚未輸入頁數,才需要計算
               cp(135) = Val(txtCP135) + Val(txtTotPage.Text)
            End If
         End If
      Else
      '2023/3/20 END
         cp(135) = txtCP135.Text
      End If
      strCon = strCon & ",cp135=" & CNULL(cp(135), True)
      '總項數
      'Modify By Sindy 2023/3/20
      If txtTotItem.Visible = True And Val(txtTotItem.Text) > 0 Then
         cp(136) = Val(txtTotItem)
      Else
      '2023/3/20 END
         cp(136) = txtCP136.Text
      End If
      strCon = strCon & ",cp136=" & CNULL(cp(136), True)
   End If
'   'Add By Sindy 2018/1/29
'   'If chkDoc(0).Value = 1 Then
'   If m_bolNotSavePageItem = False Then 'Add By Sindy 2018/4/25
'      'Add By Sindy 2018/4/23
'      If txtCP135.Enabled = True Then
'      '2018/4/23 END
'         '頁數總計
'         cp(135) = txtCP135.Text
'         strCon = strCon & ",cp135=" & CNULL(cp(135), True)
'      End If
'      'Add By Sindy 2018/4/23
'      If txtCP136.Enabled = True Then
'      '2018/4/23 END
'         '申請專利範圍項數
'         cp(136) = txtCP136.Text
'         strCon = strCon & ",cp136=" & CNULL(cp(136), True)
'      End If
'   End If
'   '2018/1/29 END
   '2018/5/21 END
   
'   cp(110) = ""
'   For ii = 0 To lstNameAgent.ListCount - 1
'      If lstNameAgent.Selected(ii) = True Then
'         'modify by sonia 2016/10/7 員工編號已可非數字需做轉換
'         'cp(110) = cp(110) & "," & lstNameAgent.ItemData(ii)
'         cp(110) = cp(110) & "," & PUB_Num2Id(lstNameAgent.ItemData(ii))
'      End If
'   Next
'   If Left(cp(110), 1) = "," Then cp(110) = Mid(cp(110), 2)
   
   'Add By Sindy 2018/4/17 + 進度備註
'   If chkDoc(1).Value = 1 Then
'      '首頁及摘要均附英文資料，減免規費800元整
'      If chkDoc(4).Value = 1 Then
'         strCon = strCon & ",cp64='" & "首頁及摘要均附英文資料，減免規費800元整;" & cp(64) & "'"
'      End If
      '有輸序列表
      'Modify By Sindy 2019/7/22
      If cp(10) <> "435" Then
      '2019/7/22 END
         If Val(txtDocCh(4)) > 0 And InStr(cp(64), "序列表" & Val(txtDocCh(4)) & "頁不納入超頁費之計算") = 0 Then
            strCon = strCon & ",cp64='" & "序列表" & Val(txtDocCh(4)) & "頁不納入超頁費之計算;" & cp(64) & "'"
         End If
      End If
'   End If
   '2018/4/17 END
   
   strCon = strCon & ",cp84=" & Val(txtCP84) '發文規費
   strCon = strCon & ",cp118='A'" '電子送件 Add By Sindy 2018/2/14
   
   cnnConnection.BeginTrans
   
On Error GoTo CheckingErr
   
   'Modify By Sindy 2019/1/16 Mark,工程師先出主動修正,程序才做中說
'   'Add By Sindy 2018/5/21 檢查此文號是否未發文未取消收文,才需要儲存資料
'   strSql = "select cp09" & _
'            " From CASEPROGRESS" & _
'            " WHERE CP09='" & strReceiveNo & "' and cp158=0 and cp159=0"
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'   If intI = 1 Then
'      '先清除此案號總頁/項數,後面SQL會將總頁/項數儲存在此筆文號中
'      strSql = "UPDATE CASEPROGRESS SET cp135=null,cp136=null,cp137=null,cp138=null" & _
'               " WHERE CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "'"
'      cnnConnection.Execute strSql
'   End If
   
   If lstNameAgent.Visible = True Then
      strCon = strCon & ",cp110='" & cp(110) & "'"
   End If
   If strCon <> "" Then
      strCon = Mid(strCon, 2)
      strSql = "UPDATE CASEPROGRESS SET " & strCon & " WHERE CP09='" & strReceiveNo & "' and cp158=0 and cp159=0"
      cnnConnection.Execute strSql
   End If
   
   'Add By Sindy 2018/4/19 回寫進度備註
   If chkDoc(1).Value = 1 Then
      '首頁及摘要均附英文資料，減免規費800元整
      If chkDoc(4).Value = 1 Then
         'Modify By Sindy 2021/7/9 改顯示字樣
'         strSql = "UPDATE CASEPROGRESS SET cp64='首頁及摘要均附英文資料，減免規費800元整;'||cp64" & _
'                  " WHERE CP01='" & pa(1) & "' and CP02='" & pa(2) & "' and CP03='" & pa(3) & "' and CP04='" & pa(4) & "'" & _
'                  " and cp10 in('201','209','235') and cp158=0 and cp159=0"
         strSql = "UPDATE CASEPROGRESS SET cp64='本案提申時未附英文說明書，但所檢附之申請書中發明名稱、申請人姓名或名稱、發明人姓名及摘要同時附有英文翻譯，故可減收申請規費800元。;'||cp64" & _
                  " WHERE CP01='" & pa(1) & "' and CP02='" & pa(2) & "' and CP03='" & pa(3) & "' and CP04='" & pa(4) & "'" & _
                  " and cp10 in('201','209','235') and cp158=0 and cp159=0"
         cnnConnection.Execute strSql
      End If
   End If
   '2018/4/19 END
   
   'Add By Sindy 2018/4/18 有勾選變更申請人之地址時,記錄變更檔
   If chkAtt(26).Value = 1 Then
      strTmp = ""
      If pa(26) <> "" Then
         m_strPA31 = GetPrjNationNumber1(ChangeCustomerL(pa(26)), "CU23")
         If m_strPA31 <> pa(31) Then
            strTmp = strTmp & ",CE23=" & CNULL(ChgSQL(m_strPA31))
         Else
            m_strPA31 = ""
         End If
         m_strPA36 = Trim(GetPrjNationNumber1(ChangeCustomerL(pa(26)), "CU24||' '||CU25||' '||CU26||' '||CU27||' '||CU28||' '||CU102"))
         If m_strPA36 <> Trim(pa(36)) Then
            strTmp = strTmp & ",CE24=" & CNULL(ChgSQL(m_strPA36))
         Else
            m_strPA36 = ""
         End If
      End If
      If pa(27) <> "" Then
         m_strPA32 = GetPrjNationNumber1(ChangeCustomerL(pa(27)), "CU23")
         If m_strPA32 <> pa(32) Then
            strTmp = strTmp & ",CE26=" & CNULL(ChgSQL(m_strPA32))
         Else
            m_strPA32 = ""
         End If
         m_strPA37 = Trim(GetPrjNationNumber1(ChangeCustomerL(pa(27)), "CU24||' '||CU25||' '||CU26||' '||CU27||' '||CU28||' '||CU102"))
         If m_strPA37 <> Trim(pa(37)) Then
            strTmp = strTmp & ",CE27=" & CNULL(ChgSQL(m_strPA37))
         Else
            m_strPA37 = ""
         End If
      End If
      If pa(28) <> "" Then
         m_strPA33 = GetPrjNationNumber1(ChangeCustomerL(pa(28)), "CU23")
         If m_strPA33 <> pa(33) Then
            strTmp = strTmp & ",CE29=" & CNULL(ChgSQL(m_strPA33))
         Else
            m_strPA33 = ""
         End If
         m_strPA38 = Trim(GetPrjNationNumber1(ChangeCustomerL(pa(28)), "CU24||' '||CU25||' '||CU26||' '||CU27||' '||CU28||' '||CU102"))
         If m_strPA38 <> Trim(pa(38)) Then
            strTmp = strTmp & ",CE30=" & CNULL(ChgSQL(m_strPA38))
         Else
            m_strPA38 = ""
         End If
      End If
      If pa(29) <> "" Then
         m_strPA34 = GetPrjNationNumber1(ChangeCustomerL(pa(29)), "CU23")
         If m_strPA34 <> pa(34) Then
            strTmp = strTmp & ",CE32=" & CNULL(ChgSQL(m_strPA34))
         Else
            m_strPA34 = ""
         End If
         m_strPA39 = Trim(GetPrjNationNumber1(ChangeCustomerL(pa(29)), "CU24||' '||CU25||' '||CU26||' '||CU27||' '||CU28||' '||CU102"))
         If m_strPA39 <> Trim(pa(39)) Then
            strTmp = strTmp & ",CE33=" & CNULL(ChgSQL(m_strPA39))
         Else
            m_strPA39 = ""
         End If
      End If
      If pa(30) <> "" Then
         m_strPA35 = GetPrjNationNumber1(ChangeCustomerL(pa(30)), "CU23")
         If m_strPA35 <> pa(35) Then
            strTmp = strTmp & ",CE35=" & CNULL(ChgSQL(m_strPA35))
         Else
            m_strPA35 = ""
         End If
         m_strPA40 = Trim(GetPrjNationNumber1(ChangeCustomerL(pa(30)), "CU24||' '||CU25||' '||CU26||' '||CU27||' '||CU28||' '||CU102"))
         If m_strPA40 <> Trim(pa(40)) Then
            strTmp = strTmp & ",CE36=" & CNULL(ChgSQL(m_strPA40))
         Else
            m_strPA40 = ""
         End If
      End If
      If strTmp <> "" Then
         strTmp = Mid(strTmp, 2)
         strExc(1) = "DELETE FROM CHANGEEVENT WHERE CE01='" & strReceiveNo & "'"
         strExc(2) = "INSERT INTO CHANGEEVENT (CE01) VALUES ('" & strReceiveNo & "')"
         strExc(3) = "UPDATE CHANGEEVENT SET " & strTmp & " WHERE CE01='" & strReceiveNo & "'"
         cnnConnection.Execute strExc(1)
         cnnConnection.Execute strExc(2)
         cnnConnection.Execute strExc(3)
      'Modify By Sindy 2018/9/20 Mark
'      Else
'         chkAtt(26).Value = 0
      End If
   End If
   '2018/4/18 END
   
   'Add By Sindy 2023/3/20 有一併主動修正者,此處不回寫基本檔,發文才回寫
   'Modify By Sindy 2023/8/18 +210製作中說:設計案在產生製作中說申請書時，若有一併主動修正，
   '   請開放輸入頁數、圖數的欄位以利程序人員產生申請書，且產生申請書時，先將頁數、圖數回寫基本檔
   'Modify By Sindy 2025/9/23 再開放 235核對中說格式,209檢視中說
   'If m_WriteNote = "Y" And cp(10) <> "210" Then
   If m_WriteNote = "Y" And cp(10) <> "210" And _
      Not (pa(8) = "3" And (cp(10) = "235" Or cp(10) = "210")) Then
      If str203CP09 <> "" Then
         '更新中說文號
         strSql = "UPDATE PageDetail SET pd20='" & strReceiveNo & "' WHERE pd01='" & str203CP09 & "'"
         cnnConnection.Execute strSql
      End If
   Else
   '2023/3/20 END
      'Added by Lydia 2018/12/27 存中文本資訊
      strSql = ""
      For Each oText In txtDocCh
         'Modified by Lydia 2019/01/10
         'If oText.Index <= 4 And oText.Tag <> oText.Text Then
         If oText.Index <> 5 And oText.Tag <> oText.Text Then
             Select Case oText.Index
                  Case 0 '摘要頁數
                       strSql = strSql & ", PA64=" & CNULL(oText.Text, True)
                  Case 1 '說明書頁數
                       strSql = strSql & ", PA65=" & CNULL(oText.Text, True)
                  Case 4 '序列表頁數
                       strSql = strSql & ", PA66=" & CNULL(oText.Text, True)
                  Case 2 '申請專利範圍頁數
                       strSql = strSql & ", PA67=" & CNULL(oText.Text, True)
                  Case 3 '圖式頁數
                       strSql = strSql & ", PA68=" & CNULL(oText.Text, True)
                  'Added by Lydia 2019/01/10
                  Case 6 '圖式圖數
                       strSql = strSql & ", PA173=" & CNULL(oText.Text, True)
             End Select
         End If
      Next
      'Add By Sindy 2023/3/14
      If txtTotItem.Visible = False And Val(txtCP136) > 0 Then
         '申請專利範圍項數(最初項數)
         strSql = strSql & ", PA172=" & CNULL(txtCP136.Text, True)
      End If
      '2023/3/14 END
      If strSql <> "" Then
           strSql = "UPDATE PATENT SET " & Mid(strSql, 2) & " WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
           strSql = "begin user_data.user_enabled:=0; " & strSql & "; end;"
           'Modified by Lydia 2021/04/27 更新來源的表單名稱 ;
           'Pub_SeekTbLog strSql '新增log
           Pub_SeekTbLog strSql, , , , Me.Caption & "(" & Me.Name & ")"
           cnnConnection.Execute strSql
      End If
      'end 2018/12/27
   End If
   
   cnnConnection.CommitTrans
   FormSave = True
   Exit Function
   
CheckingErr:
   cnnConnection.RollbackTrans
   
End Function

'Private Sub StartLetter(ByVal ET01 As String, ByVal ET03 As String)
'   Dim strTxt(110) As String, strTmp As String
'   Dim ii As Integer, jj As Integer
'   Dim strInventor As String 'Add By Sindy 2014/11/14
'
'   ii = 0
'   EndLetter ET01, strReceiveNo, ET03, strUserNum
'
'   '申請人
'   For jj = 1 To 5
'      If pa(25 + jj) <> "" Then
'         '申請人
'         strExc(0) = " SELECT C.*,N1.NA72 X1,N2.NA72 X2" & _
'            " FROM CUSTOMER C,NATION N1,NATION N2 WHERE CU01='" & Left(ChangeCustomerL(pa(25 + jj)), 8) & "'" & _
'            " and cu02='" & Mid(ChangeCustomerL(pa(25 + jj)), 9) & "' AND N1.NA01(+)=CU10 AND N2.NA01(+)=CU87"
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'         If intI = 1 Then
'            ii = ii + 1
'            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','申請人" & jj & "-國籍','" & RsTemp("X1") & "')"
'
'            ii = ii + 1
'            If RsTemp("CU15") = "0" Then
'               strTmp = "自然人"
'            Else
'               strTmp = "法人公司機關學校"
'            End If
'            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','申請人" & jj & "-身分種類','" & strTmp & "')"
'
'            ii = ii + 1
'            If RsTemp("CU10") < "011" Then
'               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                  " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','申請人" & jj & "-ID','" & RsTemp("CU11") & "')"
'            End If
'
'            ii = ii + 1
'            If RsTemp("CU15") = "0" Then
'               strTmp = "申請人" & jj & "-中文姓名"
'            Else
'               strTmp = "申請人" & jj & "-中文名稱"
'            End If
'            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strTmp & "','" & ChgSQL("" & RsTemp("CU04")) & "')"
'
'            ii = ii + 1
'            If RsTemp("CU15") = "0" Then
'               strTmp = "申請人" & jj & "-英文姓名"
'            Else
'               strTmp = "申請人" & jj & "-英文名稱"
'            End If
'            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strTmp & "','" & ChgSQL(RTrim(Trim("" & RsTemp("CU05")) & " " & Trim("" & RsTemp("CU88")) & " " & Trim("" & RsTemp("CU89")) & " " & Trim("" & RsTemp("CU90")))) & "')"
'
'            '目前抓客戶基本檔資料,等基本檔加欄位後需改抓
'            ii = ii + 1
'            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','申請人" & jj & "-居住國','" & RsTemp("X2") & "')"
'
'            ii = ii + 1
'            If RsTemp("CU10") < "011" Then
'               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                  " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','申請人" & jj & "-郵遞區號','" & RsTemp("CU112") & "')"
'            End If
'            ii = ii + 1
'            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','申請人" & jj & "-中文地址','" & ChgSQL("" & RsTemp("CU23")) & "')"
'            ii = ii + 1
'            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','申請人" & jj & "-英文地址','" & ChgSQL(RTrim(Trim("" & RsTemp("CU24")) & " " & Trim("" & RsTemp("CU25")) & " " & Trim("" & RsTemp("CU26")) & " " & Trim("" & RsTemp("CU27")) & " " & Trim("" & RsTemp("CU28")))) & "')"
'
'            If RsTemp("CU15") <> "0" Then
'               ii = ii + 1
'               If jj < 3 Then
'                  strTmp = pa(79 + 3 * (jj - 1))
'               Else
'                  strTmp = pa(109 + 3 * (jj - 1))
'               End If
'
'               If strTmp = "" Then
'                  strTmp = "後補"
'               End If
'
'               '代表人中文姓名-->非自然人時為必要欄位
'               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                  " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','申請人" & jj & "-代表人中文姓名','" & ChgSQL(strTmp) & "')"
'
'               '代表人英文姓名-->非必要欄位
'               If jj < 3 Then
'                  strTmp = pa(80 + 3 * (jj - 1))
'               Else
'                  strTmp = pa(110 + 3 * (jj - 1))
'               End If
'               If strTmp <> "" Then
'                  ii = ii + 1
'                  strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                     " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','申請人" & jj & "-代表人英文姓名','" & ChgSQL(strTmp) & "')"
'               End If
'            End If
'         End If
'      End If
'   Next
'
'   '出名代理人
'   strExc(0) = "select oa08,ST26,st02 from ouragent,staff where oa01='" & pa(1) & "' and instr('" & cp(110) & "',oa02)>0 and st01(+)=oa02 order by OA03"
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      With RsTemp
'      jj = 1
'      Do While Not .EOF
'         ii = ii + 1
'         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','代理人" & jj & "-證書字號','" & .Fields("oa08") & "')"
'         ii = ii + 1
'         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','代理人" & jj & "-ID','" & .Fields("ST26") & "')"
'         ii = ii + 1
'         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','代理人" & jj & "-中文姓名','" & PUB_ConvertNameFormat("" & .Fields("st02")) & "')"
'         jj = jj + 1
'         .MoveNext
'      Loop
'      End With
'   End If
'
'   If pa(8) = "1" Then
'      strExc(1) = "發明人"
'   ElseIf pa(8) = "2" Then
'      strExc(1) = "新型創作人"
'   Else
'      strExc(1) = "設計人"
'   End If
'   'Modify By Sindy 2014/11/14
'   If strSrvDate(1) >= 專利發明人檔啟用日 Then
'      strInventor = ""
'      strExc(0) = " SELECT IN03,IN04,IN05,IN11,NA72" & _
'                  " FROM PatentInventor,INVENTOR,NATION" & _
'                  " WHERE pi01=" + CNULL(pa(1)) + " and pi02=" + CNULL(pa(2)) + " and pi03=" + CNULL(pa(3)) + " and pi04=" + CNULL(pa(4)) & _
'                  " AND IN01=substr(pi06,1,8) AND IN02=substr(pi06,9,2)" & _
'                  " AND NA01(+)=IN11" & _
'                  " order by pi05 asc"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         RsTemp.MoveFirst
'         Do While Not RsTemp.EOF
'            If strInventor <> "" Then strInventor = strInventor & vbCrLf
'            strInventor = strInventor & "　　【" & strExc(1) & "】" & vbCrLf & _
'                                        "　　　【國籍】　　　　　　" & RsTemp("NA72") & vbCrLf & _
'                                        "　　　【中文姓名】　　　　" & ChgSQL("" & RsTemp("IN04")) & vbCrLf & _
'                                        IIf("" & RsTemp("IN05") = "", "", "　　　【英文姓名】　　　　" & ChgSQL("" & RsTemp("IN05")) & vbCrLf)
'            RsTemp.MoveNext
'         Loop
'      Else
'         strInventor = "　　【" & strExc(1) & "】" & vbCrLf & _
'                       "　　　【國籍】" & vbCrLf & _
'                       "　　　【中文姓名】" & vbCrLf
'      End If
'      ii = ii + 1
'      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','發明人資料','" & strInventor & "')"
'   Else
'   '2014/11/14 END
'      ii = ii + 1
'      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                  " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','發明人','" & strExc(1) & "')"
'      For jj = 1 To 10
'         If pa(59 + jj) <> "" Then
'            strExc(0) = " SELECT IN03,IN04,IN05,IN11,NA72" & _
'               " FROM INVENTOR,NATION WHERE IN01='" & Left(pa(59 + jj), 8) & "' AND IN02='" & Mid(pa(59 + jj), 9) & "'" & _
'               " AND NA01(+)=IN11"
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'            If intI = 1 Then
'               ii = ii + 1
'               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                  " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','發明人" & jj & "-國籍','" & RsTemp("NA72") & "')"
'
'               ii = ii + 1
'               If RsTemp("IN11") < "011" Then
'                  strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                     " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','發明人" & jj & "-ID','" & RsTemp("IN03") & "')"
'               End If
'
'               ii = ii + 1
'               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                  " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','發明人" & jj & "-中文姓名','" & ChgSQL("" & RsTemp("IN04")) & "')"
'
'               ii = ii + 1
'               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                  " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','發明人" & jj & "-英文姓名','" & ChgSQL("" & RsTemp("IN05")) & "')"
'            End If
'         End If
'      Next
'   End If
'
'   If Not ClsLawExecSQL(ii, strTxt) Then
'      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
'   End If
'End Sub

'申請書
Private Function StartLetter2(ByVal ET01 As String, ByVal ET03 As String) As Boolean
   Dim strTxt(200) As String, strTmp As String
   Dim ii As Integer, jj As Integer
   Dim strCP64 As String 'Modify By Sindy 2025/7/21
   
   EndLetter ET01, strReceiveNo, ET03, strUserNum
   
   ii = 1
   'Removed by Morgan 2018/1/10 取消不印,因可能同時補多種文件,改讓電子送件系統自行判斷--敏莉
   'strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
   '   " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','案由','" & Left(cboResonCode, 5) & "')"
   '
   'ii = ii + 1
   'end 2018/1/10
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','本所案號','" & m_CaseNo & "')"
   
   '辦理依據
   If Text5 <> "" Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','發文日期','" & ChangeTStringToTDateString(Text5) & "')"
      
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','智專字','" & Text7 & "')"
      
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','發文號','" & Text8 & "')"
   End If

   'Modify By Sindy 2017/11/15
   Call PUB_GetApplPA_EData(ET01, ET03, strReceiveNo, pa(), IIf(chkAtt(26).Value = 1, False, True))
'   For jj = 1 To 5
'      If pa(25 + jj) <> "" Then
'         '申請人
'         strExc(0) = " SELECT C.*,N1.NA72 X1,N2.NA72 X2" & _
'            " FROM CUSTOMER C,NATION N1,NATION N2 WHERE CU01='" & Left(ChangeCustomerL(pa(25 + jj)), 8) & "'" & _
'            " and cu02='" & Mid(ChangeCustomerL(pa(25 + jj)), 9) & "' AND N1.NA01(+)=CU10 AND N2.NA01(+)=CU87"
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'         If intI = 1 Then
'            ii = ii + 1
'            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','申請人" & jj & "-國籍','" & RsTemp("X1") & "')"
'
'            ii = ii + 1
'            If RsTemp("CU10") < "011" Then
'               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                  " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','申請人" & jj & "-ID','" & RsTemp("CU11") & "')"
'            End If
'
'            ii = ii + 1
'            If RsTemp("CU15") = "0" Then
'               strTmp = "申請人" & jj & "-中文姓名"
'            Else
'               strTmp = "申請人" & jj & "-中文名稱"
'            End If
'            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strTmp & "','" & ChgSQL("" & RsTemp("CU04")) & "')"
'
'            ii = ii + 1
'            If RsTemp("CU15") = "0" Then
'               strTmp = "申請人" & jj & "-英文姓名"
'            Else
'               strTmp = "申請人" & jj & "-英文名稱"
'            End If
'            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strTmp & "','" & ChgSQL(RTrim(Trim("" & RsTemp("CU05")) & " " & Trim("" & RsTemp("CU88")) & " " & Trim("" & RsTemp("CU89")) & " " & Trim("" & RsTemp("CU90")))) & "')"
'         End If
'      End If
'   Next
   
   '出名代理人
   strExc(0) = "select oa05,ST26,st02 from ouragent,staff where oa01='" & pa(1) & "' and instr('" & cp(110) & "',oa02)>0 and st01(+)=oa02 order by OA03"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      jj = 1
      Do While Not .EOF
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','代理人" & jj & "-中文姓名','" & PUB_ConvertNameFormat("" & .Fields("st02")) & "')"
         jj = jj + 1
         .MoveNext
      Loop
      End With
   End If
   
   '補正首次中文本
   If chkDoc(0).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','補正首次中文本','♀')"
'   End If
'   If chkDoc(0).Value <> 0 Then
      If txtDocCh(0).Enabled = True Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','摘要頁數','" & Val(txtDocCh(0)) & "')"
      End If
      '說明書頁數=說明書頁數+序列表
      'Modify By Sindy 2019/2/21 取消+序列表
      If txtDocCh(1).Enabled = True Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','說明書頁數','" & Val(txtDocCh(1)) & "')"
      End If
      If txtDocCh(2).Enabled = True Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','專利範圍頁數','" & Val(txtDocCh(2)) & "')"
      End If
      If txtDocCh(3).Enabled = True Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','圖式頁數','" & Val(txtDocCh(3)) & "')"
      End If
      'Modify By Sindy 2018/6/28 頁數總計要加入序列表
      'Modify By Sindy 2019/2/21 取消+序列表
      If txtCP135.Enabled = True Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','頁數總計','" & Val(txtCP135) & "')"
      End If
      If txtCP136.Enabled = True Then
         ii = ii + 1
         'Modify By Sindy 2019/12/26 +  And Val(txtPA172) > 0
         'Modify By Sindy 2023/3/21
'         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','專利範圍項數','" & IIf(txtTotItem.Visible = True And Val(txtTotItem) > 0, Val(txtTotItem), Val(txtCP136)) & "')"
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','專利範圍項數','" & Val(txtCP136) & "')"
      End If
      'If Val(txtDocCh(6)) > 0 Then
      If txtDocCh(6).Enabled = True Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','圖式圖數','" & Val(txtDocCh(6)) & "')"
      End If
   End If
   
   'Modify By Sindy 2018/5/8 + m_WriteNote = "Y" 一併主動修正
   'Modify By Sindy 2018/6/28 + Val(txtDocCh(4)) > 0
   'If chkDoc(1).Value = 1 Or chkAtt(26).Value = 1 Or m_WriteNote = "Y" Or Val(txtDocCh(4)) > 0 Then
      strTmp = ""
      'Modify By Sindy 2018/5/8
      'Modify By Sindy 2019/1/23 + 控制 實審已發文 才會有超頁超項的產生 ex:FCP-059756
      'Modify By Sindy 2019/8/27 + 201.新案翻譯(含209.檢視中說、235.核對中說格式)的電子送件"補正申請書"時實審已發文，且有超頁、超項費則在申請書帶備註
      If (m_WriteNote = "Y" Or cp(10) = "201" Or cp(10) = "209" Or cp(10) = "235") And _
         PUB_ChkCPExist(cp, "416", 2) And _
         (m_lngOverPageFee > 0 Or m_lngOverItemFee > 0) Then
         If strTmp <> "" Then strTmp = strTmp & vbCrLf & "　　　　　　　　　　　　　　　"
         'Modify By Sindy 2019/1/16 改在程序出補正時,顯示繳費金額,修正備註資訊
         'strTmp = strTmp & "需繳交的" & IIf(m_lngOverPageFee > 0, "超頁費" & IIf(m_lngOverItemFee > 0, "、超項費", ""), "超項費") & "會在同一日補呈之修正時繳納。"
         If m_WriteNote = "Y" Then '一併主動修正
            'Modify By Sindy 2021/12/24 修正實審已發文，中說+修正申請書，補述內容  FCP-065711 (暫緩)
            strTmp = strTmp & "繳費金額為" & IIf(Val(txtTotItem) <> Val(cp(136)), "修正後之", "") & IIf(m_lngOverPageFee > 0, "超頁規費" & IIf(m_lngOverItemFee > 0, "、超項規費", ""), "超項規費") & "。"
            'strTmp = strTmp & "本案同日另函提出修正申請，繳費金額為" & IIf(Val(txtPA172) <> Val(cp(136)), "修正後", "") & IIf(m_lngOverPageFee > 0, "超頁規費" & IIf(m_lngOverItemFee > 0, "、超項規費", ""), "超項規費") & "。"
         'Add By Sindy 2019/8/27
         Else
            strTmp = strTmp & "繳費金額為" & IIf(m_lngOverPageFee > 0, "超頁規費" & IIf(m_lngOverItemFee > 0, "、超項規費", ""), "超項規費") & "。"
         End If
         '2019/8/27 END
      End If
      '2018/5/8 END
      'Add By Sindy 2018/1/11
      '一併修正專利名稱
      If chkDoc(3).Value = 1 Then
'         ii = ii + 1
'         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','一併修正專利名稱','♀')"
         If strTmp <> "" Then strTmp = strTmp & vbCrLf & "　　　　　　　　　　　　　　　"
         strTmp = strTmp & "一併修正中文專利名稱" 'Modify By Sindy 2018/12/26
      End If
      '2018/1/11 END
      'Add By Sindy 2018/4/17 首頁及摘要均附英文資料，減免規費800元整
      If chkDoc(4).Value = 1 Then
         If strTmp <> "" Then strTmp = strTmp & vbCrLf & "　　　　　　　　　　　　　　　"
         'Modified by Lydia 2020/03/31 改模組A0802Query => CompNameQuery
         'Modify By Sindy 2021/7/9 改顯示字樣
         'strTmp = strTmp & "首頁及摘要均附英文資料，減免規費捌佰元整(退費支票抬頭請開" & CompNameQuery("2") & ")"
         'Modify By Sindy 2025/7/21 加抓收據號碼114DP054287(請抓新案那道規費收據編號)
         Call GetCP31isY_CP05(cp(1), cp(2), cp(3), cp(4), "cp64", strCP64)
         If InStr(strCP64, "收據號碼:") > 0 Then
            strExc(10) = Mid(strCP64, InStr(strCP64, "收據號碼:"), 16)
         Else
            strExc(10) = "收據號碼:"
         End If
         strTmp = strTmp & "本案提申時未附英文說明書，但所檢附之申請書中發明名稱、申請人姓名或名稱、發明人姓名及摘要同時附有英文翻譯，故可減收申請規費800元。(" & strExc(10) & "，退費支票抬頭請開" & CompNameQuery("2") & ")"
      End If
      '2018/4/17 END
      
      'Add By Sindy 2021/8/31
      If chkDoc(6).Value = 1 Then
         If strTmp <> "" Then strTmp = strTmp & vbCrLf & "　　　　　　　　　　　　　　　"
         strTmp = strTmp & "本案已於提出發明申請時減免規費800元整，今備函補呈英文摘要及英文專利名稱，請准予併案辦理。"
         '要加.2
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','電子稽核防止漏發文','♀')"
      End If
      'Add By Sindy 2023/4/7
      If Val(txtDecreasePageFee.Text) > 0 Or Val(txtDecreaseItemFee.Text) > 0 Then
         chkAtt(21).Value = 1
         chkAtt(22).Value = 1
'         '要加.2
'         ii = ii + 1
'         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','電子稽核防止漏發文','♀')"
         '備註:
         If m_WriteNote = "Y" Then '一併送主動修正時的備註內容
            strTmp = strTmp & "請求退還"
            If m_str938RecvNo <> "" And Val(txtDecreasePageFee.Text) > 0 Then
               strTmp = strTmp & "超頁規費「" & Format(Val(txtDecreasePageFee.Text), "###,###,##0") & "」元"
            End If
            If m_str939RecvNo <> "" And Val(txtDecreaseItemFee.Text) > 0 Then
               If Val(txtDecreasePageFee.Text) > 0 Then strTmp = strTmp & "及"
               strTmp = strTmp & "超項規費「" & Format(Val(txtDecreaseItemFee.Text), "###,###,##0") & "」元，"
            Else
               strTmp = strTmp & "，"
            End If
         Else
            strTmp = strTmp & "本案"
            If m_str938RecvNo <> "" And Val(txtDecreasePageFee.Text) > 0 Then
               strExc(0) = "select cp27 from caseprogress where cp09='" & m_str938RecvNo & "' and cp27 is not null"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  strExc(9) = TranslateKeyWord(incCNV_CHINESE_MINKO1, RsTemp.Fields("cp27"), strExc(10))
                  strTmp = strTmp & "於" & strExc(9) & "繳納超頁費，今日補呈首次中文本確認並無超頁，故請求退還超頁規費「" & Format(Val(txtDecreasePageFee.Text), "###,###,##0") & "」元，"
               End If
            End If
            If m_str939RecvNo <> "" And Val(txtDecreaseItemFee.Text) > 0 Then
               strExc(0) = "select cp27 from caseprogress where cp09='" & m_str939RecvNo & "' and cp27 is not null"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  strExc(9) = TranslateKeyWord(incCNV_CHINESE_MINKO1, RsTemp.Fields("cp27"), strExc(10))
                  strTmp = strTmp & "於" & strExc(9) & "繳納超項費，今日補呈首次中文本確認並無超項，故請求退還超項規費「" & Format(Val(txtDecreaseItemFee.Text), "###,###,##0") & "」元，"
               End If
            End If
         End If
         strTmp = strTmp & "並檢還之國庫支票抬頭請開立：「台一國際智慧財產事務所」。"
      End If
      '2023/4/7 END
      
      'Add By Sindy 2019/7/22
      If cp(10) = "435" Then '續行母案再審
         If strTmp <> "" Then strTmp = strTmp & vbCrLf & "　　　　　　　　　　　　　　　"
         strTmp = strTmp & "繳納續行母案再審之規費。"
      Else
      '2019/7/22 END
         'Add By Sindy 2018/4/18 序列表
         If Val(txtDocCh(4)) > 0 Then
            If strTmp <> "" Then strTmp = strTmp & vbCrLf & "　　　　　　　　　　　　　　　"
            strTmp = strTmp & "序列表" & Val(txtDocCh(4)) & "頁不納入超頁費之計算"
         End If
         '2018/4/18 END
      End If
      
      'Add By Sindy 2018/4/18 有勾選變更申請人之地址時,記錄變更檔
      If chkAtt(26).Value = 1 Then
         'Modify By Sindy 2018/11/23 調整顯示資料方式
         If m_strPA31 <> "" Or m_strPA36 <> "" Then
            If strTmp <> "" Then strTmp = strTmp & vbCrLf & "　　　　　　　　　　　　　　　"
            strTmp = strTmp & "變更申請人1地址為" & m_strPA31 & IIf(m_strPA36 <> "", "(" & m_strPA36 & ")", "")
         End If
         If m_strPA32 <> "" Or m_strPA37 <> "" Then
            If strTmp <> "" Then strTmp = strTmp & vbCrLf & "　　　　　　　　　　　　　　　"
            strTmp = strTmp & "變更申請人2地址為" & m_strPA32 & IIf(m_strPA37 <> "", "(" & m_strPA37 & ")", "")
         End If
         If m_strPA33 <> "" Or m_strPA38 <> "" Then
            If strTmp <> "" Then strTmp = strTmp & vbCrLf & "　　　　　　　　　　　　　　　"
            strTmp = strTmp & "變更申請人3地址為" & m_strPA33 & IIf(m_strPA38 <> "", "(" & m_strPA38 & ")", "")
         End If
         If m_strPA34 <> "" Or m_strPA39 <> "" Then
            If strTmp <> "" Then strTmp = strTmp & vbCrLf & "　　　　　　　　　　　　　　　"
            strTmp = strTmp & "變更申請人4地址為" & m_strPA34 & IIf(m_strPA39 <> "", "(" & m_strPA39 & ")", "")
         End If
         If m_strPA35 <> "" Or m_strPA40 <> "" Then
            If strTmp <> "" Then strTmp = strTmp & vbCrLf & "　　　　　　　　　　　　　　　"
            strTmp = strTmp & "變更申請人5地址為" & m_strPA35 & IIf(m_strPA40 <> "", "(" & m_strPA40 & ")", "")
         End If
      End If
      '2018/4/18 END
      'Add By Sindy 2018/6/5 可以單勾選備註
      If strTmp = "" Then 'And chkDoc(1).Value = 1
         'Modify By Sindy 2018/6/28
         'strTmp = "♀"
         strTmp = "否"
         '2018/6/28 END
      End If
      '2018/6/5 END
      If strTmp <> "" Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','備註','" & strTmp & "')"
      End If
   'End If
   
   If chkAtt(23).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','備註-補正其他文件','♀')"
   End If
   If chkAtt(24).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','備註-其他申復事項','♀')"
   End If
   'Add By Sindy 2018/7/25
   If chkDoc(5).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','其他申復事項-聲明','♀')"
   End If
   '2018/7/25 END
   If chkAtt(26).Value = 1 Or chkAtt(27).Value = 1 Or chkAtt(28).Value = 1 Or chkAtt(29).Value = 1 Or chkAtt(30).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','同時辦理事項','♀')"
   End If
   If chkAtt(26).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','變更申請人之地址','是')"
   End If
   If chkAtt(27).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','變更申請人之代理人','是')"
   End If
   If chkAtt(28).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','變更申請人之代表人','是')"
   End If
   If chkAtt(29).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','變更申請人之姓名或名稱','是')"
   End If
   If chkAtt(30).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','變更申請人之國籍','是')"
   End If
   
   ii = ii + 1
   'Modify By Sindy 2018/6/20 FCP-58925(209) 繳費金額固定為0
   'Modify By Sindy 2019/1/16 取消,改在程序出補正時,顯示繳費金額
'   If m_WriteNote = "Y" Then
'      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','繳費金額','0')"
'   Else
'   '2018/6/20 END
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','繳費金額','" & Val(txtCP84) & "')"
'   End If
   
   'Modified by Morgan 2018/1/16 附件檔名統一在案號後加".",但下載的檔案加"_"
   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-基本資料表','" & m_CaseNo & chkAtt(0).Tag & "')"
   
   '發明
   If chkAtt(1).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-發明摘要','" & m_CaseNo & chkAtt(1).Tag & "')"
   End If
   If chkAtt(2).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-發明說明書','" & m_CaseNo & chkAtt(2).Tag & "')"
   End If
   If chkAtt(11).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-序列表','" & m_CaseNo & chkAtt(11).Tag & "')"
   End If
   If chkAtt(3).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-發明申請專利範圍','" & m_CaseNo & chkAtt(3).Tag & "')"
   End If
   If chkAtt(4).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-發明圖式','" & m_CaseNo & chkAtt(4).Tag & "')"
   End If
   '新型
   If chkAtt(5).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-新型摘要','" & m_CaseNo & chkAtt(5).Tag & "')"
   End If
   If chkAtt(6).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-新型說明書','" & m_CaseNo & chkAtt(6).Tag & "')"
   End If
   If chkAtt(7).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-新型申請專利範圍','" & m_CaseNo & chkAtt(7).Tag & "')"
   End If
   If chkAtt(8).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-新型圖式','" & m_CaseNo & chkAtt(8).Tag & "')"
   End If
   '設計
   If chkAtt(9).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-設計說明書','" & m_CaseNo & chkAtt(9).Tag & "')"
   End If
   If chkAtt(10).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-設計圖式','" & m_CaseNo & chkAtt(10).Tag & "')"
   End If
   'Other
   If chkAtt(12).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-申請書','" & m_CaseNo & chkAtt(12).Tag & "')"
   End If
   If chkAtt(19).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-委任書','" & m_CaseNo & chkAtt(19).Tag & "')"
   End If
   If chkAtt(20).Value = 1 Then
      ii = ii + 1
      'Modified by Morgan 2017/11/23 chkAtt(20).Tag:Priority_US.pdf->Priority.pdf--敏莉
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-國際優先權證明文件','" & m_CaseNo & chkAtt(20).Tag & "')"
   End If
   If chkAtt(15).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-優惠期證明文件','" & m_CaseNo & chkAtt(15).Tag & "')"
   End If
   If chkAtt(16).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-國內生物材料寄存證明文件','" & m_CaseNo & chkAtt(16).Tag & "')"
   End If
   If chkAtt(17).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-國外生物材料寄存證明文件','" & m_CaseNo & chkAtt(17).Tag & "')"
   End If
   If chkAtt(18).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-生物材料獲得證明文件','" & m_CaseNo & chkAtt(18).Tag & "')"
   End If
   If chkAtt(13).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-申復書','" & m_CaseNo & chkAtt(13).Tag & "')"
   End If
   If chkAtt(14).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-再審查理由書','" & m_CaseNo & chkAtt(14).Tag & "')"
   End If
   
   If chkAtt(21).Value = 1 Or chkAtt(22).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','其他','♀')"
   End If
   If chkAtt(21).Value = 1 Then
      'Add By Sindy 2018/4/17 首頁及摘要均附英文資料，減免規費800元整
      'Modify By Sindy 2023/4/10 + 有退費時也要顯示: + Or (Val(txtDecreasePageFee.Text) > 0 Or Val(txtDecreaseItemFee.Text) > 0)
      If chkDoc(4).Value = 1 Or (Val(txtDecreasePageFee.Text) > 0 Or Val(txtDecreaseItemFee.Text) > 0) Then
         'Modify By Sindy 2025/7/21 +電子
         strTmp = "辦理退費之電子收據"
      'Add By Sindy 2021/8/31
      ElseIf chkDoc(6).Value = 1 Then
         strTmp = "英文說明書（參考用）"
      Else
         strTmp = "♀"
      End If
      '2018/4/17 END
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','其他-文件描述','" & strTmp & "')"
   End If
   If chkAtt(22).Value = 1 Then
      'Add By Sindy 2018/4/17 首頁及摘要均附英文資料，減免規費800元整
      'Modify By Sindy 2023/4/10 + 有退費時也要顯示: + Or (Val(txtDecreasePageFee.Text) > 0 Or Val(txtDecreaseItemFee.Text) > 0)
      If chkDoc(4).Value = 1 Or (Val(txtDecreasePageFee.Text) > 0 Or Val(txtDecreaseItemFee.Text) > 0) Then
         'Modify By Sindy 2021/6/15 葉敏莉:有關退費申請書電子收據的檔名
         '【附送書件】
         '【文件描述】 辦理退費之電子收據
         '【文件檔名】 FCP064078.ATT.RECEIPT.pdf
         'strTmp = m_CaseNo & ".RECEIPT.PDF"
         strTmp = m_CaseNo & ".ATT.PDF"
      'Add By Sindy 2021/8/31
      ElseIf chkDoc(6).Value = 1 Then
         strTmp = m_CaseNo & ".SEP.PDF"
      Else
         strTmp = "♀"
      End If
      '2018/4/17 END
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','其他-文件檔名','" & strTmp & "')"
   End If
   
   'Add By Sindy 2018/1/29
   If cp(10) = 實體審查 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','一併申請修正','否')"
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','一併申請誤譯訂正','否')"
   End If
   '2018/1/29 END
   
   If Not ClsLawExecSQL(ii, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   Else
      StartLetter2 = True
   End If
End Function

'Private Function ConvertHtml2Pdf(pCode As String, pFileName As String) As Boolean
'Dim program_name As String, program_path As String
'Dim stHtmlName As String
'Dim stErrFile As String, stErrText As String
'Dim process_id As Long
'Dim process_handle As Long
'
'    stHtmlName = Mid(pFileName, InStrRev(pFileName, "\") + 1)
'    program_path = "C:\TIPOHtml2Pdf\"
'    program_name = program_path & "html2pdfcmd.exe"
'    ' Start the program.
'    On Error GoTo ShellError
'
'    If Dir(program_name) = "" Then
'      MsgBox "指定路徑找不到 Pdf 轉檔程式!!", vbCritical, " ( " & stHtmlName & " ) pdf 轉檔失敗!!"
'      Exit Function
'    End If
'
'    stErrFile = program_path & "errMessage.txt"
'    If Dir(stErrFile) <> "" Then
'      Kill stErrFile
'   End If
'
'    process_id = SHELL(program_name & " " & pCode & " """ & pFileName & """", vbNormalNoFocus)
'
'    On Error GoTo 0
'
'    ' Wait for the program to finish.
'    ' Get the process handle.
'    process_handle = OpenProcess(SYNCHRONIZE, 0, process_id)
'    If process_handle <> 0 Then
'        WaitForSingleObject process_handle, INFINITE
'        CloseHandle process_handle
'    End If
'
'    If Dir(stErrFile) <> "" Then
'      frm06010301_12.TextBox1.Text = GetText(stErrFile)
'      frm06010301_12.Caption = "( " & stHtmlName & " ) " & frm06010301_12.Caption
'      frm06010301_12.Show vbModal
'    Else
'      ConvertHtml2Pdf = True
'    End If
'    Exit Function
'
'ShellError:
'    MsgBox " " & _
'        program_name & vbCrLf & _
'        Err.Description, vbOKOnly Or vbExclamation, _
'        "Error"
'End Function

Private Function GetText(pFileName As String) As String
   Dim objStream As Object
   Set objStream = CreateObject("ADODB.Stream")
   Dim stContent As String
   
On Error GoTo ErrHnd
   
   With objStream
      .Type = 2
      .Mode = 3
      .Open
      .Charset = "UTF-8" ' 或其他編碼
      '.Charset = "UTF-16"
      .LoadFromFile pFileName
      stContent = .ReadText
      ' PS : 也可透過 .SaveToFile 方法把檔案存檔
      .Close
   End With
   GetText = stContent
   
ErrHnd:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
   Set objStream = Nothing
End Function

'Private Function ConvertNameFormat(pName As String) As String
'   Dim strTmp As String
'
'   If InStr(pName, ",") = 0 Then
'      strTmp = Left(pName, 1) & "," & Mid(pName, 2)
'   Else
'      strTmp = pName
'   End If
'   ConvertNameFormat = strTmp
'
'End Function

'Added by Lydia 2020/02/21 外專：案件名稱有特殊字，開啟FCP0xxxxx.新案性質.案件名稱.doc
Private Sub CmdPA174_Click()

    If pa(1) = "" Or pa(2) = "" Or pa(3) = "" Or pa(4) = "" Then Exit Sub
    If Pub_GetPA174toFile("0", pa(1), pa(2), pa(3), pa(4), Me, frm100101_M_1) = True Then
    End If
    
End Sub
