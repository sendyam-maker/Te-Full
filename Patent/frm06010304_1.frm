VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm06010304_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "各式申請書-補件, 翻譯, 檢視中說, 製作中說, 核對中說格式"
   ClientHeight    =   5655
   ClientLeft      =   480
   ClientTop       =   975
   ClientWidth     =   8475
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   8475
   Begin VB.Frame FraPA174 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame1"
      Height          =   525
      Left            =   7650
      TabIndex        =   109
      Top             =   720
      Visible         =   0   'False
      Width           =   825
      Begin VB.CommandButton CmdPA174 
         BackColor       =   &H00C0FFFF&
         Caption         =   "特殊字"
         Height          =   280
         Left            =   0
         Style           =   1  '圖片外觀
         TabIndex        =   110
         Top             =   210
         Width           =   800
      End
      Begin VB.Label lblPA174 
         Caption         =   "有特殊字"
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   35
         TabIndex        =   111
         Top             =   0
         Width           =   765
      End
   End
   Begin VB.CheckBox Check5 
      Caption         =   "聲明中文本與申請時外文本實質內容一致"
      Height          =   255
      Left            =   3780
      TabIndex        =   108
      Top             =   5400
      Width           =   3945
   End
   Begin VB.CheckBox Check4 
      Caption         =   "首頁及摘要均附英文資料，減免規費捌佰元整(退費支票抬頭請開)"
      Height          =   255
      Left            =   360
      TabIndex        =   59
      Top             =   5130
      Width           =   8025
   End
   Begin VB.TextBox Text6 
      Height          =   270
      Left            =   7260
      MaxLength       =   1
      TabIndex        =   61
      Top             =   1320
      Width           =   300
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "frm06010304_1.frx":0000
      Left            =   1080
      List            =   "frm06010304_1.frx":000D
      Style           =   2  '單純下拉式
      TabIndex        =   80
      Top             =   660
      Width           =   615
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   6300
      TabIndex        =   62
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   7140
      TabIndex        =   63
      Top             =   60
      Width           =   1200
   End
   Begin VB.CheckBox Check51 
      Caption         =   "修正專利名稱如主旨所述"
      Height          =   255
      Left            =   360
      TabIndex        =   60
      Top             =   5400
      Width           =   2535
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3555
      Left            =   120
      TabIndex        =   78
      Top             =   1530
      Width           =   8235
      _ExtentX        =   14526
      _ExtentY        =   6271
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Page 1"
      TabPicture(0)   =   "frm06010304_1.frx":001D
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblNameAgent"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblExCode"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lbl30"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lbl5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label4"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lstNameAgent"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Text5"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Check1(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Check2(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Check1(1)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Check1(2)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Check1(4)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Check1(6)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Check2(1)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Check2(2)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Check2(4)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Check2(6)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Check2(7)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Check1(7)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Check1(27)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtIPC"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Check2(28)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Check1(28)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtExCode"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Text7"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Check1(30)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Text8"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Check3"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Check1(26)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Check1(5)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Text9"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Check1(31)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Check2(31)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).ControlCount=   34
      TabCaption(1)   =   "Page 2"
      TabPicture(1)   =   "frm06010304_1.frx":0039
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label23"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Check1(12)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Check1(14)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Check1(15)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Check2(12)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Check2(14)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Check2(15)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Check2(13)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Check1(13)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Check2(21)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Check1(21)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Check2(20)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Check1(20)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Check2(25)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Check2(24)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Check2(23)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Check1(25)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Check1(24)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Check1(23)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Check1(8)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "Check2(8)"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "chkAtt(26)"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "Check1(32)"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "Check2(32)"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).ControlCount=   24
      TabCaption(2)   =   "已刪除項目"
      TabPicture(2)   =   "frm06010304_1.frx":0055
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Check1(18)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Check1(19)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Check2(18)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Check2(19)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Check1(9)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Check2(26)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Check1(22)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Check2(22)"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Check2(3)"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Check1(3)"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "Check1(29)"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "Check2(29)"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "Check2(11)"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "Check2(10)"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "Check1(11)"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "Check1(10)"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "Check2(17)"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "Check2(16)"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "Check1(17)"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "Check1(16)"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).Control(20)=   "Check2(9)"
      Tab(2).Control(20).Enabled=   0   'False
      Tab(2).ControlCount=   21
      Begin VB.CheckBox Check2 
         Caption         =   "容後補呈"
         Height          =   255
         Index           =   32
         Left            =   -70170
         TabIndex        =   52
         Top             =   3180
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "英文說明書及圖式參考本一份"
         Height          =   255
         Index           =   32
         Left            =   -74760
         TabIndex        =   51
         Top             =   3180
         Width           =   2775
      End
      Begin VB.CheckBox chkAtt 
         Caption         =   "變更申請人之地址"
         Height          =   195
         Index           =   26
         Left            =   -74760
         TabIndex        =   50
         Top             =   2970
         Width           =   1770
      End
      Begin VB.CheckBox Check2 
         Caption         =   "容後補呈"
         Height          =   255
         Index           =   31
         Left            =   5490
         TabIndex        =   106
         Top             =   2970
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "中英文序列表電子檔光碟一份"
         Height          =   210
         Index           =   31
         Left            =   240
         TabIndex        =   105
         Top             =   3000
         Width           =   3255
      End
      Begin VB.TextBox Text9 
         BackColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   3210
         MaxLength       =   2
         TabIndex        =   22
         Top             =   2460
         Width           =   420
      End
      Begin VB.CheckBox Check2 
         Caption         =   "容後補呈"
         Height          =   255
         Index           =   9
         Left            =   -69690
         TabIndex        =   104
         Top             =   360
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "代理人委任書影本一份(正本存於第 "
         Height          =   210
         Index           =   5
         Left            =   240
         TabIndex        =   11
         Top             =   1554
         Width           =   3105
      End
      Begin VB.CheckBox Check1 
         Caption         =   "優先權證明文件電子檔（光碟片）"
         Height          =   255
         Index           =   26
         Left            =   240
         TabIndex        =   21
         Top             =   2478
         Width           =   3285
      End
      Begin VB.CheckBox Check2 
         Caption         =   "容後補呈"
         Height          =   255
         Index           =   8
         Left            =   -70170
         TabIndex        =   31
         Top             =   420
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "優惠期證明文件1份(與正本相符)"
         Height          =   255
         Index           =   8
         Left            =   -74760
         TabIndex        =   30
         Top             =   420
         Width           =   3735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "法人地位證明書正本一份"
         Height          =   255
         Index           =   16
         Left            =   -74100
         TabIndex        =   103
         Top             =   1185
         Width           =   3735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "國籍證明書正本一份"
         Height          =   255
         Index           =   17
         Left            =   -74100
         TabIndex        =   102
         Top             =   1425
         Width           =   3735
      End
      Begin VB.CheckBox Check2 
         Caption         =   "容後補呈"
         Height          =   255
         Index           =   16
         Left            =   -69510
         TabIndex        =   101
         Top             =   1185
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         Caption         =   "容後補呈"
         Height          =   255
         Index           =   17
         Left            =   -69510
         TabIndex        =   100
         Top             =   1425
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "圖卡一份"
         Height          =   255
         Index           =   10
         Left            =   -74100
         TabIndex        =   99
         Top             =   690
         Width           =   3735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "發明人拒簽之切結書正本一份"
         Height          =   255
         Index           =   11
         Left            =   -74100
         TabIndex        =   98
         Top             =   945
         Width           =   3735
      End
      Begin VB.CheckBox Check2 
         Caption         =   "容後補呈"
         Height          =   255
         Index           =   10
         Left            =   -69510
         TabIndex        =   97
         Top             =   690
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         Caption         =   "容後補呈"
         Height          =   255
         Index           =   11
         Left            =   -69510
         TabIndex        =   96
         Top             =   945
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "證書正本一份"
         Height          =   255
         Index           =   23
         Left            =   -74760
         TabIndex        =   44
         Top             =   2220
         Width           =   2775
      End
      Begin VB.CheckBox Check1 
         Caption         =   "讓與契約書正本一份"
         Height          =   255
         Index           =   24
         Left            =   -74760
         TabIndex        =   46
         Top             =   2460
         Width           =   2775
      End
      Begin VB.CheckBox Check1 
         Caption         =   "授權契約書正本一份"
         Height          =   255
         Index           =   25
         Left            =   -74760
         TabIndex        =   48
         Top             =   2700
         Width           =   2775
      End
      Begin VB.CheckBox Check2 
         Caption         =   "容後補呈"
         Height          =   255
         Index           =   23
         Left            =   -70170
         TabIndex        =   45
         Top             =   2220
         Width           =   1215
      End
      Begin VB.CheckBox Check2 
         Caption         =   "容後補呈"
         Height          =   255
         Index           =   24
         Left            =   -70170
         TabIndex        =   47
         Top             =   2460
         Width           =   1215
      End
      Begin VB.CheckBox Check2 
         Caption         =   "容後補呈"
         Height          =   255
         Index           =   25
         Left            =   -70170
         TabIndex        =   49
         Top             =   2700
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "切結書正本一份"
         Height          =   255
         Index           =   20
         Left            =   -74760
         TabIndex        =   40
         Top             =   1680
         Width           =   2775
      End
      Begin VB.CheckBox Check2 
         Caption         =   "容後補呈"
         Height          =   255
         Index           =   20
         Left            =   -70170
         TabIndex        =   41
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "變更規費參佰元整"
         Height          =   255
         Index           =   21
         Left            =   -74760
         TabIndex        =   42
         Top             =   1950
         Width           =   2775
      End
      Begin VB.CheckBox Check2 
         Caption         =   "容後補呈"
         Height          =   255
         Index           =   21
         Left            =   -70170
         TabIndex        =   43
         Top             =   1950
         Width           =   1215
      End
      Begin VB.CheckBox Check2 
         Caption         =   "容後補呈"
         Height          =   210
         Index           =   29
         Left            =   -69510
         TabIndex        =   95
         Top             =   2880
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "專利申請書首頁一份"
         Height          =   210
         Index           =   29
         Left            =   -74100
         TabIndex        =   94
         Top             =   2880
         Width           =   3735
      End
      Begin VB.CheckBox Check3 
         Caption         =   "個案"
         Height          =   255
         Left            =   2430
         TabIndex        =   9
         Top             =   1260
         Width           =   855
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   3510
         TabIndex        =   19
         Top             =   2160
         Width           =   1230
      End
      Begin VB.CheckBox Check1 
         Caption         =   "優先權證明文件影本一份(正本存於第"
         Height          =   210
         Index           =   30
         Left            =   240
         TabIndex        =   18
         Top             =   2247
         Width           =   3255
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   3390
         TabIndex        =   12
         Top             =   1470
         Width           =   1230
      End
      Begin VB.TextBox txtExCode 
         BackColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   5085
         TabIndex        =   26
         Top             =   2700
         Width           =   1230
      End
      Begin VB.CheckBox Check1 
         Caption         =   "優先權證明文件存取碼資料"
         Height          =   210
         Index           =   28
         Left            =   240
         TabIndex        =   24
         Top             =   2760
         Width           =   2565
      End
      Begin VB.CheckBox Check2 
         Caption         =   "容後補呈"
         Height          =   210
         Index           =   28
         Left            =   6435
         TabIndex        =   27
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "在國外申請日及申請案號說明頁一份"
         Height          =   255
         Index           =   3
         Left            =   -74100
         TabIndex        =   93
         Top             =   2640
         Width           =   3735
      End
      Begin VB.CheckBox Check2 
         Caption         =   "容後補呈"
         Height          =   255
         Index           =   3
         Left            =   -69510
         TabIndex        =   92
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox txtIPC 
         BackColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   1230
         TabIndex        =   29
         Top             =   3210
         Width           =   6495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "ＩＰＣ："
         Height          =   210
         Index           =   27
         Left            =   240
         TabIndex        =   28
         Top             =   3240
         Width           =   1035
      End
      Begin VB.CheckBox Check1 
         Caption         =   "優先權證明文件正本一份"
         Height          =   210
         Index           =   7
         Left            =   240
         TabIndex        =   16
         Top             =   2016
         Width           =   3570
      End
      Begin VB.CheckBox Check2 
         Caption         =   "容後補呈"
         Height          =   210
         Index           =   7
         Left            =   5490
         TabIndex        =   17
         Top             =   2016
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "寄存證明文件正本一份"
         Height          =   255
         Index           =   13
         Left            =   -74760
         TabIndex        =   34
         Top             =   930
         Width           =   3735
      End
      Begin VB.CheckBox Check2 
         Caption         =   "容後補呈"
         Height          =   255
         Index           =   13
         Left            =   -70170
         TabIndex        =   35
         Top             =   930
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         Caption         =   "容後補呈"
         Height          =   255
         Index           =   22
         Left            =   -69510
         TabIndex        =   90
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "修正規費壹仟元整"
         Height          =   255
         Index           =   22
         Left            =   -74100
         TabIndex        =   89
         Top             =   2160
         Width           =   2775
      End
      Begin VB.CheckBox Check2 
         Caption         =   "容後補呈"
         Height          =   255
         Index           =   26
         Left            =   -69510
         TabIndex        =   58
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "終止授權契約書正本一份"
         Height          =   255
         Index           =   9
         Left            =   -74100
         TabIndex        =   57
         Top             =   2400
         Width           =   2775
      End
      Begin VB.CheckBox Check2 
         Caption         =   "容後補呈"
         Height          =   255
         Index           =   19
         Left            =   -69510
         TabIndex        =   56
         Top             =   1905
         Width           =   1215
      End
      Begin VB.CheckBox Check2 
         Caption         =   "容後補呈"
         Height          =   255
         Index           =   18
         Left            =   -69510
         TabIndex        =   54
         Top             =   1665
         Width           =   1215
      End
      Begin VB.CheckBox Check2 
         Caption         =   "容後補呈"
         Height          =   255
         Index           =   15
         Left            =   -70170
         TabIndex        =   39
         Top             =   1425
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         Caption         =   "容後補呈"
         Height          =   255
         Index           =   14
         Left            =   -70170
         TabIndex        =   37
         Top             =   1185
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         Caption         =   "容後補呈"
         Height          =   255
         Index           =   12
         Left            =   -70170
         TabIndex        =   33
         Top             =   675
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         Caption         =   "容後補呈"
         Height          =   210
         Index           =   6
         Left            =   5490
         TabIndex        =   15
         Top             =   1785
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         Caption         =   "容後補呈"
         Height          =   210
         Index           =   4
         Left            =   5490
         TabIndex        =   10
         Top             =   1323
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         Caption         =   "容後補呈"
         Height          =   210
         Index           =   2
         Left            =   5490
         TabIndex        =   7
         Top             =   1092
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         Caption         =   "容後補呈"
         Height          =   210
         Index           =   1
         Left            =   5490
         TabIndex        =   5
         Top             =   861
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "圖式修正本一份"
         Height          =   255
         Index           =   19
         Left            =   -74100
         TabIndex        =   55
         Top             =   1905
         Width           =   2775
      End
      Begin VB.CheckBox Check1 
         Caption         =   "原文說明書一份"
         Height          =   255
         Index           =   18
         Left            =   -74100
         TabIndex        =   53
         Top             =   1665
         Width           =   2775
      End
      Begin VB.CheckBox Check1 
         Caption         =   "繼承證明文件正本一份"
         Height          =   255
         Index           =   15
         Left            =   -74760
         TabIndex        =   38
         Top             =   1425
         Width           =   3735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "死亡證明文件正本一份"
         Height          =   255
         Index           =   14
         Left            =   -74760
         TabIndex        =   36
         Top             =   1185
         Width           =   3735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "僱傭契約或讓與證明文件正本一份"
         Height          =   255
         Index           =   12
         Left            =   -74760
         TabIndex        =   32
         Top             =   675
         Width           =   3735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "專利申請書一份"
         Height          =   210
         Index           =   6
         Left            =   240
         TabIndex        =   14
         Top             =   1785
         Width           =   3735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "代理人委任書正本一份"
         Height          =   210
         Index           =   4
         Left            =   240
         TabIndex        =   8
         Top             =   1323
         Width           =   2205
      End
      Begin VB.CheckBox Check1 
         Caption         =   "圖式一份"
         Height          =   210
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   1092
         Width           =   3735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "設計專利說明書一份"
         Height          =   210
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   861
         Width           =   3735
      End
      Begin VB.CheckBox Check2 
         Caption         =   "容後補呈"
         Height          =   210
         Index           =   0
         Left            =   5490
         TabIndex        =   3
         Top             =   630
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "摘要、專利說明書、申請專利範圍一份"
         Height          =   210
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   630
         Width           =   3765
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   270
         Left            =   1200
         MaxLength       =   7
         TabIndex        =   0
         Top             =   330
         Width           =   1095
      End
      Begin MSForms.ListBox lstNameAgent 
         Height          =   315
         Left            =   6690
         TabIndex        =   1
         Top             =   330
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
      Begin VB.Label Label23 
         Caption         =   "(請先至客戶資料維護修改地址再產生申請書)"
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
         Height          =   195
         Left            =   -72960
         TabIndex        =   107
         Top             =   2970
         Width           =   4755
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "張（所檢送之PDF電子檔與正本相符）"
         Height          =   180
         Left            =   3720
         TabIndex        =   23
         Top             =   2490
         Width           =   3000
      End
      Begin VB.Label lbl5 
         AutoSize        =   -1  'True
         Caption         =   "號卷內)"
         Height          =   180
         Left            =   4725
         TabIndex        =   13
         Top             =   1554
         Width           =   600
      End
      Begin VB.Label lbl30 
         AutoSize        =   -1  'True
         Caption         =   "號卷內)"
         Height          =   180
         Left            =   4830
         TabIndex        =   20
         Top             =   2250
         Width           =   600
      End
      Begin VB.Label lblExCode 
         AutoSize        =   -1  'True
         Caption         =   "專利種類：發明　存取碼："
         Height          =   180
         Left            =   2880
         TabIndex        =   25
         Top             =   2760
         Width           =   2160
      End
      Begin VB.Label lblNameAgent 
         AutoSize        =   -1  'True
         Caption         =   "出名代理人"
         Height          =   180
         Left            =   5760
         TabIndex        =   91
         Top             =   360
         Width           =   900
      End
      Begin VB.Label Label2 
         Caption         =   "申請書日期"
         Height          =   255
         Left            =   240
         TabIndex        =   79
         Top             =   345
         Width           =   975
      End
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1080
      MaxLength       =   3
      TabIndex        =   67
      Top             =   150
      Width           =   550
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1632
      MaxLength       =   6
      TabIndex        =   66
      Top             =   150
      Width           =   855
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2472
      MaxLength       =   1
      TabIndex        =   65
      Top             =   150
      Width           =   255
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2712
      MaxLength       =   2
      TabIndex        =   64
      Top             =   150
      Width           =   375
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Label12"
      Height          =   180
      Index           =   7
      Left            =   4080
      TabIndex        =   88
      Top             =   1320
      Width           =   585
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Label12"
      Height          =   180
      Index           =   6
      Left            =   1230
      TabIndex        =   87
      Top             =   1320
      Width           =   570
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Label12"
      Height          =   180
      Index           =   5
      Left            =   4080
      TabIndex        =   86
      Top             =   1050
      Width           =   570
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Label12"
      Height          =   180
      Index           =   4
      Left            =   1080
      TabIndex        =   85
      Top             =   1050
      Width           =   570
   End
   Begin VB.Label Label12 
      Caption         =   "Label12"
      Height          =   240
      Index           =   3
      Left            =   1740
      TabIndex        =   84
      Top             =   690
      Width           =   5865
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Label12"
      Height          =   180
      Index           =   2
      Left            =   4080
      TabIndex        =   83
      Top             =   450
      Width           =   570
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Label12"
      Height          =   180
      Index           =   1
      Left            =   1080
      TabIndex        =   82
      Top             =   450
      Width           =   570
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "是否修改申請書內容          (Y:WORD)"
      Height          =   180
      Left            =   5550
      TabIndex        =   81
      Top             =   1320
      Width           =   2880
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      Caption         =   "案件性質:"
      Height          =   180
      Left            =   3240
      TabIndex        =   77
      Top             =   150
      Width           =   765
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "機關文號:"
      Height          =   180
      Left            =   3240
      TabIndex        =   76
      Top             =   1320
      Width           =   765
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "來函收文日:"
      Height          =   180
      Left            =   240
      TabIndex        =   75
      Top             =   1320
      Width           =   945
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Label12"
      Height          =   180
      Index           =   0
      Left            =   4080
      TabIndex        =   74
      Top             =   150
      Width           =   570
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "智權人員:"
      Height          =   180
      Left            =   3240
      TabIndex        =   73
      Top             =   1050
      Width           =   765
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "承辦人　:"
      Height          =   180
      Left            =   240
      TabIndex        =   72
      Top             =   1050
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   240
      TabIndex        =   71
      Top             =   150
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "申請案號:"
      Height          =   180
      Left            =   240
      TabIndex        =   70
      Top             =   450
      Width           =   765
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "專利號數:"
      Height          =   180
      Left            =   3240
      TabIndex        =   69
      Top             =   450
      Width           =   765
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱:"
      Height          =   180
      Left            =   240
      TabIndex        =   68
      Top             =   720
      Width           =   765
   End
End
Attribute VB_Name = "frm06010304_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/11/4 Form2.0已修改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
Option Explicit

Public strReceiveNo As String
'Modify by Morgan 2005/8/8 改用動態陣列
'Dim pa(1 To T_PA) As String
Dim pa() As String, m_CP110 As String, m_AgentName As String
Dim m_CP10 As String 'Add by Morgan 2010/1/14
Dim intWhere As Integer
Dim m_strPA31 As String, m_strPA32 As String, m_strPA33 As String, m_strPA34 As String, m_strPA35 As String 'Add By Sindy 2018/4/18
Dim m_strPA36 As String, m_strPA37 As String, m_strPA38 As String, m_strPA39 As String, m_strPA40 As String 'Add By Sindy 2018/4/18


Private Sub StartLetter(ByVal ET01 As String, ByVal ET03 As String)
Dim strTxt(1 To 30) As String, i As Integer, j As Integer, strTmp As String
Dim strCP27 As String, strCP09 As String
    
   EndLetter ET01, strReceiveNo, ET03, strUserNum
   j = 1
   For i = 0 To Me.Check1.Count - 1
      If Check1(i).Value = 1 Then
         strTmp = Check1(i).Caption
            
         'Modify by Morgan 2004/11/9 加"ＩＰＣ："27(index有調整)
         'If Check2(i).Value = 1 Then strTmp = strTmp & "　　　" & Check2(i).Caption
         If i = 27 Then
            strTmp = strTmp & ChgSQL(txtIPC.Text)
         
         'Added by Morgan 2015/8/24
         '代理人委任書影本
         ElseIf i = 5 Then
            strTmp = strTmp & Text7 & lbl5
         '優先權證明文件影本
         ElseIf i = 30 Then
            strTmp = strTmp & Text8 & lbl30
         'end 2015/8/24
         ElseIf i = 26 Then 'Add By Sindy 2018/1/31 優先權證明文件電子檔（光碟片）
            strTmp = Check1(i).Caption & Text9.Text & Label4.Caption
         'Added by Morgan 2013/12/6
         ElseIf i = 28 Then
            If Check2(i).Value = 1 Then
               strTmp = strTmp & "（期限內" & Check2(i).Caption & "）"
            Else
               'Modify By Sindy 2015/10/29 靜芳說要直接抓取優先權資料
               strExc(0) = "select * from pridate where pd01='" & pa(1) & "' and pd02='" & pa(2) & "' and pd03='" & pa(3) & "' and pd04='" & pa(4) & "'" & _
                           " order by pd05 asc"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  RsTemp.MoveFirst
                  Do While Not RsTemp.EOF
                     strTmp = strTmp & vbCrLf & String(3, "　") & "|#(粗體)" & ChgSQL("" & RsTemp.Fields("pd06")) & "　" & lblExCode & ChgSQL("" & RsTemp.Fields("pd09")) & "#|"
                     RsTemp.MoveNext
                  Loop
               Else
               '2015/10/29 END
                  strTmp = strTmp & vbCrLf & String(3, "　") & "|#(粗體)" & lblExCode & ChgSQL(txtExCode) & "#|"
               End If
            End If
         'end 2013/12/6
         Else
            'Modify by Morgan 2004/12/13 改橫式申請書
            'If Check2(i).Value = 1 Then strTmp = strTmp & "　　　" & Check2(i).Caption
            '2008/11/28 MODIFY BY SONIA
            'If Check2(i).Value = 1 Then strTmp = strTmp & "（" & Check2(i).Caption & "）"
            If Check2(i).Value = 1 Then strTmp = strTmp & "（期限內" & Check2(i).Caption & "）"
         End If
'Removed by Morgan 2015/8/24
'         Select Case i
'            'Modify by Morgan 2004/11/9 刪除原"宣誓書正本"4,"宣誓書傳真本"6,"申請權證傳真本"7;(index有調整)
'            'Case 4, 5, 6, 7
'            Case 4
'               'Modify by Morgan 2004/12/13 改橫式申請書
'               'If Check3(i).Value = 1 Then strTmp = strTmp & "　　　" & Check3(i).Caption
'               If Check3(i).Value = 1 Then strTmp = strTmp & "（" & Check3(i).Caption & "）"
'            '2004/11/9
'         End Select
         strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
            "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','補文件 V " & Format(j) & "','" & strTmp & "')"
         j = j + 1
      End If
   Next
   
   'Add By Sindy 2018/7/25
   '聲明中文本與申請時外文本實質內容一致
   If Check5.Value = 1 Then
      strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','聲明一致','♀')"
      j = j + 1
   End If
   '2018/7/25 END
   '附英文資料
   If Me.Check4.Value = vbChecked Then
      strTmp = Me.Check4.Caption
      strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附英文資料','　※" & strTmp & "')"
      j = j + 1
   End If
   If Check51.Value = 1 Then
      strTmp = Check51.Caption
      strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','修正專利名稱','　※" & strTmp & "')"
      j = j + 1
   End If
      
   'Add by Morgan 2010/1/14 申請書加超頁超項費
   'Modified by Morgan 2013/11/6 +235核對中說格式
   If m_CP10 = 翻譯 Or m_CP10 = 檢視中說 Or m_CP10 = "235" Or m_CP10 = 製作中說 Then
      strExc(1) = ""
      strTmp = ""
      strExc(0) = "SELECT NVL(SUM(CP17),0) FROM CASEPROGRESS WHERE CP01='" & pa(1) & "' AND CP02='" & pa(2) & "'" & _
         " AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' AND CP57 IS NULL AND CP10='938'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If RsTemp(0) > 0 Then
            strExc(1) = RsTemp(0)
            strTmp = "超頁規費 " & Format(RsTemp(0), DDollar) & " 元整"
         End If
      End If
      strExc(0) = "SELECT NVL(SUM(CP17),0) FROM CASEPROGRESS WHERE CP01='" & pa(1) & "' AND CP02='" & pa(2) & "'" & _
         " AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' AND CP57 IS NULL AND CP10='939'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If RsTemp(0) > 0 Then
            If strTmp <> "" Then
               strTmp = strTmp & ",超項規費" & Format(RsTemp(0), DDollar) & "元整,共計" & Format(Val(strExc(1)) + RsTemp(0), DDollar) & "元整"
            Else
               strTmp = "超項規費" & Format(RsTemp(0), DDollar) & "元整"
            End If
         End If
      End If
      
      If strTmp <> "" Then
         strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
            "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','超頁超項費用','" & strTmp & "')"
         j = j + 1
      End If
      
      'Add By Sindy 2015/12/7
      '檢查是否有延期
      strExc(0) = "select cp27,cp09,1 as sort from caseprogress" & _
                  " where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'" & _
                  " and cp10='404' and cp43='" & strReceiveNo & "'" & _
                  " Union " & _
                  "select cp27,cp09,2 as sort from caseprogress,nextprogress" & _
                  " where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'" & _
                  " and cp10='404' and cp30 is not null" & _
                  " and cp30=np22(+) and np07='202' and np01 is not null" & _
                  " Union " & _
                  "select c1.cp27,c1.cp09,3 as sort from caseprogress c1,caseprogress c2" & _
                  " where c1.cp01='" & pa(1) & "' and c1.cp02='" & pa(2) & "' and c1.cp03='" & pa(3) & "' and c1.cp04='" & pa(4) & "'" & _
                  " and c1.cp10='404' and c1.cp43 is not null" & _
                  " and c1.cp43=c2.cp09(+) and c2.cp10='202'" & _
                  " order by sort asc,cp27 desc"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         RsTemp.MoveFirst
         strCP27 = RsTemp.Fields("cp27") '有延期
         strCP09 = RsTemp.Fields("cp09")
         '檢查是否有延期受理
         strExc(0) = "select cp05,cp08 from caseprogress" & _
                     " where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'" & _
                     " and cp10='1004' and cp43='" & strCP09 & "'" & _
                     " order by cp27 desc"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            RsTemp.MoveFirst
            '有延期受理
            strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                        "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','說明一','遵照　鈞局" & Val(Left(RsTemp.Fields("cp05"), 4)) - 1911 & "年" & Mid(RsTemp.Fields("cp05"), 5, 2) & "月" & Right(RsTemp.Fields("cp05"), 2) & "日" & RsTemp.Fields("cp08") & "函辦理。')"
            j = j + 1
         Else
            '有延期
            strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                        "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','說明一','本專利案於" & Val(Left(strCP27, 4)) - 1911 & "年" & Mid(strCP27, 5, 2) & "月" & Right(strCP27, 2) & "日提出申請延期在案。')"
            j = j + 1
         End If
      Else
         '檢查是否有通知補文件 Modify By Sindy 2016/4/18:只抓通知補文件之相關總收文號為新申請案,其他案件性質之通知補文件不算
         strExc(0) = "select c1.cp05 cp05,c1.cp08 cp08,c1.cp43 cp43 from caseprogress c1, caseprogress c2" & _
                        " where c1.cp01='" & pa(1) & "' and c1.cp02='" & pa(2) & "' and c1.cp03='" & pa(3) & "' and c1.cp04='" & pa(4) & "'" & _
                        " and c1.cp10='1003' and c1.cp43 is not null" & _
                        " and c1.cp43=c2.cp09(+) and c2.cp10 in(" & NewCasePtyList & ")"
         If Label12(6) <> "" Then
            strExc(0) = strExc(0) & " and c1.cp05<=" & DBDATE(Label12(6))
         End If
         strExc(0) = strExc(0) & " order by c1.cp05 desc"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            RsTemp.MoveFirst
            strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                        "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','說明一','遵照　鈞局" & Val(Left(RsTemp.Fields("cp05"), 4)) - 1911 & "年" & Mid(RsTemp.Fields("cp05"), 5, 2) & "月" & Right(RsTemp.Fields("cp05"), 2) & "日" & RsTemp.Fields("cp08") & "函辦理。')"
            j = j + 1
         Else
            '申請發文日
            strExc(0) = "select cp27,decode(pa08,'1','發明','2','發明','3','設計',pa08) from caseprogress,patent" & _
                        " where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'" & _
                        " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)" & _
                        " and cp10 in(" & NewCasePtyList & ")"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               RsTemp.MoveFirst
               'Modify By Sindy 2016/1/5
               If IsNull(RsTemp.Fields("cp27")) Then
                  strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','說明一','本專利案於　年　月　日提出" & RsTemp.Fields(1) & "申請在案。')"
               Else
               '2016/1/5 END
                  strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','說明一','本專利案於" & Val(Left(RsTemp.Fields("cp27"), 4)) - 1911 & "年" & Mid(RsTemp.Fields("cp27"), 5, 2) & "月" & Right(RsTemp.Fields("cp27"), 2) & "日提出" & RsTemp.Fields(1) & "申請在案。')"
               End If
               j = j + 1
            End If
         End If
      End If
      '2015/12/7 END
   End If
   
   If m_CP10 = 補文件 Then
      'Modified by Morgan 2017/10/5 歸卷可能有一筆以上要抓最後發文的日期 Ex:FCP-056981--敏莉
      strExc(0) = "select cp08,ed08 from caseprogress,edocument where CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' AND CP10 IN ('1003','1004','101','102','103') AND ed11(+)=cp09 ORDER BY CP05 DESC, ED08 desc"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If Not IsNull(RsTemp("ED08")) Then
            strTmp = "" & RsTemp("CP08")
            'modify by sonia 2019/11/26 incCNV_CHINESE_MINKO改用incCNV_CHINESE_MINKO1
            strTmp = TranslateKeyWord(incCNV_CHINESE_MINKO1, TransDate(RsTemp("ED08"), 1), "") & strTmp
            strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','來函文號','" & strTmp & "')"
            j = j + 1
         End If
      End If
   'Add By Sindy 2017/9/18
   '若點新案翻譯or核對中說格式or檢視中說or製作中說等案件性質產生申請書時，
   '先判斷最近一道程序順序為1.延期受理  2. 發明申請or設計申請or新型申請 ，
   '若有延期受理先抓延期受理的智慧局函號，
   '若無則抓發明申請or設計申請or新型申請的智慧局函號。
   Else
      '檢查是否有延期受理
      strExc(0) = "select cp05,cp08,ed08 from caseprogress,edocument" & _
                  " where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'" & _
                  " and cp10='1004' AND ed11(+)=cp09" & _
                  " order by cp05 desc"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         RsTemp.MoveFirst
         '若有延期受理先抓延期受理的智慧局函號
         If Not IsNull(RsTemp("ED08")) Then
            strTmp = "" & RsTemp("cp08")
            'modify by sonia 2019/11/26 incCNV_CHINESE_MINKO改用incCNV_CHINESE_MINKO1
            strTmp = TranslateKeyWord(incCNV_CHINESE_MINKO1, TransDate(RsTemp("ED08"), 1), "") & strTmp
            strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','來函文號','" & strTmp & "')"
            j = j + 1
         End If
      Else
         '若無則抓發明申請or設計申請or新型申請的智慧局函號
         strExc(0) = "select cp27,decode(pa08,'1','發明','2','發明','3','設計',pa08),cp08,ed08 from caseprogress,patent,edocument" & _
                     " where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'" & _
                     " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)" & _
                     " and cp10 in(" & NewCasePtyList & ") AND ed11(+)=cp09" & _
                     " ORDER BY CP05 DESC"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            RsTemp.MoveFirst
            If Not IsNull(RsTemp("ED08")) Then
               strTmp = "" & RsTemp("cp08")
               'modify by sonia 2019/11/26 incCNV_CHINESE_MINKO改用incCNV_CHINESE_MINKO1
               strTmp = TranslateKeyWord(incCNV_CHINESE_MINKO1, TransDate(RsTemp("ED08"), 1), "") & strTmp
               strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','來函文號','" & strTmp & "')"
               j = j + 1
            End If
         End If
      End If
      '2017/9/18 END
   End If
   
   'Add By Sindy 2018/4/18 變更申請人地址
   If chkAtt(26).Value = 1 Then
      strTmp = ""
      If m_strPA31 <> "" Then
         If strTmp <> "" Then strTmp = strTmp & vbCrLf & "　"
         strTmp = strTmp & "變更申請人1中文地址為" & m_strPA31
      End If
      If m_strPA36 <> "" Then
         If strTmp <> "" Then strTmp = strTmp & vbCrLf & "　"
         strTmp = strTmp & "變更申請人1英文地址為" & m_strPA36
      End If
      If m_strPA32 <> "" Then
         If strTmp <> "" Then strTmp = strTmp & vbCrLf & "　"
         strTmp = strTmp & "變更申請人2中文地址為" & m_strPA32
      End If
      If m_strPA37 <> "" Then
         If strTmp <> "" Then strTmp = strTmp & vbCrLf & "　"
         strTmp = strTmp & "變更申請人2英文地址為" & m_strPA37
      End If
      If m_strPA33 <> "" Then
         If strTmp <> "" Then strTmp = strTmp & vbCrLf & "　"
         strTmp = strTmp & "變更申請人3中文地址為" & m_strPA33
      End If
      If m_strPA38 <> "" Then
         If strTmp <> "" Then strTmp = strTmp & vbCrLf & "　"
         strTmp = strTmp & "變更申請人3英文地址為" & m_strPA38
      End If
      If m_strPA34 <> "" Then
         If strTmp <> "" Then strTmp = strTmp & vbCrLf & "　"
         strTmp = strTmp & "變更申請人4中文地址為" & m_strPA34
      End If
      If m_strPA39 <> "" Then
         If strTmp <> "" Then strTmp = strTmp & vbCrLf & "　"
         strTmp = strTmp & "變更申請人4英文地址為" & m_strPA39
      End If
      If m_strPA35 <> "" Then
         If strTmp <> "" Then strTmp = strTmp & vbCrLf & "　"
         strTmp = strTmp & "變更申請人5中文地址為" & m_strPA35
      End If
      If m_strPA40 <> "" Then
         If strTmp <> "" Then strTmp = strTmp & vbCrLf & "　"
         strTmp = strTmp & "變更申請人5英文地址為" & m_strPA40
      End If
      If strTmp <> "" Then
         strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','變更申請人地址','　" & strTmp & "')"
         j = j + 1
      End If
   End If
   '2018/4/18 END
   
   If Not ClsLawExecSQL(j - 1, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If
End Sub

Private Sub Check2_Click(Index As Integer)
 On Error Resume Next
   If Check2(Index).Value = 1 Then
      Check1(Index).Value = 1
      'Check3(Index).Value = 0
   End If
End Sub

'Private Sub Check3_Click(Index As Integer)
'   If Check3(Index).Value = 1 Then Check2(Index).Value = 0
'End Sub
'Add by Morgan 2005/8/8
Private Function FormSave() As Boolean
Dim strTmp As String
   
On Error GoTo ErrorHandler
   
   cnnConnection.BeginTrans
   
   If lstNameAgent.Visible = True Then
      strSql = " UPDATE CASEPROGRESS SET cp110=" & CNULL(m_CP110) & " WHERE CP09='" & strReceiveNo & "'"
      cnnConnection.Execute strSql
   End If
   
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
      Else
         chkAtt(26).Value = 0
      End If
   End If
   '2018/4/18 END
   
   'Add By Sindy 2018/9/14 回寫進度備註
   If Check4.Value = 1 Then
      '首頁及摘要均附英文資料，減免規費800元整
      strSql = "UPDATE CASEPROGRESS SET cp64='首頁及摘要均附英文資料，減免規費800元整;'||cp64" & _
               " WHERE CP01='" & pa(1) & "' and CP02='" & pa(2) & "' and CP03='" & pa(3) & "' and CP04='" & pa(4) & "'" & _
               " and cp10 in('201','209','235') and cp158=0 and cp159=0"
      cnnConnection.Execute strSql
   End If
   
   cnnConnection.CommitTrans
   FormSave = True
   
ErrorHandler:
   If Err.Number <> 0 Then
    cnnConnection.RollbackTrans
   End If
End Function

'Add by Morgan 2005/8/8
Private Function TxtValidate() As Boolean
   Dim Cancel As Boolean
   If lstNameAgent.Visible = True Then
      Cancel = False
      lstNameAgent_Validate Cancel
      If Cancel = True Then
         SSTab1.Tab = 0
         lstNameAgent.SetFocus
         Exit Function
      End If
   End If
   
   'Added by Morgan 2013/12/6
   If Check1(28).Value = 1 And Check2(28).Value = 0 Then
      'Modify By Sindy 2015/10/29 Mark:靜芳說要直接抓取優先權資料
'      If txtExCode = "" Then
'         MsgBox "請輸入交換碼！", vbExclamation
'         txtExCode.SetFocus
'         Exit Function
'      End If
   End If
   'end 2013/12/6
   
   'Added by Morgan 2015/8/24
   If Check1(5).Value = 1 Then
      If Text7 = "" Then
         MsgBox "請輸入正本申請號！", vbInformation
         Text7.SetFocus
         Exit Function
      ElseIf Text7 = pa(11) Then
         MsgBox "正本申請號不可為本案！", vbInformation
         Text7.SetFocus
         Exit Function
      Else
         strExc(0) = "SELECT 1 FROM PATENT WHERE PA11='" & Text7 & "' AND PA09='000' AND INSTR(PA26||PA27||PA28||PA29||PA30,'" & ChangeCustomerL(pa(26)) & "')>0"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 0 Then
            MsgBox "正本申請號輸入錯誤！申請人必須與本案相同。", vbInformation
            Text7.SetFocus
            Exit Function
         End If
      End If
   End If
   If Check1(30).Value = 1 Then
      If Text8 = "" Then
         MsgBox "請輸入正本申請號！", vbInformation
         Text8.SetFocus
         Exit Function
      ElseIf Text8 = pa(11) Then
         MsgBox "正本申請號不可為本案！", vbInformation
         Text8.SetFocus
         Exit Function
      Else
         strExc(0) = "SELECT 1 FROM PATENT,PRIDATE A WHERE PA11='" & Text8 & "' AND PA09='000' AND PD01(+)=PA01 AND PD02(+)=PA02 AND PD03(+)=PA03 AND PD04(+)=PA04 AND EXISTS(SELECT * FROM PRIDATE B WHERE B.PD01='" & pa(1) & "' AND B.PD02='" & pa(2) & "' AND B.PD03='" & pa(3) & "' AND B.PD04='" & pa(4) & "' AND B.PD05=A.PD05 AND B.PD06=A.PD06 AND B.PD07=A.PD07)"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 0 Then
            MsgBox "正本申請號輸入錯誤！優先權號必須與本案相同。", vbInformation
            Text8.SetFocus
            Exit Function
         End If
      End If
   End If
   'end 2015/8/24
   'Add By Sindy 2018/2/2 優先權證明文件電子檔（光碟片）
   If Check1(26).Value = vbChecked Then
      If Text9 = "" Then
         MsgBox "請輸入光碟片張數！", vbInformation
         Text9.SetFocus
         Exit Function
      End If
   End If
   '2018/2/2 END
   
   TxtValidate = True
End Function

Private Sub cmdOK_Click(Index As Integer)
Dim bolChk As Boolean, i As Integer
Dim stET02 As String 'Add by Morgan 2004/12/2
Dim strCaseData As String 'Add By Sindy 2015/12/11
Dim strBookDate As String 'Add By Sindy 2015/12/11
   
   If Index = 0 Then
      bolChk = False
        'Modify By Cheng 2003/01/03
'      For i = 0 To 27
      For i = 0 To Me.Check1.Count - 1
         If Check1(i).Value = 1 Then
            bolChk = True
            Exit For
         End If
      Next
      If bolChk = False Then
         MsgBox "請選擇欲補之文件 !", vbCritical
         Exit Sub
      End If
      If Text6 = "Y" Then
         bolChk = True
      Else
         bolChk = False
      End If
      'Add by Morgan 2005/8/8
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
      
      stET02 = "00"
      '補文件         00
      StartLetter "01", stET02
      strLetterDate = Text5.Text
      NowPrint strReceiveNo, "01", stET02, bolChk, strUserNum, 0
      
      'Add By Sindy 2015/12/7
      'Modify By Sindy 2016/12/27 敏莉:當勾選”代理人委任書正本一份”出譯文,但若有勾選”容後補呈”就不需要出譯文
      'If Check1(4).Value = 1 Then
      If Check1(4).Value = 1 And Check2(4).Value = 0 Then
      '2016/12/27 END
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
            NowPrint strReceiveNo, "01", "31", bolChk, strUserNum, 0
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
            NowPrint strReceiveNo, "01", "32", bolChk, strUserNum, 0
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
            NowPrint strReceiveNo, "01", "33", bolChk, strUserNum, 0
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
            NowPrint strReceiveNo, "01", "34", bolChk, strUserNum, 0
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
            NowPrint strReceiveNo, "01", "35", bolChk, strUserNum, 0
         End If
      End If
      '2015/12/7 END
      
      frm060103_1.Show
      ' 90.08.27 modify by louis (回到原畫面要清除畫面)
      frm060103_1.ClearForm
   Else
      frm060103_1.Show
   End If
   Unload Me
End Sub

Private Sub Combo1_Click()
   Select Case Combo1
      Case "中"
         Label12(3) = pa(5)
      Case "英"
         Label12(3) = pa(6)
      'Modified by Lydia 2022/04/25 「日文名稱」改為「外文名稱」
      Case "外"
         Label12(3) = pa(7)
   End Select
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = 國外_FC
   With frm060103_1
      Text1 = .Text1
      Text2 = .Text2
      Text3 = .Text3
      Text4 = .Text4
      strReceiveNo = .Tag
   End With
   'Add by Morgan 2005/8/8
   ReDim pa(TF_PA)
   ReadPatent
   'Add by Morgan 2005/8/8
   '加出名代理人清單供勾選
   lstNameAgent.Clear
   PUB_SetOurAgent lstNameAgent, pa(), m_CP110, , True
   'Added by Sindy 2021/5/10 如果一開始將ListBox拉到需要的大小，字型會自動放大；所以畫面預設為一列高度，Form_Load才放大到需要的大小
   lstNameAgent.Height = 1100
   lstNameAgent.Width = 1300

   Combo1.ListIndex = 0
   Text6 = "Y"
   Text5.Text = strSrvDate(2)
   
   SSTab1.TabVisible(2) = False 'Add By Sindy 2018/1/31
   SSTab1.Tab = 0 'Add By Sindy 2018/1/31
   
   FraPA174.BackColor = &H8000000F 'Added by Lydia 2020/02/21
   
   'Modified by Lydia 2020/03/31 改模組A0802Query => CompNameQuery
   Me.Check4.Caption = "首頁及摘要均附英文資料，減免規費捌佰元整(退費支票抬頭請開" & CompNameQuery("2") & ")" 'Add By Sindy 2020/3/30
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm06010304_1 = Nothing
End Sub

Private Sub ReadPatent()
 Dim rsTemp1 As New ADODB.Recordset, Lbl As Label
   For Each Lbl In Label12
      Lbl = ""
   Next
   pa(1) = Text1
   pa(2) = Text2
   pa(3) = Text3
   pa(4) = Text4
   If ClsPDReadPatentDatabase(pa(), intWhere) Then 'edit by nickc 2007/02/02 不用 dll 了 If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
      Text5 = pa(10)
      Label12(1) = pa(11)
      Label12(2) = pa(22)
      Label12(3) = pa(5)
      lblExCode = Replace(lblExCode, "發明", PUB_GetPatentKindName(pa(8), 台灣國家代號)) 'Added by Morgan 2013/12/6
   End If
   strExc(0) = "select cpm03,staff.st02 as st1,staff1.st02 as st2," & _
      "cp43,CP110,cp10 from caseprogress,casepropertymap,staff," & _
      "staff staff1 where cp09='" & strReceiveNo & "' AND cp01=cpm01(+) and cp10=cpm02(+) and cp14=staff.st01(+) and " & _
      "cp13=staff1.st01(+)"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   With RsTemp
   If intI = 1 Then
      m_CP110 = "" & .Fields("CP110")
      m_CP10 = "" & .Fields("CP10")
      If Not IsNull(.Fields(0)) Then Label12(0) = .Fields(0)
      If Not IsNull(.Fields(1)) Then Label12(4) = .Fields(1)
      If Not IsNull(.Fields(2)) Then Label12(5) = .Fields(2)
      If Not IsNull(.Fields(3)) Then
         strExc(0) = "SELECT CP05,CP08 FROM CASEPROGRESS WHERE CP09='" & .Fields(3) & "'"
         intI = 1
         Set rsTemp1 = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If Not IsNull(rsTemp1.Fields(0)) Then Label12(6) = TransDate(rsTemp1.Fields(0), 1)
            If Not IsNull(rsTemp1.Fields(1)) Then Label12(7) = rsTemp1.Fields(1)
         End If
      End If
   End If
   End With
   'Modify By Sindy 2018/1/8 Mark
'   'Add By Sindy 2017/5/15
'   '專利種類PA08='3'且案件性質CP10='210'製作中說時,第二選項之 '設計圖說一式二份'請改為'設計專利說明書及圖式一式二份'
'   If pa(8) = "3" And m_CP10 = "210" Then
'      Check1(1).Caption = "設計專利說明書及圖式一式二份"
'   Else
'      Check1(1).Caption = "設計圖說一式二份"
'   End If
'   '2017/5/15 END

   'Added by Lydia 2018/10/19 原本是在電子送件有自,紙本送件也要自動勾選(ex.FCP-59244) ;
   'P.X若工程師提申後從命名系統修改專利名稱則自動加註"一併修改專利名稱",待中說或補文件發文後自動將註記的欄位清空。
   strExc(0) = "select cp09,cp10,tct01,tct15 from caseprogress,transcasetitle where CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' AND cp31='Y' and cp09=tct01(+) "
   intI = 1
   Set rsTemp1 = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If "" & rsTemp1.Fields("tct01") <> "" And "" & rsTemp1.Fields("tct15") = "Y" Then
          Check51.Value = vbChecked
      End If
   End If
   'end 2018/10/19
   
   'Added by Lydia 2020/02/21 預設「名稱有特殊字」
   FraPA174.Visible = False
   If pa(1) = "FCP" Or pa(1) = "P" Then
       If pa(174) = "Y" Then
          FraPA174.Visible = True
       End If
   End If
   'end 2020/02/21
   
End Sub


Private Sub Text5_GotFocus()
  TextInverse Text5
End Sub

Private Sub Text5_Validate(Cancel As Boolean)
   Cancel = Not ChkLetterDate(Text5.Text)
   If Cancel = True Then TextInverse Text5
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 89 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text9_GotFocus()
  TextInverse Text9
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub txtExCode_GotFocus()
   TextInverse txtExCode
End Sub

Private Sub txtExCode_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtIPC_GotFocus()
   TextInverse txtIPC
End Sub
'Add by Morgan 2005/8/8
'檢查並設定cp110資料
Private Sub lstNameAgent_Validate(Cancel As Boolean)
   Dim ii As Integer
   Cancel = True
   m_CP110 = "": m_AgentName = ""
   For ii = 0 To lstNameAgent.ListCount - 1
      If lstNameAgent.Selected(ii) = True Then
         'modify by sonia 2016/10/7 員工編號已可非數字需做轉換
         'm_CP110 = m_CP110 & "," & lstNameAgent.ItemData(ii)
         'Modify By Sindy 2021/5/10
         'm_CP110 = m_CP110 & "," & PUB_Num2Id(lstNameAgent.ItemData(ii))
         m_CP110 = m_CP110 & "," & PUB_GetItemData(lstNameAgent.Tag, ii)
         m_AgentName = m_AgentName & "、" & lstNameAgent.List(ii)
         '2021/5/10 END
         Cancel = False
      End If
   Next
   If Cancel = True Then
      MsgBox "出名代理人不可空白！", vbExclamation
   Else
      If Left(m_CP110, 1) = "," Then m_CP110 = Mid(m_CP110, 2)
      m_AgentName = Mid(m_AgentName, 2) 'Add By Sindy 2021/5/10
   End If
End Sub

'Added by Lydia 2020/02/21 外專：案件名稱有特殊字，開啟FCP0xxxxx.新案性質.案件名稱.doc
Private Sub CmdPA174_Click()

    If pa(1) = "" Or pa(2) = "" Or pa(3) = "" Or pa(4) = "" Then Exit Sub
    If Pub_GetPA174toFile("0", pa(1), pa(2), pa(3), pa(4), Me, frm100101_M_1) = True Then
    End If
    
End Sub
