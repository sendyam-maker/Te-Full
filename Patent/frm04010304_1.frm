VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm04010304_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "各式申請書-補件, 實體審查"
   ClientHeight    =   5730
   ClientLeft      =   40
   ClientTop       =   2260
   ClientWidth     =   8950
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   8950
   Begin TabDlg.SSTab SSTab1 
      Height          =   3855
      Left            =   60
      TabIndex        =   25
      Top             =   1860
      Width           =   8865
      _ExtentX        =   15646
      _ExtentY        =   6791
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "紙本送件"
      TabPicture(0)   =   "frm04010304_1.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(1)=   "Label18"
      Tab(0).Control(2)=   "Check1(11)"
      Tab(0).Control(3)=   "Text5"
      Tab(0).Control(4)=   "Check1(10)"
      Tab(0).Control(5)=   "Check1(9)"
      Tab(0).Control(6)=   "Check1(8)"
      Tab(0).Control(7)=   "Check1(7)"
      Tab(0).Control(8)=   "Check1(6)"
      Tab(0).Control(9)=   "Check1(5)"
      Tab(0).Control(10)=   "Check1(4)"
      Tab(0).Control(11)=   "Check1(3)"
      Tab(0).Control(12)=   "Check1(2)"
      Tab(0).Control(13)=   "Check1(1)"
      Tab(0).Control(14)=   "Check1(0)"
      Tab(0).Control(15)=   "Text6"
      Tab(0).ControlCount=   16
      TabCaption(1)   =   "電子送件"
      TabPicture(1)   =   "frm04010304_1.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label17"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Shape2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label22"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label21"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label20"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label19"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label16(0)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label4"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label27"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label10"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label6"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label8"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Label14"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Label23"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Shape1"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Label28"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Label24"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Label1(71)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Label1(75)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "txtDocCh(2)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "txtDocCh(6)"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "txtCP136"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "txtCP135"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "txtDocCh(3)"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "txtDocCh(1)"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "txtDocCh(0)"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "chkDoc(0)"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "txtCP84"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "Text9"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "Text8"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "Text7"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "Text10"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "Check2"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "chkAtt(30)"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "chkAtt(29)"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "chkAtt(28)"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "chkAtt(27)"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).Control(37)=   "chkAtt(26)"
      Tab(1).Control(37).Enabled=   0   'False
      Tab(1).Control(38)=   "chkAtt(22)"
      Tab(1).Control(38).Enabled=   0   'False
      Tab(1).Control(39)=   "chkAtt(21)"
      Tab(1).Control(39).Enabled=   0   'False
      Tab(1).Control(40)=   "chkDoc(2)"
      Tab(1).Control(40).Enabled=   0   'False
      Tab(1).Control(41)=   "chkAtt(19)"
      Tab(1).Control(41).Enabled=   0   'False
      Tab(1).Control(42)=   "chkAtt(20)"
      Tab(1).Control(42).Enabled=   0   'False
      Tab(1).Control(43)=   "txtDocCh(4)"
      Tab(1).Control(43).Enabled=   0   'False
      Tab(1).ControlCount=   44
      Begin VB.TextBox txtDocCh 
         Height          =   270
         Index           =   4
         Left            =   3060
         MaxLength       =   4
         TabIndex        =   81
         Top             =   2040
         Width           =   420
      End
      Begin VB.CheckBox chkAtt 
         Caption         =   "國際優先權證明文件"
         Height          =   195
         Index           =   20
         Left            =   6240
         TabIndex        =   78
         Tag             =   ".PRI.pdf"
         Top             =   2280
         Width           =   2295
      End
      Begin VB.CheckBox chkAtt 
         Caption         =   "委任書"
         Height          =   195
         Index           =   19
         Left            =   6240
         TabIndex        =   77
         Tag             =   ".POA.pdf"
         Top             =   2010
         Width           =   1860
      End
      Begin VB.CheckBox chkDoc 
         Caption         =   "其他"
         Height          =   195
         Index           =   2
         Left            =   6240
         TabIndex        =   76
         Top             =   2550
         Width           =   660
      End
      Begin VB.CheckBox chkAtt 
         Caption         =   "文件描述"
         Enabled         =   0   'False
         Height          =   195
         Index           =   21
         Left            =   7050
         TabIndex        =   75
         Top             =   2550
         Width           =   1200
      End
      Begin VB.CheckBox chkAtt 
         Caption         =   "文件檔名"
         Enabled         =   0   'False
         Height          =   195
         Index           =   22
         Left            =   7050
         TabIndex        =   74
         Tag             =   ".ATT.pdf"
         Top             =   2760
         Width           =   1200
      End
      Begin VB.CheckBox chkAtt 
         Caption         =   "變更申請人之地址"
         Height          =   195
         Index           =   26
         Left            =   3810
         TabIndex        =   50
         Top             =   1710
         Width           =   1770
      End
      Begin VB.CheckBox chkAtt 
         Caption         =   "變更申請人之代理人"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   27
         Left            =   3810
         TabIndex        =   51
         Top             =   1980
         Width           =   2340
      End
      Begin VB.CheckBox chkAtt 
         Caption         =   "變更申請人之代表人"
         Height          =   195
         Index           =   28
         Left            =   3810
         TabIndex        =   52
         Top             =   2250
         Width           =   2340
      End
      Begin VB.CheckBox chkAtt 
         Caption         =   "變更申請人之姓名或名稱"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   29
         Left            =   3810
         TabIndex        =   53
         Top             =   2520
         Width           =   2340
      End
      Begin VB.CheckBox chkAtt 
         Caption         =   "變更申請人之國籍"
         Height          =   195
         Index           =   30
         Left            =   3810
         TabIndex        =   54
         Top             =   2790
         Width           =   2340
      End
      Begin VB.CheckBox Check2 
         Caption         =   "基本資料表"
         Height          =   195
         Left            =   6240
         TabIndex        =   72
         Tag             =   ".CONTACT.pdf"
         Top             =   1755
         Width           =   1425
      End
      Begin VB.TextBox Text10 
         Height          =   270
         Left            =   1680
         MaxLength       =   7
         TabIndex        =   57
         Top             =   585
         Width           =   1005
      End
      Begin VB.TextBox Text7 
         Height          =   270
         Left            =   2820
         MaxLength       =   15
         TabIndex        =   58
         Text            =   "一(二)"
         Top             =   855
         Width           =   1485
      End
      Begin VB.TextBox Text8 
         Height          =   270
         Left            =   4800
         MaxLength       =   11
         TabIndex        =   59
         Top             =   855
         Width           =   1260
      End
      Begin VB.TextBox Text9 
         Alignment       =   2  '置中對齊
         Appearance      =   0  '平面
         BackColor       =   &H8000000F&
         BorderStyle     =   0  '沒有框線
         Height          =   180
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   900
         Width           =   375
      End
      Begin VB.TextBox txtCP84 
         Height          =   270
         Left            =   7560
         MaxLength       =   7
         TabIndex        =   49
         Top             =   585
         Width           =   1140
      End
      Begin VB.CheckBox chkDoc 
         Caption         =   "中文本資訊"
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   60
         Top             =   1545
         Width           =   1230
      End
      Begin VB.TextBox txtDocCh 
         Height          =   270
         Index           =   0
         Left            =   3060
         MaxLength       =   4
         TabIndex        =   61
         Top             =   1500
         Width           =   420
      End
      Begin VB.TextBox txtDocCh 
         Height          =   270
         Index           =   1
         Left            =   3060
         MaxLength       =   4
         TabIndex        =   62
         Top             =   1770
         Width           =   420
      End
      Begin VB.TextBox txtDocCh 
         Height          =   270
         Index           =   3
         Left            =   3060
         MaxLength       =   4
         TabIndex        =   64
         Top             =   2595
         Width           =   420
      End
      Begin VB.TextBox txtCP135 
         Height          =   270
         Left            =   3060
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   65
         Top             =   2880
         Width           =   420
      End
      Begin VB.TextBox txtCP136 
         Height          =   270
         Left            =   3060
         MaxLength       =   4
         TabIndex        =   66
         Top             =   3150
         Width           =   420
      End
      Begin VB.TextBox txtDocCh 
         Height          =   270
         Index           =   6
         Left            =   3060
         TabIndex        =   67
         Top             =   3420
         Width           =   420
      End
      Begin VB.TextBox txtDocCh 
         Height          =   270
         Index           =   2
         Left            =   3060
         MaxLength       =   4
         TabIndex        =   63
         Top             =   2310
         Width           =   420
      End
      Begin VB.TextBox Text6 
         Height          =   270
         Left            =   -72660
         MaxLength       =   1
         TabIndex        =   39
         Top             =   2700
         Width           =   300
      End
      Begin VB.CheckBox Check1 
         Caption         =   "專利說明書一式3份"
         Height          =   255
         Index           =   0
         Left            =   -74340
         TabIndex        =   38
         Top             =   870
         Width           =   2055
      End
      Begin VB.CheckBox Check1 
         Caption         =   "申請專利範圍修正本一式3份"
         Height          =   255
         Index           =   1
         Left            =   -74340
         TabIndex        =   37
         Top             =   1170
         Width           =   2775
      End
      Begin VB.CheckBox Check1 
         Caption         =   "舉發證據正本1份"
         Height          =   255
         Index           =   2
         Left            =   -74340
         TabIndex        =   36
         Top             =   1470
         Width           =   2415
      End
      Begin VB.CheckBox Check1 
         Caption         =   "原文說明書3份"
         Height          =   255
         Index           =   3
         Left            =   -74340
         TabIndex        =   35
         Top             =   1770
         Width           =   2055
      End
      Begin VB.CheckBox Check1 
         Caption         =   "圖式一式3份"
         Height          =   255
         Index           =   4
         Left            =   -74340
         TabIndex        =   34
         Top             =   2070
         Width           =   2055
      End
      Begin VB.CheckBox Check1 
         Caption         =   "再審查理由一式2份"
         Height          =   255
         Index           =   5
         Left            =   -71460
         TabIndex        =   33
         Top             =   870
         Width           =   2055
      End
      Begin VB.CheckBox Check1 
         Caption         =   "說明書修正本一式3份"
         Height          =   255
         Index           =   6
         Left            =   -71460
         TabIndex        =   32
         Top             =   1170
         Width           =   2895
      End
      Begin VB.CheckBox Check1 
         Caption         =   "代理人委任書正本1份"
         Height          =   255
         Index           =   7
         Left            =   -71460
         TabIndex        =   31
         Top             =   1470
         Width           =   2655
      End
      Begin VB.CheckBox Check1 
         Caption         =   "申請權證明書1份"
         Height          =   255
         Index           =   8
         Left            =   -71460
         TabIndex        =   30
         Top             =   1770
         Width           =   2055
      End
      Begin VB.CheckBox Check1 
         Caption         =   "優先權證明文件正本"
         Height          =   255
         Index           =   9
         Left            =   -71460
         TabIndex        =   29
         Top             =   2070
         Width           =   2085
      End
      Begin VB.CheckBox Check1 
         Caption         =   "修正規費300元整"
         Height          =   255
         Index           =   10
         Left            =   -71460
         TabIndex        =   28
         Top             =   2370
         Width           =   2055
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   -73290
         MaxLength       =   7
         TabIndex        =   27
         Top             =   480
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "符合減免資格證明文件影本 "
         Height          =   255
         Index           =   11
         Left            =   -74340
         TabIndex        =   26
         Top             =   2370
         Width           =   2685
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
         Index           =   75
         Left            =   2200
         TabIndex        =   83
         Top             =   2085
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "序列表："
         Height          =   180
         Index           =   71
         Left            =   1530
         TabIndex        =   82
         Top             =   2085
         Width           =   735
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "(無發文日期, 辦理依據整行不顯示)"
         ForeColor       =   &H000000C0&
         Height          =   180
         Left            =   2730
         TabIndex        =   80
         Top             =   630
         Width           =   2730
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "附件文書"
         Height          =   180
         Left            =   6240
         TabIndex        =   79
         Top             =   1500
         Width           =   720
      End
      Begin VB.Shape Shape1 
         Height          =   2055
         Left            =   6150
         Top             =   1440
         Width           =   2610
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "同時辦理事項"
         Height          =   180
         Left            =   3570
         TabIndex        =   73
         Top             =   1470
         Width           =   1080
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "辦理依據:"
         Height          =   180
         Left            =   480
         TabIndex        =   71
         Top             =   390
         Width           =   765
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "發文日期："
         Height          =   180
         Left            =   735
         TabIndex        =   70
         Top             =   630
         Width           =   900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "發文字號：（        ）智專                                     字第                               號"
         Height          =   180
         Left            =   735
         TabIndex        =   69
         Top             =   900
         Width           =   5580
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "範例：（ 102 ）智專   一(二)15172   字第  10241450220  號"
         Height          =   180
         Left            =   735
         TabIndex        =   68
         Top             =   1170
         Width           =   4530
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "繳費金額："
         Height          =   180
         Left            =   6615
         TabIndex        =   55
         Top             =   630
         Width           =   900
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "摘要頁數:"
         Height          =   180
         Left            =   1530
         TabIndex        =   48
         Top             =   1545
         Width           =   765
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "說明書頁數:"
         Height          =   180
         Index           =   0
         Left            =   1530
         TabIndex        =   47
         Top             =   1815
         Width           =   945
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "圖式頁數:"
         Height          =   180
         Left            =   1530
         TabIndex        =   46
         Top             =   2640
         Width           =   765
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "頁數總計:"
         Height          =   180
         Left            =   1530
         TabIndex        =   45
         Top             =   2925
         Width           =   765
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "申請專利範圍項數:"
         Height          =   180
         Left            =   1530
         TabIndex        =   44
         Top             =   3195
         Width           =   1485
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "圖式圖數:"
         Height          =   180
         Left            =   1530
         TabIndex        =   43
         Top             =   3465
         Width           =   765
      End
      Begin VB.Shape Shape2 
         Height          =   2320
         Left            =   90
         Top             =   1440
         Width           =   3420
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "申請專利範圍頁數:"
         Height          =   180
         Left            =   1530
         TabIndex        =   42
         Top             =   2355
         Width           =   1485
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "是否修改申請書內容          (Y:WORD)"
         Height          =   180
         Left            =   -74340
         TabIndex        =   41
         Top             =   2700
         Width           =   2880
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "申請書日期:"
         Height          =   180
         Left            =   -74340
         TabIndex        =   40
         Top             =   510
         Width           =   1065
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   405
      Index           =   0
      Left            =   6810
      TabIndex        =   22
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   405
      Index           =   1
      Left            =   7635
      TabIndex        =   23
      Top             =   70
      Width           =   1200
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1050
      MaxLength       =   3
      TabIndex        =   3
      Top             =   240
      Width           =   495
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1530
      MaxLength       =   6
      TabIndex        =   2
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2370
      MaxLength       =   1
      TabIndex        =   1
      Top             =   240
      Width           =   255
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2610
      MaxLength       =   2
      TabIndex        =   0
      Top             =   240
      Width           =   375
   End
   Begin MSForms.ListBox lstNameAgent 
      Height          =   1020
      Left            =   7200
      TabIndex        =   21
      Top             =   780
      Width           =   1500
      VariousPropertyBits=   746586139
      ScrollBars      =   2
      DisplayStyle    =   2
      Size            =   "2646;1799"
      MatchEntry      =   0
      ListStyle       =   1
      MultiSelect     =   1
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1050
      TabIndex        =   14
      Top             =   780
      Width           =   6030
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "10636;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblNameAgent 
      AutoSize        =   -1  'True
      Caption         =   "出名代理人"
      Height          =   180
      Left            =   7290
      TabIndex        =   24
      Top             =   570
      Width           =   900
   End
   Begin MSForms.Label Label12 
      Height          =   180
      Index           =   7
      Left            =   1290
      TabIndex        =   20
      Top             =   1620
      Width           =   5580
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "9842;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   180
      Index           =   6
      Left            =   1290
      TabIndex        =   19
      Top             =   1380
      Width           =   1530
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "2699;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   180
      Index           =   5
      Left            =   4050
      TabIndex        =   18
      Top             =   1140
      Width           =   1170
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "2064;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   180
      Index           =   4
      Left            =   1050
      TabIndex        =   17
      Top             =   1140
      Width           =   1800
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "3175;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   180
      Index           =   2
      Left            =   4050
      TabIndex        =   16
      Top             =   570
      Width           =   2370
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "4180;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   180
      Index           =   1
      Left            =   1050
      TabIndex        =   15
      Top             =   540
      Width           =   1920
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "3387;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      Caption         =   "案件性質:"
      Height          =   180
      Left            =   3210
      TabIndex        =   13
      Top             =   240
      Width           =   765
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "機關文號:"
      Height          =   180
      Left            =   210
      TabIndex        =   12
      Top             =   1620
      Width           =   765
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "來函收文日:"
      Height          =   180
      Left            =   210
      TabIndex        =   11
      Top             =   1380
      Width           =   945
   End
   Begin MSForms.Label Label12 
      Height          =   180
      Index           =   0
      Left            =   4050
      TabIndex        =   10
      Top             =   240
      Width           =   2340
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "4128;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "智權人員:"
      Height          =   180
      Left            =   3210
      TabIndex        =   9
      Top             =   1140
      Width           =   765
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "承辦人:"
      Height          =   180
      Left            =   210
      TabIndex        =   8
      Top             =   1140
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Index           =   0
      Left            =   210
      TabIndex        =   7
      Top             =   240
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "申請案號:"
      Height          =   180
      Left            =   210
      TabIndex        =   6
      Top             =   540
      Width           =   765
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "專利號數:"
      Height          =   180
      Left            =   3210
      TabIndex        =   5
      Top             =   570
      Width           =   765
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱:"
      Height          =   180
      Left            =   210
      TabIndex        =   4
      Top             =   780
      Width           =   765
   End
End
Attribute VB_Name = "frm04010304_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/10 改成Form2.0 (Combo1,lstNameAgent,Label12)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/10 日期欄已修改
'Modified by Morgan 2015/6/2 --玲玲
'修改 Check1(9) 優先權證明文件正本及首頁影本各1份、首頁中譯本2份
Option Explicit

Public strReceiveNo As String
'Modify by Morgan 2005/8/1 改用動態陣列
'Dim pa(1 To T_PA) As String
Dim pa() As String
Dim cp() As String 'Add By Sindy 2018/6/29
Dim m_CP110 As String
Dim m_CP22 As String
Dim intWhere As Integer
Public iFrom As Integer '0=內專,1=承辦人 Add by Morgan 2011/9/22
Public oParentForm As Form '呼叫的Form Add by Morgan 2011/9/22
Dim m_CP10 As String '案件性質 Add by Amy 2014/08/13
Dim m_CaseNo As String 'Add By Sindy 2018/6/29
Dim m_AppType As String 'Add By Sindy 2018/6/29
Dim oText  As Control  'Added by Lydia 2018/12/27
Dim m_bolIsEng As Boolean 'Added by Morgan 2023/5/25

Private Sub StartLetter(ByVal ET01 As String, ByVal ET03 As String)
 Dim strTxt() As String, i As Integer, j As Integer
 Dim oChkeck As CheckBox
   EndLetter ET01, strReceiveNo, ET03, strUserNum
   j = 1
   'Modify by Morgan 2011/9/22 改寫法項目調整不必再改
   For Each oChkeck In Check1
      If oChkeck.Value = 1 Then
         ReDim Preserve strTxt(j)
         strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
            "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','補文件 V " & Format(j) & "','" & oChkeck.Caption & "')"
         j = j + 1
      End If
   Next
   'edit by nickc 2007/02/05 不用 dll 了
   'If Not objLawDll.ExecSQL(j - 1, strTxt) Then
   If Not ClsLawExecSQL(j - 1, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If
End Sub

'Add by Morgan 2005/8/1
Private Function FormSave() As Boolean

On Error GoTo ErrorHandler
   
   cnnConnection.BeginTrans
   
   'Add By Sindy 2020/4/13
   strSql = ""
   If m_AppType = "3" Then
      cp(118) = "A"
   Else
      cp(118) = ""
   End If
   strSql = strSql & ",cp118=" & CNULL(cp(118))
   '2020/4/13 END
   If lstNameAgent.Visible = True Then
      cp(110) = m_CP110
      strSql = strSql & ",cp110=" & CNULL(m_CP110)
   End If
   strSql = " UPDATE CASEPROGRESS SET " & Mid(strSql, 2) & ",cp22=" & CNULL(m_CP22) & " WHERE CP09='" & strReceiveNo & "' and cp158=0 and cp159=0"
   cnnConnection.Execute strSql
   
   'Modify By Sindy 2018/6/29
   If m_AppType <> "3" Then
   '2018/6/29 END
      'Add by Amy 2014/08/13 P台灣案電子化
      'Modified by Morgan 2015/1/12 工程師不要轉pdf
      If P台灣案電子化啟用日 <= Val(strSrvDate(1)) And iFrom = 0 Then
      If ExistCheck("AppForm", "AF01", strReceiveNo, "", False) = False Then
           '新增申請書轉檔記錄
           PUB_AddAppForm strReceiveNo
      End If
      End If
      'end 2014/08/13
   End If
   
   'Added by Lydia 2018/12/27 存中文本資訊
   strSql = ""
   For Each oText In txtDocCh
      'Modified by Lydia 2019/01/10
      'If oText.Index <= 4 And oText.Tag <> oText.Text Then
      If (oText.Index <= 4 Or oText.Index = 6) And oText.Tag <> oText.Text Then
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
   If strSql <> "" Then
        strSql = "UPDATE PATENT SET " & Mid(strSql, 2) & " WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
        strSql = "begin user_data.user_enabled:=0; " & strSql & "; end;"
        Pub_SeekTbLog strSql '新增log
        cnnConnection.Execute strSql
   End If
   'end 2018/12/27
   
   cnnConnection.CommitTrans
   
   FormSave = True
   
ErrorHandler:
   If Err.NUMBER <> 0 Then
    cnnConnection.RollbackTrans
   End If
End Function

Private Sub chkDoc_Click(Index As Integer)
   If Index = 2 Then
      If chkDoc(2).Value = 0 Then
         chkAtt(21).Enabled = False
         chkAtt(21).Value = 0
         chkAtt(22).Enabled = False
         chkAtt(22).Value = 0
      Else
         chkAtt(21).Enabled = True
         chkAtt(22).Enabled = True
      End If
   End If
End Sub

Private Sub cmdok_Click(Index As Integer)
Dim bolChk As Boolean, i As Integer
Dim Cancel As Boolean
Dim oCheck As CheckBox 'Added by Morgan 2011/11/4
Dim strFolder As String, strFileName As String 'Add By Sindy 2018/6/29
   
   If Index = 0 Then
      'Modify By Sindy 2018/6/29
      If m_AppType = "3" Then
      Else
      '2018/6/29 END
         bolChk = False
         'Modified by Morgan 2011/11/4 考慮項目會增減以後就不必再改
         'For i = 0 To 13
         '   If Check1(i).Value = 1 Then
         For Each oCheck In Check1
            If oCheck.Value = 1 Then
         'end 2011/11/4
               bolChk = True
               Exit For
            End If
         Next
         If bolChk = False Then
            MsgBox "請選擇欲補之文件 !", vbCritical
            Exit Sub
         End If
      End If
      'Add by Morgan 2005/8/1
      If lstNameAgent.Visible = True Then
         Cancel = False
         lstNameAgent_Validate Cancel
         If Cancel = True Then
            If lstNameAgent.Enabled = True Then
               lstNameAgent.SetFocus
            End If
            Exit Sub
         End If
      End If
      
      If FormSave = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Sub
      '2005/8/1---
      
      'Modify By Sindy 2018/6/29
      If m_AppType = "3" Then
         m_CaseNo = PUB_FCPCaseNo2FileName(pa(1), pa(2), pa(3), pa(4))
         'Modified by Morgan 2023/5/25 增加工程師從案件進度維護呼叫時開啟Word
         If m_bolIsEng Then
            strFileName = ""
         Else
            strFolder = PUB_Getdesktop
            strFolder = strFolder & "\" & m_CaseNo
            If Dir(strFolder, vbDirectory) = "" Then
               MkDir strFolder
            End If
            strFileName = strFolder & "\" & m_CaseNo & ".data"
         End If
         'end 2023/5/25
         
         If cp(10) = 實體審查 Then
            '2.申請書
            If StartLetter2("01", "03") = False Then Exit Sub
            NowPrint strReceiveNo, "01", "03", False, strUserNum, , , True, strExc(9)
            'strFileName = strFolder & "\" & m_CaseNo & "發明專利實體審查申請書"
            'Modify By Sindy 2020/1/6
            'strFileName = strFolder & "\" & m_CaseNo & ".data" 'Removed by Morgan 2023/5/25 移到上面
            'Call PUB_MakeDoc(strExc(9), strFileName)
         Else
            '2.申請書
            If StartLetter2("01", "02") = False Then Exit Sub
            NowPrint strReceiveNo, "01", "02", False, strUserNum, , , True, strExc(9)
            'strFileName = strFolder & "\" & m_CaseNo & "專利補正文件申請書"
            'Modify By Sindy 2020/1/6
            'strFileName = strFolder & "\" & m_CaseNo & ".data" 'Removed by Morgan 2023/5/25 移到上面
         End If
         If Check2.Value = 1 Then
            '1.基本資料
            StartLetterPA_EData "01", "14", strReceiveNo, pa, cp, True, True ', IIf(ChkAtt(26).Value = 1, True, False)
            NowPrint strReceiveNo, "01", "14", False, strUserNum, , , True, strExc(10)
            'strFileName = strFolder & "\" & m_CaseNo & ".contact"
            Call PUB_MakeDoc(strExc(9) & Chr(12) & strExc(10), strFileName, False)
         Else
            Call PUB_MakeDoc(strExc(9), strFileName)
         End If
         If m_bolIsEng Then g_WordAp.Activate 'Added by Morgan 2023/5/25
      Else
      '2018/6/29 END
         If Text6 = "Y" Then
            bolChk = True
         Else
            bolChk = False
         End If
         '補文件         00
         StartLetter "01", "00"
         strLetterDate = Text5.Text
         NowPrint strReceiveNo, "01", "00", bolChk, strUserNum, 0, , , , , , , , , , , , strReceiveNo
         'Modify by Amy 2014/08/13 P台灣案電子化
         If P台灣案電子化啟用日 <= Val(strSrvDate(1)) Then
         If bolChk = True Then
            'Modified by Morgan 2015/1/7
            '工程師仍然開啟Word(doc要放歷程)
            If iFrom = 0 Then
               frm1105_1.m_RecNo = strReceiveNo
               'Modify By Sindy 2022/5/11 流水號要足6碼
               frm1105_1.m_PdfName = Text1 & Text2 & IIf(Text3 & Text4 = "000", "", "-" & Text3 & "-" & Text4) & "." & m_CP10 & ".DATA.PDF"
               frm1105_1.Show
            End If 'Added by Morgan 2015/1/7
         End If
         End If
      End If
      'end 2014/08/13
      'Modify by Morgan 2011/9/22 配合承辦人系統也要用
      'frm040103_1.Show
      'frm040103_1.ClearForm
      
      If Not m_bolIsEng Then 'Added by Morgan 2023/5/25
         oParentForm.Show
         oParentForm.ClearForm
      End If
      'end 2011/9/22
      
   'Modifie by Morgan 2023/5/25
   'Else
   ElseIf Not m_bolIsEng Then
   'end 2023/5/25
   
      'Modify by Morgan 2011/9/22 配合承辦人系統也要用
      'frm040103_1.Show
      oParentForm.Show
   End If
   Unload Me
End Sub


Private Sub Form_Activate()
   If iFrom = 1 Then
      lstNameAgent.Enabled = False
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = 國內
   'Modify by Morgan 2011/9/22 配合承辦人系統也要用
   'With frm040103_1
   With oParentForm
      'Added by Morgan 2023/5/25
      If .Name = "frm090201_2" Then
         m_bolIsEng = True
         Text1 = SystemNumber(Trim(.lbl1(7).Caption), 1)
         Text2 = SystemNumber(Trim(.lbl1(7).Caption), 2)
         Text3 = SystemNumber(Trim(.lbl1(7).Caption), 3)
         Text4 = SystemNumber(Trim(.lbl1(7).Caption), 4)
         strReceiveNo = .lbl1(3)
      Else
      'end 2023/5/25
         Text1 = .Text1
         Text2 = .Text2
         Text3 = .Text3
         Text4 = .Text4
         strReceiveNo = .Tag
      End If
   End With
   'Add by Morgan 2005/8/1
   ReDim pa(TF_PA)
   ReDim cp(TF_CP) 'Add By Sindy 2018/6/29
   ReadPatent
   'Add by Morgan 2005/8/1
   '台灣加出名代理人清單供勾選,原是否出名欄位不顯示
   lstNameAgent.Clear
   If pa(9) = "000" Then
      PUB_SetOurAgent lstNameAgent, pa(), m_CP110, , True 'Modified by Morgan 2021/12/10 +傳入bForm2=True
      lstNameAgent.Visible = True
      lblNameAgent.Visible = True
   Else
      lstNameAgent.Visible = False
      lblNameAgent.Visible = False
   End If
   '2005/8/1 END
   Combo1.ListIndex = 0
   Text5.Text = strSrvDate(2)
   
   'Add By Sindy 2018/6/29
   'Modified by Morgan 2023/5/25
   'm_AppType = oParentForm.Text6
   If m_bolIsEng Then
      m_AppType = "3"
   Else
      m_AppType = oParentForm.Text6
   End If
   'end 2023/5/25
   If m_AppType = "3" Then
      SSTab1.Tab = 1
      SSTab1.TabVisible(0) = False
   Else
      SSTab1.Tab = 0
      SSTab1.TabVisible(1) = False
   End If
   '2018/6/29 END
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm04010304_1 = Nothing
End Sub

Private Sub ReadPatent()
 Dim rsTemp1 As New ADODB.Recordset, Lbl As Object
   
   For Each Lbl In Label12
      Lbl.Caption = ""
   Next
   'Added by Lydia 2018/12/27
   For Each oText In txtDocCh
      oText.Text = ""
   Next
   'end 2018/12/27
   
   pa(1) = Text1
   pa(2) = Text2
   pa(3) = Text3
   pa(4) = Text4
   If ClsPDReadPatentDatabase(pa, intWhere) Then  'edit by nickc 2007/02/02 不用 dll 了  If objPublicData.ReadPatentDatabase(pA, intWhere) Then
      Label12(1) = pa(11)
      Label12(2) = pa(22)
      AddCboName Combo1, pa(5), pa(6), pa(7)
   End If
   
   cp(9) = strReceiveNo
   If PUB_ReadCaseProgressDatabase(cp(), intWhere) Then
      txtCP84 = cp(84)
   End If
   
   'Added by Lydia 2018/12/27 中文本資訊-各項頁數
   txtDocCh(0).Text = pa(64) '摘要頁數
   txtDocCh(1).Text = pa(65) '說明書頁數
   txtDocCh(4).Text = pa(66) '序列表頁數
   txtDocCh(2).Text = pa(67) '申請專利範圍頁數
   txtDocCh(3).Text = pa(68) '圖式頁數
   'Added by Lydia 2019/01/10
   txtCP136.Text = pa(172) '申請專利範圍項數(最初項數)
   txtDocCh(6).Text = pa(173) '圖式圖數
   'end 2019/01/10
   If Val(pa(64)) + Val(pa(65)) + Val(pa(66)) + Val(pa(67)) + Val(pa(68)) > 0 Then
       chkDoc(0).Value = 1
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
   
   'Modify by Amy 2014/08/13 +CP10
   strExc(0) = "select cpm03,staff.st02 as st1,staff1.st02 as st2," & _
      "cp43,CP110,CP10 from caseprogress,casepropertymap,staff,staff staff1 where " & _
      "cp09='" & strReceiveNo & "' and cp01=cpm01(+) and cp10=cpm02(+) and " & _
      "cp14=staff.st01(+) and cp13=staff1.st01(+)"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   With RsTemp
   If intI = 1 Then
      m_CP110 = "" & .Fields("CP110")
      m_CP10 = .Fields("CP10") 'Add by Amy 2014/08/13
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
   
   '來函文號:
   '歸卷公文會有一筆以上，考慮可能有紙本公文保留CP05的判斷
   'strExc(0) = "select cp08,ed08 from caseprogress,edocument where CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' AND CP10 IN ('1201','1004','101','102','103') AND ed11(+)=cp09 ORDER BY NVL(ED08,CP05) DESC"
   strExc(0) = "select cp08,NVL(ED08,CP05) ED08 from caseprogress,edocument where CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' AND ed11(+)=cp09 AND cp09='" & strReceiveNo & "' ORDER BY NVL(ED08,CP05) DESC"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If Not IsNull(RsTemp("ED08")) Then
         Text10 = RsTemp("ED08") - 19110000
         Text9 = Val(Text10) \ 10000
      End If
         If Not IsNull(RsTemp("cp08")) Then
            strExc(0) = Mid(RsTemp("cp08"), InStr(RsTemp("cp08"), "智專") + 2, Len(RsTemp("cp08")))
            Text7 = Mid(strExc(0), 1, InStr(strExc(0), "字") - 1)
            strExc(0) = Replace(strExc(0), Text7 & "字第", "")
            Text8 = Mid(strExc(0), 1, InStr(RsTemp("cp08"), "號") - 1)
         End If
      'End If
   End If
End Sub

Private Sub Text5_GotFocus()
  TextInverse Text5
End Sub

Private Sub Text5_Validate(Cancel As Boolean)
   Cancel = Not ChkLetterDate(Text5.Text)
   If Cancel = True Then TextInverse Text5
End Sub

Private Sub Text6_GotFocus()
  TextInverse Text6
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 89 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub
'Add by Morgan 2005/8/1
'檢查並設定cp110資料
Private Sub lstNameAgent_Validate(Cancel As Boolean)
   Dim ii As Integer, bolCheck As Boolean
   bolCheck = False
   m_CP110 = ""
   For ii = 0 To lstNameAgent.ListCount - 1
      If lstNameAgent.Selected(ii) = True Then
         'modify by sonia 2016/10/7 員工編號已可非數字需做轉換
         'm_CP110 = m_CP110 & "," & lstNameAgent.ItemData(ii)
         'Modified by Morgan 2021/12/10 Forms2.0 改用模組
         'm_CP110 = m_CP110 & "," & PUB_Num2Id(lstNameAgent.ItemData(ii))
         m_CP110 = m_CP110 & "," & PUB_GetItemData(lstNameAgent.Tag, ii)
         bolCheck = True
      End If
   Next
   If Left(m_CP110, 1) = "," Then m_CP110 = Mid(m_CP110, 2)
   If bolCheck = True Then
      m_CP22 = ""
   Else
      m_CP22 = "N"
      If MsgBox("未勾選代理人，確定不出名？", vbYesNo + vbDefaultButton2 + vbInformation) = vbNo Then
         Cancel = True
      End If
   End If
End Sub

'申請書
Private Function StartLetter2(ByVal ET01 As String, ByVal ET03 As String) As Boolean
   Dim strTxt(200) As String, strTmp As String
   Dim ii As Integer, jj As Integer
   
   ii = 0
   EndLetter ET01, strReceiveNo, ET03, strUserNum
   
   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','本所案號','" & m_CaseNo & "')"
   
   '辦理依據
   If Text10 <> "" Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','發文日期','" & ChangeTStringToTDateString(Text10) & "')"
      
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','智專字','" & Text7 & "')"
      
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','發文號','" & Text8 & "')"
   End If

   Call PUB_GetApplPA_EData(ET01, ET03, strReceiveNo, pa(), IIf(chkAtt(26).Value = 1, False, True))
   
   '出名代理人
'   strExc(0) = "select oa05,ST26,st02 from ouragent,staff where oa01='" & pa(1) & "' and instr('" & cp(110) & "',oa02)>0 and st01(+)=oa02 order by OA03"
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      With RsTemp
'      jj = 1
'      Do While Not .EOF
'         ii = ii + 1
'         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','代理人" & jj & "-中文姓名','" & PUB_ConvertNameFormat("" & .Fields("st02")) & "')"
'         jj = jj + 1
'         .MoveNext
'      Loop
'      End With
'   End If
   'Modify By Sindy 2020/4/8 申請書:出名代理人
   Call PUB_ReadPToAppBaseData(pa(1), pa(2), pa(3), pa(4), 1, cp(110), ET01, strReceiveNo, ET03, ii, strTxt())
   
   If chkDoc(0).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','補正首次中文本','♀')"

      If txtDocCh(0).Enabled = True Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','摘要頁數','" & Val(txtDocCh(0)) & "')"
      End If
      If txtDocCh(1).Enabled = True Then
         ii = ii + 1
         '說明書頁數=說明書頁數+序列表
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
      If txtCP135.Enabled = True Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','頁數總計','" & Val(txtCP135) & "')"
      End If
      If txtCP136.Enabled = True Then
         ii = ii + 1
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
   
   If chkAtt(26).Value = 1 Or chkAtt(27).Value = 1 Or chkAtt(28).Value = 1 Or chkAtt(29).Value = 1 Or chkAtt(30).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','備註-同時辦理事項','♀')"
   End If
   If chkAtt(26).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','備註-變更申請人之地址','是')"
   End If
   If chkAtt(27).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','備註-變更申請人之代理人','是')"
   End If
   If chkAtt(28).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','備註-變更申請人之代表人','是')"
   End If
   If chkAtt(29).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','備註-變更申請人之姓名或名稱','是')"
   End If
   If chkAtt(30).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','備註-變更申請人之國籍','是')"
   End If
   
   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','繳費金額','" & Val(txtCP84) & "')"
   'Modify By Sindy 2020/4/9 有繳費金額就要帶出收據抬頭
   If Val(txtCP84) > 0 Then
      Call PUB_ReadPToAppBaseData(pa(1), pa(2), pa(3), pa(4), 3, , ET01, strReceiveNo, ET03, ii, strTxt())
   End If
   
   If Check2.Value = 0 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-基本資料表','未變更本案基本資料')"
   Else
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-基本資料表','" & m_CaseNo & Check2.Tag & "')"
   End If
   
   'Added by Morgan 2023/6/5
   '補中文明書(244)預設附件
   If m_CP10 = "244" Then
      '發明
      If pa(8) = "1" Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-發明摘要','" & m_CaseNo & ".inv_ABSTRACT.pdf')"
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-發明說明書','" & m_CaseNo & ".inv_DESCRIPTION.pdf')"
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-發明申請專利範圍','" & m_CaseNo & ".inv_CLAIMS.pdf')"
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-發明圖式','" & m_CaseNo & ".FIG.pdf')"
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-中文本原始檔1','" & m_CaseNo & ".inv.docx')"
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-中文本原始檔2','" & m_CaseNo & "dwg.pdf')"
            
      '新型
      ElseIf pa(8) = "2" Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-新型摘要','" & m_CaseNo & ".utl_ABSTRACT.pdf')"
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-新型說明書','" & m_CaseNo & ".utl_DESCRIPTION.pdf')"
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-新型申請專利範圍','" & m_CaseNo & ".utl_CLAIMS.pdf')"
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-新型圖式','" & m_CaseNo & ".FIG.pdf')"
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-中文本原始檔1','" & m_CaseNo & ".utl.docx')"
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-中文本原始檔2','" & m_CaseNo & "dwg.pdf')"
      
      End If
   End If
   'end 2023/6/5
   
   'Add By Sindy 2018/7/9
   If chkAtt(19).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-委任書','" & m_CaseNo & chkAtt(19).Tag & "')"
   End If
   If chkAtt(20).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-國際優先權證明文件','" & m_CaseNo & chkAtt(20).Tag & "')"
   End If
   If chkAtt(21).Value = 1 Or chkAtt(22).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','其他','♀')"
   End If
   If chkAtt(21).Value = 1 Then
'      'Add By Sindy 2018/4/17 首頁及摘要均附英文資料，減免規費800元整
'      If chkDoc(4).Value = 1 Then
'         strTmp = "辦理退費之收據"
'      Else
         strTmp = "♀"
'      End If
'      '2018/4/17 END
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','其他-文件描述','" & strTmp & "')"
   End If
   If chkAtt(22).Value = 1 Then
'      'Add By Sindy 2018/4/17 首頁及摘要均附英文資料，減免規費800元整
'      If chkDoc(4).Value = 1 Then
'         strTmp = m_CaseNo & ".RECEIPT.PDF"
'      Else
         strTmp = "♀"
'      End If
'      '2018/4/17 END
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','其他-文件檔名','" & m_CaseNo & chkAtt(22).Tag & "')"
   End If
   '2018/7/9 END
   
   If cp(10) = 實體審查 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','一併申請修正','否')"
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','一併申請誤譯訂正','否')"
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','備註','否')"
   End If
   
   If Not ClsLawExecSQL(ii, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   Else
      StartLetter2 = True
   End If
End Function

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

   'Memo by Lydia 2018/12/27 序列表(不算超頁費=不算錢),所以不加入進度檔的總頁數
   If Index <= 3 Then
      txtCP135 = Val(txtDocCh(0)) + Val(txtDocCh(1)) + Val(txtDocCh(2)) + Val(txtDocCh(3))
      'Added by Lydia 2018/12/27 有輸入頁數,預設勾中文本資訊
      If Val(txtCP135) > 0 And chkDoc(0).Value = vbUnchecked Then
          chkDoc(0).Value = vbChecked
      End If
      'end 2018/12/27
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

'Private Sub txtCP84_Validate(Cancel As Boolean)
'   If pa(9) = "000" Then
'      If Val(txtCP84.Text) <> Val(cp(17)) And Val(txtCP84.Text) <> Val(txtCP84.Tag) Then
'         If MsgBox("發文規費【" & txtCP84.Text & "】與收文規費【" & cp(17) & "】不同，確定要繼續！", vbInformation + vbYesNo + vbDefaultButton2) = vbYes Then
'            txtCP84.Tag = txtCP84.Text
'         Else
'            txtCP84_GotFocus
'            Cancel = True
'         End If
'      End If
'   End If
'End Sub

Private Sub Text10_GotFocus()
   TextInverse Text10
   CloseIme
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text10_Validate(Cancel As Boolean)
   If Text10 <> "" Then
      If Not ChkDate(Text10) Then
         Cancel = True
      ElseIf Val(Text10) > Val(strSrvDate(2)) Then
         MsgBox "發文日期不可大於系統日！"
         Cancel = True
      Else
         Text9 = Val(Text10) \ 10000
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

