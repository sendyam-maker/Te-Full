VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060120 
   BorderStyle     =   1  '單線固定
   Caption         =   "客戶提供文件"
   ClientHeight    =   6660
   ClientLeft      =   828
   ClientTop       =   972
   ClientWidth     =   8244
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   8244
   Begin VB.CommandButton cmdTCN13 
      BackColor       =   &H00C0FFFF&
      Caption         =   "確定無對應英/中說"
      Height          =   330
      Left            =   4680
      Style           =   1  '圖片外觀
      TabIndex        =   78
      Top             =   630
      Width           =   1755
   End
   Begin VB.TextBox textTCN13 
      Height          =   300
      Left            =   2544
      MaxLength       =   1
      TabIndex        =   77
      Top             =   168
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox textTCN16 
      Height          =   300
      Left            =   1356
      MaxLength       =   1
      TabIndex        =   75
      Top             =   140
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "卷宗區"
      Height          =   375
      Index           =   2
      Left            =   5520
      TabIndex        =   40
      Top             =   120
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "外文本"
      Height          =   375
      Index           =   1
      Left            =   6360
      TabIndex        =   38
      Top             =   120
      Width           =   800
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "尋找(&F)"
      Height          =   330
      Left            =   2865
      TabIndex        =   2
      Top             =   630
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   7200
      TabIndex        =   39
      Top             =   120
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   375
      Index           =   0
      Left            =   4680
      TabIndex        =   37
      Top             =   120
      Width           =   800
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4560
      Left            =   120
      TabIndex        =   41
      Top             =   2010
      Width           =   8055
      _ExtentX        =   14203
      _ExtentY        =   8043
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "說明書"
      TabPicture(0)   =   "frm060120.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "補文件 ＆ 資訊"
      TabPicture(1)   =   "frm060120.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  '沒有框線
         Caption         =   "Frame2"
         Height          =   4035
         Left            =   -74880
         TabIndex        =   62
         Top             =   400
         Width           =   7815
         Begin VB.CheckBox Chk1 
            Caption         =   "11.其他"
            Height          =   255
            Index           =   33
            Left            =   0
            TabIndex        =   33
            Top             =   3330
            Width           =   1935
         End
         Begin VB.CheckBox Chk1 
            Caption         =   "10.非WTO會員國之住所證明"
            Height          =   375
            Index           =   31
            Left            =   0
            TabIndex        =   31
            Top             =   2775
            Width           =   1695
         End
         Begin VB.CheckBox Chk1 
            Caption         =   "9.發明人資訊"
            Height          =   255
            Index           =   29
            Left            =   0
            TabIndex        =   29
            Top             =   2220
            Width           =   1455
         End
         Begin VB.CheckBox Chk1 
            Caption         =   "8.申請人資訊"
            Height          =   255
            Index           =   27
            Left            =   0
            TabIndex        =   27
            Top             =   1665
            Width           =   1455
         End
         Begin VB.CheckBox Chk1 
            Caption         =   "7.代表人資訊"
            Height          =   255
            Index           =   25
            Left            =   0
            TabIndex        =   25
            Top             =   1110
            Width           =   1455
         End
         Begin VB.CheckBox Chk1 
            Caption         =   "6.委任狀"
            Height          =   255
            Index           =   23
            Left            =   0
            TabIndex        =   23
            Top             =   555
            Width           =   1455
         End
         Begin VB.CheckBox Chk1 
            Caption         =   "5.優先權證明書"
            Height          =   255
            Index           =   21
            Left            =   0
            TabIndex        =   21
            Top             =   30
            Width           =   1695
         End
         Begin MSForms.TextBox txtCSD 
            Height          =   645
            Index           =   34
            Left            =   2520
            TabIndex        =   34
            Top             =   3330
            Width           =   4800
            VariousPropertyBits=   -1466941413
            MaxLength       =   200
            ScrollBars      =   2
            Size            =   "8467;1129"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtCSD 
            Height          =   520
            Index           =   32
            Left            =   2520
            TabIndex        =   32
            Top             =   2775
            Width           =   4800
            VariousPropertyBits=   -1466941413
            MaxLength       =   60
            ScrollBars      =   2
            Size            =   "8467;917"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtCSD 
            Height          =   520
            Index           =   30
            Left            =   2520
            TabIndex        =   30
            Top             =   2220
            Width           =   4800
            VariousPropertyBits=   -1466941413
            MaxLength       =   60
            ScrollBars      =   2
            Size            =   "8467;917"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtCSD 
            Height          =   520
            Index           =   28
            Left            =   2520
            TabIndex        =   28
            Top             =   1665
            Width           =   4800
            VariousPropertyBits=   -1466941413
            MaxLength       =   60
            ScrollBars      =   2
            Size            =   "8467;917"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtCSD 
            Height          =   520
            Index           =   26
            Left            =   2520
            TabIndex        =   26
            Top             =   1110
            Width           =   4800
            VariousPropertyBits=   -1466941413
            MaxLength       =   60
            ScrollBars      =   2
            Size            =   "8467;917"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtCSD 
            Height          =   520
            Index           =   24
            Left            =   2520
            TabIndex        =   24
            Top             =   555
            Width           =   4800
            VariousPropertyBits=   -1466941413
            MaxLength       =   60
            ScrollBars      =   2
            Size            =   "8467;917"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtCSD 
            Height          =   520
            Index           =   22
            Left            =   2520
            TabIndex        =   22
            Top             =   0
            Width           =   4800
            VariousPropertyBits=   -1466941413
            MaxLength       =   60
            ScrollBars      =   2
            Size            =   "8467;917"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "(備註：                                                                                                             )"
            Height          =   180
            Index           =   15
            Left            =   1920
            TabIndex        =   69
            Top             =   3367
            Width           =   5565
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "(備註：                                                                                                             )"
            Height          =   180
            Index           =   14
            Left            =   1920
            TabIndex        =   68
            Top             =   2775
            Width           =   5565
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "(備註：                                                                                                             )"
            Height          =   180
            Index           =   13
            Left            =   1920
            TabIndex        =   67
            Top             =   2220
            Width           =   5565
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "(備註：                                                                                                             )"
            Height          =   180
            Index           =   11
            Left            =   1920
            TabIndex        =   66
            Top             =   1665
            Width           =   5565
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "(備註：                                                                                                             )"
            Height          =   180
            Index           =   10
            Left            =   1920
            TabIndex        =   65
            Top             =   1110
            Width           =   5565
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "(備註：                                                                                                             )"
            Height          =   180
            Index           =   9
            Left            =   1920
            TabIndex        =   64
            Top             =   555
            Width           =   5565
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "(備註：                                                                                                             )"
            Height          =   180
            Index           =   12
            Left            =   1920
            TabIndex        =   63
            Top             =   60
            Width           =   5565
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  '沒有框線
         Caption         =   "Frame1"
         Height          =   4095
         Left            =   120
         TabIndex        =   48
         Top             =   400
         Width           =   7815
         Begin VB.ComboBox CmbFL 
            Height          =   300
            Index           =   3
            Left            =   1560
            Style           =   2  '單純下拉式
            TabIndex        =   19
            Top             =   3680
            Width           =   5100
         End
         Begin VB.ComboBox CmbFL 
            Height          =   300
            Index           =   2
            Left            =   1560
            Style           =   2  '單純下拉式
            TabIndex        =   15
            Top             =   2670
            Width           =   5100
         End
         Begin VB.ComboBox CmbFL 
            Height          =   300
            Index           =   1
            Left            =   1560
            Style           =   2  '單純下拉式
            TabIndex        =   11
            Top             =   1640
            Width           =   5100
         End
         Begin VB.ComboBox CmbFL 
            Height          =   300
            Index           =   0
            Left            =   1560
            Style           =   2  '單純下拉式
            TabIndex        =   7
            Top             =   540
            Width           =   5100
         End
         Begin VB.CommandButton cmdFile 
            Caption         =   "上傳檔案"
            Height          =   330
            Index           =   0
            Left            =   6780
            TabIndex        =   8
            Top             =   555
            Width           =   920
         End
         Begin VB.CommandButton cmdFile 
            Caption         =   "上傳檔案"
            Height          =   330
            Index           =   1
            Left            =   6780
            TabIndex        =   12
            Top             =   1640
            Width           =   920
         End
         Begin VB.CommandButton cmdFile 
            Caption         =   "上傳檔案"
            Height          =   330
            Index           =   2
            Left            =   6780
            TabIndex        =   16
            Top             =   2670
            Width           =   920
         End
         Begin VB.CommandButton cmdFile 
            Caption         =   "上傳檔案"
            Height          =   330
            Index           =   3
            Left            =   6780
            TabIndex        =   20
            Top             =   3675
            Width           =   920
         End
         Begin VB.CheckBox Chk1 
            Caption         =   "4.簡(繁)體中說"
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   17
            Top             =   3120
            Width           =   1935
         End
         Begin VB.CheckBox Chk1 
            Caption         =   "3.英說(參考/翻譯用)"
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   13
            Top             =   2080
            Width           =   1935
         End
         Begin VB.CheckBox Chk1 
            Caption         =   "2.替換版原文說明書"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   9
            Top             =   1040
            Width           =   1935
         End
         Begin VB.CheckBox Chk1 
            Caption         =   "1.原文說明書"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   5
            Top             =   0
            Width           =   1455
         End
         Begin MSForms.TextBox txtCSD 
            Height          =   300
            Index           =   14
            Left            =   0
            TabIndex        =   70
            Top             =   600
            Width           =   460
            VariousPropertyBits=   671105051
            Size            =   "811;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtCSD 
            Height          =   300
            Index           =   15
            Left            =   1560
            TabIndex        =   10
            Top             =   1320
            Width           =   3915
            VariousPropertyBits=   671105051
            MaxLength       =   60
            Size            =   "6906;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtCSD 
            Height          =   300
            Index           =   16
            Left            =   0
            TabIndex        =   71
            Top             =   1680
            Width           =   460
            VariousPropertyBits=   671105051
            Size            =   "811;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtCSD 
            Height          =   300
            Index           =   17
            Left            =   1560
            TabIndex        =   14
            Top             =   2355
            Width           =   3915
            VariousPropertyBits=   671105051
            MaxLength       =   60
            Size            =   "6906;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtCSD 
            Height          =   300
            Index           =   18
            Left            =   0
            TabIndex        =   72
            Top             =   2640
            Width           =   460
            VariousPropertyBits=   671105051
            Size            =   "811;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtCSD 
            Height          =   300
            Index           =   19
            Left            =   1560
            TabIndex        =   18
            Top             =   3360
            Width           =   3915
            VariousPropertyBits=   671105051
            MaxLength       =   60
            Size            =   "6906;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtCSD 
            Height          =   300
            Index           =   13
            Left            =   1560
            TabIndex        =   6
            Top             =   240
            Width           =   3915
            VariousPropertyBits=   671105051
            MaxLength       =   60
            Size            =   "6906;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtCSD 
            Height          =   300
            Index           =   20
            Left            =   0
            TabIndex        =   73
            Top             =   3600
            Width           =   460
            VariousPropertyBits=   671105051
            Size            =   "811;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "*.ORI.REP2.PDF"
            ForeColor       =   &H000000FF&
            Height          =   180
            Index           =   16
            Left            =   6440
            TabIndex        =   61
            Top             =   1320
            Width           =   1230
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "檔案路徑："
            Height          =   180
            Index           =   3
            Left            =   660
            TabIndex        =   60
            Top             =   3720
            Width           =   900
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "檔案路徑："
            Height          =   180
            Index           =   2
            Left            =   660
            TabIndex        =   59
            Top             =   2745
            Width           =   900
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "檔案路徑："
            Height          =   180
            Index           =   1
            Left            =   660
            TabIndex        =   58
            Top             =   1680
            Width           =   900
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "檔案路徑："
            Height          =   180
            Index           =   0
            Left            =   660
            TabIndex        =   57
            Top             =   600
            Width           =   900
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "(備註：                                                                                        )"
            Height          =   180
            Index           =   1
            Left            =   960
            TabIndex        =   56
            Top             =   300
            Width           =   4620
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "(備註：                                                                                         )"
            Height          =   180
            Index           =   2
            Left            =   960
            TabIndex        =   55
            Top             =   1380
            Width           =   4665
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "(備註：                                                                                        )"
            Height          =   180
            Index           =   3
            Left            =   960
            TabIndex        =   54
            Top             =   2415
            Width           =   4620
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "(備註：                                                                                        )"
            Height          =   180
            Index           =   4
            Left            =   960
            TabIndex        =   53
            Top             =   3420
            Width           =   4620
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "檔名：*.ORI.PDF"
            ForeColor       =   &H000000FF&
            Height          =   180
            Index           =   5
            Left            =   2040
            TabIndex        =   52
            Top             =   0
            Width           =   1320
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "檔名：*.ORI.REP.PDF，上傳後會自動加流水號，例如：*.ORI.REP1.PDF"
            ForeColor       =   &H000000FF&
            Height          =   180
            Index           =   6
            Left            =   2040
            TabIndex        =   51
            Top             =   1040
            Width           =   5610
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "檔名：*.ENSP.MSG"
            ForeColor       =   &H000000FF&
            Height          =   180
            Index           =   7
            Left            =   2040
            TabIndex        =   50
            Top             =   2085
            Width           =   1485
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "檔名：*.CNSP.MSG"
            ForeColor       =   &H000000FF&
            Height          =   180
            Index           =   8
            Left            =   2040
            TabIndex        =   49
            Top             =   3120
            Width           =   1500
         End
      End
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   330
      Left            =   900
      TabIndex        =   76
      Top             =   1020
      Width           =   7275
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "12832;582"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblTCN16 
      Caption         =   "新案暫不認領:                 (Y:是)"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   36
      TabIndex        =   74
      Top             =   168
      Visible         =   0   'False
      Width           =   2904
   End
   Begin MSForms.TextBox txtPA 
      Height          =   300
      Index           =   75
      Left            =   900
      TabIndex        =   47
      Top             =   1680
      Width           =   1005
      VariousPropertyBits=   671105055
      MaxLength       =   9
      Size            =   "1773;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtPA 
      Height          =   300
      Index           =   26
      Left            =   900
      TabIndex        =   46
      Top             =   1365
      Width           =   1005
      VariousPropertyBits=   671105055
      MaxLength       =   9
      Size            =   "1773;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtPA 
      Height          =   300
      Index           =   4
      Left            =   2460
      TabIndex        =   4
      Top             =   630
      Width           =   375
      VariousPropertyBits=   671105051
      MaxLength       =   2
      Size            =   "661;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtPA 
      Height          =   300
      Index           =   3
      Left            =   2220
      TabIndex        =   3
      Top             =   630
      Width           =   255
      VariousPropertyBits=   671105051
      MaxLength       =   1
      Size            =   "450;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtPA 
      Height          =   300
      Index           =   2
      Left            =   1380
      TabIndex        =   1
      Top             =   630
      Width           =   855
      VariousPropertyBits=   671105051
      MaxLength       =   6
      Size            =   "1508;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtPA 
      Height          =   300
      Index           =   1
      Left            =   900
      TabIndex        =   0
      Top             =   630
      Width           =   495
      VariousPropertyBits=   671105051
      MaxLength       =   3
      Size            =   "873;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label27 
      Height          =   255
      Index           =   26
      Left            =   1920
      TabIndex        =   45
      Top             =   1395
      Width           =   5955
      VariousPropertyBits=   27
      Caption         =   "Label27"
      Size            =   "10504;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Lbl 
      Caption         =   "申請人1："
      Height          =   240
      Index           =   26
      Left            =   30
      TabIndex        =   44
      Top             =   1410
      Width           =   855
   End
   Begin MSForms.Label Label27 
      Height          =   255
      Index           =   75
      Left            =   1920
      TabIndex        =   43
      Top             =   1710
      Width           =   5955
      VariousPropertyBits=   27
      Caption         =   "Label27"
      Size            =   "10504;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Lbl 
      Caption         =   "代理人："
      Height          =   240
      Index           =   75
      Left            =   30
      TabIndex        =   42
      Top             =   1710
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "案件名稱："
      Height          =   240
      Left            =   30
      TabIndex        =   36
      Top             =   1080
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      Height          =   240
      Left            =   30
      TabIndex        =   35
      Top             =   690
      Width           =   900
   End
End
Attribute VB_Name = "frm060120"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/10/07 改成Form2.0 ; txtPA(index)、Label27(index)
'Create by Lydia 2018/02/01 客戶提供文件
Option Explicit

Public m_strSaveFiles As String
Private Const FLmax As Integer = 150 '限制檔案名稱長度
Dim intWhere As Integer
'Modfied by Lydia 2023/06/14 改抓專利基本檔和新案進度檔
'Dim pa(1 To 4) As String '本所案號
Dim pa() As String, cp() As String
Dim m_PA08 As String '專利種類
Dim m_PA09 As String '申請國家
Dim m_PA150 As String '工程師組別
Dim m_CP09 As String  '新增-D類收文號

Dim oText As Control
Dim oLabel As Control
Dim oCheck As CheckBox
Dim intJ As Integer
Dim strDesc(1 To 11) As String '項目名稱
Dim strDescType(1 To 4) As String '副檔名
Dim tmpArr As Variant
Dim bUpdENSP As Boolean, bUpdCNSP As Boolean 'Added by Lydia 2018/02/12 是否覆蓋\\English_vers的檔案
Dim newCP09 As String, newCP10 As String 'Added by Lydia 2018/03/07 新申請案進度的收文號和案件性質
Dim m_MinNP08 As String 'Added by Lydia 2021/02/22 對應之智慧局期限 'Modified by Lydia 2021/07/27 改成所限: m_MinNP09=>m_MinNP08
'Add By Sindy 2022/6/15
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Public m_strCP01 As String, m_strCP02 As String, m_strCP03 As String, m_strCP04 As String
Public m_RDate As String
Dim m_Done As Boolean
Dim m_PrevForm As Form
'2022/6/15 END
'Added by Lydia 2022/08/02
Dim m_bolFMP As Boolean '是否為FMP案
'Modified by Morgan 2025/1/21 P案程序人員工作改依智權區域分配，FMP目前設定仍為98012，程式還是改用函數抓以免將來有變動
'Private Const strToFMP As String = "98012" 'FMP案處理人員
Dim strToFMP As String
'end 2025/1/21
'Added by Lydia 2023/06/14
Dim m_TCT01 As String '命名記錄之收文號=新申請案進度的收文號
Dim m_TCT04 As String, m_TCT05 As String '命名記錄之工程師主管、主管確認日期
Dim bolMailTcn13 As Boolean '無對應英/中說=>直接發email

'Add By Sindy 2022/6/15
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Sub cmdExit_Click()
   'Mark by Lydia 2018/02/02
   'If cmdOK(0).Visible = True Then
   If cmdOK(0).Enabled = True Then
      If CheckInput = True Then
         If MsgBox("你並未存檔，確定離開嗎?", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
            Exit Sub
         End If
      End If
   End If
   
   Unload Me
End Sub

Private Function CheckToDir(ByVal nInx As Integer, Optional ByRef cInA As Integer) As Boolean
Dim strToDir As String
   
   CheckToDir = False
   'Modified by Lydia 2018/05/09 +系統別
   'strToDir = Pub_GetFCPcaseFilePath(txtPA(2), , txtPA(1)) 'Remove by Lydia 2021/12/06 (109/4/6)已將\\Typing2的"English_Vers"和"專利案件"的案件資料夾，全部搬到原始檔區
    'Added by Lydia 2020/01/20 專利案件和English_Vers檔案：判斷檔案上傳目的地，外文本->原始檔區
    If cmdOK(1).Tag = "" Then
      If PUB_ChkCPExist(pa, cntEnglish_Vers, , strExc(1), , "D") = True Then 'English_Vers992
          cmdOK(1).Caption = "原始檔"
          cmdOK(1).Tag = strExc(1)
      End If
    End If
    'end 2020/01/20
    Select Case nInx
         Case 0, 14
              cInA = 14
              'Added by Lydia 2020/01/20 檢查原始檔區
              If InStr(cmdOK(1).Caption, "原始檔") > 0 Then
                    If ChkCPFisExists(cmdOK(1).Tag, strDescType(1), "2") = True Then
                         MsgBox txtPA(1) & "-" & txtPA(2) & "在〔原始檔區〕的已存在原文說明書(*" & strDescType(1) & ") !", vbInformation, "稽核-" & strDesc(1) & "(*" & strDescType(1) & ")"
                         Exit Function
                    End If
              Else
              'end 2020/01/20
                    If Dir(strToDir & "\*" & strDescType(1)) <> "" Then
                         MsgBox txtPA(1) & "-" & txtPA(2) & "在" & strToDir & "的資料夾已存在原文說明書(*" & strDescType(1) & ") !", vbInformation, "稽核-" & strDesc(1) & "(*" & strDescType(1) & ")"
                         Exit Function
                    End If
              End If 'Added by Lydia 2020/01/20
         Case 1, 16
              cInA = 16
              'Added by Lydia 2020/01/20 檢查原始檔區
              If InStr(cmdOK(1).Caption, "原始檔") > 0 Then
                    If ChkCPFisExists(cmdOK(1).Tag, strDescType(1), "2") = False Then
                         MsgBox txtPA(1) & "-" & txtPA(2) & "在〔原始檔區〕的不存在原文說明書(*" & strDescType(1) & ") !", vbInformation, "稽核-" & strDesc(1) & "(*" & strDescType(1) & ")"
                         Exit Function
                    End If
              Else
              'end 2020/01/20
                    ''Modified by Lydia 2018/05/09 +系統別
                    'Remove by Lydia 2021/12/06 (109/4/6)已將\\Typing2的"English_Vers"和"專利案件"的案件資料夾，全部搬到原始檔區
                    'strExc(1) = Pub_GetFCPcaseFilePath(txtPA(2), , txtPA(1))
                    'If Dir(strToDir & "\*" & strDescType(1)) = "" Then
                    '     MsgBox txtPA(1) & "-" & txtPA(2) & "在" & strToDir & "的資料夾不存在原文說明書(*" & strDescType(1) & ") !", vbInformation, "稽核-" & strDesc(2) & "(*" & strDescType(2) & ")"
                    '     Exit Function
                    'End If
                    'end 2021/12/06
              End If 'Added by Lydia 2020/01/20
         Case 2, 18
              cInA = 18
              'Added by Lydia 2020/01/20 檢查原始檔區
              If InStr(cmdOK(1).Caption, "原始檔") > 0 Then
                    If ChkCPFisExists(cmdOK(1).Tag, strDescType(3), "2") = True Then
                         If bUpdENSP = False Then
                            If MsgBox(txtPA(1) & "-" & txtPA(2) & "在〔原始檔區〕已存在" & strDesc(3) & "(*" & strDescType(3) & ") ，是否覆蓋檔案？", vbInformation + vbYesNo + vbDefaultButton2, "稽核-" & strDesc(3) & "(*" & strDescType(3) & ")") = vbNo Then
                                 Exit Function
                            Else
                                 bUpdENSP = True
                            End If
                         End If
                    End If
              Else
              'end 2020/01/20
                    If Dir(strToDir & "\*" & strDescType(3)) <> "" Then
                         'Modified by Lydia 2018/02/12 改成是否覆蓋 (因為僅供參考,所以只要保留最新版)
                         'MsgBox txtPA(1) & "-" & txtPA(2) & "在" & strToDir & "的資料夾已存在" & strDesc(3) & "(*" & strDescType(3) & ") !", vbInformation, "稽核-" & strDesc(3) & "(*" & strDescType(3) & ")"
                         If bUpdENSP = False Then
                              If MsgBox(txtPA(1) & "-" & txtPA(2) & "在" & strToDir & "的資料夾已存在" & strDesc(3) & "(*" & strDescType(3) & ") ，是否覆蓋檔案？", vbInformation + vbYesNo + vbDefaultButton2, "稽核-" & strDesc(3) & "(*" & strDescType(3) & ")") = vbNo Then
                                  Exit Function
                              Else
                                   bUpdENSP = True
                              End If
                         End If
                    End If
              End If 'Added by Lydia 2020/01/20
         Case 3, 20
              cInA = 20
              'Added by Lydia 2020/01/20 檢查原始檔區
              If InStr(cmdOK(1).Caption, "原始檔") > 0 Then
                    'Modified by Lydia 2024/12/30 debug ; strDescType(3)->strDescType(4)
                    If ChkCPFisExists(cmdOK(1).Tag, strDescType(4), "2") = True Then
                         If bUpdCNSP = False Then
                            If MsgBox(txtPA(1) & "-" & txtPA(2) & "在〔原始檔區〕已存在" & strDesc(4) & "(*" & strDescType(4) & ") ，是否覆蓋檔案？", vbInformation + vbYesNo + vbDefaultButton2, "稽核-" & strDesc(4) & "(*" & strDescType(4) & ")") = vbNo Then
                                 Exit Function
                            Else
                                 bUpdCNSP = True
                            End If
                         End If
                    End If
              Else
              'end 2020/01/20
                    If Dir(strToDir & "\*" & strDescType(4)) <> "" Then
                         'Modified by Lydia 2018/02/12 改成是否覆蓋 (因為僅供參考,所以只要保留最新版)
                         'MsgBox txtPA(1) & "-" & txtPA(2) & "在" & strToDir & "的資料夾已存在" & strDesc(4) & "(*" & strDescType(4) & ") !", vbInformation, "稽核-" & strDesc(4) & "(*" & strDescType(4) & ")"
                         If bUpdCNSP = False Then
                              If MsgBox(txtPA(1) & "-" & txtPA(2) & "在" & strToDir & "的資料夾已存在" & strDesc(4) & "(*" & strDescType(4) & ") ，是否覆蓋檔案？", vbInformation + vbYesNo + vbDefaultButton2, "稽核-" & strDesc(4) & "(*" & strDescType(4) & ")") = vbNo Then
                                  Exit Function
                              Else
                                   bUpdCNSP = True
                              End If
                         End If
                    End If
              End If
    End Select
   
   CheckToDir = True
End Function

'上傳檔案-來源
Private Sub cmdFile_Click(Index As Integer)
Dim stFileName As String
Dim sFile
Dim errTxt As String
Dim backErr As String

On Error GoTo ErrHnd
    
   'Modified by Lydia 2023/06/14 Trim(txtPA(5) & txtPA(6) & txtPA(7)) => Trim(Combo1.Tag)
   If txtPA(1) & txtPA(2) & txtPA(3) & txtPA(4) <> pa(1) & pa(2) & pa(3) & pa(4) _
                  Or Trim(pa(1) & pa(2) & pa(3) & pa(4)) = "" _
                  Or Trim(Combo1.Tag) = "" Then
           MsgBox "請先執行尋找本所案號資料 !", vbCritical
           Exit Sub
   End If
   
   'Added by Lydia 2023/06/14
   If textTCN13.Text = "0" Or textTCN13.Tag = "0" Then
      strExc(9) = ""
      If Trim(txtCSD(18)) <> "" Then strExc(9) = strExc(9) & ", 英說(參考/翻譯用)"
      If Trim(txtCSD(20)) <> "" Then strExc(9) = strExc(9) & ", 簡(繁)體中說"
      If strExc(9) <> "" Then
         MsgBox "外文本的對應英/中說已設定為無文件， " & vbCrLf & "又上傳" & Mid(strExc(9), 2), vbCritical
         Exit Sub
      End If
   End If
   'end 2023/06/14
   
   If CheckToDir(Index, intJ) = False Then Exit Sub
         
   'Modified by Lydia 2018/02/26 改讀取combo
   'Me.m_strSaveFiles = txtCSD(intJ).Text
   strExc(1) = ""
   If CmbFL(Index).ListCount > 0 Then
       For intI = 0 To CmbFL(Index).ListCount - 1
          strExc(1) = strExc(1) & IIf(strExc(1) <> "", "&", "") & CmbFL(Index).List(intI)
       Next intI
   End If
   Me.m_strSaveFiles = strExc(1)
   'end 2018/02/26
   Call frm090801_8.SetParent(Me, FLmax, True, "上傳檔案-" & strDesc(Index + 1) & "(*" & strDescType(Index + 1) & ")")
   frm090801_8.m_strSaveFiles = Me.m_strSaveFiles
   frm090801_8.lblCaseNo = ""
   frm090801_8.lblCaseNo.Visible = False '避開檔名要有案號的檢查
   frm090801_8.Label1.Visible = True
   frm090801_8.Label1.Caption = "本所案號： " & pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4)
   frm090801_8.Label4.Visible = False
   frm090801_8.bolNotPDF = True
   frm090801_8.Show vbModal
   
   backErr = cmdFile(Index).Tag 'Added by Lydia 2018/02/26 保留前一次檢查結果
   cmdFile(Index).Tag = ""
   If txtCSD(intJ).Text <> Me.m_strSaveFiles Then
       Call SetCmbList(True, Index, intJ, Me.m_strSaveFiles, errTxt)
       If errTxt <> "" Then
            MsgBox errTxt, vbCritical, "稽核-" & strDesc(Index + 1) & "(*" & strDescType(Index + 1) & ")"
            cmdFile(Index).Tag = errTxt
            cmdFile(Index).SetFocus
       End If
   'Added by Lydia 2018/02/26 保留前一次檢查結果
   Else
       cmdFile(Index).Tag = backErr
   'end 2018/02/26
   End If
   If txtCSD(intJ).Text <> "" And Chk1(Index).Value = vbUnchecked Then
      Chk1(Index).Value = vbChecked
   End If
   
   Exit Sub
ErrHnd:
   If Err.Number <> 32755 Then
      MsgBox Err.Description
   End If
End Sub

Private Sub SetCmbList(ByVal bolSet As Boolean, ByVal inK As Integer, ByVal inJ As Integer, ByVal strFlist As String, ByRef ChkErr As String)
'bolSet　是否重置Combo和Textbox
'inK       Button的Index => 項目再+1
'inJ        Textbox的Index
'strFlist   傳入檔案串列
'ChkErr  回傳錯誤訊息
Dim tmpBol As Boolean
Dim strT1 As String
Dim intA As Integer
Dim chkFile As String
     
     If bolSet = True Then
         CmbFL(inK).Clear
         txtCSD(inJ).Text = ""
     End If
     ChkErr = ""
     tmpBol = False
     If strFlist <> "" Then
         chkFile = strDescType(inK + 1)
         tmpArr = Empty
         tmpArr = Split(strFlist, "&")
         For intA = 0 To UBound(tmpArr)
              strT1 = "" & tmpArr(intA)
              If Trim(strT1) <> "" Then
                   If bolSet = True Then
                        CmbFL(inK).AddItem strT1, intA
                        If InStrRev(strT1, " (") > 0 Then
                            strT1 = Left(strT1, InStrRev(strT1, " (") - 1)
                        End If
                        txtCSD(inJ).Text = txtCSD(inJ).Text & strT1 & "&"
                        If chkFile <> "" Then
                            '是否符合檔案名稱
                            If tmpBol = False Then
                                 If Right(UCase(strT1), Len(chkFile)) = UCase(chkFile) Then
                                     tmpBol = True
                                 End If
                            End If
                            'Added by Lydia 2020/01/20 限制單個檔名長度
                            strExc(1) = GetFileName(strT1)
                            If GetTextLength(strExc(1)) >= 75 Then
                               ChkErr = ChkErr & IIf(ChkErr <> "", vbCrLf, "") & strDesc(inK + 1) & "的檔名超過75字元，請檢查 !" & vbCrLf & strExc(1)
                            End If
                            
                            '檢查是否有其他檔案
                            For intI = 1 To 4
                                 If strDescType(intI) <> chkFile Then
                                     If Right(UCase(strT1), Len(strDescType(intI))) = UCase(strDescType(intI)) Then
                                         ChkErr = ChkErr & IIf(ChkErr <> "", vbCrLf, "") & "上傳檔案不該有" & strDesc(intI) & "(*" & strDescType(intI) & ") !"
                                     End If
                                 End If
                            Next intI
                        End If
                   End If
              End If
         Next intA

         If bolSet = True Then
             If tmpBol = False Then
                ChkErr = ChkErr & IIf(ChkErr <> "", vbCrLf, "") & "上傳檔案沒有" & strDesc(inK + 1) & "(*" & UCase(chkFile) & ") !"
             End If
             CmbFL(inK).ListIndex = 0
             txtCSD(inJ).Text = Mid(txtCSD(inJ).Text, 1, Len(txtCSD(inJ)) - 1)
         End If
     End If
     
End Sub

Private Sub ChkCmbList(ByVal inX As Integer, ByVal inE As Integer, ByRef iPreList As String, ByRef ChkErr As String)
'inX       項目的index
'inE        Textbox的Index
'iPreList 已讀取過的檔案名稱
'ChkErr  回傳錯誤訊息
Dim strT1 As String
Dim intA As Integer
Dim chkFile As String
Dim TempList As String

    ChkErr = ""
    chkFile = strDescType(inX)
    tmpArr = Empty
    tmpArr = Split(txtCSD(inE).Text, "&")
    For intA = 0 To UBound(tmpArr)
         strT1 = "" & tmpArr(intA)
         If Trim(strT1) <> "" Then
                '檢查檔案是否存在
                If Dir(strT1) = "" Then
                    ChkErr = ChkErr & IIf(ChkErr <> "", vbCrLf, "") & strDesc(inX) & "的下列檔案路徑不正確，請檢查 !" & vbCrLf & strT1
                End If
                '去除路徑的檔名
                strExc(1) = strT1
                If InStrRev(strExc(1), " (") > 0 Then strExc(1) = Left(strExc(1), InStrRev(strExc(1), " (") - 1)
                If InStrRev(strExc(1), "\") > 0 Then strExc(1) = Mid(strExc(1), InStrRev(strExc(1), "\") + 1)
                'Added by Lydia 2020/01/20 限制單個檔名長度
                If GetTextLength(strExc(1)) >= 75 Then
                   ChkErr = ChkErr & IIf(ChkErr <> "", vbCrLf, "") & strDesc(inX) & "的檔名超過75字元，請檢查 !" & vbCrLf & strExc(1)
                End If
                
                '檔案重覆
                'Modified by Lydia 2018/03/06 ";"改成"&"
                If InStr(iPreList & "&" & iPreList, strExc(1)) > 0 Then
                    ChkErr = ChkErr & IIf(ChkErr <> "", vbCrLf, "") & strDesc(inX) & "的下列檔案與之前的上傳檔名重覆，請檢查 !" & vbCrLf & strT1
                End If
                TempList = TempList & strExc(1) & "&"
                'end 2018/03/06
         End If
    Next intA
    TempList = Mid(TempList, 1, Len(TempList) - 1)
    If GetTextLength(TempList) > FLmax Then
       ChkErr = ChkErr & IIf(ChkErr <> "", vbCrLf, "") & "去除檔案路徑的檔案名稱總長度超過" & FLmax & "字元 !"
    End If
     'Modified by Lydia 2018/03/06 ";"改成"&"
    iPreList = iPreList & IIf(iPreList <> "", "&", "") & TempList
End Sub

Private Sub cmdok_Click(Index As Integer)
Dim hLocalFile As Long 'Added by Lydia 2018/06/21
    
   'Added by Lydia 2020/02/26 先檢查
   If Index = 0 Or Index = 2 Then  '確定+卷宗區: 檢查卷宗區
        If PUB_CheckFormExist("frm100101_L") Then
            MsgBox "請先關閉共同查詢〔卷宗區〕畫面！"
            Exit Sub
        End If
   ElseIf Index = 1 Then
       If InStr(cmdOK(Index).Caption, "原始檔") > 0 Then
            If PUB_CheckFormExist("frm100101_M") Then
                MsgBox "請先關閉共同查詢〔原始檔區〕畫面！"
                Exit Sub
            End If
       End If
   End If
   'end 2020/02/26
   
   Select Case Index
      Case 0 '確定
         'Added by Lydia 2021/10/07 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字;因為後面有Me.Enabled = False 會造成無法判斷欄位，所以放在這裡
         If PUB_ChkUniText(Me, , True, "TextBox") = False Then
             Exit Sub
         End If
         'end 2021/10/07
         
         'Add By Sindy 2022/6/15
         If m_strIR01 <> "" Then
            If m_strCP01 & m_strCP02 & m_strCP03 & m_strCP04 <> txtPA(1) & txtPA(2) & txtPA(3) & txtPA(4) Then
               MsgBox "信件輸入必須與信件本所案號(" & m_strCP01 & m_strCP02 & m_strCP03 & m_strCP04 & ")一致！"
               Exit Sub
            End If
            If PUB_ChkFileOpening2(Forms(0).Tmpfrm04010519.m_strFullFileName, "後續才能一併歸卷！") = True Then
               Exit Sub
            End If
         End If
         '2022/6/15 END
         
         'Added by Lydia 2023/02/24 外專新案認領
         If textTCN16.Visible = True And textTCN16 = "Y" Then
             'Modified by Lydia 2024/04/18 加註動作
             'If MsgBox("若文件已齊備，是否取消「新案暫不認領」？", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
             If MsgBox("若文件已齊備，是否取消「新案暫不認領」？" & vbCrLf & vbCrLf & "選「是」回到輸入畫面，手動清除「新案暫不認領」=Y；" & vbCrLf & "選「否」繼續存檔作業。", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
                textTCN16.SetFocus
                Call textTCN16_GotFocus
                Exit Sub
             End If
         End If
         'end 2023/02/24
         'Added by Lydia 2023/06/14 避免案號查詢到存檔之間的資料已更新
         If strSrvDate(1) >= 外專新案認領啟用日 And m_TCT01 <> "" Then
            'Modified by Lydia 2023/06/17 改成模組
            m_TCT04 = GetTCT04(m_TCT01, m_TCT05)
            
            If InStr("1,2", textTCN13.Tag) > 0 And m_TCT04 <> "" And m_TCT05 = "" And (txtCSD(18) <> "" Or txtCSD(20) <> "") Then
               MsgBox "請先等工程師完成現階段命名作業!", vbExclamation, "非英說案件"
               Exit Sub
            End If
         End If
         'end 2023/06/14
         
         'Added by Morgan 2025/1/25
         If m_bolFMP And strSrvDate(1) >= P業務區劃分啟用日 Then
            strToFMP = PUB_GetPHandler(pa(1) & pa(2) & pa(3) & pa(4))
         Else
            strToFMP = "98012"
         End If
         'end 2025/1/25
         
         Screen.MousePointer = vbHourglass
         Me.Enabled = False 'Added by Lydia 2018/10/05 有使用者連續按2下確定
         If CheckInput = False Then
             MsgBox "尚未輸入資料 !", vbCritical
             GoTo JumpDefault
         End If
         '重新檢查欄位有效性
         If TxtValidate = False Then GoTo JumpDefault
         
         '上傳檔案
         If MoveEngVersFile = False Then GoTo JumpDefault
         
         'Me.Enabled = False 'Mark by Lydia 2018/10/05
         If FormSave = False Then
              MsgBox "存檔失敗，請洽系統管理員 !", vbCritical
              Me.Enabled = True
              GoTo JumpDefault
         End If
         'Me.Enabled = True  'Mark by Lydia 2018/10/05
         
         '發Email通知
         If GetAutoEmail = False Then
             MsgBox "發信失敗，請改人工通知 !", vbCritical
         End If
         PUB_SendMailCache 'Added by Lydia 2023/05/03
         
         Me.Enabled = True  'Added by Lydia 2018/10/05
         SSTab1.Tab = 0
         '跳到卷宗區
          'Mark by Lydia 2018/02/02 保留移動
         'cmdOK(0).Visible = False
         'cmdOK(2).Visible = True '卷宗區
         cmdOK(0).Enabled = False
         Frame1.Enabled = False
         Frame2.Enabled = False
         cmdTCN13.Enabled = False 'Added by Lydia 2023/06/14
         'Add By Sindy 2022/6/15
         If Me.m_strIR01 <> "" Then
            Screen.MousePointer = vbDefault
            If Not m_PrevForm Is Nothing Then
               Call m_PrevForm.GoNext
            End If
            Unload Me
            Exit Sub
         Else
         '2022/6/15 END
            Call cmdok_Click(2)
         End If
         
      Case 1 '外文本
'Modified by Lydia 2018/03/23 無權限的錯誤要改訊息
'On Error Resume Next
On Error GoTo ErrHand01
         If Len(txtPA(2)) < 6 Then
            MsgBox "請輸入本所案號 !"
            txtPA(2).SetFocus
            txtPA_GotFocus 2
            Exit Sub
         Else
            'Remove by Lydia 2018/03/23
            'If Pub_StrUserSt03 <> "M51" And Left(Pub_StrUserSt03, 1) <> "F" Then
            '      If MsgBox("非國外部人員無權限進入\\English_Vers，是否繼續開啟？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
            '           Exit Sub
            '      End If
            'End If
            'end 2018/03/23
            
            'Added by Lydia 2020/01/20 開啟[原始檔區]
            If InStr(cmdOK(Index).Caption, "原始檔") > 0 Then
                If cmdOK(Index).Tag = "" Then
                    MsgBox txtPA(1) & "-" & txtPA(2) & "在〔原始檔區〕的English_Vers收文號不存在!", vbInformation
                Else
                    frm100101_M.m_strKey = cmdOK(Index).Tag '多筆總收文號
                    frm100101_M.SetParent Me
                    If frm100101_M.QueryData = True Then
                       frm100101_M.Show
                       Me.Hide
                    End If
                End If
            Else
            'end 2020/01/20
                'Modified by Lydia 2018/05/09 +系統別
                'Remove by Lydia 2021/12/06 (109/4/6)已將\\Typing2的"English_Vers"和"專利案件"的案件資料夾，全部搬到原始檔區
                'strExc(1) = Pub_GetFCPcaseFilePath(txtPA(2), , txtPA(1))
                'If Dir(strExc(1) & "\*.*") <> "" Then
                ''     'Modified by Lydia 2018/06/21 用檔案總管開啟放置1~2分鐘後,檔案總管會出錯(ex. A2037, A4041)
                 '    'SHELL "Explorer.exe " & strExc(1), vbNormalFocus  '開啟案件資料夾
                 '    ShellExecute hLocalFile, "explore", strExc(1), vbNullString, vbNullString, 1
               ' Else
               '      MsgBox txtPA(1) & "-" & txtPA(2) & "在" & strExc(1) & "的資料夾不存在或無檔案!", vbInformation
               ' End If
                'end 2021/12/06
            End If 'Added by Lydia 2020/01/20
         End If
         
      Case 2 '卷宗區
            If cmdOK(Index).Visible = True And txtPA(2) <> "" And txtPA(1) & txtPA(2) & txtPA(3) & txtPA(4) = pa(1) & pa(2) & pa(3) & pa(4) Then
                Me.Enabled = False
                Screen.MousePointer = vbHourglass
                frm100101_L.m_strKey = pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4)
                frm100101_L.SetParent Me
                If frm100101_L.QueryData = True Then
                   frm100101_L.Show
                   Me.Hide
                End If
                Screen.MousePointer = vbDefault
                Me.Enabled = True
            End If
   End Select
   
JumpDefault:
   Me.Enabled = True  'Added by Lydia 2018/10/05
   Screen.MousePointer = vbDefault
   
   Exit Sub

'Added by Lydia 2018/03/23
ErrHand01:
    If Err.Number <> 0 Then
         '全部錯誤訊息統一
         'Modified by Lydia 2018/05/09 +系統別
         'Modified by Lydia 2021/12/06 統一
         'MsgBox "無法讀取" & Pub_GetFCPcaseFilePath(pa(2), , pa(1)) & "，請通知電腦中心！", vbCritical
         If Index = 1 Or Index = 2 Then
              MsgBox "無法讀取" & IIf(Index = 1, "外文本", "卷宗區") & "，請通知電腦中心！", vbCritical
         End If
         'end 2021/12/06
         Resume Next
    End If
'end 2018/03/23
End Sub

Private Sub cmdFind_Click()
Dim tmpBol As Boolean 'Added by Lydia 2022/06/21

 SSTab1.Tab = 0
   
   If txtPA(1) = "" Or Len(txtPA(2)) <> 6 Then
      MsgBox "本所案號輸入錯誤，請重新輸入 !", vbCritical
      txtPA(1).SetFocus
      Exit Sub
   End If
   If txtPA(3) = "" Then txtPA(3) = "0"
   If txtPA(4) = "" Then txtPA(4) = "00"
   m_bolFMP = False  'Added by Lydia 2022/08/02
   'Added by Lydia 2022/06/21 外專後續案收文，請開放P的寰華案也可以操作
   If txtPA(1) = "P" Then
      'Modified by Lydia 2022/07/27 開放P(非寰華案)=FMP案的收文權限
      'tmpBol = PUB_FMPtoCheck(1, 2, Pub_strUserST05, txtPA(1), txtPA(2), txtPA(3), txtPA(4))
      tmpBol = PUB_ChkIsFMP(txtPA(1), txtPA(2), txtPA(3), txtPA(4))
      If tmpBol = False Then
          'Modified by Lydia 2022/07/27 開放P(非寰華案)=FMP案的收文權限
          'MsgBox "只可收文寰華案！"
          MsgBox "只可收文寰華案／FMP案！"
          txtPA(2).SetFocus
          Exit Sub
      End If
      'Added by Lydia 2022/08/02
      tmpBol = PUB_FMPtoCheck(1, 2, Pub_strUserST05, txtPA(1), txtPA(2), txtPA(3), txtPA(4))
      If tmpBol = False Then
          m_bolFMP = True
      End If
      'end 2022/08/02
   End If
   'end 2022/06/21
     
   pa(1) = txtPA(1):   pa(2) = txtPA(2)
   pa(3) = txtPA(3):   pa(4) = txtPA(4)
   
   FormClear
   
   txtPA(1) = pa(1):     txtPA(2) = pa(2)
   txtPA(3) = pa(3):     txtPA(4) = pa(4)
   
   '客戶名稱:中->英->日 ; 代理人名稱: 英->中->日
   'Modified by Lydia 2019/11/01 +PA27~PA30
   strExc(0) = "SELECT PA01,PA02,PA03,PA04,PA05,PA06,PA07,PA08,PA09,PA150,PA26,PA75," & _
                     " NVL(CU04,NVL(CU05,CU06)) CNAME,NVL(FA05,NVL(FA04,FA06)) FNAME,PA27,PA28,PA29,PA30" & _
                     " FROM PATENT,CUSTOMER,FAGENT" & _
                     " WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4)) & _
                     " AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+)" & _
                     " AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+)"
    intI = 0
    Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
    If intI = 1 Then
        'Added by Lydia 2019/11/01 利益衝突案件：逐案號判斷
        If strSrvDate(1) >= XY特殊權限啟用日 And XY特殊權限範圍 <> "" Then
            If PUB_ChkCufaByCase(Me.Name, pa(1), pa(1) & pa(2) & pa(3) & pa(4), "" & RsTemp.Fields("PA26") & "," & RsTemp.Fields("PA27") & "," & RsTemp.Fields("PA28") & "," & RsTemp.Fields("PA29") & "," & RsTemp.Fields("PA30"), "" & RsTemp.Fields("PA75")) = False Then
                MsgBox "本案為限閱案件！", vbInformation, MsgText(1110)
                txtPA(2).SetFocus
                txtPA_GotFocus 2
                GoTo JumpToExit
            End If
        End If
        'end 2019/11/01
        
        'Modified by Lydia 2023/06/14
        'txtPA(5) = "" & RsTemp.Fields("PA05")
        'txtPA(6) = "" & RsTemp.Fields("PA06")
        'txtPA(7) = "" & RsTemp.Fields("PA07")
        AddCboName Combo1, "" & RsTemp.Fields("PA05"), "" & RsTemp.Fields("PA06"), "" & RsTemp.Fields("PA07")
        Combo1.Tag = "" & RsTemp.Fields("PA05") & RsTemp.Fields("PA06") & RsTemp.Fields("PA07")
        'end 2023/06/14
        m_PA08 = "" & RsTemp.Fields("PA08")
        m_PA09 = "" & RsTemp.Fields("PA09")
        m_PA150 = "" & RsTemp.Fields("PA150")
        txtPA(26).Text = "" & RsTemp.Fields("PA26")
        txtPA(75).Text = "" & RsTemp.Fields("PA75")
        Label27(26).Caption = "" & RsTemp.Fields("CNAME")
        Label27(75).Caption = "" & RsTemp.Fields("FNAME")
        cmdOK(0).Enabled = True
         'Mark by Lydia 2018/02/02 保留移動
        'cmdOK(0).Visible = True
        'cmdOK(2).Visible = False '卷宗區
        cmdOK(2).Enabled = True
        cmdOK(1).Enabled = True
        'Added by Lydia 2018/03/07 抓新申請案進度
        'Modified by Lydia 2023/06/14 併入下面
        'strExc(0) = "select cp09,cp10 from caseprogress where " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
        '                  " and cp31='Y' "
        'intI = 0
        'Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
        'If intI = 1 Then
        '     newCP09 = "" & RsTemp.Fields("cp09")
        '     newCP10 = "" & RsTemp.Fields("cp10")
        'End If
        ''end 2018/03/07
        
        'Added by Lydia 2020/01/20 專利案件和English_Vers檔案：判斷檔案上傳目的地，外文本->原始檔區
        If PUB_ChkCPExist(pa, cntEnglish_Vers, , strExc(1), , "D") = True Then 'English_Vers992
            cmdOK(1).Caption = "原始檔"
            cmdOK(1).Tag = strExc(1)
        End If
        If strSrvDate(1) >= XY特殊權限啟用日by檔案 Then
           cmdOK(1).Caption = "原始檔"
        End If
        'end 2020/01/20
    End If
    
   'Added by Lydia 2023/02/24 外專新案認領：暫不認領
   lblTCN16.Visible = False: textTCN16.Visible = False
   'FraTCN13.Visible = False 'Added by Lydia 2023/06/14
   cmdTCN13.Visible = False: cmdTCN13.Enabled = True  'Added by Lydia 2023/05/24
   'If strSrvDate(1) >= 外專新案認領啟用日 Then 'Mark by Lydia 2023/06/14
       'Modified by Lydia 2023/06/14 + tct05,tcn13,tcn23
       strExc(0) = "select cp09,cp10,tct01,tct04,tct05,tcn16,tcn13,tcn23 from caseprogress,trackingcasename,transcasetitle " & _
                "where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' " & _
                "and cp31='Y' and cp09=tct01(+) and cp09=tcn05(+) "
       intI = 1
       Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
       If intI = 1 Then
           '暫不認領TCN16=Y , 同時新案未分案TCT04=null
           If "" & RsTemp.Fields("tcn16") = "Y" And "" & RsTemp.Fields("tct01") <> "" And "" & RsTemp.Fields("tct04") = "" Then
               lblTCN16.Visible = True: textTCN16.Visible = True
               textTCN16.Text = "" & RsTemp.Fields("tcn16")
               textTCN16.Tag = textTCN16.Text
           End If
           'Added by Lydia 2023/06/14 外文本的對應英/中說
           newCP09 = "" & RsTemp.Fields("cp09")
           newCP10 = "" & RsTemp.Fields("cp10")
           m_TCT01 = "" & RsTemp.Fields("tct01")  '有命名記錄
           m_TCT04 = "" & RsTemp.Fields("tct04")
           m_TCT05 = "" & RsTemp.Fields("tct05")
           textTCN13.Text = "" & RsTemp.Fields("tcn13")
           textTCN13.Tag = "" & RsTemp.Fields("tcn13")
           '待處理增加按鈕，排除處於急件認領TCN23=0的狀態
           If "" & RsTemp.Fields("tcn13") = "2" And ("" & RsTemp.Fields("tcn23") <> "0" Or m_TCT04 <> "") Then
              cmdTCN13.Visible = True
           End If
           'end 2023/06/14
       End If
   'End If
   'end 2023/02/24
   
   'Added by Lydia 2023/06/14
   If pa(1) = "P" Then
       intWhere = 國內
   Else
       intWhere = 國外_FC
   End If
   If PUB_ReadPatentDatabase(pa, intWhere, False) Then
   End If
   cp(9) = newCP09
   If PUB_ReadCaseProgressDatabase(cp(), intWhere, False) Then
   End If
   'end 2023/06/14
   
JumpToExit: 'Added by Lydia 2019/11/01

End Sub

Private Sub Form_Activate()
   'Added by Sindy 2022/6/25
   If m_strIR01 <> "" And m_Done = False Then
      txtPA(1) = m_strCP01
      txtPA(2) = m_strCP02
      txtPA(3) = m_strCP03
      txtPA(4) = m_strCP04
      m_Done = True
      Me.cmdFind.Value = True
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2022/6/25 END
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = 國外_FC
   FormClear
   SSTab1.Tab = 0
   SendKeys "{Tab}"

   txtPA(26).BackColor = &H8000000F
   txtPA(75).BackColor = &H8000000F
   Frame1.BackColor = &H8000000F
   Frame2.BackColor = &H8000000F
   
   'Mark by Lydia 2018/02/02 保留移動
   'cmdOK(2).Top = cmdOK(0).Top
   'cmdOK(2).Left = cmdOK(0).Left
   '項目名稱
   strDesc(1) = "原文說明書"
   strDescType(1) = ".ORI.PDF"
   strDesc(2) = "替換版原文說明書"
   strDescType(2) = ".ORI.REP.PDF"
   strDesc(3) = "英說(參考/翻譯用)"
   'Modified by Lydia 2018/02/12 Elaine說改.msg, 這樣才可以修改附件內容
   'strDescType(3) = ".ENSP.PDF"
   strDescType(3) = ".ENSP.MSG"
   strDesc(4) = "簡(繁)體中說"
   'Modified by Lydia 2018/02/12 Elaine說改.msg, 這樣才可以修改附件內容
   'strDescType(4) = ".CNSP.PDF"
   strDescType(4) = ".CNSP.MSG"
   strDesc(5) = "優先權證明書"
   strDesc(6) = "委任狀"
   strDesc(7) = "代表人資訊"
   strDesc(8) = "申請人資訊"
   strDesc(9) = "發明人資訊"
   strDesc(10) = "非WTO會員國之住所證明"
   strDesc(11) = "其他"
   
   '隱藏路徑Textbox
   'If Pub_StrUserSt03 <> "M51" Then
        txtCSD(14).Visible = False
        txtCSD(16).Visible = False
        txtCSD(18).Visible = False
        txtCSD(20).Visible = False
   'End If
   'Added by Lydia 2023/06/14
   cmdTCN13.Visible = False
   ReDim pa(TF_PA) As String
   ReDim cp(TF_CP) As String
   'end 2023/06/14
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Sindy 2022/6/15
   If Not m_PrevForm Is Nothing Then
      Set m_PrevForm = Nothing
   End If
   '2022/6/15 END
   PUB_SendMailCache 'Added by Lydia 2023/02/24
   
   Set frm060120 = Nothing
End Sub

Private Function FormSave() As Boolean
   Dim strCon1 As String, strCon2 As String
   Dim strCP64 As String
   Dim strAddMenu As String 'Added by Lydia 2018/03/06
   
On Error GoTo CheckingErr
   cnnConnection.BeginTrans
   
   '新增D類-客戶提供文件1920: 期限=系統日+2工作天(不含當日)且承辦人掛程序
   m_CP09 = AutoNo("D", 6)
   'Added by Lydia 2022/08/02 判斷FMP案
   If m_bolFMP = True Then
      strExc(1) = strToFMP
   Else
   'end 2022/08/02
      strExc(1) = PUB_GetFCPHandler(pa(1), pa(2), pa(3), pa(4))
   End If 'Added by Lydia 2022/08/02
   strExc(6) = CompWorkDay(3, strSrvDate(1))
   
   'Added by Lydia 2021/02/22 抓補文件之最早法限：特定文字設定參考basPublic.PUB_SetCombo202
   strExc(5) = ""
   For intJ = 21 To 31 Step 2
      If Chk1(intJ).Value = vbChecked Then
         Select Case intJ
             Case 21 '優先權證明書
                  strExc(5) = strExc(5) & " or instr(np15,'優先權') > 0 "
             Case 23 '委任狀
                  strExc(5) = strExc(5) & " or instr(np15,'委任書') > 0 "
             Case 25 '代表人資訊
                  strExc(5) = strExc(5) & " or instr(np15,'代表人') > 0 "
             Case 27 '申請人資訊
                  strExc(5) = strExc(5) & " or instr(np15,'申請人') > 0 "
             Case 29 '發明人資訊
                  strExc(5) = strExc(5) & " or instr(np15,'發明人') > 0 "
             Case 31  '非WTO會員國之住所證明
                  strExc(5) = strExc(5) & " or instr(np15,'國籍證明') > 0 "
         End Select
      End If
   Next intJ
   If strExc(5) <> "" Then
        'Modified by Lydia 2021/07/27 np09=>np08
        strSql = "select min(np08) ndate from nextprogress where np02='" & pa(1) & "' and np03='" & pa(2) & "' and np04='" & pa(3) & "' and np05='" & pa(4) & "' and np06 is null and np07='202' and (" & Mid(strExc(5), 4) & ") "
        intI = 1
        Set RsTemp = ClsLawReadRstMsg(intI, strSql)
        If intI = 1 Then
            m_MinNP08 = "" & RsTemp.Fields("ndate")
            If m_MinNP08 <> "" Then
                '急件：若 +3個工作天期限大於智慧局期限，則以智慧局期限(所限)
                If m_MinNP08 < strSrvDate(1) Then  '期限超過，當天完成
                    m_MinNP08 = strSrvDate(1)
                Else
                    If strExc(6) <= m_MinNP08 Then
                        m_MinNP08 = ""
                    End If
                End If
            End If
        End If
   End If
   If m_MinNP08 <> "" Then strExc(6) = m_MinNP08
   'end 2021/02/22
   'Added by Lydia 2021/07/27 「說明書」頁籤1-4項以中說收文的本所期限為智慧局期限
   If Chk1(0).Value = 1 Or Chk1(1).Value = 1 Or Chk1(2).Value = 1 Or Chk1(3).Value = 1 Then
        strSql = "select cp09,cp06,cp07 from caseprogress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp159=0 and cp10 in ('201','209','210','242','235') "
        intI = 1
        Set RsTemp = ClsLawReadRstMsg(intI, strSql)
        If intI = 1 Then
            m_MinNP08 = "" & RsTemp.Fields("cp06")
            If m_MinNP08 <> "" Then
                '急件：若 +3個工作天期限大於智慧局期限，則以智慧局期限(所限)
                If m_MinNP08 < strSrvDate(1) Then  '期限超過，當天完成
                    m_MinNP08 = strSrvDate(1)
                Else
                    If strExc(6) <= m_MinNP08 Then
                        m_MinNP08 = ""
                    End If
                End If
                If m_MinNP08 <> "" Then strExc(6) = m_MinNP08
            End If
        End If
   End If
   'end 2021/07/27
   
   'Modified by Lydia 2021/07/27 改成「承辦期限+所限」
   'strSql = "insert into caseprogress( cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp09,cp10,cp12,cp13,cp14 ) " & _
                    "values ('" & pa(1) & "', '" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'," & strSrvDate(1) & "," & strExc(6) & "," & strExc(6) & _
                     ",'" & m_CP09 & "','1920','" & Pub_StrUserSt03 & "','" & strUserNum & "','" & strExc(1) & "')"
   strSql = "insert into caseprogress( cp01,cp02,cp03,cp04,cp05,cp06,cp48,cp09,cp10,cp12,cp13,cp14 ) " & _
                    "values ('" & pa(1) & "', '" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'," & strSrvDate(1) & "," & strExc(6) & "," & strExc(6) & _
                     ",'" & m_CP09 & "','1920','" & Pub_StrUserSt03 & "','" & strUserNum & "','" & strExc(1) & "')"
   cnnConnection.Execute strSql, intI
   
   'Added by Lydia 2018/03/06 新增卷宗區(CASE.Menu)
   strExc(1) = Format(ServerTime, "000000")
   strSql = "insert into casepaperpdf(cpp01,cpp02,cpp03,CPP05,CPP06,CPP07,cpp08,cpp09,cpp10)" & _
        " values('" & m_CP09 & "','" & m_CP09 & "." & FCP提供文件 & "',0,'" & strUserNum & "'," & strSrvDate(1) & "," & strExc(1) & "," & strSrvDate(1) & "," & strExc(1) & ",'Y')"
   cnnConnection.Execute strSql, intI
   'end 2018/03/06
          
   strCon1 = "insert into CustSupportDoc(CSD01 , CSD02 , CSD03 , CSD04 , CSD05 , CSD06 , CSD07 , CSD08"
   strCon2 = "'" & pa(1) & "' , '" & pa(2) & "' , '" & pa(3) & "' , '" & pa(4) & "' , '" & m_CP09 & "' , '" & strUserNum & "' , " & strSrvDate(1) & " , " & Left(Format(ServerTime, "000000"), 4) & " "
   
   strCP64 = ""
   For intJ = 13 To 34
         strExc(2) = ""
         strCon1 = strCon1 & " , CSD" & Format(intJ, "00")
         Select Case intJ
               Case 13, 15, 17, 19, 22, 24, 26, 28, 30, 32, 34 '備註
                     If IsEmpty(txtCSD(intJ).Text) = True Then
                          strCon2 = strCon2 & ", NULL "
                     Else
                          strCon2 = strCon2 & ", '" & ChgSQL(txtCSD(intJ).Text) & "' "
                     End If
               Case 14, 16, 18, 20 '檔案名稱
                     If txtCSD(intJ).Tag = "" Then
                          strCon2 = strCon2 & ", NULL "
                     Else
                          strCon2 = strCon2 & ", '" & ChgSQL(txtCSD(intJ).Tag) & "' "
                          Select Case intJ
                               Case 14: strExc(2) = strDesc(1)
                               Case 16: strExc(2) = strDesc(2)
                               Case 18: strExc(2) = strDesc(3)
                               Case 20: strExc(2) = strDesc(4)
                          End Select
                          If strExc(2) <> "" Then strCP64 = strCP64 & strExc(2) & ";"
                     End If
               Case 21, 23, 25, 27, 29, 31, 33 'Check項目
                     If Chk1(intJ).Value = vbUnchecked Then
                          strCon2 = strCon2 & ", NULL "
                     Else
                          strCon2 = strCon2 & ", 'Y' "
                          Select Case intJ
                               Case 21: strExc(2) = strDesc(5)
                               Case 23: strExc(2) = strDesc(6)
                               Case 25: strExc(2) = strDesc(7)
                               Case 27: strExc(2) = strDesc(8)
                               Case 29: strExc(2) = strDesc(9)
                               Case 31: strExc(2) = strDesc(10)
                               Case 33: strExc(2) = strDesc(11)
                          End Select
                          If strExc(2) <> "" Then strCP64 = strCP64 & strExc(2) & ";"
                     End If
         End Select
   Next intJ
   
   strSql = strCon1 & ") values (" & strCon2 & ")"
   
   cnnConnection.Execute strSql, intI
   If strCP64 <> "" Then
       strSql = "update caseprogress set cp64='" & strCP64 & "' where cp09='" & m_CP09 & "' "
       cnnConnection.Execute strSql, intI
   End If
   
   'Added by Lydia 2020/10/14 Murgitroyd呈送期限設定: 檢視中說 （代理人提供簡體中說）: 收到簡體中說21日內完成中說並報告。
   If Chk1(3).Value = 1 Then
      strExc(0) = Pub_GetSpecMan("外專MURGITROYD設定")
      If strExc(0) <> "" And InStr(strExc(0), txtPA(75)) > 0 And txtPA(75) <> "" Then
          'Modified by Lydia 2021/01/06 (109/12/28)檢視中說改為收到簡體中說+14個日曆天再往前推3個工作天
          'strExc(1) = CompWorkDay(4, CompDate(2, 21, strSrvDate(1)), 1)
          strExc(1) = CompWorkDay(4, CompDate(2, 14, strSrvDate(1)), 1)
          '自動帶至進度備註：為Murgitroyd案需xx月xx日（收到簡體中說＋14個日曆天再往前推3個工作天）完成中說並報告
          strExc(4) = "為Murgitroyd案需" & ChangeWStringToTDateString(strExc(1)) & "完成中說並報告"
          'Added by Lydia 2022/08/02 判斷FMP案
          If m_bolFMP = True Then
               strExc(3) = strToFMP
          Else
          'end 2022/08/02
              strExc(3) = PUB_GetFCPHandler(pa(1), pa(2), pa(3), pa(4))
          End If 'Added by Lydia 2022/08/02
          'Modified by Lydia 2021/05/20 在行事曆事由增加[解除管制不通知]，排除" 解除人員非建立行事曆人員會發email通知建立人員"行事曆已被解除管制"
          If PUB_AddFCPStaffCalendar(strExc(1), "1", strExc(3), strExc(4) & "[解除管制不通知]", strExc(3), "1", pa(1), pa(2), pa(3), pa(4)) = True Then
              'Modified by Lydia 2021/01/06 (109/12/28)承辦期限設為收到簡體中說+14個日曆天再往前推3個工作天
              'strSql = "Update CaseProgress set cp64=" & CNULL(strExc(4)) & "||';'||cp64 where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp10='209' and cp159=0 "
              'cnnConnection.Execute strSql
              If PUB_ChkCPExist(pa, "209", 1, strExc(9)) = True Then
                   strSql = "Update CaseProgress set cp64=" & CNULL(strExc(4)) & "||';'||cp64, cp48='" & strExc(1) & "' where cp09='" & strExc(9) & "' "
                   cnnConnection.Execute strSql
                   '更新文件齊備日
                   strSql = "Update Engineerprogress Set EP06='" & strSrvDate(1) & "' where ep02='" & strExc(9) & "' "
                   cnnConnection.Execute strSql
              End If
              'end 2021/01/06
          End If
      End If
   End If
   'end 2020/10/14
   
   'Add by Sindy 2022/6/15
   If m_strIR01 <> "" Then
      PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm060120", m_CP09
   End If
   '2022/6/15 END
   
   'Added by Lydia 2023/02/24 外專新案認領
   'Modified by Lydia 2023/06/14 外專新案認領：暫不認領＋非英說案之確收
   'If textTCN16.Visible = True And m_TCT01 <> "" Then
   '   If textTCN16.Tag = "Y" And textTCN16.Text = "" Then
   '       strSql = "Update TrackingCaseName Set TCN16=null Where TCN05='" & m_TCT01 & "' "
   '       cnnConnection.Execute strSql
   '       If PUB_UpdateTCNstate(IIf(textTCN16.Tag = "Y", "1", "2"), pa(1) & pa(2) & pa(3) & pa(4)) = False Then
   '          GoTo CheckingErr
   '       End If
   '   End If
   'End If
   ''end 2023/02/24
   If m_TCT01 <> "" Then
      strExc(6) = ""
      If textTCN16.Tag <> textTCN16.Text And textTCN16.Visible = True Then
         strExc(6) = strExc(6) & ", TCN16=" & CNULL(textTCN16)
      End If
      If textTCN13.Tag <> "" Then
         '1=有文件, 2=待確定
         If (textTCN13.Tag = "1" Or textTCN13.Tag = "2") And Trim(txtCSD(18)) <> "" Or Trim(txtCSD(20)) <> "" Then
            textTCN13.Text = "3" '確定已收文件
         End If
         If textTCN13.Tag <> textTCN13.Text Then
            strExc(6) = strExc(6) & ", TCN13=" & CNULL(textTCN13)
         End If
      End If
      If strExc(6) <> "" Then
          strExc(7) = ""
          strSql = "Update TrackingCaseName Set " & Mid(strExc(6), 2) & " Where TCN05='" & m_TCT01 & "' "
          cnnConnection.Execute strSql
          'Added by Lydia 2024/11/15 客戶提供文件寫入TCN13=3確定已收文件，同時記錄收件日
          If textTCN13.Text = "3" And textTCN13.Tag <> textTCN13.Text Then
             strSql = "Update TrackingCaseName Set TCN26=" & strSrvDate(1) & " Where TCN05='" & m_TCT01 & "' AND NVL(TCN26,0)=0 "
             cnnConnection.Execute strSql
          End If
          'end 2024/11/15
          'Added by Lydia 2025/01/17 暫不認領TCN16=Y的取消日期記錄在命名記錄TransCaseTitle.TCT121,TCT122
          If textTCN16.Text = "" And textTCN16.Tag <> textTCN16.Text Then
             strSql = "Update TransCaseTitle Set TCT121=TO_CHAR(SYSDATE,'YYYYMMDD'), TCT122=TO_CHAR(SYSDATE,'HH24MI') Where TCT01='" & m_TCT01 & "' "
             cnnConnection.Execute strSql
          End If
          'end 2025/01/17
          
          '暫不認領=>進入認領階段
          If InStr(UCase(strExc(6)), "TCN16") > 0 And textTCN16.Tag = "Y" Then
            If PUB_UpdateTCNstate("1", pa(1) & pa(2) & pa(3) & pa(4)) = False Then
               GoTo CheckingErr
            End If
          Else
            If textTCN13.Tag <> textTCN13.Text Then
               If m_TCT04 = "" Then  '認領階段中間
                  strExc(1) = PUB_GetEngGrpMan(strExc(2))
                  strExc(3) = PUB_GetTCNmTitle(pa(1), pa(2), pa(3), pa(4), pa(10), textTCN13.Text, "")
                  strExc(3) = strExc(3) & "，請協助確認組別，謝謝！"
                  strExc(9) = ""
                  If Trim(txtCSD(18)) <> "" Then strExc(9) = strExc(9) & ", 英說(參考/翻譯用)"
                  If Trim(txtCSD(20)) <> "" Then strExc(9) = strExc(9) & ", 簡(繁)體中說"
                  If textTCN13.Text = "3" Then
                     strExc(4) = "已收" & Mid(strExc(9), 2) & "，請至原始檔區查看文件。" & vbCrLf & vbCrLf
                  Else
                     strExc(4) = "確定無外文本的對應英/中說。" & vbCrLf & vbCrLf
                  End If
                  strExc(4) = strExc(4) & "代理人：" & txtPA(75) & " " & Label27(75).Caption & vbCrLf & _
                                 "申請人：" & txtPA(26) & " " & Label27(26).Caption
                  strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                                " values( '" & strUserNum & "','" & strExc(1) & "',to_char(sysdate,'yyyymmdd')" & _
                               ",to_char(sysdate,'hh24miss'),'" & strExc(3) & "','" & ChgSQL(strExc(4)) & "',null)"
                  cnnConnection.Execute strSql
               Else
                  If textTCN13.Tag <> "0" And bolMailTcn13 = False Then '排除已執行「確定無對應英/中說」
                     '原本設定為"有or 待確定"+第一次認領階段完成=>判斷是否進入第二認領
                     If PUB_UpdateReTCN(pa, cp, True) = True Then
                     End If
                  End If
               End If
            End If
          End If
      End If
   End If
   'end 2023/06/14
   
   cnnConnection.CommitTrans
   FormSave = True

CheckingErr:

   If Err.Number <> 0 Then
      cnnConnection.RollbackTrans
      MsgBox Err.Description, vbCritical
   End If
   
End Function

' 清除資料表
Private Sub FormClear()

   For Each oText In txtPA
      oText.Text = ""
   Next

   For Each oText In txtCSD
      oText.Text = ""
      oText.Tag = ""
      oText.Locked = False
   Next
   
   For Each oLabel In Label27
      oLabel.Caption = ""
   Next
   
   For Each oCheck In Chk1
      oCheck.Value = vbUnchecked
   Next
   
   'Added by Lydia 2018/02/12
   bUpdENSP = False
   bUpdCNSP = False
   'end 2018/02/12
   
   CmbFL(0).Clear
   CmbFL(1).Clear
   CmbFL(2).Clear
   CmbFL(3).Clear
   
   txtPA(1) = strSysKind

   cmdOK(0).Enabled = False
    'Mark by Lydia 2018/02/02 保留移動
   'cmdOK(0).Visible = True
   cmdOK(1).Enabled = False
   cmdOK(1).Tag = "" 'Added by Lydia 2020/01/20
   
    'Mark by Lydia 2018/02/02 保留移動
   'cmdOK(2).Visible = False '卷宗區
   cmdOK(2).Enabled = False
   Frame1.Enabled = True
   Frame2.Enabled = True
   m_CP09 = ""
   
   cmdFile(0).Enabled = True
   cmdFile(1).Enabled = True
   cmdFile(2).Enabled = True
   cmdFile(3).Enabled = True
   'Added by Lydia 2018/03/07
   newCP09 = ""
   newCP10 = ""
   m_MinNP08 = "" 'Added by Lydia 2021/02/22
   
   'Added by Lydia 2023/02/24 外專新案認領：暫不認領
   textTCN16.Text = ""
   textTCN16.Tag = ""
   'Added by Lydia 2023/06/14 外專新案認領：外文本的對應英/中說
   textTCN13.Text = ""
   textTCN13.Tag = ""
   m_TCT01 = ""
   m_TCT04 = ""
   m_TCT05 = ""
   bolMailTcn13 = False
End Sub

Private Function TxtValidate() As Boolean
Dim Cancel As Boolean
Dim intP As Integer
Dim strMsg As String, strList As String

   TxtValidate = False
   
   Cancel = False
   txtPA_Validate 1, Cancel
   If Cancel = True Then
       Exit Function
   End If
    
    'Added by Lydia 2018/04/24 重抓工程師組別(ex.FCP58620承辦先查詢，之後才請程序改組別，沒有重新查詢，發給舊的工程師主管。)
    strSql = "SELECT PA150 FROM PATENT " & _
                     " WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
    intI = 1
    Set RsTemp = ClsLawReadRstMsg(intI, strSql)
    If intI = 1 Then
         m_PA150 = "" & RsTemp.Fields("pa150")
    End If
    'Modified by Lydia 2023/03/03：外專新案認領：判斷「暫不認領」 lblTCN16.Visible = False
    If lblTCN16.Visible = False And m_PA150 = "" And (txtCSD(15).Text <> "" Or CmbFL(1).ListCount > 0) Then
        MsgBox "工程師組別目前為退程序，請通知程序到新案建檔設定工程師組別後，再重新執行尋找本所案號資料！", vbCritical
        Exit Function
    End If
    'end 2018/04/24
    
   If txtPA(2) = "" Or txtPA(3) = "" Or txtPA(4) = "" Then
      MsgBox "請輸入本所案號 !", vbCritical
      Exit Function
   'Modified by Lydia 2023/06/14 Trim(txtPA(5) & txtPA(6) & txtPA(7)) => Trim(Combo1.Tag)
   ElseIf txtPA(1) & txtPA(2) & txtPA(3) & txtPA(4) <> pa(1) & pa(2) & pa(3) & pa(4) _
                  Or Trim(pa(1) & pa(2) & pa(3) & pa(4)) = "" _
                  Or Trim(Combo1.Tag) = "" Then
      MsgBox "請先執行尋找本所案號資料 !", vbCritical
      Exit Function
   End If

   '檢查輸入資料
   Cancel = False
   For Each oText In txtCSD
        txtCSD_Validate oText.Index, Cancel
        If Cancel = True Then
            Exit Function
        End If
   Next
   
   '檢查檔案是否存在,是否正確,是否重覆
   For intI = 0 To 3
       If cmdFile(intI).Tag <> "" Then
            MsgBox cmdFile(intI).Tag, vbCritical, strDesc(intI + 1)
            If cmdFile(intI).Enabled = True Then 'Added by Lydia 2023/06/17
               cmdFile(intI).SetFocus
            End If 'Added by Lydia 2023/06/17
            Exit Function
       End If
   Next intI
   intI = 0
   strList = ""
   For intJ = 14 To 20 Step 2
       If txtCSD(intJ).Text <> "" Then
            If CheckToDir(intI) = False Then
                 cmdFile(intI).SetFocus
                 Exit Function
            End If
       End If
        If Trim(txtCSD(intJ)) <> "" Then
             strMsg = ""
             Call ChkCmbList(intI + 1, intJ, strList, strMsg)
             If strMsg <> "" Then
                  MsgBox strMsg, vbCritical, "稽核-" & strDesc(intI + 1) & "(*" & strDescType(intI + 1) & ")"
                  If cmdFile(intI).Enabled = True Then 'Added by Lydia 2023/06/17
                     cmdFile(intI).SetFocus
                  End If  'Added by Lydia 2023/06/17
                  Exit Function
             End If
        ElseIf Trim(txtCSD(intJ - 1)) <> "" Then
                  MsgBox "未上傳" & strDesc(intI + 1) & "，備註不可輸入 !", vbCritical
                  If txtCSD(intJ - 1).Enabled = True Then 'Added by Lydia 2023/06/17 ex.FCP-069729沒有上傳檔案,但是有寫備註
                     txtCSD(intJ - 1).SetFocus
                     txtCSD_GotFocus intJ - 1
                  End If 'Added by Lydia 2023/06/17
                  Exit Function
        End If
        intI = intI + 1
   Next intJ
   
   TxtValidate = True
End Function

Private Function CheckInput() As Boolean
     CheckInput = False
     
     For Each oText In txtCSD
          If Trim(oText.Text) <> "" Then
               CheckInput = True
               Exit Function
          End If
     Next
     For Each oCheck In Chk1
          If oCheck.Value = vbChecked Then
               CheckInput = True
               Exit Function
          End If
     Next
End Function

Private Sub txtCSD_GotFocus(Index As Integer)
   TextInverse txtCSD(Index)
   CloseIme
End Sub

'Remove by Lydia 2021/10/07
'Private Sub txtCSD_KeyPress(Index As Integer, KeyAscii As Integer)
'   'Modified by Lydia 2018/03/12 因為會輸入密碼,所以不限大寫
'   'KeyAscii = UpperCase(KeyAscii)
'End Sub
'end 2021/10/07

Private Sub txtCSD_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   If txtCSD(Index).Text <> "" Then
       txtCSD(Index).ToolTipText = PUB_StringFilter(txtCSD(Index).Text)
   End If
End Sub

Private Sub txtCSD_Validate(Index As Integer, Cancel As Boolean)
Dim iMax As Integer
    Select Case Index
         Case 13, 14
              If txtCSD(Index).Text <> "" And Chk1(0).Value = vbUnchecked Then
                 Chk1(0).Value = vbChecked
              End If
         Case 15, 16
              If txtCSD(Index).Text <> "" And Chk1(1).Value = vbUnchecked Then
                 Chk1(1).Value = vbChecked
              End If
         Case 17, 18
              If txtCSD(Index).Text <> "" And Chk1(2).Value = vbUnchecked Then
                 Chk1(2).Value = vbChecked
              End If
         Case 19, 20
              If txtCSD(Index).Text <> "" And Chk1(3).Value = vbUnchecked Then
                 Chk1(3).Value = vbChecked
              End If
         Case 22, 24, 26, 28, 30, 32, 34
              If txtCSD(Index).Text <> "" And Chk1(Index - 1).Value = vbUnchecked Then
                 Chk1(Index - 1).Value = vbChecked
              End If
    End Select

    If InStr("14,16,18,20", Format(Index, "00")) > 0 Then '檔案名稱改用模組判斷
    Else
        If CheckLengthIsOK(txtCSD(Index).Text, txtCSD(Index).MaxLength) = False Then
             Cancel = True
             If txtCSD(Index).Enabled = True Then 'Added by Lydia 2023/06/17
                txtCSD(Index).SetFocus
                txtCSD_GotFocus Index
             End If   'Added by Lydia 2023/06/17
             Exit Sub
        End If
    End If
    Exit Sub
   
End Sub

Private Sub txtPA_GotFocus(Index As Integer)
   TextInverse txtPA(Index)
   CloseIme
End Sub

'Modified by Lydia 2021/10/07 改成Form 2.0
'Private Sub txtPA_KeyPress(Index As Integer, KeyAscii As Integer)
Private Sub txtPA_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtPA_Validate(Index As Integer, Cancel As Boolean)
    If Index = 1 Then
        'Modified by Lydia 2022/06/21 +P案(FMP案)
        If txtPA(Index) <> "FCP" And txtPA(Index) <> "P" Then
            MsgBox "系統類別錯誤，請重新輸入 !", vbCritical
            txtPA(Index).SetFocus
            Cancel = True
        End If
    End If
End Sub

'上傳檔案
Private Function MoveEngVersFile() As Boolean
Dim strNewName As String, strOldName As String
Dim strFileType As String, strTemp As String
Dim strToDir As String
Dim strErrList As String
Dim inX As Integer
Dim inK As Integer
Dim strErr1 As String
'Added by Lydia 2020/01/20
Dim nCP09 As String
Dim nFileName As String

On Error GoTo FileErrHand

      MoveEngVersFile = True
      If txtCSD(14) & txtCSD(16) & txtCSD(18) & txtCSD(20) = "" Then Exit Function
      
      MoveEngVersFile = False
      
      'Remove by Lydia 2021/12/06 (109/4/6)已將\\Typing2的"English_Vers"和"專利案件"的案件資料夾，全部搬到原始檔區
'      If InStr(cmdOK(1).Caption, "原始檔") = 0 Then  'Added by Lydia 2020/01/20 +判斷
'        'Modified by Lydia 2018/05/09 +系統別
'        strToDir = Pub_GetFCPcaseFilePath(txtPA(2), , txtPA(1))
'        '建立English_vers資料夾(前3碼)
'        If txtPA(1) = "FCP" Then 'Added by Lydia 2018/05/09 限FCP案
'          If Dir(Pub_GetFCPcaseFilePath(txtPA(2), True), vbDirectory) = "" Then
'               MkDir Pub_GetFCPcaseFilePath(txtPA(2), True)
'          End If
'        End If 'end 2018/05/09
'        '建立English_vers資料夾
'        If Dir(strToDir, vbDirectory) = "" Then
'             MkDir strToDir
'        End If
'      End If 'Added by Lydia 2020/01/20
      'end 2021/12/06
      
      intI = 0
      For intJ = 14 To 20 Step 2
          intI = intI + 1
          strErr1 = ""
          If txtCSD(intJ).Text <> "" And txtCSD(intJ).Tag = "" Then
              tmpArr = Empty
              tmpArr = Split(txtCSD(intJ).Text, "&")
              For inK = 0 To UBound(tmpArr)
                   If Trim(tmpArr(inK)) <> "" Then
                        If Dir(tmpArr(inK)) <> "" Then
                             strOldName = tmpArr(inK)
                             'Added by Lydia 2020/01/20 專利案件和English_Vers檔案：上傳到原始檔區
                             'Modified by Lydia 2020/03/18 +原始檔
                             'If strSrvDate(1) >= XY特殊權限啟用日by檔案 Then
                             If strSrvDate(1) >= XY特殊權限啟用日by檔案 Or InStr(cmdOK(1).Caption, "原始檔") > 0 Then
                                nCP09 = cmdOK(1).Tag
                                strNewName = strOldName
                                'English_Vers992 : 預設承辦人2=操作者,除了替換本自動+流水號(A),若有重覆檔案直接刪除(D)
                                If PUB_UploadCPFfile("2", strNewName, pa(1), pa(2), pa(3), pa(4), cntEnglish_Vers, nCP09, , IIf(Right(UCase(strOldName), Len(".ORI.REP.PDF")) = ".ORI.REP.PDF", "A", "D"), False, strExc(6), nFileName) = True Then
                                    txtCSD(intJ).Tag = txtCSD(intJ).Tag & IIf(txtCSD(intJ).Tag <> "", "&", "") & nFileName '記錄檔名
                                    'Added by Lydia 2023/01/04 FCP-068521客戶提供文件清單有&&，在處理作業會程式出錯
                                    strExc(1) = Replace(txtCSD(intJ).Tag, "&&", "&")
                                    txtCSD(intJ).Tag = strExc(1)
                                    'end 2023/01/04
                                    If cmdOK(1).Tag <> nCP09 Then  '變更瀏覽按鈕
                                         cmdOK(1).Caption = "原始檔"
                                         cmdOK(1).Tag = nCP09
                                    End If
                                Else
                                     strErr1 = strErr1 & IIf(strErr1 <> "", vbCrLf, "") & strExc(6)
                                End If
                             Else
                             'end 2020/01/20
                                'Remove by Lydia 2018/03/23 因為來源字串本身不包含" (檔案大小 KB)" ,所以不需截取" ("以前
                                'If InStrRev(strOldName, " (") > 0 Then strOldName = Left(strOldName, InStrRev(strOldName, " (") - 1)
                                If InStrRev(strOldName, "\") > 0 Then strOldName = Mid(strOldName, InStrRev(strOldName, "\") + 1)
                                If InStr(UCase(strOldName), strDescType(intI)) > 0 Then
                                       strTemp = Mid(strOldName, 1, InStr(UCase(strOldName), strDescType(intI)) - 1)  '只抓副檔名之前的名稱
                                       strNewName = strTemp & strDescType(intI)
                                       '替換版原文說明書要自動+流水號
                                       If Right(UCase(strNewName), Len(".ORI.REP.PDF")) = ".ORI.REP.PDF" Then
                                           inX = 1
                                           'Added by Lydia 2018/03/07 替換版上傳後自動更名為案號.副檔名
                                           'Modified by Lydia 2018/03/09 改FCP本所案號轉檔名
                                           strTemp = PUB_FCPCaseNo2FileName(pa(1), pa(2), pa(3), pa(4))
                                       Else
                                           inX = 0
                                           'Added by Lydia 2018/03/07 原文說明書上傳後自動更名
                                           If Right(UCase(strNewName), Len(FcpTcnFKey02)) = FcpTcnFKey02 Then
                                               'Modified by Lydia 2018/03/09 改FCP本所案號轉檔名
                                               'Modified by Lydia 2018/03/16 會議後通知,拿掉ForeignSpec
                                               'strNewName = PUB_FCPCaseNo2FileName(pa(1), pa(2), pa(3), pa(4)) & IIf(newCP10 <> "", "." & newCP10, "") & ".ForeignSpec" & FcpTcnFKey02
                                               strNewName = PUB_FCPCaseNo2FileName(pa(1), pa(2), pa(3), pa(4)) & IIf(newCP10 <> "", "." & newCP10, "") & FcpTcnFKey02
                                           End If
                                           'end 2018/03/07
                                       End If
                                Else
                                       inX = 0
                                       strNewName = strOldName
                                End If
                                
                                If inX > 0 Then
                                    strNewName = strTemp & ".ORI.REP1.PDF"
                                    strExc(1) = Dir(strToDir & "\" & strNewName)
                                    If strExc(1) <> "" Then
                                         inX = 1
                                         Do While strExc(1) <> ""
                                              strNewName = strTemp & ".ORI.REP" & inX + 1 & ".PDF"
                                              strExc(1) = Dir(strToDir & "\" & strNewName)
                                              inX = inX + 1
                                         Loop
                                    End If
                                End If
                                '複製檔案到English_vers資料夾 (若English_vers已有檔案,會直接覆蓋)
                                FileCopy tmpArr(inK), strToDir & "\" & strNewName
                                'Modified by Lydia 2018/03/06 ";"改成"&"
                                txtCSD(intJ).Tag = txtCSD(intJ).Tag & IIf(txtCSD(intJ).Tag <> "", "&", "") & strNewName '記錄檔名
                             End If 'Added by Lydia 2020/01/20
                        Else
                             strErr1 = strErr1 & IIf(strErr1 <> "", vbCrLf, "") & tmpArr(inK)
                        End If
                   End If
              Next inK
              If strErr1 = "" Then
                    '上傳檔案後,先鎖住
                    cmdFile(intI - 1).Enabled = False
              Else
                    strErrList = strErrList & IIf(strErrList <> "", vbCrLf, "") & strDesc(intI) & ":" & vbCrLf & strErr1
              End If
          End If
      Next intJ
      
      If strErrList <> "" Then
        'Added by Lydia 2020/01/20 專利案件和English_Vers檔案：上傳到原始檔區
        If strSrvDate(1) >= XY特殊權限啟用日by檔案 Then
            MsgBox "無法寫入〔原始檔區 〕\English_Vers(" & nCP09 & ")\" & strErrList & "，請通知電腦中心！", vbCritical
        Else
        'end 2020/01/20
            MsgBox "下列檔案路徑不正確，請檢查 !" & strErrList, vbCritical, "複製到\\English_Vers"
        End If 'Added by Lydia 2020/01/20
      Else
           MoveEngVersFile = True
      End If
      
      Exit Function
      
FileErrHand:
     If Err.Number <> 0 Then
        'Added by Lydia 2020/01/20 專利案件和English_Vers檔案：上傳到原始檔區
        If strSrvDate(1) >= XY特殊權限啟用日by檔案 Then
            MsgBox "無法寫入〔原始檔區 〕\English_Vers(" & nCP09 & ")\" & strErrList & "，請通知電腦中心！", vbCritical
        Else
        'end 2020/01/20
           'Modified by Lydia 2018/03/23 全部錯誤訊息統一
           'MsgBox Err.Description
           MsgBox "無法寫入" & strToDir & "，請通知電腦中心！", vbCritical
        End If 'Added by Lydia 2020/01/20
     End If
End Function

'發Email通知相關人員
Private Function GetAutoEmail() As Boolean
Dim stTO As String, stSub As String, stContent As String
Dim m_Sub As String, m_GrpMan As String
Dim m_list As String '合併通知項目
Dim inX As Integer
Dim strCC As String 'Add By Sindy 2022/6/20

GetAutoEmail = False

    'Added by Lydia 2022/08/02 判斷FMP案
    If m_bolFMP = True Then
        stTO = strToFMP
    Else
    'end 2022/08/02
        stTO = PUB_GetFCPHandler(pa(1), pa(2), pa(3), pa(4)) '預設-通知程序
    End If 'Added by Lydia 2022/08/02
    If stTO = "" Then Exit Function
    
   'Modify By Sindy 2022/6/20 + 承辦人的2級主管為副本收件者
   strExc(0) = "SELECT ST52 FROM staff WHERE ST01='" & strUserNum & "' and ST52 is not null"
   intI = 1: strCC = ""
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If Not IsNull(RsTemp.Fields("ST52")) Then
         strCC = RsTemp.Fields("ST52")
      End If
   End If
   '2022/6/20 END
    
On Error GoTo ErrHandle:

    m_Sub = pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) <> "000", "-" & pa(3) & "-" & pa(4), "")
    m_GrpMan = ""
    If m_PA150 <> "" And (txtCSD(15) <> "" Or txtCSD(16).Tag <> "" Or txtCSD(17) <> "" Or txtCSD(18).Tag <> "") Then
      '抓工程師-主管
      'Modified by Lydia 2019/01/19
      m_GrpMan = Pub_GetFCPGrpMan(m_PA150)
      'Added by Lydia 2022/10/12 特殊情況之指定職代
      m_GrpMan = PUB_GetStateForMan(m_GrpMan)
    End If
    
    '01.原文說明書：
    If txtCSD(13) <> "" Or txtCSD(14).Tag <> "" Then
         'Modified by Lydia 2025/01/14 備註移到內文
         'stSub = m_Sub & "客戶提供文件－" & strDesc(1) & "ORI" & IIf(txtCSD(13) <> "", "(備註:" & Trim(txtCSD(13).Text) & ")", "")
         stSub = m_Sub & "客戶提供文件－" & strDesc(1) & "ORI"
         'Added by Lydia 2023/06/14 配合外專新案認領 ---- from Phoebe 5/24 Email
         If strSrvDate(1) >= 外專新案認領啟用日 Then
            If textTCN16.Visible = True And textTCN16.Text = "Y" Then
                stContent = "收原文說明書，維持暫不認領狀態。"
            Else
                stContent = "收原文說明書，系統將請工程師進行認領流程，待收到工程師系統命名完成之信函，" & vbCrLf & _
                                 "即可新案提申或待工程師完成提申前告代(修正)再提申。"
            End If
         Else
         'end 2023/06/14
            stContent = "收原文說明書，請見主旨備註的分組訊息，如無顯示分組訊息，" & vbCrLf & _
                              "待收到組別確認結果 (由承辦確認)，建完組別之後請工程師上系統命名。"
         End If  'Added by Lydia 2023/05/24
         
         'Added by Lydia 2025/01/14 備註移到內文
         If txtCSD(13) <> "" Then stContent = stContent & vbCrLf & vbCrLf & "備註:" & Trim(txtCSD(13).Text)
         
         'Modify By Sindy 2022/6/20 + 副本收件者
         PUB_SendMail strUserNum, stTO, "", stSub, vbCrLf & stContent & vbCrLf, , , , , , strCC
    End If
    '02.替換版原文說明書：
    If txtCSD(15) <> "" Or txtCSD(16).Tag <> "" Then
         tmpArr = Empty
         'Modified by Lydia 2018/03/06 ";"改成"&"
         tmpArr = Split(txtCSD(16).Tag, "&")
         strExc(2) = ""
         For intI = 0 To UBound(tmpArr)
             strExc(1) = UCase("" & tmpArr(intI))
             If strExc(1) <> "" And InStr(strExc(1), ".ORI.REP") > 0 Then
                 strExc(2) = strExc(1)
                 Exit For
             End If
         Next intI
         'Modified by Lydia 2025/01/14 備註移到內文
         'stSub = m_Sub & "客戶提供文件－" & strDesc(2) & "ORI.REP" & IIf(txtCSD(15) <> "", "(備註:" & Trim(txtCSD(15).Text) & ")", "")
         stSub = m_Sub & "客戶提供文件－" & strDesc(2) & "ORI.REP"
         'Modified by Lydia 2018/05/10
         'stContent = "收替換版原文說明書（檔名：" & strExc(2) & "），請上系統重新命名。" '請上系統檢視說明書內容
         'Added by Lydia 2020/01/20 +判斷
         If InStr(cmdOK(1).Caption, "原始檔") > 0 Then
            stContent = "收替換版原文說明書（檔名：" & strExc(2) & "），請上〔原始檔區〕檢視說明書內容。"
         Else
         'end 2020/01/20
            stContent = "收替換版原文說明書（檔名：" & strExc(2) & "），請上系統檢視說明書內容。"
         End If 'Added by Lydia 2020/01/20
         
         'Added by Lydia 2025/01/14 備註移到內文
         If txtCSD(15) <> "" Then stContent = stContent & vbCrLf & vbCrLf & "備註:" & Trim(txtCSD(15).Text)
         
         'Added by Lydia 2018/05/10 收件人+命名人員
         strExc(3) = ""
         strSql = "select tct01,tct10 from caseprogress,transcasetitle where " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                      " and cp31='Y' and cp09=tct01(+) "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
              RsTemp.MoveFirst
              strExc(3) = "" & RsTemp.Fields("tct10")
         End If
         'Modified by Lydia 2018/05/10  收件人+命名人員
         'PUB_SendMail strUserNum, stTO & IIf(m_GrpMan <> "", ";" & m_GrpMan, ""), "", stSub, vbCrLf & stContent & vbCrLf
         'Modify By Sindy 2022/6/20 + 副本收件者
         PUB_SendMail strUserNum, stTO & IIf(m_GrpMan <> "", ";" & m_GrpMan, "") & IIf(strExc(3) <> "", ";" & strExc(3), ""), "", stSub, vbCrLf & stContent & vbCrLf, , , , , , strCC
    End If
    '03.英說(參考/翻譯用)：
    If txtCSD(17) <> "" Or txtCSD(18).Tag <> "" Then
         '工程師(抓承辦人為外專工程師的最大收文號)
         strExc(2) = ""
         'Modified by Lydia 2019/08/02 排除F4102 (FCP年費不續辦)
         strSql = "select cp05,cp09,cp14 from caseprogress,staff where " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                      " and cp159=0 and cp14=st01(+) and cp14<>'F4102' and st03='F21' order by cp05 desc,cp09 desc,cp14"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
              RsTemp.MoveFirst
              strExc(2) = "" & RsTemp.Fields("cp14")
         End If
         'Modified by Lydia 2025/01/14 備註移到內文
         'stSub = m_Sub & "客戶提供文件－" & strDesc(3) & "ENSP" & IIf(txtCSD(17) <> "", "(備註:" & Trim(txtCSD(17).Text) & ")", "")
         stSub = m_Sub & "客戶提供文件－" & strDesc(3) & "ENSP"
         'Added by Lydia 2023/06/17 區別非英說案
         If m_TCT01 <> "" And (textTCN13.Tag = "1" Or textTCN13.Tag = "2") And bolMailTcn13 = False And textTCN13.Tag <> textTCN13.Text Then
            stContent = "1.收英說，供日後參考/翻譯用。"
            'Modified by Lydia 2024/03/25 改成判斷認領階段的主管人數；
            'strExc(1) = GetTCT04(m_TCT01)
            'stContent = stContent & vbCrLf & "2.非英說提申案，" & IIf(m_TCT04 <> "" And strExc(1) = m_TCT04, "不需重新認領", "已【進入認領階段】") & "。"
            stContent = stContent & vbCrLf & "2.非英說提申案，" & GetReTCNstatus(m_TCT01)
            'end 2024/03/25
         Else
         'end 2023/06/17
            stContent = "收英說，供日後參考/翻譯用。"
         End If  'Added by Lydia 2023/06/17
         'Added by Lydia 2024/11/15
         If InStr(stContent, "不需重新認領") > 0 And m_TCT04 <> "" Then
            stContent = stContent & vbCrLf & "3.請工程師重新命名"
         End If
         'end 2024/11/13
         
         'Added by Lydia 2025/01/14 備註移到內文
         If txtCSD(17) <> "" Then stContent = stContent & vbCrLf & vbCrLf & "備註:" & Trim(txtCSD(17).Text)
         
         'Modify By Sindy 2022/6/20 + 副本收件者
         PUB_SendMail strUserNum, stTO & IIf(strExc(2) <> "", ";" & strExc(2), "") & IIf(m_GrpMan <> "", ";" & m_GrpMan, ""), "", stSub, vbCrLf & stContent & vbCrLf, , , , , , strCC
    End If
    
    '04.簡(繁)體中說：
    If txtCSD(19) <> "" Or txtCSD(20).Tag <> "" Then
         'Modified by Lydia 2025/01/14 備註移到內文
         'stSub = m_Sub & "客戶提供文件－" & strDesc(4) & "CNSP" & IIf(txtCSD(19) <> "", "(備註:" & Trim(txtCSD(19).Text) & ")", "")
         stSub = m_Sub & "客戶提供文件－" & strDesc(4) & "CNSP"
         'Added by Lydia 2023/06/17 區別非英說案
         If m_TCT01 <> "" And (textTCN13.Tag = "1" Or textTCN13.Tag = "2") And bolMailTcn13 = False And textTCN13.Tag <> textTCN13.Text Then
            stContent = "1.收簡 (繁) 體中說，請分案檢視中說 /核對中說格式。"
            'Modified by Lydia 2024/03/25 改成判斷認領階段的主管人數；
            'strExc(1) = GetTCT04(m_TCT01)
            'stContent = stContent & vbCrLf & "2.非英說提申案，" & IIf(m_TCT04 <> "" And strExc(1) = m_TCT04, "不需重新認領", "已【進入認領階段】") & "。"
            stContent = stContent & vbCrLf & "2.非英說提申案，" & GetReTCNstatus(m_TCT01)
            'end 2024/03/25
         Else
         'end 2023/06/17
            stContent = "收簡 (繁) 體中說，請分案檢視中說 /核對中說格式。"
         End If 'Added by Lydia 2023/06/17
         'Added by Lydia 2024/11/15
         If InStr(stContent, "不需重新認領") > 0 And m_TCT04 <> "" Then
            stContent = stContent & vbCrLf & "3.請工程師重新命名"
         End If
         'end 2024/11/13
         
         'Added by Lydia 2025/01/14 備註移到內文
         If txtCSD(19) <> "" Then stContent = stContent & vbCrLf & vbCrLf & "備註:" & Trim(txtCSD(19).Text)
         
         'Modify By Sindy 2022/6/20 + 副本收件者
         PUB_SendMail strUserNum, stTO, "", stSub, vbCrLf & stContent & vbCrLf, , , , , , strCC
    End If
    
    '05.~11.備註
    inX = 21 '起始欄位位置
    strExc(2) = ""
    stContent = ""
    m_list = ""
    For intJ = 5 To 11
         If Chk1(inX).Value = vbChecked Or Trim(txtCSD(inX + 1)) <> "" Then
            m_list = m_list & IIf(m_list <> "", ";", "") & IIf(intJ = 10 Or intJ = 11, Mid(strDesc(intJ), 1, 4), Mid(strDesc(intJ), 1, 3))
            stContent = stContent & Format(intJ, "00") & ". " & strDesc(intJ) & "(備註:" & txtCSD(inX + 1) & ")" & vbCrLf
         End If
         inX = inX + 2
    Next
    If m_list <> "" Then
         stSub = m_Sub & "客戶提供文件－補文件＆資訊:" & m_list
         stContent = vbCrLf & "收申請相關資訊/文件 (內容見下列備註/卷宗區)，請補呈。" & vbCrLf & vbCrLf & stContent
         strExc(2) = PUB_GetFCPProSup(stTO)
         'Modified by Lydia 2018/03/07 程序主管改為副本
         'PUB_SendMail strUserNum, stTO & IIf(strExc(2) <> "", ";" & strExc(2), ""), "", stSub, stContent
         'Modified by Lydia 2021/02/22 加註：◎急件！
         'PUB_SendMail strUserNum, stTO, "", stSub, stContent, , , , , , strExc(2)
         'Modified by Lydia 2022/06/08 急件才需要CC給程序主管
         'PUB_SendMail strUserNum, stTo, "", IIf(m_MinNP08 <> "", "◎急件！", "") & stSub, stContent, , , , , , strExc(2)
         'Modify By Sindy 2022/6/20 + 承辦人的2級主管為副本收件者
         PUB_SendMail strUserNum, stTO, "", IIf(m_MinNP08 <> "", "◎急件！", "") & stSub, stContent, , , , , , IIf(m_MinNP08 <> "", strExc(2) & ";" & strCC, strCC)
    End If
    
GetAutoEmail = True
Exit Function

ErrHandle:
     Exit Function
End Function

'Added by Lydia 2018/03/12 跳離開,直接查詢
Private Sub txtPA_LostFocus(Index As Integer)
    If Index = 2 Then
        If txtPA(Index).Text <> "" Then
           If Len(txtPA(Index)) = 6 And txtPA(Index).Text <> pa(2) Then
                 'Call cmdFind_Click 'Remove by Lydia 2018/03/23 按enter自動尋找
           ElseIf Len(txtPA(Index)) <> 6 Then
                 MsgBox "本所案號請輸入6碼!! '"
                 txtPA(Index).SetFocus
                 txtPA_GotFocus Index
           End If
        End If
    End If
End Sub

'Added by Lydia 2020/01/20 檢查原始檔區的檔案是否存在相同檔案
Private Function ChkCPFisExists(ByVal iCp09 As String, Optional ByVal iTxt As String, Optional ByVal iType As String = "1", Optional ByRef outMaxCPF02 As String) As Boolean
Dim intQ As Integer
Dim rsQuery As New ADODB.Recordset
Dim strQ1 As String
    
'iCP09 : 收文號
'iTxt   :查詢的檔名
'iType : 1=相同檔名；2=相似副檔名

    ChkCPFisExists = False
    outMaxCPF02 = ""
    If iCp09 = "" Then Exit Function
    
    strQ1 = "select cpf02 from casepaperfile where cpf01='" & iCp09 & "' "
    If iTxt <> "" Then
        If iType = "1" Then '1=相同檔名
            strQ1 = strQ1 & "and upper(cpf02) ='" & UCase(iTxt) & "' "
        ElseIf iType = "2" Then '2=相似副檔名
            strQ1 = strQ1 & "and upper(cpf02) like '%" & UCase(iTxt) & "' "
        End If
    End If
    strQ1 = strQ1 & "order by cpf08 desc, cpf09 desc "
    intQ = 1
    Set rsQuery = ClsLawReadRstMsg(intQ, strQ1)
    If intQ = 1 Then
        ChkCPFisExists = True
        outMaxCPF02 = "" & rsQuery.Fields("cpf02")
    End If
    Set rsQuery = Nothing
End Function

'Added by Lydia 2023/02/24
Private Sub textTCN16_GotFocus()
   TextInverse textTCN16
End Sub

'Added by Lydia 2023/02/24
Private Sub textTCN16_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
      KeyAscii = 0
      Beep
   End If
End Sub

'Added by Lydia 2023/06/14
Private Sub cmdTCN13_Click()
   
   If txtPA(1) & txtPA(2) & txtPA(3) & txtPA(4) <> pa(1) & pa(2) & pa(3) & pa(4) _
                  Or Trim(pa(1) & pa(2) & pa(3) & pa(4)) = "" _
                  Or Trim(Combo1.Tag) = "" Then
           MsgBox "請先執行尋找本所案號資料 !", vbCritical
           Exit Sub
   End If
   
   If m_TCT04 <> "" And m_TCT05 = "" Then
      MsgBox "請先等工程師完成現階段命名作業!", vbExclamation, "非英說案件"
      Exit Sub
   End If
   strExc(9) = ""
   If Trim(txtCSD(18)) <> "" Then strExc(9) = strExc(9) & ", 英說(參考/翻譯用)"
   If Trim(txtCSD(20)) <> "" Then strExc(9) = strExc(9) & ", 簡(繁)體中說"
   If strExc(9) <> "" Then
      MsgBox "已上傳" & Mid(strExc(9), 2) & "，" & vbCrLf & "請先確定是否要上傳檔案！", vbCritical
      Exit Sub
   End If
   If MsgBox("是否無對應英/中說？", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
      strSql = "Update TrackingCaseName Set TCN13='4' Where TCN05='" & m_TCT01 & "' "
      cnnConnection.Execute strSql
      textTCN13.Text = "4"  '確定無文件
      textTCN13.Tag = textTCN13.Text
      cmdTCN13.Enabled = False
      If PUB_UpdateReTCN(pa, cp, True) = True Then
         PUB_SendMailCache
         bolMailTcn13 = True
         'Added by Lydia 2023/06/17 確定無文件=>比照確收文件,發email通知程序
         If m_TCT01 <> "" Then
            strExc(0) = pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) <> "000", "-" & pa(3) & "-" & pa(4), "") & "客戶提供文件〔非英說案: 確定無參考本〕"
            'Modified by Lydia 2024/03/25 改成判斷認領階段的主管人數；
            'strExc(1) = GetTCT04(m_TCT01)
            'strExc(2) = "非英說提申案，" & IIf(m_TCT04 <> "" And strExc(1) = m_TCT04, "不需重新認領", "已【進入認領階段】") & "。"
            strExc(2) = "非英說提申案，" & GetReTCNstatus(m_TCT01)
            'end 2024/03/25
            strExc(3) = PUB_GetFCPHandler(pa(1), pa(2), pa(3), pa(4))
            If strExc(3) <> "" Then
               PUB_SendMail strUserNum, strExc(3), "", strExc(0), strExc(2)
            End If
         End If
         'end 2023/06/17
      End If
   End If
End Sub

'Added by Lydia 2023/06/17
Private Function GetTCT04(ByVal pTCT01 As String, Optional ByRef pTCT05 As String) As String
Dim strQ As String, intQ As Integer
Dim rsQuery As New ADODB.Recordset
  
   GetTCT04 = ""
   
   If pTCT01 <> "" Then
      strQ = "select tct04, tct05 from transcasetitle where tct01='" & pTCT01 & "' "
      intQ = 1
      Set rsQuery = ClsLawReadRstMsg(intQ, strQ)
      If intQ = 1 Then
         GetTCT04 = "" & rsQuery.Fields("TCT04")
         pTCT05 = "" & rsQuery.Fields("TCT05")
      End If
   End If
   
   Set rsQuery = Nothing
End Function

'Added by Lydia 2024/03/25 判斷認領階段的主管;
'收英文參考本通知Email原本是依照當時主管與現在主管是否相同帶出內文；
'改成依照當時認領主管(=Y)的人數，只有一位主管內文帶出「不需重新認領」，有二位以上則帶出「已【進入認領階段】」。
Private Function GetReTCNstatus(ByVal pTCT01 As String) As String
Dim strQ As String, intQ As Integer
Dim rsQuery As New ADODB.Recordset
     
   GetReTCNstatus = ""
   If pTCT01 <> "" Then
      strQ = "select tcn23,tcn24 from trackingcasename where tcn05='" & pTCT01 & "' "
      intQ = 1
      Set rsQuery = ClsLawReadRstMsg(intQ, strQ)
      If intQ = 1 Then
         If "" & rsQuery.Fields("tcn23") = "9" Or "" & rsQuery.Fields("tcn24") = "Y" Then
            GetReTCNstatus = "不需重新認領" '9=程序人員設定, TCN24=最高主管進行核判
         Else
            strQ = ""
            Select Case "" & rsQuery.Fields("tcn23")
               Case "0", "1", "4" '0=急件認領,1=主管認領
                  strQ = "" & rsQuery.Fields("tcn23")
               Case "2" '2=主管+職代認領
                  strQ = "1"
               Case "3" '協調認領認領
                  strQ = "2"
            End Select
            
            If strQ <> "" Then
               strQ = "select count(*) as cnt, sum(decode(tfa05,'Y',1,0)) as t1 from transfeeassign where tfa01='" & pTCT01 & "' and tfa09='" & strQ & "' "
               intQ = 1
               Set rsQuery = ClsLawReadRstMsg(intQ, strQ)
               If intQ = 1 Then
                  If Val("" & rsQuery.Fields("cnt")) >= FCPforEngNum And Val("" & rsQuery.Fields("t1")) = 1 Then
                     GetReTCNstatus = "不需重新認領"
                  Else
                     GetReTCNstatus = "已【進入認領階段】"
                  End If
               End If
            End If
         End If
      End If
   End If
   
   Set rsQuery = Nothing
End Function
