VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060121_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "客戶提供文件處理"
   ClientHeight    =   6640
   ClientLeft      =   830
   ClientTop       =   980
   ClientWidth     =   8260
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6640
   ScaleWidth      =   8260
   Begin VB.TextBox txtResult 
      Height          =   300
      Left            =   1200
      MaxLength       =   1
      TabIndex        =   0
      Top             =   1725
      Width           =   375
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   960
      TabIndex        =   52
      Text            =   "Combo1"
      Top             =   690
      Width           =   6735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "卷宗區"
      Height          =   375
      Index           =   2
      Left            =   5520
      TabIndex        =   2
      Top             =   120
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "外文本"
      Height          =   375
      Index           =   1
      Left            =   6360
      TabIndex        =   3
      Top             =   120
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   7200
      TabIndex        =   4
      Top             =   120
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   375
      Index           =   0
      Left            =   4680
      TabIndex        =   1
      Top             =   120
      Width           =   800
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4530
      Left            =   90
      TabIndex        =   29
      Top             =   2070
      Width           =   8055
      _ExtentX        =   14199
      _ExtentY        =   7990
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "說明書"
      TabPicture(0)   =   "frm060121_1.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "CmbFL(3)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "CmbFL(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "CmbFL(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "CmbFL(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "補文件 ＆ 資訊"
      TabPicture(1)   =   "frm060121_1.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      Begin VB.ComboBox CmbFL 
         Height          =   300
         Index           =   0
         Left            =   1680
         TabIndex        =   68
         Text            =   "CmbFL"
         Top             =   960
         Width           =   6060
      End
      Begin VB.ComboBox CmbFL 
         Height          =   300
         Index           =   1
         Left            =   1680
         TabIndex        =   67
         Text            =   "CmbFL"
         Top             =   2040
         Width           =   6060
      End
      Begin VB.ComboBox CmbFL 
         Height          =   300
         Index           =   2
         Left            =   1680
         TabIndex        =   66
         Text            =   "CmbFL"
         Top             =   3075
         Width           =   6060
      End
      Begin VB.ComboBox CmbFL 
         Height          =   300
         Index           =   3
         Left            =   1680
         TabIndex        =   65
         Text            =   "CmbFL"
         Top             =   4110
         Width           =   6060
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  '沒有框線
         Caption         =   "Frame2"
         Height          =   3975
         Left            =   -74880
         TabIndex        =   44
         Top             =   400
         Width           =   7815
         Begin VB.CheckBox Chk1 
            Caption         =   "11.其他"
            Height          =   255
            Index           =   33
            Left            =   0
            TabIndex        =   25
            Top             =   3270
            Width           =   1935
         End
         Begin VB.CheckBox Chk1 
            Caption         =   "10.非WTO會員國之住所證明"
            Height          =   375
            Index           =   31
            Left            =   0
            TabIndex        =   23
            Top             =   2725
            Width           =   1695
         End
         Begin VB.CheckBox Chk1 
            Caption         =   "9.發明人資訊"
            Height          =   255
            Index           =   29
            Left            =   0
            TabIndex        =   21
            Top             =   2180
            Width           =   1455
         End
         Begin VB.CheckBox Chk1 
            Caption         =   "8.申請人資訊"
            Height          =   255
            Index           =   27
            Left            =   0
            TabIndex        =   19
            Top             =   1635
            Width           =   1455
         End
         Begin VB.CheckBox Chk1 
            Caption         =   "7.代表人資訊"
            Height          =   255
            Index           =   25
            Left            =   0
            TabIndex        =   17
            Top             =   1090
            Width           =   1455
         End
         Begin VB.CheckBox Chk1 
            Caption         =   "6.委任狀"
            Height          =   255
            Index           =   23
            Left            =   0
            TabIndex        =   15
            Top             =   545
            Width           =   1455
         End
         Begin VB.CheckBox Chk1 
            Caption         =   "5.優先權證明書"
            Height          =   255
            Index           =   21
            Left            =   0
            TabIndex        =   13
            Top             =   0
            Width           =   1695
         End
         Begin MSForms.TextBox txtCSD 
            Height          =   645
            Index           =   34
            Left            =   2520
            TabIndex        =   26
            Top             =   3270
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
            TabIndex        =   24
            Top             =   2725
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
            TabIndex        =   22
            Top             =   2180
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
            TabIndex        =   20
            Top             =   1635
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
            TabIndex        =   18
            Top             =   1090
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
            TabIndex        =   16
            Top             =   545
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
            TabIndex        =   14
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
            TabIndex        =   51
            Top             =   3270
            Width           =   5565
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "(備註：                                                                                                             )"
            Height          =   180
            Index           =   14
            Left            =   1920
            TabIndex        =   50
            Top             =   2725
            Width           =   5565
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "(備註：                                                                                                             )"
            Height          =   180
            Index           =   13
            Left            =   1920
            TabIndex        =   49
            Top             =   2180
            Width           =   5565
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "(備註：                                                                                                             )"
            Height          =   180
            Index           =   11
            Left            =   1920
            TabIndex        =   48
            Top             =   1635
            Width           =   5565
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "(備註：                                                                                                             )"
            Height          =   180
            Index           =   10
            Left            =   1920
            TabIndex        =   47
            Top             =   1090
            Width           =   5565
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "(備註：                                                                                                             )"
            Height          =   180
            Index           =   9
            Left            =   1920
            TabIndex        =   46
            Top             =   545
            Width           =   5565
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "(備註：                                                                                                             )"
            Height          =   180
            Index           =   12
            Left            =   1920
            TabIndex        =   45
            Top             =   0
            Width           =   5565
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  '沒有框線
         Caption         =   "Frame1"
         Height          =   4095
         Left            =   120
         TabIndex        =   30
         Top             =   390
         Width           =   7815
         Begin VB.CheckBox Chk1 
            Caption         =   "4.簡(繁)體中說"
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   11
            Top             =   3120
            Width           =   1935
         End
         Begin VB.CheckBox Chk1 
            Caption         =   "3.英說(參考/翻譯用)"
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   9
            Top             =   2080
            Width           =   1935
         End
         Begin VB.CheckBox Chk1 
            Caption         =   "2.替換版原文說明書"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   7
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
            TabIndex        =   74
            Top             =   420
            Width           =   465
            VariousPropertyBits=   671105051
            Size            =   "1931;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtCSD 
            Height          =   300
            Index           =   16
            Left            =   0
            TabIndex        =   73
            Top             =   1500
            Width           =   465
            VariousPropertyBits=   671105051
            Size            =   "1931;503"
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
            Top             =   2505
            Width           =   465
            VariousPropertyBits=   671105051
            Size            =   "1931;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtCSD 
            Height          =   300
            Index           =   20
            Left            =   0
            TabIndex        =   71
            Top             =   3540
            Width           =   465
            VariousPropertyBits=   671105051
            Size            =   "1931;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtCSD 
            Height          =   300
            Index           =   15
            Left            =   1560
            TabIndex        =   8
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
            Index           =   17
            Left            =   1560
            TabIndex        =   10
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
            Index           =   19
            Left            =   1560
            TabIndex        =   12
            Top             =   3390
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
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "*.ORI.REP2.PDF"
            ForeColor       =   &H000000FF&
            Height          =   180
            Index           =   16
            Left            =   6440
            TabIndex        =   43
            Top             =   1320
            Width           =   1230
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "檔案名稱："
            Height          =   180
            Index           =   3
            Left            =   660
            TabIndex        =   42
            Top             =   3750
            Width           =   900
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "檔案名稱："
            Height          =   180
            Index           =   2
            Left            =   690
            TabIndex        =   41
            Top             =   2715
            Width           =   900
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "檔案名稱："
            Height          =   180
            Index           =   1
            Left            =   660
            TabIndex        =   40
            Top             =   1680
            Width           =   900
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "檔案名稱："
            Height          =   180
            Index           =   0
            Left            =   660
            TabIndex        =   39
            Top             =   600
            Width           =   900
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "(備註：                                                                                        )"
            Height          =   180
            Index           =   1
            Left            =   960
            TabIndex        =   38
            Top             =   300
            Width           =   4620
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "(備註：                                                                                         )"
            Height          =   180
            Index           =   2
            Left            =   960
            TabIndex        =   37
            Top             =   1380
            Width           =   4665
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "(備註：                                                                                        )"
            Height          =   180
            Index           =   3
            Left            =   990
            TabIndex        =   36
            Top             =   2415
            Width           =   4620
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "(備註：                                                                                        )"
            Height          =   180
            Index           =   4
            Left            =   960
            TabIndex        =   35
            Top             =   3450
            Width           =   4620
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "檔名：*.ORI.PDF"
            ForeColor       =   &H000000FF&
            Height          =   180
            Index           =   5
            Left            =   2040
            TabIndex        =   34
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
            TabIndex        =   33
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
            TabIndex        =   32
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
            TabIndex        =   31
            Top             =   3120
            Width           =   1500
         End
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "(1-發文,2-內部收文)"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   6
      Left            =   1680
      TabIndex        =   70
      Top             =   1748
      Width           =   1545
   End
   Begin VB.Label lblClose 
      AutoSize        =   -1  'True
      Caption         =   "內部收文號 : "
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3480
      TabIndex        =   69
      Top             =   1748
      Width           =   1035
   End
   Begin MSForms.Label Lbl3 
      Height          =   255
      Index           =   6
      Left            =   3330
      TabIndex        =   64
      Top             =   330
      Width           =   1215
      VariousPropertyBits=   27
      Caption         =   "Lbl3"
      Size            =   "2143;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "收文號:"
      Height          =   255
      Left            =   2640
      TabIndex        =   63
      Top             =   330
      Width           =   585
   End
   Begin MSForms.Label Lbl3 
      Height          =   255
      Index           =   5
      Left            =   6840
      TabIndex        =   62
      Top             =   1748
      Width           =   1095
      VariousPropertyBits=   27
      Caption         =   "Lbl3"
      Size            =   "1931;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Lbl3 
      Height          =   255
      Index           =   4
      Left            =   1920
      TabIndex        =   61
      Top             =   1395
      Width           =   6135
      VariousPropertyBits=   27
      Caption         =   "Lbl3"
      Size            =   "10821;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Lbl3 
      Height          =   255
      Index           =   3
      Left            =   960
      TabIndex        =   60
      Top             =   1395
      Width           =   930
      VariousPropertyBits=   27
      Caption         =   "Lbl3"
      Size            =   "1640;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Lbl3 
      Height          =   255
      Index           =   2
      Left            =   1920
      TabIndex        =   59
      Top             =   1065
      Width           =   6135
      VariousPropertyBits=   27
      Caption         =   "Lbl3"
      Size            =   "10821;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Lbl3 
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   58
      Top             =   1065
      Width           =   930
      VariousPropertyBits=   27
      Caption         =   "Lbl3"
      Size            =   "1640;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Lbl3 
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   57
      Top             =   330
      Width           =   1455
      VariousPropertyBits=   27
      Caption         =   "Lbl3"
      Size            =   "2566;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "承辦人員 : "
      Height          =   255
      Index           =   4
      Left            =   5880
      TabIndex        =   56
      Top             =   1748
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "處理方式："
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   55
      Top             =   1748
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "代理人 : "
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   54
      Top             =   1395
      Width           =   675
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "申請人1 : "
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   53
      Top             =   1065
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱:"
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   28
      Top             =   720
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   330
      Width           =   765
   End
End
Attribute VB_Name = "frm060121_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/10/07 改成Form2.0 ; txtPA(index)、lbl3(index)
'Create by Lydia 2018/02/01 客戶提供文件處理
Option Explicit

Dim intWhere As Integer
Dim mPrevForm As Form
Dim pa(1 To 4) As String '本所案號
Dim m_PA08 As String '專利種類
Dim m_PA09 As String '申請國家
Dim m_PA150 As String '工程師組別
Dim m_CSD05 As String  'D類收文號

Dim oText As Control
Dim oLabel As Control
Dim oCheck As CheckBox
Dim intJ As Integer
Dim strDesc(1 To 11) As String '項目名稱
Dim strDescType(1 To 4) As String '副檔名
Dim tmpArr As Variant
Dim mRole As String 'Added by Lydia 2018/03/06  U=處理 ; Q=查看
Dim m_TCTchk As String 'Add By Sindy 2023/11/13


Public Sub SetParent(ByVal fm As Form, ByVal CNo As String, ByVal CP09 As String)
   Set mPrevForm = fm
   Call ChgCaseNo(CNo, pa)
   m_CSD05 = CP09
   
   'Added by Lydia 2018/03/06
   If TypeName(mPrevForm) = "frm060121" Then
        mRole = "U"
   Else
        mRole = "Q"
   End If
   'end 2018/03/06
End Sub

Private Sub cmdExit_Click()
   If cmdOK(0).Enabled = True Then
        If MsgBox("你並未存檔，確定離開嗎?", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
           Exit Sub
        End If
   End If
   
   Unload Me
End Sub

'先做內部收文，後存檔
Public Sub GetB202(ByVal stCP09 As String)
    If Len(stCP09) < 9 Then '內部收文202收文號
        MsgBox "內部收文未存檔 !", vbCritical
        Me.Show
        Unload Me
        Exit Sub
    End If
    If FormSave(txtResult, stCP09) = False Then
        Me.Show
    Else
        Unload Me
    End If
End Sub

Private Sub cmdOK_Click(Index As Integer)
Dim hLocalFile As Long 'Added by Lydia 2018/06/21
Dim strF22Emp As String 'Add By Sindy 2024/1/2
   
   'Added by Lydia 2020/02/26 先檢查
   If Index = 2 Then
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
      
         If txtResult <> "1" And txtResult <> "2" Then
               MsgBox "請輸入處理方式(1或2) !", vbCritical
               txtResult.SetFocus
               txtResult_GotFocus
               Exit Sub
         End If
         
         'Added by Lydia 2018/03/05 從新案建檔改到這裡，控制要輸入組別
         If (txtCSD(13).Text <> "" Or CmbFL(0).ListCount > 0) And m_PA150 = "" Then
               MsgBox "請到新案建檔，輸入工程師組別!", vbCritical
               Exit Sub
         End If
         'end 2018/03/05
         
         '處理方式1-發文
         If txtResult = "1" Then
               If FormSave(txtResult) = False Then Exit Sub
               'Add By Sindy 2023/11/13 淑華說，客戶提供文件發文時，發Mail通知分案人員進行209.檢視中說／235.核對中說格式分案作業
               If strSrvDate(1) >= 外專承辦歷程啟用日 And Chk1(3).Value = 1 Then '4.簡(繁)體中說
                  strSql = "select cp06,cp07,cp48,cp10 from caseprogress" & _
                           " where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'" & _
                           " and cp10 in('209','235') and cp27||cp57 is null and cp14 is null"
                  CheckOC
                  With adoRecordset
                     .CursorLocation = adUseClient
                     .Open strSql, cnnConnection, adOpenForwardOnly, adLockReadOnly
                     If .RecordCount > 0 Then
                        strExc(0) = "select cp06,cp07,cp48,cp10,tct118,nvl(tct10,tct04) as grpman" & _
                           " from caseprogress,transcasetitle" & _
                           " where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'" & _
                           " and cp31='Y' and tct01=cp09"
                        intI = 1
                        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                        If intI = 1 Then
                           m_TCTchk = "" & RsTemp.Fields("grpman") '命名記錄是否分組/工程師
                        End If
                        '案件性質名稱
                        strExc(10) = ""
                        strF22Emp = PUB_GetFCPHandler(pa(1), pa(2), pa(3), pa(4)) '外專程序管制人 Add By Sindy 2024/1/2
                        Call ClsPDGetCasePropertyL(1, pa(1), .Fields("cp10"), strExc(10))
                        'Modify By Sindy 2024/2/1 Winfrey:取消檢視中說工程師 =>strExc(10) & "工程師：" & GetPrjSalesNM(m_TCTchk) & vbCrLf
                        '因為隨著人員的離職或是後續主動修正或其他道以被分配為不同的工程師，導致帶出來的人與實際分案的人員會有不一致的情形，若要再改抓取的條件可能要想很多規則(目前是抓命名人員)，故建議刪除此欄位就好。
                        strExc(9) = "To " & GetPrjSalesNM(Pub_GetSpecMan("C")) & "," & vbCrLf & vbCrLf & _
                                    pa(1) & "-" & pa(2) & "今日客戶提供文件已發文，請分案" & strExc(10) & vbCrLf & _
                                    "本所期限：" & ChangeWStringToTDateString("" & .Fields("cp06")) & vbCrLf & _
                                    "法定期限：" & ChangeWStringToTDateString("" & .Fields("cp07")) & vbCrLf & vbCrLf & _
                                    "To " & GetPrjSalesNM(strF22Emp) & "," & vbCrLf & vbCrLf & _
                                    "請至案件資料及案件進度查詢新增承辦歷程交中打室排版。" & vbCrLf & vbCrLf & _
                                    "謝謝" & vbCrLf
                        strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                           "values('" & strUserNum & "','" & Pub_GetSpecMan("C") & ";" & strF22Emp & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
                           ",'【請分案" & strExc(10) & "】Our Ref: " & pa(1) & "-" & pa(2) & " [INCOM." & .Fields("cp10") & "]'" & _
                           ",'" & ChgSQL(strExc(9)) & "','" & strUserNum & "')"
                        cnnConnection.Execute strSql, intI
                     End If
                  End With
               End If
               '2023/11/13 END
               Unload Me
               
         '處理方式2-內部收文
         ElseIf txtResult = "2" Then
                If CheckUse("frm010001", strExec) = True Then
                    Me.Hide
                    frm010001.intChoose = 1
                    frm010001.intReceiveKind = 0
                    frm010001.intModifyKind = 0
                    Set frm010001.mPrevForm = Me
                    frm010001.Caption = "內部收文－新增"
                    frm010001.lblReciveCode.Caption = 內部收文
                    frm010001.txtSystem = pa(1)
                    frm010001.txtCode(0) = pa(2)
                    frm010001.txtCode(1) = pa(3)
                    frm010001.txtCode(2) = pa(4)
                    frm010001.txtCaseProperty = "202" '案件性質-補文件
                    frm010001.m_GetB202CP09 = "B" 'Added by Lydia 2021/02/22 客戶提供文件做內部收文
                Else
                    Exit Sub
                End If
               Call frm010001.cmdOK_Click(0)
         End If
         
      Case 1 '外文本
'Modified by Lydia 2018/03/23 無權限的錯誤要改訊息
'On Error Resume Next
On Error GoTo ErrHand01

            'Added by Lydia 2020/01/20 開啟[原始檔區]
            If InStr(cmdOK(Index).Caption, "原始檔") > 0 Then
                If cmdOK(Index).Tag = "" Then
                    MsgBox pa(1) & "-" & pa(2) & "在〔原始檔區〕的English_Vers收文號不存在!", vbInformation
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
'                strExc(1) = Pub_GetFCPcaseFilePath(pa(2), , pa(1))
'                'Remove by Lydia 2018/03/23
'                'If Pub_StrUserSt03 <> "M51" And Left(Pub_StrUserSt03, 1) <> "F" Then
'                '      If MsgBox("非國外部人員無權限進入\\English_Vers，是否繼續開啟？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
'                '           Exit Sub
'                '      End If
'                'End If
'                'end 2018/03/23
'                If Dir(strExc(1) & "\*.*") <> "" Then
'                     'Modified by Lydia 2018/06/21 用檔案總管開啟放置1~2分鐘後,檔案總管會出錯(ex. A2037, A4041)
'                     'SHELL "Explorer.exe " & strExc(1), vbNormalFocus  '開啟案件資料夾
'                     ShellExecute hLocalFile, "explore", strExc(1), vbNullString, vbNullString, 1
'                Else
'                     MsgBox pa(1) & "-" & pa(2) & "在" & strExc(1) & "的資料夾不存在或無檔案!", vbInformation
'                End If
                'end 2021/12/06
            End If 'Added by Lydia 2020/01/20
      Case 2 '卷宗區
            If cmdOK(Index).Visible = True Then
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

'Modified by Lydia 2018/03/06 改成Function
'Private function ReadData()
Public Function ReadData() As Boolean
 Dim rsRd As New ADODB.Recordset
 
 SSTab1.Tab = 0

   FormClear
    
   Lbl3(0).Caption = pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4)
   
   '客戶名稱:中->英->日 ; 代理人名稱: 英->中->日
   strExc(0) = "SELECT PA01,PA02,PA03,PA04,PA05,PA06,PA07,PA08,PA09,PA150,PA26,PA75," & _
                     " NVL(CU04,NVL(CU05,CU06)) CNAME,NVL(FA05,NVL(FA04,FA06)) FNAME" & _
                     " FROM PATENT,CUSTOMER,FAGENT" & _
                     " WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4)) & _
                     " AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+)" & _
                     " AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+)"
    intI = 0
    Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
    If intI = 1 Then
        Combo1.AddItem "中:" & RsTemp.Fields("PA05")
        Combo1.AddItem "英:" & RsTemp.Fields("PA06")
        Combo1.AddItem "日:" & RsTemp.Fields("PA07")
        Combo1.ListIndex = 0
        m_PA08 = "" & RsTemp.Fields("PA08")
        m_PA09 = "" & RsTemp.Fields("PA09")
        m_PA150 = "" & RsTemp.Fields("PA150")
        Lbl3(1).Caption = "" & RsTemp.Fields("PA26")
        Lbl3(2).Caption = "" & RsTemp.Fields("CNAME")
        Lbl3(3).Caption = "" & RsTemp.Fields("PA75")
        Lbl3(4).Caption = "" & RsTemp.Fields("FNAME")
        'Added by Lydia 2020/01/20 專利案件和English_Vers檔案：判斷檔案上傳目的地，外文本->原始檔區
        If PUB_ChkCPExist(pa, cntEnglish_Vers, , strExc(1), , "D") = True Then 'English_Vers992
            cmdOK(1).Caption = "原始檔"
            cmdOK(1).Tag = strExc(1)
        End If
        'Mark by Lydia 2020/03/18 以收文為準
        'If strSrvDate(1) >= XY特殊權限啟用日by檔案 Then
        '   cmdOK(1).Caption = "原始檔"
        'End If
        'end 2020/01/20
    End If
    'Modified by Lydia 2018/03/08 +處理人員
    strExc(0) = "select CustSupportDoc.*,s1.ST02,s2.st02 csd10n from CustSupportDoc,Staff s1, Staff s2 " & _
                     "where csd01='" & pa(1) & "' and csd02='" & pa(2) & "' and csd03='" & pa(3) & "' and csd04='" & pa(4) & "' " & _
                     "and csd05='" & m_CSD05 & "' and csd06=s1.st01(+) and csd10=s2.st01(+) "
    intI = 1
    Set rsRd = ClsLawReadRstMsg(intI, strExc(0))
    If intI = 1 Then
         With rsRd
              ReadData = True 'Added by Lydia 2018/03/06
              'Modified by Lydia 2018/03/06 判斷是否可處理
              If mRole = "U" Then
                    If Val("" & .Fields("csd11")) > 0 Then
                         MsgBox "該記錄已處理 ! "
                         cmdOK(0).Enabled = False
                         txtResult.Locked = True 'Added by Lydia 2018/03/06
                    End If
              Else
                    cmdOK(0).Enabled = False
                    cmdOK(0).Visible = False
                    cmdOK(2).Visible = False
                    txtResult.Visible = False
                    Label2(6).Visible = False
                    lblClose.Visible = True
                    'Modified by Lydia 2018/03/08 發文處理的顯示方式
                    If "" & rsRd.Fields("csd09") <> "" Then
                         lblClose.Caption = "內部收文號：" & rsRd.Fields("csd09")
                    ElseIf Val("" & .Fields("csd11")) > 0 Then
                         lblClose.Caption = "發文人員：" & rsRd.Fields("csd10n") & "　　發文時間：" & CFDate(TransDate(rsRd.Fields("csd11"), 1)) & " " & Format(rsRd.Fields("csd12"), "00:00")
                    Else
                         lblClose.Caption = "未處理"
                    End If
                    'end 2018/03/08
              End If
              'end 2018/03/06
              
              For intJ = 13 To 34
                    Select Case intJ
                         Case 21, 23, 25, 27, 29, 31, 33 'Check項目
                                 If "" & .Fields(intJ - 1) = "Y" Then
                                       Chk1(intJ).Value = vbChecked
                                 End If
                         Case Else
                                 txtCSD(intJ).Text = "" & .Fields(intJ - 1)
                    End Select
              Next intJ
              Lbl3(5).Caption = "" & .Fields("st02") '建檔人=承辦
              Lbl3(6).Caption = "" & .Fields("csd05") '收文號
              If Trim(txtCSD(13) & txtCSD(14)) <> "" Then Chk1(0).Value = vbChecked
              If Trim(txtCSD(15) & txtCSD(16)) <> "" Then Chk1(1).Value = vbChecked
              If Trim(txtCSD(17) & txtCSD(18)) <> "" Then Chk1(2).Value = vbChecked
              If Trim(txtCSD(19) & txtCSD(20)) <> "" Then Chk1(3).Value = vbChecked
              
              intI = 0
              For intJ = 14 To 20 Step 2
                    If txtCSD(intJ).Text <> "" Then
                         'Modify By Sindy 2024/1/17 改為共用函數
                         'Call SetCmbList(intI, txtCSD(intJ).Text)
                         Call PUB_SetCmbList(CmbFL(intI), txtCSD(intJ).Text)
                    End If
                    intI = intI + 1
              Next intJ
         End With
    End If
    Set rsRd = Nothing
End Function

Private Sub Form_Activate()
    'Modified by Lydia 2018/03/06
    'Me.txtResult.SetFocus
    If mRole = "U" Then Me.txtResult.SetFocus
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = 國外_FC
   FormClear
   SSTab1.Tab = 0
   SendKeys "{Tab}"

   Frame1.BackColor = &H8000000F
   Frame2.BackColor = &H8000000F
   
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
    
   'Remove by Lydia 2018/03/06 加入卷宗區，改呼叫
   'Call ReadData
   
   'Added by Lydia 2018/03/07
   If mRole = "Q" Then
       Me.Caption = "客戶提供文件-查詢"
   End If
   lblClose.Left = txtResult.Left
End Sub

Private Sub Form_Unload(Cancel As Integer)
   PUB_SendMailCache 'Added by Sindy 2023/11/14
   'Modified by Lydia 2018/03/06 判斷母表單
   'If TypeName(mPrevForm) <> "Nothing" Then
   If mRole = "U" Then
       Call mPrevForm.Command1_Click
       mPrevForm.Show
   'Added by Lydia 2018/03/06
   'Modifed by Lydia 2021/02/22 加判斷mPrevForm
   ElseIf TypeName(mPrevForm) <> "Nothing" Then
        mPrevForm.Show
   'end 2018/03/06
   End If
   
   Set frm060121_1 = Nothing
End Sub

Private Function FormSave(ByVal aKind As String, Optional ByVal strB202 As String = "") As Boolean
   Dim strCon1 As String, strCon2 As String

On Error GoTo CheckingErr
   cnnConnection.BeginTrans
    '處理方式-1.發文：將案件進度檔D類收文客戶提供文件(1920)的發文日和客戶提供文件記錄的處理日更新為系統日，並且刪除D類進度檔；
    If aKind = "1" Then
          strSql = "Update caseprogress set cp27=" & strSrvDate(1) & " where cp09 = '" & m_CSD05 & "'  and cp10='1920' "
          cnnConnection.Execute strSql, intI
          strSql = "Update CustSupportDoc set csd10='" & strUserNum & "' , csd11=" & strSrvDate(1) & ", csd12=" & Left(Format(ServerTime, "000000"), 4) & " where csd05='" & m_CSD05 & "' "
          cnnConnection.Execute strSql, intI
          strSql = "delete from caseprogress where cp09 = '" & m_CSD05 & "' and cp10='1920' "
          cnnConnection.Execute strSql, intI
          'Added by Lydia 2021/09/02 有點選4.簡(繁)體中說並且處理方式選1.發文時，將下一程序202(客戶提供中說（簡/繁）)，上Y，將期限銷掉。ex.FCP065570
          If Chk1(3).Value = 1 Then
               strSql = "Update NextProgress Set NP06='Y' where np02='" & pa(1) & "' and np03='" & pa(2) & "' and np04='" & pa(3) & "' and np05='" & pa(4) & "' and np06 is null and np07='202' and instr(np15,'客戶提供中說（簡/繁）') > 0 "
               Pub_SeekTbLog strSql
               cnnConnection.Execute strSql, intI
          End If
          'end 2021/09/02
    '處理方式-2.內部收文：會跳到內部收文畫面，新增B類收文202補文件後會刪除原D類收文，更新客戶提供文件記錄的處理收文號=B類收文號和處理日=系統日；並且將原本掛在D類收文號的卷宗區資料改掛到B類收文202的收文號。
    ElseIf aKind = "2" And strB202 <> "" Then
          '加案號,有5,6碼之分
          strExc(1) = IIf(pa(3) <> "0", "-" & pa(3), "") & IIf(pa(4) <> "00", "-" & pa(4), "")
          'Modified by Lydia 2018/04/18 by Lydia 2018/04/18 轉到補文件，PDF檔案仍保留在客戶提供文件之下，不分紙本送件or電子送件，若有需要再由程序手動移檔
          'strSql = "Update CasePaperPdf set CPP02=REPLACE(CPP02,'" & pa(1) & Val(pa(2)) & strExc(1) & ".1920.','" & pa(1) & Val(pa(2)) & strExc(1) & ".202.') " & _
                      "where CPP01='" & m_CSD05 & "'  and instr(CPP02,'" & pa(1) & Val(pa(2)) & strExc(1) & ".1920.') > 0  "
          strSql = "Update CasePaperPdf set CPP02=REPLACE(CPP02,'" & pa(1) & Val(pa(2)) & strExc(1) & ".1920.','" & pa(1) & Val(pa(2)) & strExc(1) & ".202.') " & _
                      "where CPP01='" & m_CSD05 & "'  and instr(CPP02,'" & pa(1) & Val(pa(2)) & strExc(1) & ".1920.') > 0 and upper(cpp02) not like '%.PDF' "
          cnnConnection.Execute strSql, intI
          'Modified by Lydia 2018/04/18
          'strSql = "Update CasePaperPdf set CPP02=REPLACE(CPP02,'" & pa(1) & Format(pa(2), "000000") & strExc(1) & ".1920.','" & pa(1) & Format(pa(2), "000000") & strExc(1) & ".202.') " & _
                       "where CPP01='" & m_CSD05 & "'  and instr(CPP02,'" & pa(1) & Format(pa(2), "000000") & strExc(1) & ".1920.') > 0  "
          strSql = "Update CasePaperPdf set CPP02=REPLACE(CPP02,'" & pa(1) & Format(pa(2), "000000") & strExc(1) & ".1920.','" & pa(1) & Format(pa(2), "000000") & strExc(1) & ".202.') " & _
                       "where CPP01='" & m_CSD05 & "'  and instr(CPP02,'" & pa(1) & Format(pa(2), "000000") & strExc(1) & ".1920.') > 0 and upper(cpp02) not like '%.PDF' "
          cnnConnection.Execute strSql, intI
          '更換收文號
          'Modified by Lydia 2018/03/07 排除.menu
          'strSql = "Update CasePaperPdf set CPP01='" & strB202 & "' where CPP01='" & m_CSD05 & "' "
          'Modified by Lydia 2018/04/18
          'strSql = "Update CasePaperPdf set CPP01='" & strB202 & "' where CPP01='" & m_CSD05 & "' and instr(upper(CPP02)," & CNULL(UCase(FCP提供文件)) & ") = 0 "
          strSql = "Update CasePaperPdf set CPP01='" & strB202 & "' where CPP01='" & m_CSD05 & "' and instr(upper(CPP02)," & CNULL(UCase(FCP提供文件)) & ") = 0 and upper(cpp02) not like '%.PDF' "
          cnnConnection.Execute strSql, intI
          'end 2018/04/18
          'Added by Lydia 2021/09/09 有點選3.英說(參考/翻譯用)將此收文補文件的進度所限及法限，更新為新案翻譯的所限及法限(可參FCP065191)
          If Chk1(2).Value = 1 Then
               strSql = "select cp06, cp07 from caseprogress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp10='201' and cp159=0 "
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strSql)
               If intI = 1 Then
                   If "" & RsTemp.Fields("cp06") & RsTemp.Fields("cp07") <> "" Then
                      strSql = "Update CaseProgress set CP06=" & CNULL(RsTemp.Fields("cp06"), True) & ", CP07=" & CNULL(RsTemp.Fields("cp07"), True) & " where cp09='" & strB202 & "' "
                      cnnConnection.Execute strSql, intI
                   End If
               End If
          End If
          'end 2021/09/09
          strSql = "Update CustSupportDoc set csd09='" & strB202 & "', csd10='" & strUserNum & "' , csd11=" & strSrvDate(1) & ", csd12=" & Left(Format(ServerTime, "000000"), 4) & " where csd05='" & m_CSD05 & "' "
          cnnConnection.Execute strSql, intI
          strSql = "delete from caseprogress where cp09 = '" & m_CSD05 & "' and cp10='1920' "
          cnnConnection.Execute strSql, intI
    End If
    
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

   For Each oText In txtCSD
      oText.Text = ""
      oText.Tag = ""
      oText.Locked = True
   Next

   For Each oLabel In Lbl3
      oLabel.Caption = ""
   Next
   
   For Each oCheck In Chk1
      oCheck.Value = vbUnchecked
   Next
   
   txtResult.Text = ""
   Combo1.Clear
   CmbFL(0).Clear
   CmbFL(1).Clear
   CmbFL(2).Clear
   CmbFL(3).Clear
   
   cmdOK(0).Enabled = True
   cmdOK(1).Tag = "" 'Added by Lydia 2020/01/20
   'Added by Lydia 2018/03/06
   lblClose.Caption = "" '內部收文號
   lblClose.Visible = False
   txtResult.Visible = True
   Label2(6).Visible = True
   cmdOK(0).Visible = True
   cmdOK(2).Visible = True
   'end 2018/03/06
End Sub

Private Sub txtCSD_GotFocus(Index As Integer)
   TextInverse txtCSD(Index)
   CloseIme
End Sub

'Modified by Lydia 2021/10/07 改成Form 2.0
'Private Sub txtCSD_KeyPress(Index As Integer, KeyAscii As Integer)
Private Sub txtCSD_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCSD_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   If txtCSD(Index).Text <> "" Then
       txtCSD(Index).ToolTipText = PUB_StringFilter(txtCSD(Index).Text)
   End If
End Sub

Private Sub txtResult_GotFocus()
  TextInverse txtResult
End Sub

Private Sub txtResult_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If (KeyAscii > 50 Or KeyAscii < 49) And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub
