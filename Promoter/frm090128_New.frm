VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090128_New 
   BorderStyle     =   1  '單線固定
   Caption         =   "查名單(網中)明細作業"
   ClientHeight    =   8436
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9432
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8436
   ScaleWidth      =   9432
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  '沒有框線
      Height          =   372
      Left            =   5232
      TabIndex        =   124
      Top             =   1848
      Width           =   4188
      Begin VB.CheckBox ChkS2 
         Caption         =   "團體標章"
         ForeColor       =   &H00FF0000&
         Height          =   250
         Index           =   0
         Left            =   24
         TabIndex        =   94
         Top             =   72
         Width           =   1140
      End
      Begin VB.CheckBox ChkS3 
         Caption         =   "僅查本所代理"
         ForeColor       =   &H000000FF&
         Height          =   250
         Index           =   1
         Left            =   1152
         TabIndex        =   95
         Top             =   72
         Width           =   1404
      End
      Begin VB.CheckBox ChkS3 
         Caption         =   "包含無效或核駁"
         ForeColor       =   &H000000FF&
         Height          =   250
         Index           =   2
         Left            =   2616
         TabIndex        =   96
         Top             =   72
         Width           =   1692
      End
   End
   Begin VB.PictureBox G_SeekPicColor 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   10560
      ScaleHeight     =   197
      ScaleMode       =   3  '像素
      ScaleWidth      =   246
      TabIndex        =   123
      Top             =   2640
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Index           =   1
      Left            =   3144
      MaxLength       =   7
      TabIndex        =   84
      Top             =   1152
      Width           =   756
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Index           =   0
      Left            =   1152
      MaxLength       =   7
      TabIndex        =   83
      Top             =   1152
      Width           =   756
   End
   Begin VB.CommandButton cmdRoute 
      BackColor       =   &H00C0FFFF&
      Caption         =   "輸入"
      Height          =   348
      Left            =   8544
      Style           =   1  '圖片外觀
      TabIndex        =   106
      Top             =   3984
      Width           =   708
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3900
      Left            =   48
      TabIndex        =   61
      Top             =   4488
      Width           =   9348
      _ExtentX        =   16489
      _ExtentY        =   6879
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   420
      TabMaxWidth     =   4410
      TabCaption(0)   =   "查名結果"
      TabPicture(0)   =   "frm090128_New.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblCN(5)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblCN(4)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblCN(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblCN(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "textDB(39)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "textDB(40)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "textDB(41)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "textDB(42)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "FR11"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Frame1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "FraTMA65"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "網中查名"
      TabPicture(1)   =   "frm090128_New.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1(18)"
      Tab(1).Control(1)=   "LBL1(15)"
      Tab(1).Control(2)=   "lblCN(7)"
      Tab(1).Control(3)=   "lblCN(6)"
      Tab(1).Control(4)=   "textDB(9)"
      Tab(1).Control(5)=   "textDB(6)"
      Tab(1).Control(6)=   "textDB(44)"
      Tab(1).Control(7)=   "textDB(5)"
      Tab(1).Control(8)=   "Frame3"
      Tab(1).ControlCount=   9
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  '沒有框線
         Caption         =   "Frame3"
         Height          =   1836
         Left            =   -74928
         TabIndex        =   125
         Top             =   1080
         Width           =   9204
         Begin VB.CommandButton cmdUp 
            Caption         =   "^"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   18
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   8904
            TabIndex        =   40
            Top             =   432
            Width           =   280
         End
         Begin VB.CommandButton cmdDown 
            Caption         =   "v"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   18
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   8904
            TabIndex        =   41
            Top             =   792
            Width           =   280
         End
         Begin VB.CommandButton cmdRemove 
            BackColor       =   &H00C0C0FF&
            Caption         =   "刪除<"
            Height          =   345
            Left            =   3240
            Style           =   1  '圖片外觀
            TabIndex        =   39
            Top             =   504
            Width           =   732
         End
         Begin VB.CommandButton cmdAdd 
            BackColor       =   &H00C0FFFF&
            Caption         =   "加入>"
            Height          =   345
            Left            =   3240
            Style           =   1  '圖片外觀
            TabIndex        =   38
            Top             =   48
            Width           =   732
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MGRD1 
            Height          =   1692
            Left            =   4032
            TabIndex        =   130
            Top             =   24
            Width           =   4812
            _ExtentX        =   8488
            _ExtentY        =   2985
            _Version        =   393216
            Cols            =   5
            FixedCols       =   0
            FormatString    =   "順序|檢索中文　　|檢索英文　　|檢索日文　　|檢索記號"
            _NumberOfBands  =   1
            _Band(0).Cols   =   5
         End
         Begin VB.Label lblCN 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFC0&
            Caption         =   "檢索中文："
            Height          =   180
            Index           =   8
            Left            =   24
            TabIndex        =   129
            Top             =   96
            Width           =   900
         End
         Begin MSForms.TextBox txtFM2 
            Height          =   324
            Index           =   0
            Left            =   996
            TabIndex        =   34
            Top             =   24
            Width           =   2196
            VariousPropertyBits=   -1476378597
            MaxLength       =   100
            ScrollBars      =   2
            Size            =   "3881;572"
            Value           =   "一二三四五六七八九十"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtFM2 
            Height          =   324
            Index           =   3
            Left            =   996
            TabIndex        =   37
            Top             =   1080
            Width           =   2196
            VariousPropertyBits=   -1467989989
            MaxLength       =   100
            ScrollBars      =   2
            Size            =   "3873;572"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label lblCN 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFC0&
            Caption         =   "檢索記號："
            Height          =   180
            Index           =   11
            Left            =   24
            TabIndex        =   128
            Top             =   1152
            Width           =   900
         End
         Begin VB.Label lblCN 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFC0&
            Caption         =   "檢索英文："
            Height          =   180
            Index           =   9
            Left            =   24
            TabIndex        =   127
            Top             =   432
            Width           =   900
         End
         Begin VB.Label lblCN 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFC0&
            Caption         =   "檢索日文："
            Height          =   180
            Index           =   10
            Left            =   24
            TabIndex        =   126
            Top             =   768
            Width           =   900
         End
         Begin MSForms.TextBox txtFM2 
            Height          =   324
            Index           =   1
            Left            =   996
            TabIndex        =   35
            Top             =   360
            Width           =   2196
            VariousPropertyBits=   -1467989989
            MaxLength       =   100
            ScrollBars      =   2
            Size            =   "3873;572"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtFM2 
            Height          =   324
            Index           =   2
            Left            =   996
            TabIndex        =   36
            Top             =   720
            Width           =   2196
            VariousPropertyBits=   -1467989989
            MaxLength       =   100
            ScrollBars      =   2
            Size            =   "3873;572"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
      End
      Begin VB.Frame FraTMA65 
         Caption         =   "覆核區"
         Height          =   1092
         Left            =   48
         TabIndex        =   108
         Top             =   2760
         Width           =   9204
         Begin VB.CheckBox ChkTMA67 
            Caption         =   "已排除近似"
            ForeColor       =   &H000000C0&
            Height          =   180
            Index           =   2
            Left            =   4944
            TabIndex        =   11
            Top             =   216
            Width           =   1188
         End
         Begin VB.CheckBox ChkTMA67 
            Caption         =   "否（需確認客戶關係）"
            ForeColor       =   &H000000C0&
            Height          =   180
            Index           =   1
            Left            =   2616
            TabIndex        =   10
            Top             =   216
            Width           =   2076
         End
         Begin VB.CheckBox ChkTMA69 
            Caption         =   "經上級核可先提申再補同意書"
            Height          =   225
            Index           =   1
            Left            =   1464
            TabIndex        =   14
            Top             =   816
            Width           =   2760
         End
         Begin VB.CheckBox ChkTMA69 
            Caption         =   "經上級核可代理"
            Height          =   225
            Index           =   0
            Left            =   1464
            TabIndex        =   13
            Top             =   552
            Width           =   1944
         End
         Begin VB.CommandButton cmdSaveD 
            Caption         =   "存檔"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   8.4
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Index           =   1
            Left            =   8088
            Style           =   1  '圖片外觀
            TabIndex        =   12
            Top             =   168
            Width           =   960
         End
         Begin VB.CheckBox ChkTMA67 
            Caption         =   "是"
            ForeColor       =   &H000000C0&
            Height          =   180
            Index           =   0
            Left            =   1824
            TabIndex        =   9
            Top             =   216
            Width           =   540
         End
         Begin MSForms.TextBox textDB 
            Height          =   492
            Index           =   68
            Left            =   5400
            TabIndex        =   15
            Top             =   552
            Width           =   3696
            VariousPropertyBits=   -1466939365
            MaxLength       =   300
            Size            =   "6519;868"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "協商流程結果："
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   2
            Left            =   96
            TabIndex        =   111
            Top             =   552
            Width           =   1260
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "意見："
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   9
            Left            =   4824
            TabIndex        =   110
            Top             =   576
            Width           =   540
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "與本所近似或相同："
            ForeColor       =   &H000000C0&
            Height          =   180
            Index           =   8
            Left            =   96
            TabIndex        =   109
            Top             =   216
            Width           =   1620
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "附件區："
         Height          =   1596
         Left            =   5304
         TabIndex        =   87
         Top             =   1128
         Width           =   3936
         Begin VB.ListBox lstAtt 
            Height          =   924
            Index           =   0
            IntegralHeight  =   0   'False
            ItemData        =   "frm090128_New.frx":0038
            Left            =   96
            List            =   "frm090128_New.frx":003F
            MultiSelect     =   2  '進階多重選取
            Sorted          =   -1  'True
            TabIndex        =   107
            TabStop         =   0   'False
            Top             =   192
            Width           =   3780
         End
         Begin VB.CommandButton cmdOpenAtt 
            Caption         =   "開啟"
            Height          =   345
            Index           =   0
            Left            =   144
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   1176
            Width           =   675
         End
         Begin VB.CommandButton cmdSaveAtt 
            Caption         =   "下載"
            Height          =   345
            Index           =   0
            Left            =   1632
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   1176
            Width           =   675
         End
         Begin VB.CommandButton cmdAddAtt 
            Caption         =   "新增"
            Height          =   345
            Index           =   0
            Left            =   2376
            TabIndex        =   23
            Top             =   1176
            Width           =   675
         End
         Begin VB.CommandButton cmdRemAtt 
            Caption         =   "刪除"
            Height          =   345
            Index           =   0
            Left            =   3144
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   1176
            Width           =   675
         End
         Begin VB.CommandButton cmdSelect 
            Caption         =   "全選"
            Height          =   345
            Index           =   0
            Left            =   888
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   1176
            Width           =   675
         End
      End
      Begin VB.Frame FR11 
         Caption         =   "查覆區"
         Height          =   1596
         Left            =   48
         TabIndex        =   71
         Top             =   1128
         Width           =   5244
         Begin VB.CommandButton cmdSaveD 
            Caption         =   "存檔"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   8.4
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Index           =   0
            Left            =   3576
            Style           =   1  '圖片外觀
            TabIndex        =   7
            Top             =   576
            Width           =   960
         End
         Begin VB.CheckBox ChkTMA16 
            Caption         =   "是"
            Height          =   180
            Left            =   1608
            TabIndex        =   4
            Top             =   192
            Width           =   564
         End
         Begin VB.ComboBox Cbo1 
            Height          =   276
            IntegralHeight  =   0   'False
            ItemData        =   "frm090128_New.frx":004B
            Left            =   3672
            List            =   "frm090128_New.frx":004D
            TabIndex        =   5
            Text            =   "Cbo1"
            Top             =   168
            Width           =   1485
         End
         Begin MSForms.TextBox textDB 
            Height          =   540
            Index           =   15
            Left            =   672
            TabIndex        =   8
            Top             =   960
            Width           =   4404
            VariousPropertyBits=   -1466939365
            MaxLength       =   300
            Size            =   "7768;952"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox textDB 
            Height          =   444
            Index           =   17
            Left            =   1248
            TabIndex        =   6
            Top             =   480
            Width           =   2196
            VariousPropertyBits=   -1467987941
            MaxLength       =   30
            Size            =   "3873;783"
            FontName        =   "新細明體"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label LBL1 
            Caption         =   "近似本所申請號/審定號："
            Height          =   372
            Index           =   9
            Left            =   72
            TabIndex        =   117
            Top             =   528
            Width           =   1212
         End
         Begin VB.Label LBL1 
            AutoSize        =   -1  'True
            Caption         =   "是否與本所近似："
            Height          =   180
            Index           =   16
            Left            =   72
            TabIndex        =   74
            Top             =   216
            Width           =   1440
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "意見："
            Height          =   180
            Index           =   1
            Left            =   72
            TabIndex        =   73
            Top             =   984
            Width           =   540
         End
         Begin VB.Label LBL1 
            Caption         =   "標章查覆結果："
            Height          =   216
            Index           =   19
            Left            =   2376
            TabIndex        =   72
            Top             =   192
            Width           =   1260
         End
      End
      Begin MSForms.TextBox textDB 
         Height          =   300
         Index           =   5
         Left            =   -71712
         TabIndex        =   31
         Top             =   336
         Width           =   996
         VariousPropertyBits=   -1467987941
         MaxLength       =   7
         Size            =   "1757;529"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textDB 
         Height          =   300
         Index           =   44
         Left            =   -73992
         TabIndex        =   33
         Top             =   672
         Width           =   8052
         VariousPropertyBits=   -1467987941
         MaxLength       =   50
         Size            =   "14203;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textDB 
         Height          =   300
         Index           =   6
         Left            =   -69432
         TabIndex        =   32
         Top             =   336
         Width           =   3468
         VariousPropertyBits=   -1467987941
         MaxLength       =   50
         Size            =   "6117;529"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textDB 
         Height          =   300
         Index           =   9
         Left            =   -73992
         TabIndex        =   30
         Top             =   336
         Width           =   996
         VariousPropertyBits=   -1467987941
         MaxLength       =   7
         Size            =   "1757;529"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textDB 
         Height          =   396
         Index           =   42
         Left            =   6048
         TabIndex        =   3
         Top             =   672
         Width           =   3168
         VariousPropertyBits=   -1467987941
         Size            =   "5588;698"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textDB 
         Height          =   300
         Index           =   41
         Left            =   6048
         TabIndex        =   1
         Top             =   360
         Width           =   3168
         VariousPropertyBits=   -1467987941
         MaxLength       =   10
         Size            =   "5588;529"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textDB 
         Height          =   396
         Index           =   40
         Left            =   1272
         TabIndex        =   2
         Top             =   672
         Width           =   3168
         VariousPropertyBits=   -1467987941
         Size            =   "5588;698"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textDB 
         Height          =   300
         Index           =   39
         Left            =   1272
         TabIndex        =   0
         Top             =   360
         Width           =   3168
         VariousPropertyBits=   -1467987941
         MaxLength       =   10
         Size            =   "5588;529"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label lblCN 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "網中編號："
         Height          =   180
         Index           =   6
         Left            =   -70368
         TabIndex        =   114
         Top             =   390
         Width           =   900
      End
      Begin VB.Label lblCN 
         BackColor       =   &H00FFFFC0&
         Caption         =   "近似本所申請號/審定號："
         Height          =   408
         Index           =   3
         Left            =   96
         TabIndex        =   113
         Top             =   672
         Width           =   1116
      End
      Begin VB.Label lblCN 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "文字近似度："
         Height          =   180
         Index           =   2
         Left            =   96
         TabIndex        =   70
         Top             =   384
         Width           =   1080
      End
      Begin VB.Label lblCN 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "圖形近似度："
         Height          =   180
         Index           =   4
         Left            =   4872
         TabIndex        =   69
         Top             =   384
         Width           =   1080
      End
      Begin VB.Label lblCN 
         BackColor       =   &H00FFFFC0&
         Caption         =   "近似本所申請號/審定號："
         Height          =   408
         Index           =   5
         Left            =   4872
         TabIndex        =   68
         Top             =   672
         Width           =   1116
      End
      Begin VB.Label lblCN 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "系統備註："
         Height          =   180
         Index           =   7
         Left            =   -74928
         TabIndex        =   67
         Top             =   720
         Width           =   900
      End
      Begin VB.Label LBL1 
         Caption         =   "送出日期："
         Height          =   240
         Index           =   15
         Left            =   -72624
         TabIndex        =   66
         Top             =   390
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "分發日期："
         Height          =   180
         Index           =   18
         Left            =   -74928
         TabIndex        =   63
         Top             =   390
         Width           =   900
      End
   End
   Begin VB.CommandButton cmdKD 
      BackColor       =   &H00C0FFFF&
      Caption         =   "更換圖片"
      Height          =   345
      Index           =   2
      Left            =   8376
      Style           =   1  '圖片外觀
      TabIndex        =   103
      Top             =   3264
      Width           =   900
   End
   Begin VB.PictureBox tmpKeyPic1 
      Height          =   1680
      Left            =   5664
      ScaleHeight     =   136
      ScaleMode       =   3  '像素
      ScaleWidth      =   210
      TabIndex        =   59
      Top             =   2232
      Width           =   2568
      Begin VB.Image tmpKeyImg1 
         Height          =   1560
         Left            =   264
         Top             =   72
         Width           =   1632
      End
   End
   Begin VB.CommandButton cmdKD 
      BackColor       =   &H00C0FFFF&
      Caption         =   "另開視窗"
      Height          =   345
      Index           =   0
      Left            =   8376
      Style           =   1  '圖片外觀
      TabIndex        =   101
      Top             =   2424
      Width           =   900
   End
   Begin VB.CommandButton cmdKD 
      BackColor       =   &H00C0FFFF&
      Caption         =   "下載"
      Height          =   345
      Index           =   1
      Left            =   8376
      Style           =   1  '圖片外觀
      TabIndex        =   102
      Top             =   2832
      Width           =   900
   End
   Begin VB.CheckBox ChkS3 
      Caption         =   "是"
      Height          =   250
      Index           =   0
      Left            =   8544
      TabIndex        =   88
      Top             =   1201
      Width           =   492
   End
   Begin VB.CommandButton cmdTo 
      BackColor       =   &H00C0FFC0&
      Caption         =   "收文"
      Height          =   345
      Left            =   5868
      Style           =   1  '圖片外觀
      TabIndex        =   55
      Top             =   72
      Width           =   1000
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   345
      Left            =   8424
      TabIndex        =   54
      Top             =   72
      Width           =   900
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "下一張單(&N)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   7.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3864
      Style           =   1  '圖片外觀
      TabIndex        =   53
      Top             =   72
      Width           =   1000
   End
   Begin VB.CommandButton cmdSend 
      BackColor       =   &H00C0C0FF&
      Caption         =   "送出(&O)"
      Height          =   345
      Left            =   6960
      Style           =   1  '圖片外觀
      TabIndex        =   52
      Top             =   72
      Width           =   1200
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "上一張單(&P)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   7.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2832
      Style           =   1  '圖片外觀
      TabIndex        =   51
      Top             =   72
      Width           =   1000
   End
   Begin VB.CommandButton cmdSendMail 
      BackColor       =   &H00C0FFFF&
      Caption         =   "通知送件"
      Height          =   345
      Left            =   4932
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   50
      Top             =   72
      Width           =   900
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8880
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSForms.TextBox textDB 
      Height          =   300
      Index           =   7
      Left            =   2952
      TabIndex        =   100
      Top             =   3648
      Width           =   924
      VariousPropertyBits=   -1467987941
      MaxLength       =   7
      Size            =   "1630;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textDB 
      Height          =   300
      Index           =   4
      Left            =   5520
      TabIndex        =   77
      Top             =   528
      Width           =   996
      VariousPropertyBits=   -1467987941
      MaxLength       =   7
      Size            =   "1757;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblData 
      Height          =   252
      Index           =   35
      Left            =   1056
      TabIndex        =   122
      Top             =   264
      Width           =   1452
      ForeColor       =   16711680
      VariousPropertyBits=   27
      Size            =   "2561;444"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label LBL1 
      Caption         =   "發  文  日："
      Height          =   240
      Index           =   7
      Left            =   2160
      TabIndex        =   121
      Top             =   1200
      Width           =   972
   End
   Begin VB.Label LBL1 
      Caption         =   "通知送件日："
      Height          =   240
      Index           =   6
      Left            =   48
      TabIndex        =   120
      Top             =   1188
      Width           =   1116
   End
   Begin MSForms.Label lblData 
      Height          =   252
      Index           =   43
      Left            =   4896
      TabIndex        =   119
      Top             =   3672
      Width           =   612
      ForeColor       =   192
      VariousPropertyBits=   27
      Size            =   "1080;444"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblData 
      Height          =   252
      Index           =   25
      Left            =   984
      TabIndex        =   118
      Top             =   3672
      Width           =   996
      ForeColor       =   192
      VariousPropertyBits=   27
      Size            =   "1757;444"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textDB 
      Height          =   300
      Index           =   32
      Left            =   6024
      TabIndex        =   86
      Top             =   1152
      Width           =   756
      VariousPropertyBits=   -1467987941
      MaxLength       =   7
      Size            =   "1333;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textDB 
      Height          =   300
      Index           =   31
      Left            =   5016
      TabIndex        =   85
      Top             =   1152
      Width           =   756
      VariousPropertyBits=   -1467987941
      MaxLength       =   7
      Size            =   "1333;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textDB 
      Height          =   444
      Index           =   28
      Left            =   5328
      TabIndex        =   105
      Top             =   3960
      Width           =   3180
      VariousPropertyBits=   -1467987941
      MaxLength       =   50
      Size            =   "5609;783"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textDB 
      Height          =   444
      Index           =   26
      Left            =   624
      TabIndex        =   104
      Top             =   3960
      Width           =   3660
      VariousPropertyBits=   -1467987941
      MaxLength       =   50
      Size            =   "6456;783"
      FontName        =   "細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textDB 
      Height          =   300
      Index           =   38
      Left            =   8472
      TabIndex        =   92
      Top             =   1512
      Width           =   408
      VariousPropertyBits=   -1467987941
      MaxLength       =   3
      Size            =   "720;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
      ParagraphAlign  =   2
   End
   Begin MSForms.TextBox textDB 
      Height          =   300
      Index           =   37
      Left            =   7224
      TabIndex        =   91
      Top             =   1512
      Width           =   408
      VariousPropertyBits=   -1467987941
      MaxLength       =   3
      Size            =   "720;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
      ParagraphAlign  =   2
   End
   Begin MSForms.TextBox textDB 
      Height          =   300
      Index           =   36
      Left            =   6000
      TabIndex        =   90
      Top             =   1512
      Width           =   408
      VariousPropertyBits=   -1467987941
      MaxLength       =   3
      Size            =   "720;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
      ParagraphAlign  =   2
   End
   Begin MSForms.TextBox textDB 
      Height          =   444
      Index           =   33
      Left            =   936
      TabIndex        =   99
      Top             =   3192
      Width           =   4500
      VariousPropertyBits=   -1467987941
      MaxLength       =   50
      Size            =   "7937;783"
      FontName        =   "細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textDB 
      Height          =   444
      Index           =   18
      Left            =   936
      TabIndex        =   98
      Top             =   2736
      Width           =   4500
      VariousPropertyBits=   -1467987941
      MaxLength       =   100
      Size            =   "7937;783"
      FontName        =   "細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textDB 
      Height          =   444
      Index           =   24
      Left            =   936
      TabIndex        =   97
      Top             =   2280
      Width           =   4500
      VariousPropertyBits=   -1467987941
      MaxLength       =   250
      Size            =   "7937;783"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textDB 
      Height          =   444
      Index           =   23
      Left            =   624
      TabIndex        =   93
      Top             =   1800
      Width           =   4500
      VariousPropertyBits=   -1467987941
      MaxLength       =   250
      Size            =   "7937;783"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textDB 
      Height          =   300
      Index           =   22
      Left            =   624
      TabIndex        =   89
      Top             =   1488
      Width           =   3600
      VariousPropertyBits=   -1467987941
      MaxLength       =   30
      Size            =   "6350;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textDB 
      Height          =   300
      Index           =   65
      Left            =   7680
      TabIndex        =   82
      Top             =   840
      Width           =   672
      VariousPropertyBits=   -1467987941
      MaxLength       =   6
      Size            =   "1185;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textDB 
      Height          =   300
      Index           =   14
      Left            =   5520
      TabIndex        =   81
      Top             =   840
      Width           =   996
      VariousPropertyBits=   -1467987941
      MaxLength       =   7
      Size            =   "1757;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textDB 
      Height          =   300
      Index           =   11
      Left            =   7680
      TabIndex        =   78
      Top             =   528
      Width           =   996
      VariousPropertyBits=   -1467987941
      MaxLength       =   7
      Size            =   "1757;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textDB 
      Height          =   300
      Index           =   12
      Left            =   3360
      TabIndex        =   80
      Top             =   840
      Width           =   996
      VariousPropertyBits=   -1467987941
      MaxLength       =   7
      Size            =   "1757;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textDB 
      Height          =   300
      Index           =   1
      Left            =   3360
      TabIndex        =   76
      Top             =   528
      Width           =   996
      VariousPropertyBits=   -1467987941
      MaxLength       =   9
      Size            =   "1757;529"
      Value           =   "H11300001"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textDB 
      Height          =   300
      Index           =   10
      Left            =   816
      TabIndex        =   79
      Top             =   840
      Width           =   672
      VariousPropertyBits=   -1467987941
      MaxLength       =   6
      Size            =   "1185;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textDB 
      Height          =   300
      Index           =   8
      Left            =   816
      TabIndex        =   75
      Top             =   528
      Width           =   672
      VariousPropertyBits=   -1467987941
      MaxLength       =   6
      Size            =   "1185;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "委查筆數：中文　　　筆，英文　　　筆，圖形　　　筆"
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   16
      Left            =   4656
      TabIndex        =   116
      Top             =   1536
      Width           =   4692
   End
   Begin VB.Label Label2 
      Caption         =   "～"
      Height          =   228
      Left            =   5784
      TabIndex        =   115
      Top             =   1200
      Width           =   228
   End
   Begin VB.Label lblCN 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
      Caption         =   "回寫結果："
      Height          =   180
      Index           =   1
      Left            =   3984
      TabIndex        =   112
      Top             =   3672
      Width           =   900
   End
   Begin VB.Label LBL1 
      Alignment       =   1  '靠右對齊
      Caption         =   "智權備註："
      Height          =   240
      Index           =   17
      Left            =   48
      TabIndex        =   65
      Top             =   3240
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "指定商品/服務："
      Height          =   384
      Index           =   25
      Left            =   48
      TabIndex        =   64
      Top             =   2328
      Width           =   840
   End
   Begin VB.Label LBL1 
      Alignment       =   1  '靠右對齊
      Caption         =   "查詢區間："
      Height          =   240
      Index           =   12
      Left            =   4104
      TabIndex        =   62
      Top             =   1188
      Width           =   900
   End
   Begin MSForms.TextBox textCUID 
      Height          =   240
      Left            =   48
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   24
      Width           =   2736
      VariousPropertyBits=   671105055
      ForeColor       =   16711680
      Size            =   "4826;423"
      Value           =   "CREATE : "
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "是否進行全類檢索："
      Height          =   252
      Index           =   12
      Left            =   6888
      TabIndex        =   58
      Top             =   1200
      Width           =   1620
   End
   Begin MSForms.Label lblData 
      Height          =   252
      Index           =   65
      Left            =   8400
      TabIndex        =   57
      Top             =   864
      Width           =   756
      VariousPropertyBits=   27
      Size            =   "1333;444"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblData 
      Height          =   252
      Index           =   10
      Left            =   1512
      TabIndex        =   56
      Top             =   864
      Width           =   756
      VariousPropertyBits=   27
      Size            =   "1333;444"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblData 
      Height          =   252
      Index           =   8
      Left            =   1512
      TabIndex        =   49
      Top             =   552
      Width           =   756
      VariousPropertyBits=   27
      Size            =   "1333;444"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblCN 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
      Caption         =   "回寫日期："
      Height          =   180
      Index           =   0
      Left            =   2040
      TabIndex        =   48
      Top             =   3672
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "圖形路徑："
      ForeColor       =   &H000000C0&
      Height          =   240
      Index           =   7
      Left            =   4416
      TabIndex        =   47
      Top             =   4056
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "文字："
      ForeColor       =   &H000000C0&
      Height          =   240
      Index           =   6
      Left            =   48
      TabIndex        =   46
      Top             =   4056
      Width           =   636
   End
   Begin VB.Label Label1 
      Caption         =   "檢索方式："
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   2
      Left            =   48
      TabIndex        =   45
      Top             =   3672
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "類別："
      Height          =   240
      Index           =   3
      Left            =   48
      TabIndex        =   44
      Top             =   1536
      Width           =   600
   End
   Begin VB.Label Label1 
      Caption         =   "組群："
      Height          =   240
      Index           =   1
      Left            =   48
      TabIndex        =   43
      Top             =   1848
      Width           =   600
   End
   Begin VB.Label Label1 
      Caption         =   "客戶名稱："
      Height          =   240
      Index           =   5
      Left            =   48
      TabIndex        =   42
      Top             =   2784
      Width           =   900
   End
   Begin VB.Label LBL1 
      Alignment       =   1  '靠右對齊
      Caption         =   "查覆期限："
      Height          =   240
      Index           =   4
      Left            =   6744
      TabIndex        =   29
      Top             =   540
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "查名單號："
      Height          =   240
      Index           =   4
      Left            =   2424
      TabIndex        =   28
      Top             =   564
      Width           =   900
   End
   Begin VB.Label LBL1 
      Caption         =   "查名人："
      Height          =   240
      Index           =   1
      Left            =   48
      TabIndex        =   27
      Top             =   852
      Width           =   720
   End
   Begin VB.Label Label1 
      Caption         =   "送出期限："
      Height          =   240
      Index           =   19
      Left            =   2424
      TabIndex        =   26
      Top             =   852
      Width           =   900
   End
   Begin VB.Label LBL1 
      Caption         =   "委查人："
      Height          =   240
      Index           =   8
      Left            =   48
      TabIndex        =   25
      Top             =   564
      Width           =   720
   End
   Begin VB.Label LBL1 
      Caption         =   "委查日期："
      Height          =   240
      Index           =   3
      Left            =   4608
      TabIndex        =   19
      Top             =   540
      Width           =   900
   End
   Begin VB.Label LBL1 
      Alignment       =   1  '靠右對齊
      Caption         =   "查覆日期："
      Height          =   240
      Index           =   5
      Left            =   4608
      TabIndex        =   18
      Top             =   828
      Width           =   900
   End
   Begin VB.Label LBL1 
      Caption         =   "覆核主管："
      Height          =   240
      Index           =   2
      Left            =   6744
      TabIndex        =   17
      Top             =   828
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   48
      TabIndex        =   16
      Top             =   288
      Width           =   996
   End
End
Attribute VB_Name = "frm090128_New"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create by Lydia 2024/10/09
Option Explicit

Dim m_NoList As String '傳入單號(多筆)
Dim m_NoIdx As Integer '操作的單號index
Dim iStiu As Integer  '狀態 :0查詢 1編輯
Dim mSeqNo As String

'使用者角色 U:查名人 Q:委查人(包含非分配的查名人) M:覆核主管 A:查名單維護(限電腦中心)
Dim R_type As String '依條件判斷權限記錄在cmdSend.tag
Dim mbolCall As Boolean  '外部呼叫
Private Const fileMax As Integer = 99  '預設最大附件數
Dim fileNow As Integer '目前附件最大流水號
'設定可使用表單
Private nfrm090129 As Form
Private nfrm090131 As Form

Dim m_PrevForm As Form '前一畫面
Dim m_TMQApp As String '收文已勾選的單號(接洽單Form_Unload使用)

Dim intJ As Integer, strTmp1 As String, tmpArr As Variant
Dim rsAD As New ADODB.Recordset
Dim rsTmp1 As New ADODB.Recordset
Dim intLastRow As Integer

Dim oObj As Control

'附件宣告區
Dim m_AttachPath As String
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

Dim m_TMA01 As String '查名單號
Dim m_TMA02 As String '資料來源
Private Const m_ATMF02 As String = "3"   '查名結果附件
Dim m_STMF03 As String '圖形查詢附件序號：當時將舊資料匯入是將原本的附件直接匯入(TMF03=01,02)，後來定義00=圖形查詢附件

Dim ShowCP09 As String '目前進度的總收文號(可能是多次申請)
Dim ShowCP(1 To 4) As String, ShowCP14 As String, ShowCP27 As String, ShowCP57 As String '目前進度的案號資料

Dim FirstCP09 As String '總收文號(第一次收文=TMA34)
Dim FirstCP(1 To 4) As String, FirstCP14 As String, FirstCP27 As String, FirstCP57 As String '第一次收文之本所案號
Dim FirstCPP02t As String '符合規則的卷宗區檔名開頭
Dim mPrevTM1215 As String '保留上次輸入的申請號/審定號
Dim bolSave As String '是否已存檔
Dim mbolSend As Boolean '是否查覆完畢

Dim m_TMA13 As String '委查人自請撤回
Dim m_TMA20 As String '1-圖體標章; 113/10/4 關閉2-證明標章
Dim m_TMA25 As String '檢索方式TMA25：1-文字, 2-圖形, 3-文字＋圖形
Dim m_TMA27 As String '圖形查名：Y=有匯入圖片
Dim m_TMA71 As String '(原)查名單號(114.4.14):1.由原查名單轉入 2.HM開頭，表示為急件=人工處理TMQapp.TQA01=TMQ18
Dim m_TMA72 As String '委查人附件已讀
Dim bolModify As Boolean '查覆完畢再次修改(先有mbolSend=true,之後有修改內容)
Dim bolModCheck As Boolean '筆數已變更
Dim strMod(0 To 2)  As String '再次修改的記錄mail(0=主旨,1=修改歷程,2=收件人)
'Dim strAgree As String 'Added by Lydia 2016/06/27 內商核可人員---不使用
Dim strPreAgree As String 'Added by Lydia 2022/05/25 內商查名覆核人員
Dim bolChgTMA69 As Boolean 'Added by Lydia 2016/09/12 覆核結果更改
'用在OpenDocument
Private Const ERROR_FILE_NO_ASSOCIATION As Long = 31
Private Const ERROR_FILE_NOT_FOUND As Long = 2
Private Const ERROR_PATH_NOT_FOUND As Long = 3
Private Const ERROR_FILE_SUCCESS As Long = 32
Private Const ERROR_BAD_FORMAT As Long = 11

Dim stIdList As String 'Added by Lydia 2019/08/12 創新業務組成員可操作清單(WXX部門的人可以操作自已部門所有人的資料,例W10所有人都可操作W1001，W20所有人都可操作W2001。
'開放特殊設定權限
Dim bolSpecMan As Boolean  '是否為特殊設定檔人員
Dim strSpecCode As String '特殊設定檔設定代號
Dim strGrpTmp1 As String, strGrpTmp2 As String
Private Const outType As String = "JPG" '網中系統限制圖片只能為JPG檔
Dim colORD As Integer, colWord(0 To 3) As Integer  'Grid欄位值
Dim strTestReceiver As String 'Added by Lydia 2025/04/21 開放測試模式，指定測試收件者

Public Sub SetParent(ByRef pForm As Form, ByVal bolCall As Boolean, ByVal pNoList As String, ByVal pNoIdx As Integer, ByVal pType As String, ByVal pStiu As String, Optional ByVal pCP09 As String)
   
   Set m_PrevForm = pForm
   
   mbolCall = bolCall  '外部呼叫
   m_NoList = pNoList  '傳入單號(多筆)
   m_NoIdx = pNoIdx    '操作的單號index
   R_type = pType      '使用者角色 U:查名人 Q:委查人(包含非分配的查名人) M:覆核主管 A:查名單維護(限電腦中心)
   iStiu = pStiu       '狀態 :0查詢 1編輯(原本>>查名單維護(限電腦中心))
   ShowCP09 = pCP09    '傳入目前進度的總收文號(可能是多次申請)

End Sub

Private Function IsSaveData() As Boolean
Dim strTmpA As String
   
   IsSaveData = False
   
   If cmdSaveD(0).Visible = True And cmdSaveD(0).Enabled = True And textDB(1) <> "" Then
      If SaveDetailAns(0) = "F" Then
         Exit Function
      End If
   End If
   If cmdSaveD(1).Visible = True And cmdSaveD(1).Enabled = True And textDB(1) <> "" Then
      If SaveDetailAns(1) = "F" Then
         Exit Function
      End If
   End If
   
   IsSaveData = True
End Function

'*****依權限設定欄位*****
Private Sub FormEnabled()
Dim tmpBol As Boolean

   '查名單維護
   If R_type = "A" Then '查名單維護A：先預設全部可以變更
      cmdSend.Enabled = True
      cmdTo.Visible = False
      cmdSendMail.Visible = False
   Else
      If iStiu = 1 Then
         cmdSend.Enabled = True
         If InStr(cmdSend.Caption, "送出") > 0 And Val(textDB(5)) > 0 Then
            cmdSend.Enabled = False
         End If
      Else
         cmdSend.Enabled = False
      End If
      
      If R_type = "M" Or R_type = "A" Then
         FraTMA65.Visible = True
      Else
         FraTMA65.Visible = False
      End If
      
      cmdSaveD(0).Visible = False
      cmdSaveD(1).Visible = False
      
      Cbo1.Locked = True
      
      For Each oObj In textDB
         oObj.Locked = True
      Next

      For Each oObj In Text1
         oObj.Locked = True
      Next
      Frame2.Enabled = False  'ChkS2(index), ChkS3(index)
      ChgObjEnabled ChkTMA16, False
      ChgObjEnabled ChkTMA67(0), False
      ChgObjEnabled ChkTMA67(1), False
      ChgObjEnabled ChkTMA67(2), False
      ChgObjEnabled ChkTMA69(0), False
      ChgObjEnabled ChkTMA69(1), False

      '外部呼叫隱藏按鈕
      If mbolCall = True Then
         cmdSend.Visible = False: cmdTo.Visible = False
         cmdSendMail.Visible = False
         If Len(m_NoList) <= 10 Then
            cmdPrevious.Visible = False: cmdNext.Visible = False
         End If
      Else
         Select Case R_type
            Case "U" '查名人
               If m_TMA13 = "" And iStiu = 1 Then
                  If InStr(cmdSend.Caption, "查覆完畢") > 0 Then
                     Frame1.Visible = True
                     Cbo1.Locked = False
                     cmdSaveD(0).Visible = True
                     textDB(17).Locked = False
                     textDB(15).Locked = False
                     ChgObjEnabled ChkTMA16, True
                  ElseIf InStr(cmdSend.Caption, "送出") > 0 Then
                     Frame1.Visible = False
                  End If
               End If
               cmdTo.Visible = False
               cmdSendMail.Visible = False
            Case "M" '覆核主管
               If m_TMA13 = "" Then
                  cmdSaveD(1).Visible = True
                  textDB(68).Locked = False
                  textDB(17).Locked = False 'Added by Lydia 2023/07/06 開放覆核主管可修改「申請號/審定號」
                  ChgObjEnabled ChkTMA67(0), True
                  ChgObjEnabled ChkTMA67(1), True
                  ChgObjEnabled ChkTMA67(2), True
                  ChgObjEnabled ChkTMA69(0), True
                  ChgObjEnabled ChkTMA69(1), True
               Else
                   MsgBox "委查人已撤回查名單 ！", vbCritical
                   cmdSend.Enabled = False
               End If
               cmdTo.Visible = False
               cmdSendMail.Visible = False
            Case "Q" '委查人
               If mbolCall = False And iStiu = 1 Then
                  cmdSend.Visible = True
               Else
                  cmdSend.Visible = False
               End If
         End Select
      End If

      'Modified by Lydia 2016/04/06 +已收文判斷(若已收文,明細的收文按鈕不能按,但是查覆區可以按)
      'Modified by Lydia 2019/08/12 增加創新業務組成員可互相操作
      'Modified by Lydia 2019/12/25 開放特殊設定權限
      If (Pub_StrUserSt03 = "M51" Or InStr(stIdList, textDB(8).Text) > 0 Or (bolSpecMan = True And InStr(strSpecCode, textDB(8).Text) > 0)) _
                And mbolSend = True And ShowCP09 = "" And m_TMA13 = "" Then
          cmdTo.Enabled = True
          cmdSendMail.Enabled = False
      Else
          cmdTo.Enabled = False
          'Modified by Lydia 2019/08/12 增加創新業務組成員可互相操作
          'Modified by Lydia 2019/12/25 開放特殊設定權限
          If Text1(0) & Text1(1) = "" And ShowCP09 <> "" And (InStr(stIdList, textDB(8).Text) > 0 Or (bolSpecMan = True And InStr(strSpecCode, textDB(8).Text) > 0)) Then
             cmdSendMail.Enabled = True
          Else
             cmdSendMail.Enabled = False
          End If
      End If
   End If
   
   '顯示附件(新增,刪除)
   If iStiu = 0 Or R_type = "Q" Then
      cmdAddAtt(0).Visible = False
      cmdRemAtt(0).Visible = False
   Else
      'Added by Lydia 2021/11/02 判斷覆核主管不可刪除附件; ex.T-234290的HB0040687、HB0040686查名附件不存在，推測是在覆核或核可階段刪除
                                               '另外增加刪除附件的操作者非建立人員，另外寫log(在basUpdate.PUB_TMQAFileDel)。
      If R_type = "M" Then
          cmdRemAtt(0).Visible = False
      Else '以下包含U查名人員，A電腦中心維護
          cmdAddAtt(0).Visible = True
          cmdRemAtt(0).Visible = True
      End If
   End If
   
End Sub

Private Sub FormReset()

   For Each oObj In textDB 'Table：直接操作
      oObj.Text = ""
      oObj.Tag = ""
      oObj.Locked = False
   Next
   For Each oObj In txtFM2
      oObj.Text = ""
      oObj.Tag = ""
      oObj.Locked = False
   Next
   For Each oObj In Text1
      oObj.Text = ""
      oObj.Tag = ""
      oObj.Locked = False
   Next
   For Each oObj In lblData 'Table：內容經過轉換
      oObj.Caption = ""
   Next
   For Each oObj In ChkS2
      oObj.Value = False
      oObj.Tag = ""
   Next
   For Each oObj In ChkS3
      oObj.Value = False
      oObj.Tag = ""
   Next
   ChkTMA16.Value = False
   ChkTMA16.Tag = ""
   For Each oObj In ChkTMA67
      oObj.Value = False
      oObj.Tag = ""
   Next
   For Each oObj In ChkTMA69
      oObj.Value = False
      oObj.Tag = ""
   Next
   
   Call SetCombo("0", Cbo1)
   Cbo1.Text = "": Cbo1.Tag = ""
   lstAtt(0).Clear
   
   tmpKeyPic1.Visible = False
   For Each oObj In cmdKD
      oObj.Visible = False
      oObj.Tag = ""
   Next
   
   mPrevTM1215 = ""
   cmdSend.Caption = "查覆完畢(&O)"
   cmdSend.Tag = ""
End Sub

'*****設定查覆下拉選單 or 取得查覆結果*****
Private Sub SetCombo(ByVal pType As String, ByRef cmbN As ComboBox, Optional ByRef pVAL01 As String, Optional ByRef pVAL02 As String)
Dim intX As Integer
    
   If pType = "0" Then  '預設下拉清單
      tmpArr = Empty
      tmpArr = Split(PUB_GetTMQans("2", False), ",")
      cmbN.Clear
     '增加空白
      cmbN.AddItem ""
      For intX = 0 To UBound(tmpArr)
          cmbN.AddItem Mid(tmpArr(intX), 3, Len(tmpArr(intX)) - 2)
      Next intX
   ElseIf pType = "1" Or pType = "2" Then  '轉換
      pVAL02 = ""
      If pVAL01 <> "" Then
         tmpArr = Empty
         tmpArr = Split(PUB_GetTMQans("2", False), ",")
         For intX = 0 To UBound(tmpArr)
            'Value轉文字
            If pType = 1 And Trim(Mid(Trim(tmpArr(intX)), 1, 2)) = pVAL01 Then
               pVAL02 = Trim(Mid(Trim(tmpArr(intX)), 3))
               Exit For
            End If
            '文字轉Value
            If pType = 2 And Trim(Mid(Trim(tmpArr(intX)), 3)) = pVAL01 Then
               pVAL02 = Trim(Mid(Trim(tmpArr(intX)), 1, 2))
               Exit For
            End If
         Next intX
         If pVAL02 = "" Then
            pVAL01 = ""
         End If
      End If
   End If
End Sub

'*****檢查查覆結果*****
Private Function CheckCombo(ByRef cmbN As ComboBox, Optional ByRef nKind As String = "") As Boolean
Dim iX As Integer

   CheckCombo = False
   tmpArr = Empty
  
   tmpArr = Split(PUB_GetTMQans("2", False), ",")
   For iX = 0 To UBound(tmpArr)
       If Mid(tmpArr(iX), 3, Len(tmpArr(iX)) - 2) = cmbN.Text Then
          CheckCombo = True: Exit Function
       End If
   Next iX

   '可以空白
   If cmbN.Text <> "" Then
      MsgBox "請選取正確的查覆結果", vbInformation
   Else
      CheckCombo = True
   End If
End Function


Private Sub SetGrd(Optional ByVal pReset As Boolean = False)
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer

   arrGridHeadText = Array("V", "順序", "中文", "英文", "日文", "記號")
   arrGridHeadWidth = Array(200, 500, 950, 950, 950, 950)

   MGRD1.Visible = False
   MGRD1.Cols = UBound(arrGridHeadText) + 1
   If pReset = True Then
      MGRD1.Clear
      MGRD1.Rows = 2
   End If
   
   For iRow = 0 To MGRD1.Cols - 1
      MGRD1.row = 0
      MGRD1.col = iRow
      MGRD1.Text = arrGridHeadText(iRow)
      MGRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      MGRD1.CellAlignment = flexAlignCenterCenter
   Next
   If colORD = 0 Then
      colORD = PUB_MGridGetId("順序", MGRD1)
      colWord(0) = PUB_MGridGetId("中文", MGRD1)
      colWord(1) = PUB_MGridGetId("英文", MGRD1)
      colWord(2) = PUB_MGridGetId("日文", MGRD1)
      colWord(3) = PUB_MGridGetId("記號", MGRD1)
   End If
   MGRD1.Visible = True
End Sub

Public Function QueryData() As Boolean

On Error GoTo ErrQuery:
QueryData = False

   If m_NoList = "" Or m_NoIdx < 0 Then
      Exit Function
   Else
      tmpArr = Empty
      tmpArr = Split(m_NoList, ",")
      If UBound(tmpArr) < m_NoIdx Then
         Exit Function
      Else
         strTmp1 = " SELECT A1.*,NVL(S4.ST02,'網中') AS TMA03N,S1.ST02 AS TMA08N,S2.ST02 AS TMA10N,S3.ST02 AS TMA65N,CP01,CP02,CP03,CP04,CP10,CP27,CP14,CP57,TQC07" & _
                    " FROM TMQAPPFORM A1,STAFF S1, STAFF S2 ,STAFF S3, STAFF S4, CASEPROGRESS, TMQCASEMAP " & _
                    " WHERE TMA01='" & tmpArr(m_NoIdx) & "' AND TMA08=S1.ST01(+) AND TMA10=S2.ST01(+) AND TMA65=S3.ST01(+) AND TMA03=S4.ST01(+) " & _
                    " AND TMA34=CP09(+) AND TQC02(+)=TMA34 AND TQC03(+)=TMA01 "
         intJ = 1
         Set rsAD = ClsLawReadRstMsg(intJ, strTmp1)
         If intJ = 1 Then
            iStiu = 0  '預設不可編輯
            bolModify = False: bolModCheck = False
            strMod(0) = "": strMod(1) = "": strMod(2) = ""
            If "" & rsAD.Fields("TMA14") = "" Then
               mbolSend = False
            Else
               mbolSend = True '已查覆完畢
            End If
            cmdSend.BackColor = &HC0C0FF
            '開放在查覆完畢後並且智權人員未收文的情況,可以再修改 ->排除已撤回,排除已發文
            If R_type = "U" And (Pub_StrUserSt03 = "M51" Or strUserNum = "" & rsAD.Fields("TMA10")) And "" & rsAD.Fields("TMA13") = "" And Trim("" & rsAD.Fields("CP27")) = "" Then
               '不限階段
               If Not "" & rsAD.Fields("TMA14") = "" Then
                  iStiu = 1: cmdSend.Tag = "U"
                  cmdSend.Caption = "修改完畢(&O)"
                  cmdSend.BackColor = &HFF00& '查覆完畢後再修改，按鈕變綠色
               Else
                  cmdSend.Caption = "查覆完畢(&O)"
                  If "" & rsAD.Fields("TMA14") = "" Then
                     iStiu = 1: cmdSend.Tag = "U"
                  End If
               End If
            ElseIf R_type = "M" Then
                  cmdSend.Caption = "覆核完畢(&O)"
                  iStiu = 1
                  cmdSend.Tag = "M"
            ElseIf R_type = "Q" Then
                  cmdSend.Caption = "撤　回" '申請的所有委查單尚未輸入結果，可自請撤回申請
                  If Len("" & rsAD.Fields("TMA13") & rsAD.Fields("TMA14") & rsAD.Fields("TMA65")) = 0 Then
                     iStiu = 1
                     cmdSend.Tag = "Q"
                  End If
            ElseIf R_type = "A" Then
                     cmdSend.Caption = "維護完畢(&O)"
                     iStiu = 1
                     cmdSend.Tag = "A"
            End If
            'Added by Lydia 2019/08/12 創新業務組成員可操作清單
            If R_type = "Q" And stIdList = "" Then
                stIdList = PUB_GetSalesList(textDB(8).Text, , , , , strGrpTmp1, strGrpTmp2)
                If InStr(stIdList, "W") = 0 Or Left(strGrpTmp1, 1) <> "W" Then
                    stIdList = CNULL(strUserNum) '非創新業務組
                End If
            End If
            '*****設定欄位值*****
            For Each oObj In textDB
               Select Case oObj.Index
                  Case 4, 5, 7
                     'CREATE DATE=委查日期TMA04、送出日期TMA05、回寫日期TMA07
                     oObj.Text = TransDate(Format("" & rsAD.Fields("TMA" & Format(oObj.Index, "00")), "yyyymmdd"), 1)
                  Case 9, 11, 12, 14, 31, 32
                     '分發日期TMA09、查覆期限TMA11、送出期限TMA12、查覆日期TMA14、查詢區間-起始日期TMA31、查詢區間-終止日期TMA32
                     oObj.Text = TransDate("" & rsAD.Fields("TMA" & Format(oObj.Index, "00")), 1)
                  Case Else
                     oObj.Text = "" & rsAD.Fields("TMA" & Format(oObj.Index, "00"))
                     '委查人員TMA08N、查名人TMA10N、覆核人員TMA65N
                     If InStr("08,10,65", Format(oObj.Index, "00")) > 0 Then
                        lblData(oObj.Index) = "" & rsAD.Fields("TMA" & Format(oObj.Index, "00") & "N")
                     End If
               End Select
               oObj.Tag = oObj.Text
            Next
            
            m_TMA01 = textDB(1)
            m_TMA02 = "" & rsAD.Fields("TMA02") '資料來源：人工輸入=1，網中轉入=2
            m_TMA13 = "" & rsAD.Fields("TMA13")  '委查人是否撤回
            m_TMA71 = "" & rsAD.Fields("TMA71")  '(原)查名單號/(原)查名單申請編號(114.4.14新增)
            m_TMA72 = "" & rsAD.Fields("TMA72") '委查人附件已讀
            
            '全類檢索TMA21
            If "" & rsAD.Fields("TMA21") = "Y" Then ChkS3(0).Value = 1
            ChkS3(0).Tag = "" & rsAD.Fields("TMA21")
            
            '查詢資料範圍(TMA29)：1-全部, 2-僅查本所代理
            If "" & rsAD.Fields("TMA29") = "2" Then ChkS3(1).Value = 1
            ChkS3(1).Tag = "" & rsAD.Fields("TMA29")
            
            '是否包含無效或核駁資料(TMA30)
            If "" & rsAD.Fields("TMA30") = "Y" Then ChkS3(2).Value = 1
            ChkS3(2).Tag = "" & rsAD.Fields("TMA30")
            
            '團體標章/證明標章(TMA20)：1-團體標章, 2=證明標章(113/10/4 關閉證明標章「9999」代碼)
            m_TMA20 = "" & rsAD.Fields("TMA20")
            If m_TMA20 = "1" Then ChkS2(0).Value = 1
            
            If m_TMA20 = "" Then
               LBL1(16).Visible = True:  ChkTMA16.Visible = True
               LBL1(19).Visible = False:  Cbo1.Visible = False
            Else
               LBL1(16).Visible = False: ChkTMA16.Visible = False
               LBL1(19).Visible = True: Cbo1.Visible = True
               '標章查覆結果
               strExc(2) = ""
               If "" & rsAD.Fields("TMA19") <> "" Then
                  Call SetCombo("1", Cbo1, "" & rsAD.Fields("TMA19"), strExc(2))
                  Cbo1.Text = strExc(2)
               End If
               Cbo1.Tag = strExc(2)
            End If
            
            '檢索方式TMA25：1-文字, 2-圖形, 3-文字＋圖形
            m_TMA25 = "" & rsAD.Fields("TMA25")
            lblData(25).Caption = IIf(m_TMA25 = "1", "文字", IIf(m_TMA25 = "2", "圖形", "文字＋圖形"))
            '回寫結果TMA43=(網中)商標查詢系統結果
            lblData(43).Caption = IIf("" & rsAD.Fields("TMA43") <> "", IIf("" & rsAD.Fields("TMA43") = "Y", "成功", "失敗"), "")
            
            '是否與本所近似TMA16
            If "" & rsAD.Fields("TMA16") = "Y" Then ChkTMA16.Value = 1
            ChkTMA16.Tag = "" & rsAD.Fields("TMA16")
            
            '覆核是否與本所近似TMA67
            ChkTMA67(0).Tag = "" & rsAD.Fields("TMA67")
            If "" & rsAD.Fields("TMA67") = "Y" Then  'Y=是，需進行協商流程
                ChkTMA67(0).Value = 1
            ElseIf "" & rsAD.Fields("TMA67") = "N" Then  'N=否（需確認客戶關係）
                ChkTMA67(1).Value = 1
            ElseIf "" & rsAD.Fields("TMA67") = "A" Then  'A=已排除近似
                ChkTMA67(2).Value = 1
            End If
            '協商流程結果TMA69
            ChkTMA69(0).Tag = "" & rsAD.Fields("TMA69")
            If "" & rsAD.Fields("TMA69") = "1" Then  '1.經上級核可代理
               ChkTMA69(0).Value = 1
            ElseIf "" & rsAD.Fields("TMA69") = "2" Then  '2.經上級核可先提申再補同意書
               ChkTMA69(1).Value = 1
            End If
            
            '總收文號(第一次收文=TMA34)
            FirstCP09 = "" & rsAD.Fields("TMA34")
            If "" & rsAD.Fields("TMA35") <> "" Then
               Call ChgCaseNo("" & rsAD.Fields("TMA35"), FirstCP)
            ElseIf FirstCP09 <> "" Then
               FirstCP(1) = "" & rsAD.Fields("CP01")
               FirstCP(2) = "" & rsAD.Fields("CP02")
               FirstCP(3) = "" & rsAD.Fields("CP03")
               FirstCP(4) = "" & rsAD.Fields("CP04")
               FirstCP14 = "" & rsAD.Fields("CP14")
               FirstCP27 = "" & rsAD.Fields("CP27")
               FirstCP57 = "" & rsAD.Fields("CP57")
            End If
            
            '傳入目前進度的總收文號>總收文號(第一次收文)
            If ShowCP09 <> "" Then
               strExc(0) = "select CP01,CP02,CP03,CP04,CP27,CP14,CP57,TQC07 from caseprogress,TMQCASEMAP " & _
                                 "where cp09='" & ShowCP09 & "' AND CP09=TQC02 AND TQC03='" & m_TMA01 & "' "
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  Text1(0).Text = ChangeWStringToTString("" & RsTemp.Fields("TQC07"))
                  Text1(1).Text = ChangeWStringToTString("" & RsTemp.Fields("CP27"))
                  ShowCP(1) = "" & RsTemp.Fields("CP01")
                  ShowCP(2) = "" & RsTemp.Fields("CP02")
                  ShowCP(3) = "" & RsTemp.Fields("CP03")
                  ShowCP(4) = "" & RsTemp.Fields("CP04")
                  ShowCP14 = "" & RsTemp.Fields("CP14")
                  ShowCP57 = "" & RsTemp.Fields("CP57")
                  lblData(35) = ShowCP(1) & "-" & ShowCP(2) & "-" & ShowCP(3) & "-" & ShowCP(4)
               End If
            End If
            If Trim(lblData(35).Caption) = "" And FirstCP09 <> "" Then
               ShowCP09 = FirstCP09
               Text1(0).Text = ChangeWStringToTString("" & rsAD.Fields("TQC07"))
               Text1(1).Text = ChangeWStringToTString("" & rsAD.Fields("CP27"))
               lblData(35) = FirstCP(1) & "-" & FirstCP(2) & "-" & FirstCP(3) & "-" & FirstCP(4)
            End If
            'Added by Lydia 2025/04/15 (原)查名單號/(原)查名單申請編號(114.4.14新增)
            If m_TMA71 <> "" And Left(m_TMA71, 3) < "HB4" And Left(m_TMA71, 2) <> "HM" Then
               Label1(0) = "舊查名單："
               lblData(35) = m_TMA71
            End If
            
            '圖形查詢附件序號
            strExc(1) = "select min(tmf03) mno from tmqappfile where tmf01='" & m_TMA01 & "' and nvl(tmf02,'" & m_ATMF02 & "')<>'" & m_ATMF02 & "' and nvl(tmf03,'02') < '02' and instr(upper(tmf10),'.JPG') > 0"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(1))
            If intI = 1 Then
               m_STMF03 = "" & RsTemp.Fields("mno")
            Else
               m_STMF03 = "00"
            End If

            '非標章查名：先經過送出->啟動網中查名，等網中回寫後Trigger觸發查覆期限TMA11->查名人員查看後再上查覆
            If R_type = "U" And m_TMA02 = "1" And m_TMA13 = "" Then  '1234 拿掉And iStiu = 1，查名主管看到的才會一致
               If m_TMA20 <> "" Then  '標章查名
                  cmdSend.Caption = "查覆完畢(&O)"
               Else
                  If Val(textDB(7)) = 0 Then
                     cmdSend.Caption = "送出(&O)"
                     SSTab1.Tab = 1
                  Else
                     cmdSend.Caption = "查覆完畢(&O)"
                  End If
               End If
            End If
            
            FormEnabled '欄位編輯控制 '1234 移到下面？
                         
            '查名區+卷宗區：顯示／隱藏覆核結果
            If "" & rsAD.Fields("TMA67") <> "" Or R_type = "M" Or R_type = "A" Then
               FraTMA65.Visible = True
               Me.Height = 8860
               Me.SSTab1.Height = 3900
            End If
            
            '文字查名不可輸入查名路徑
            textDB(28).Locked = True
            cmdRoute.Caption = "顯示"
            cmdRoute.Visible = False
            If m_TMA25 <> "1" Then
               If cmdRoute.Tag = "M" And Val(textDB(7)) = 0 And Val(textDB(14)) = 0 And m_TMA13 = "" Then
                  cmdRoute.Caption = "輸入"
               End If
               cmdRoute.Visible = True
            End If
            
            '圖形查名：Y=有匯入圖片
            m_TMA27 = "" & rsAD.Fields("TMA27")
            If m_TMA27 = "Y" Then
               tmpKeyPic1.Visible = True
               cmdKD(0).Visible = True
               cmdKD(1).Visible = True
               If (R_type = "U" Or R_type = "M" Or R_type = "A") And "" & rsAD.Fields("TMA05") = "" Then 'TMA05啟動(網中)商標查詢系統
                  cmdKD(2).Visible = True
               End If
               Call PUB_KillTempFile(strUserNum & "\H*." & outType)
               Call PUB_KillTempFile(strUserNum & "\H*." & LCase(outType))
               strExc(1) = m_AttachPath & "\" & m_TMA01 & m_TMA02 & m_STMF03 & "." & outType
               'P.S. 當時將舊資料匯入是將原本的附件直接匯入(TMF03=01,02)，後來定義00=圖形查詢附件
               If PUB_TMQAppFileGet(m_AttachPath, strExc(1), m_TMA01, m_TMA02, m_STMF03, strExc(4)) = False Then
                  MsgBox "無法儲存檔案[ " & strExc(1) & " ]！"
                  Exit Function
               Else
                  cmdKD(0).Tag = m_TMA01 & m_TMA02 & m_STMF03
                  cmdKD(2).Tag = strExc(4)  '載入圖片的Create ID|Date|Time
                  Set G_SeekPicColor.Picture = pvGetStdPicture(Trim(strExc(1)))
                  '固定PictureBox中的image,載入圖片後調整圖片大小
                  Call Pub_PicToObj(Trim(strExc(1)), G_SeekPicColor, tmpKeyPic1, tmpKeyImg1)
               End If
            End If
            
            'AI文字檢索1~5
            Call SetGrd(True)
            If m_TMA25 = "2" Then
               Frame3.Visible = False
            Else
               Frame3.Visible = True
               If (InStr(cmdSend.Caption, "送出") > 0 And Val(textDB(5)) = 0) Or R_type = "A" Then
                  cmdAdd.Visible = True: cmdRemove.Visible = True
                  cmdUp.Visible = True: cmdDown.Visible = True
               Else
                  cmdAdd.Visible = False: cmdRemove.Visible = False
                  cmdUp.Visible = False: cmdDown.Visible = False
               End If
               Call SWordForRead
            End If
            textCUID.Text = "CREATE : " & rsAD.Fields("TMA03N") & "  " & ChangeWStringToTDateString(Format(rsAD.Fields("TMA04"), "YYYYMMDD")) & "  " & Format(rsAD.Fields("TMA04"), "HH:MM ")
             
            '*****查名結果附件*****
            lstAtt(0).Clear
            'Modified by Lydia 2025/04/15 網中回寫缺少:TMF04,TMF06,TMF07,TMF10
            'strExc(0) = "select tmf01,tmf02,tmf03,tmf04,tmf10 from tmqappfile where tmf01='" & m_TMA01 & "' and tmf02='" & m_ATMF02 & "' " & IIf(m_STMF03 <> "", "and tmf03>'" & m_STMF03 & "' ", "") & " order by tmf03 "
            strExc(0) = "select tmf01,tmf02,tmf03,nvl(tmf04,'未知大小') tmf04,nvl(tmf10,substr(tmf09,18,16)) tmf10,tmf08 from tmqappfile where tmf01='" & m_TMA01 & "' and tmf02='" & m_ATMF02 & "' " & IIf(m_STMF03 <> "", "and tmf03>'" & m_STMF03 & "' ", "") & " order by tmf03 "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               With RsTemp
                  intJ = 0
                  Do While Not .EOF
                     strExc(1) = Dir(m_AttachPath & "\" & RsTemp.Fields("tmf10"))
                     If strExc(1) <> "" Then
                        If PUB_ChkFileOpening(m_AttachPath & "\" & RsTemp.Fields("tmf10")) = True Then
                           MsgBox m_AttachPath & "\" & RsTemp.Fields("tmf10") & vbCrLf & "檔案正在使用中（請關閉），方可繼續操作？", vbExclamation
                           Exit Function
                        End If
                        Kill m_AttachPath & "\" & RsTemp.Fields("tmf10")
                     End If
                     'Modified by Lydia 2025/04/15 網中回寫缺少
                     'strExc(1) = "" & RsTemp.Fields("tmf10") & " (" & Round(Val("" & RsTemp.Fields("tmf04")) / 1024, 2) & " KB)"
                     strExc(1) = "" & RsTemp.Fields("tmf10") & " (" & IIf("" & RsTemp.Fields("tmf04") = "未知大小", "" & RsTemp.Fields("tmf04"), Round(Val("" & RsTemp.Fields("tmf04")) / 1024, 2) & " KB") & " )" & IIf("" & RsTemp.Fields("tmf08") <> "Y", " (未讀)", "")
                     lstAtt(0).AddItem strExc(1), intJ
                     lstAtt(0).ItemData(0) = 0
                     .MoveNext
                     intJ = intJ + 1
                  Loop
               End With
            End If
            If lstAtt(0).ListCount > 0 Then SetListScroll lstAtt(0)
             '申請的所有委查單尚未輸入結果，可自請撤回申請
            If R_type = "Q" And cmdSend.Caption = "撤　回" And iStiu = 1 Then
                If lstAtt(0).ListCount > 0 Then cmdSend.Visible = False
            End If
            
         Else
            MsgBox "查無此查名單號資料！", vbCritical, "查覆明細作業"
            GoTo ExitClose
         End If
              
         QueryData = True
      End If 'If TmpArr(m_NoIdx) = ""
   End If 'If m_NoList = ""

ErrQuery:
   If Err.Number <> 0 Then
      MsgBox Err.Description
   End If
ExitClose:
    Set rsAD = Nothing

End Function

Private Sub Cbo1_LostFocus()
   If iStiu = 1 Then
      If (Trim(Cbo1.Text) = TMQ_近似T1 Or Trim(Cbo1.Text) = TMQ_近似T2) And Trim(Cbo1.Text) <> Cbo1.Tag And Trim(textDB(17).Text) = "" Then
         textDB(17).Text = mPrevTM1215
      End If
   End If
End Sub

Private Sub Cbo1_Validate(Cancel As Boolean)
   If Trim(Cbo1.Text) <> "" Then
      If CheckCombo(Cbo1) = False Then
         Cbo1.SetFocus
         Cancel = True
      End If
   End If
End Sub

Private Sub ChkTMA67_Click(Index As Integer)
Dim pIdx As Integer
   
   If ChkTMA67(Index).Value = 1 Then
      pIdx = Index
      'Y=是，需進行協商流程TMA69
      If pIdx <> 0 Then
         For Each oObj In ChkTMA69
            oObj.Value = False
         Next
      End If
   Else
      pIdx = -1
   End If
   If pIdx >= 0 Then
      For Each oObj In ChkTMA67
         If oObj.Index <> pIdx Then
            oObj.Value = False
         End If
      Next
   End If
End Sub

Private Sub ChkTMA69_Click(Index As Integer)
Dim pIdx As Integer
   
   If ChkTMA69(Index).Value = 1 Then
      pIdx = Index
      'Y=是，需進行協商流程TMA69
      ChkTMA67(0).Value = 1
   Else
      pIdx = -1
   End If
   If pIdx >= 0 Then
      For Each oObj In ChkTMA69
         If oObj.Index <> pIdx Then
            oObj.Value = False
         End If
      Next
   End If
End Sub

'結束
Private Sub cmdExit_Click()

   If IsSaveData = False Then Exit Sub
   
   '判斷不查，有修改筆數(結果為不查)
   If bolModify = True And bolModCheck = False And strMod(1) <> "" And textDB(1) <> "" Then
      If InStr(strMod(1), "不查") > 0 Then
         Call cmdSend_Click
      End If
   End If
   '結束前檢查是否有按覆核完畢(發mail通知)
   If bolChgTMA69 = True Then
      If MsgBox("覆核結果有更改，是否要繼續執行覆核完畢？", vbYesNo + vbDefaultButton1) = vbYes Then
         Call cmdSend_Click
      End If
   End If
   
   Unload Me
End Sub

Private Sub cmdKD_Click(Index As Integer)
Dim stFileName As String, stFullName As String
  
   If Index < 2 Then Screen.MousePointer = vbHourglass
   
   stFileName = m_TMA01 & m_TMA02 & m_STMF03 & "." & outType
   
   Select Case Index
      Case 0 '另開視窗
         If Not nfrm090129 Is Nothing Then
            nfrm090129.SetParent Me, m_TMA01, Val(m_TMA02), outType, m_STMF03
            nfrm090129.Show
            If nfrm090129.iStiu = 1 Then
            Else
               Unload nfrm090129
            End If
         End If
      Case 1 '下載
         stFullName = GetSaveName(stFileName)
         If stFullName <> "" Then
            If Dir(stFullName) <> "" Then
               If MsgBox("檔案[ " & stFullName & " ]已存在是否要覆蓋?？", vbYesNo + vbDefaultButton2) = vbNo Then
                  stFullName = ""
               End If
            End If
            If stFullName <> "" Then
               If PUB_TMQAppFileGet(Replace(stFullName, "\" & stFileName, ""), stFullName, m_TMA01, m_TMA02, m_STMF03) = False Then
                  MsgBox "無法儲存檔案[ " & stFullName & " ]！"
                  GoTo RunExit
               Else
                  MsgBox "下載完成！" & vbCrLf & stFullName, vbInformation, "圖形查詢附件"
               End If
            End If
         End If
          
      Case 2 '更換圖片
         If ChkIsTest = True Then Exit Sub 'Added by Lydia 2025/04/15 檢查資料是否可維護
         
         If MsgBox("是否要更換圖片？", vbYesNo + vbDefaultButton2 + vbInformation, "圖形查詢附件") = vbYes Then
            frmPic001.oCP01 = m_TMA01
            frmPic001.oCP02 = "0"
            frmPic001.oCP03 = m_TMA02
            frmPic001.oCP04 = m_STMF03
            Set frmPic001.oPic = G_SeekPicColor
            Set frmPic001.oImg = tmpKeyImg1
            Set frmPic001.UpForm = Me
            frmPic001.oRtPic = False
            frmPic001.m_TMQ = "A"  '與原查名單區別
            frmPic001.cmdOK(4).Visible = False
            frmPic001.cmdOK(5).Visible = False
            frmPic001.cmdOK(6).Visible = False
            frmPic001.cmdOK(7).Visible = False
            frmPic001.cmdOK(2).Caption = "存檔(&O)"
            frmPic001.cmdOK(3).Caption = "離開(&X)"
            frmPic001.Label11.Caption = "選擇圖片"
            frmPic001.cmdOK(0).Left = frmPic001.cmdOK(0).Left - 250
            frmPic001.cmdOK(1).Left = frmPic001.cmdOK(1).Left - 250
            frmPic001.cmdOK(2).Left = frmPic001.cmdOK(2).Left - 250
            frmPic001.cmdOK(3).Left = frmPic001.cmdOK(3).Left - 250
            frmPic001.Width = 3800
            MoveFormToCenter frmPic001
            frmPic001.SetSeekCmdok
            Unload frmpic002
            frmPic001.Show vbModal
            
            '重置圖片
            stFileName = m_AttachPath & "\" & m_TMA01 & m_TMA02 & m_STMF03 & "." & outType
            If PUB_TMQAppFileGet(m_AttachPath, stFileName, m_TMA01, m_TMA02, m_STMF03, strExc(4)) = False Then
               MsgBox "無法儲存檔案[ " & stFileName & " ]！"
               cmdKD(2).Tag = ""
               GoTo RunExit
            End If
            cmdKD(2).Tag = strExc(4) '載入圖片的Create ID|Date|Time
            Set G_SeekPicColor.Picture = pvGetStdPicture(Trim(stFileName))
            '固定PictureBox中的image,載入圖片後調整圖片大小
            Call Pub_PicToObj(Trim(stFileName), G_SeekPicColor, tmpKeyPic1, tmpKeyImg1)
            If R_type = "A" Then
               strSql = GetModETitle("，圖形查名有變更！")
               If Trim(textDB(10)) <> "" Then
                  PUB_SendMail strUserNum, Trim(textDB(10)), "", strSql, "同主旨"
               End If
            End If
         End If
   End Select

RunExit:
   Screen.MousePointer = vbDefault

End Sub

Private Sub cmdNext_Click()

  If IsSaveData = True Then
     FormReset
     m_NoIdx = m_NoIdx + 1
     If QueryData() = False Then
        If m_NoIdx > 0 Then m_NoIdx = m_NoIdx - 1
        Call cmdExit_Click
     End If
  End If
End Sub

Private Sub cmdPrevious_Click()
  
  If IsSaveData = True Then
     FormReset
     If m_NoIdx > 0 Then
        m_NoIdx = m_NoIdx - 1
        If QueryData() = False Then
           If m_NoIdx > 0 Then m_NoIdx = m_NoIdx + 1
           Call cmdExit_Click
        End If
     Else
        Call cmdExit_Click
     End If
  End If
End Sub

Private Sub cmdSaveD_Click(Index As Integer)
Dim tmpAns As String
   '合併欄位檢查和存檔
   tmpAns = SaveDetailAns(Index)
   
End Sub

'*****存檔(查覆/覆核)*****
Private Function SaveDetailAns(ByVal iType As String) As String
Dim tmpAns As String, tmpBol As Boolean
Dim strUpd As String

On Error GoTo ErrSaveDetail

   Select Case iType
      Case 0 '送出/查覆
         If m_TMA20 <> "" And Cbo1.Visible = True And Cbo1.Locked = False Then
            If CheckCombo(Cbo1) = False Then GoTo ErrSaveDetail
            If Trim(Cbo1.Text) = TMQ_近似T1 Or Trim(Cbo1.Text) = TMQ_近似T2 Then
               tmpAns = "標章查覆結果為" & Cbo1.Text
            End If
         End If
         If tmpAns = "" And Val(textDB(7)) > 0 And ChkTMA16.Value = 1 Then
            tmpAns = "查名結果與本所近似"
         End If

         If tmpAns <> "" Then
            If Trim(textDB(17).Text) = "" Then
               MsgBox "請輸入" & tmpAns & "的本所案件之申請號或審定號！", vbCritical
               textDB(17).SetFocus
               GoTo ErrSaveDetail
            Else
               Call textDB_Validate(17, tmpBol)
               If tmpBol = True Then
                  textDB(17).SetFocus
                  GoTo ErrSaveDetail
               End If
            End If
         Else
            If Trim(textDB(17).Text) <> "" Then
               MsgBox "非與本所案件近似，請勿輸入申請號或審定號！", vbCritical
               textDB(17).SetFocus
               GoTo ErrSaveDetail
            End If
         End If 'If tmpAns <> "" Then
         
         If ChkTMA16.Value = 1 Then
            strExc(1) = "Y"
         Else
            strExc(1) = ""
         End If
         If ChkTMA16.Tag = strExc(1) And Trim(Cbo1.Tag) = Cbo1.Text And textDB(17).Tag = textDB(17).Tag And textDB(15).Tag = textDB(15).Text And textDB(28).Tag = textDB(28).Text Then
            SaveDetailAns = "N"
         Else  '有異動
            strUpd = ""
            '是否與本所近似
            If strExc(1) <> ChkTMA16.Tag Then
               strUpd = strUpd & ", TMA16=" & CNULL(strExc(1))
            End If
            '標章查覆結果
            If Trim(Cbo1.Tag) <> Cbo1.Text Then
               Call SetCombo("2", Cbo1, Cbo1.Text, strExc(2))
               strUpd = strUpd & ", TMA19=" & CNULL(strExc(2))
               strMod(1) = strMod(1) & "標章查覆結果由" & CNULL(Cbo1.Tag) & "->" & CNULL(Cbo1.Text) & " ;" & vbCrLf
            End If
            '查覆近似本所申請號／審定號
            If textDB(17).Text <> textDB(17).Tag Then
               strUpd = strUpd & ", TMA17=" & CNULL(ChgSQL(textDB(17)))
               strMod(1) = strMod(1) & " 申請號／審定號由" & Replace(CNULL(textDB(17).Tag), "NULL", "空白") & "->" & CNULL(textDB(17).Text) & " ;" & vbCrLf
            End If
            '查覆(查名)意見
            If textDB(15).Text <> textDB(15).Tag Then
               strUpd = strUpd & ", TMA15=" & CNULL(ChgSQL(textDB(15)))
               strMod(1) = strMod(1) & " 查名意見由" & Replace(CNULL(textDB(15).Tag), "NULL", "空白") & "->" & CNULL(textDB(15).Text) & " ;" & vbCrLf
            End If
            If textDB(28).Text <> textDB(28).Tag And cmdRoute.Caption = "輸入" Then
               strUpd = strUpd & ", TMA28=" & CNULL(ChgSQL(textDB(28)))
               strMod(1) = strMod(1) & " 圖形路徑由" & Replace(CNULL(textDB(28).Tag), "NULL", "空白") & "->" & CNULL(textDB(28).Text) & " ;" & vbCrLf
            End If
            If mbolSend = True And R_type <> "A" Then
               '已收文案件提示(第一次收文案件)
               strMod(0) = GetModETitle
               strMod(2) = IIf(strMod(2) = "", textDB(8).Text, strMod(2))
            End If
            If strUpd <> "" Then
               cnnConnection.BeginTrans
                  strSql = "Update TMQAppForm Set " & Mid(strUpd, 2) & " Where TMA01='" & m_TMA01 & "' "
                  cnnConnection.Execute strSql
               cnnConnection.CommitTrans
               SaveDetailAns = "T"
               ChkTMA16.Tag = strExc(1)
               Cbo1.Tag = Cbo1.Text
               textDB(17).Tag = textDB(17).Text
               textDB(15).Tag = textDB(15).Text
            End If
         End If
      Case 1 '覆核
         If ChkTMA67(0).Value = 1 Then
            strExc(1) = "Y"
         ElseIf ChkTMA67(1).Value = 1 Then
            strExc(1) = "N"
         ElseIf ChkTMA67(2).Value = 1 Then
            strExc(1) = "A"
         Else
            strExc(1) = ""
         End If
         If ChkTMA69(0).Value = 1 Then
            strExc(2) = "1"
         ElseIf ChkTMA69(1).Value = 1 Then
            strExc(2) = "2"
         Else
            strExc(2) = ""
         End If
         If strExc(1) = "Y" And Trim(textDB(17)) = "" Then
            MsgBox "請輸入近似本所案件之申請號或審定號！", vbCritical
            textDB(17).SetFocus
            GoTo ErrSaveDetail
         End If
         
         'Added by Lydia 2023/07/06 開放覆核主管可修改「申請號/審定號」TMA17
         If ChkTMA67(0).Tag = strExc(1) And ChkTMA69(0).Tag = strExc(2) And textDB(17).Tag = textDB(17).Text And textDB(68).Tag = textDB(68).Text Then
            SaveDetailAns = "N"
         Else  '有異動
            strUpd = ""
            '覆核是否與本所近似或相同
            If strExc(1) <> ChkTMA67(0).Tag Then
               strUpd = strUpd & ", TMA67=" & CNULL(strExc(1))
            End If
            '協商流程結果
            If strExc(2) <> ChkTMA69(0).Tag Then
               strUpd = strUpd & ", TMA69=" & CNULL(strExc(2))
            End If
            '覆核意見
            If textDB(68).Text <> textDB(68).Tag Then
               strUpd = strUpd & ", TMA68=" & CNULL(ChgSQL(textDB(68)))
            End If
            '查覆近似本所申請號／審定號
            If textDB(17).Text <> textDB(17).Tag Then
               strUpd = strUpd & ", TMA17=" & CNULL(ChgSQL(textDB(17)))
            End If
            If strUpd <> "" Then
               strSql = "Update TMQAppForm Set " & Mid(strUpd, 2) & " Where TMA01='" & m_TMA01 & "' "
               cnnConnection.BeginTrans
                  If bolModify = True Or R_type = "A" Then Pub_SeekTbLog strSql
                  cnnConnection.Execute strSql
               cnnConnection.CommitTrans
               SaveDetailAns = "T"
               ChkTMA67(0).Tag = strExc(1)
               ChkTMA69(0).Tag = strExc(2)
               textDB(17).Tag = textDB(17).Text
               textDB(68).Tag = textDB(68).Text
               If Trim(textDB(65)) <> "" Then bolChgTMA69 = True
            End If
         End If
   End Select
   
   Exit Function
   
ErrSaveDetail:
   SaveDetailAns = "F"
   If strUpd <> "" And Err.Number <> 0 Then
      cnnConnection.RollbackTrans
      MsgBox IIf(iType = "0", "查覆區", "覆核區") & "存檔失敗:" & Err.Description
   End If
End Function

'Added by Lydia 2016/04/29
Private Sub cmdSendMail_Click()

On Error GoTo ErrHand01

   'Modified by Lydia 2019/08/12 增加創新業務組成員可互相操作
   'Modified by Lydia 2019/12/25 開放特殊設定權限
   If cmdSendMail.Visible = True Then
      strExc(1) = "N"
      If InStr(stIdList, textDB(8)) > 0 Then
          strExc(1) = ""
      '代理-總經理、A7
      ElseIf bolSpecMan = True And InStr(strSpecCode, textDB(8)) > 0 Then
          strExc(1) = ""
      End If
      If strExc(1) = "N" Then
         MsgBox "無權限！", vbCritical
         Exit Sub
      End If
   End If
   If FirstCP09 = "" Then
      MsgBox "委查單: " & m_TMA01 & " 未收文", vbCritical
      Exit Sub
   End If

   If FirstCP09 <> ShowCP09 And ShowCP09 <> "" Then
      strExc(2) = "本所案號: " & ShowCP(1) & "-" & ShowCP(2) & IIf(ShowCP(3) & ShowCP(4) = "000", "", "-" & ShowCP(3) & "-" & ShowCP(4))
   Else
      strExc(2) = "本所案號: " & FirstCP(1) & "-" & FirstCP(2) & IIf(FirstCP(3) & FirstCP(4) = "000", "", "-" & FirstCP(3) & "-" & FirstCP(4))
   End If

   If (ShowCP(1) <> "" And ShowCP57 <> "") Or (ShowCP(1) = "" And FirstCP57 <> "") Then
      MsgBox strExc(2) & " 已取消收文", vbCritical
      Exit Sub
   End If
   If (ShowCP(1) <> "" And ShowCP14 = "") Or (ShowCP(1) = "" And FirstCP14 = "") Then
      MsgBox strExc(2) & " 未分案", vbCritical
      Exit Sub
   Else
      '判斷所有查名單是否查覆完畢
      If PUB_TMACheckOver(ShowCP09) = False Then
         Exit Sub
      End If
      'Added by Lydia 2025/04/21 開放測試模式，指定測試收件者
      If strTestReceiver <> "" Then
         PUB_SendMail strUserNum, strTestReceiver, "", "原收件者：" & PUB_ReadUserData(IIf(FirstCP09 = ShowCP09 Or ShowCP09 = "", FirstCP14, ShowCP14), True) & vbCrLf & vbCrLf & strExc(2) & "案，經智權人員確認，請送件！", vbCrLf & "如主旨"
      Else
      'end 2024/04/18
         PUB_SendMail strUserNum, IIf(FirstCP09 = ShowCP09 Or ShowCP09 = "", FirstCP14, ShowCP14), "", strExc(2) & "案，經智權人員確認，請送件！", vbCrLf & "如主旨"
      End If
   End If

    '同一收文號，只通知一次
    'Memo by Lydia 2016/07/07 若有追加查名結果，可再通知
    cnnConnection.BeginTrans
       strSql = "UPDATE TMQCASEMAP SET TQC07=" & strSrvDate(1) & " WHERE TQC02=" & CNULL(ShowCP09) & " AND TQC07 IS NULL "
       cnnConnection.Execute strSql, intI
    cnnConnection.CommitTrans
   Text1(0).Text = strSrvDate(2)
   cmdSendMail.Enabled = False

   Exit Sub
ErrHand01:

   MsgBox Err.Description, vbCritical
   cnnConnection.RollbackTrans
End Sub

Private Sub cmdTo_Click()

   If IsSaveData = True Then
      If cmdSend.Enabled = False Then
         m_TMQApp = ""
         PubShowNextData
      End If
   End If
End Sub

Private Sub Form_Load()
   
   If R_type = "M" Or R_type = "A" Then
      Me.Height = 8860
      Me.SSTab1.Height = 3900
   Else
      Me.Height = 7880
      Me.SSTab1.Height = 2940
   End If
   
   MoveFormToCenter Me
   
   m_AttachPath = App.path & "\" & strUserNum
   If Dir(m_AttachPath, vbDirectory) = "" Then
      MkDir m_AttachPath
   End If
   
   Call PUB_GetTMQans("1", True)
   '開啟核可案狀態strAgree---不使用

   strPreAgree = Pub_GetSpecMan("內商查名覆核人員")

   SSTab1.Tab = 0
   textCUID.BackColor = &H8000000F
   Frame2.BackColor = &H8000000F
   Frame3.BackColor = &H8000000F
   
   FormReset
   
   '開放特殊設定權限
    If CheckLevel(strUserNum, "總經理業務工作代理人員") = True Then
       bolSpecMan = True
       strSpecCode = Pub_GetSpecMan("總經理員工編號")
   '開放專利處部份智權同仁資料給彥葶代為處理
   ElseIf CheckLevel(strUserNum, "A8") = True Then
        bolSpecMan = True
        strSpecCode = Pub_GetSpecMan("A7")
   End If
 
   Set nfrm090129 = Forms(0).GetForm("frm090129")
   
   cmdRoute.Tag = ""
   Set nfrm090131 = Forms(0).GetForm("frm090131")
   If Not nfrm090131 Is Nothing Then
      If R_type = "U" Or R_type = "A" Then
         cmdRoute.Tag = "M"
      Else
         cmdRoute.Tag = "Q"
      End If
   End If
   
   Set nfrm090129 = Forms(0).GetForm("frm090129")
   
   'Added by Lydia 2025/04/21 開放測試模式，指定測試收件者
   If strSrvDate(1) <= "20991231" Then
      strTestReceiver = Pub_GetSpecMan("協助檢查網中查名單")
   End If
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

   'Added by Lydia 2019/10/23 查覆完成後，再進入查名單明細(修改)，刪除附件按結束沒有檢查附件是否存在。
   'ex.T-221140和T-221141的查名結果附件被刪: 非發證和核駁閉卷,所以非批次刪除(TMQ20未上註記); 推測可能查覆完成後，查名人員刪除附件。
   If Val(textDB(14)) > 0 And m_TMA13 = "" And R_type <> "Q" And cmdSend.Visible = True And cmdSend.Enabled = True Then
      If Cbo1.Visible = True And InStr("無,不查", Trim(Cbo1.Text)) > 0 Then
      Else
         strSql = "select count(*) cnt from tmqappfile where tmf01='" & m_TMA01 & "' and tmf02='" & m_ATMF02 & "' and instr(upper(tmf10),'.PDF')>0 "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            If Val("" & RsTemp.Fields("cnt")) = 0 Then
                MsgBox "尚未新增查覆附件，請確認資料的正確性！", vbCritical
                Cancel = True
            End If
         End If
      End If
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim rsB As New ADODB.Recordset
Dim strB As String
   
   '查覆完畢後,再次修改需寄信給相關人員(委查人,近似本所案的智權人員)
   If bolModify = True And strMod(1) <> "" Then
       'Added by Lydia 2016/04/06 已收文加寄給承辦人(所有申請案的承辦人
       If FirstCP09 <> "" Or ShowCP09 <> "" Then
          'Modified by Lydia 2016/07/06 改TQC07
          'strB = "select distinct(nvl(cp14,'')) from caseprogress where cp09 in (select tqc02 from tmqcasemap where tqc03='" & m_TMA01 & "' and not(tqc02 is null) and tqc07 is null) and cp57 is null"
          strB = "select distinct(nvl(cp14,'')) from caseprogress where cp09 in (select tqc02 from tmqcasemap where tqc03='" & m_TMA01 & "' and not(tqc02 is null)) and cp57 is null"
          intI = 1
          Set rsB = ClsLawReadRstMsg(intI, strB)
          If intI = 1 Then
             rsB.MoveFirst
             Do While Not rsB.EOF
                If rsB(0) <> "" And InStr(strMod(2), "" & rsB(0)) = 0 Then
                   strMod(2) = strMod(2) & ";" & rsB(0)
                End If
                rsB.MoveNext
             Loop
             
          End If
       End If
       strExc(6) = vbCrLf & vbCrLf & "變更內容如下列：" & vbCrLf & vbCrLf & strMod(1)
       'Added by Lydia 2025/04/21 開放測試模式，指定測試收件者
       If strTestReceiver <> "" Then
           PUB_SendMail strUserNum, strTestReceiver, "", strMod(0), "原收件者：" & PUB_ReadUserData(strMod(2), True) & vbCrLf & vbCrLf & strExc(6)
       Else
       'end 2025/04/21
           PUB_SendMail strUserNum, strMod(2), "", strMod(0), strExc(6)
       End If
   End If

   Set frm090128_New = Nothing
   If TypeName(m_PrevForm) <> "Nothing" Then
        m_PrevForm.Show
        If m_PrevForm.Name = "frm090127_New" Then
           Call m_PrevForm.QueryData
        End If
   End If

   Set nfrm090129 = Nothing
   Set nfrm090131 = Nothing
   
   Set rsAD = Nothing
   Set rsTmp1 = Nothing
   Set m_PrevForm = Nothing
End Sub

Private Sub cmdSend_Click()
Dim bolConn As Boolean
Dim strUpd As String

   '檢查明細是否已存檔
   If IsSaveData = False Then Exit Sub

   If TxtValidate = False Then Exit Sub
   
   If ChkIsTest = True Then Exit Sub 'Added by Lydia 2025/04/15 檢查資料是否可維護
   
On Error GoTo ErrHand
   
   Select Case cmdSend.Tag
   Case "U" '查覆完畢,修改完畢(未收文前查名人員可修改)
       If ProcSaveU = True Then  '整合存檔和原本的CloseMail
          PUB_SendMailCache
       End If
      
   Case "M" '覆核完畢
       If ProcSaveM = True Then  '整合存檔和原本的CloseMail
          bolChgTMA69 = False
          
          PUB_SendMailCache
          
          Call cmdNext_Click
       End If

   Case "Q" '撤回
        'TxtValidate已包含詢問
        strExc(0) = ""
        If Left(m_TMA71, 1) = "H" And Left(m_TMA71, 2) <> "HM" Then
            strExc(1) = "select tma01,tma34 from tmqappform where tma71='" & m_TMA71 & "' "
        Else
            strExc(1) = "select tma01,tma34 from tmqappform where tma01='" & m_TMA01 & "' "
        End If
        intI = 1
        Set RsTemp = ClsLawReadRstMsg(intI, strExc(1))
        If intI = 1 Then
           RsTemp.MoveFirst
           cnnConnection.BeginTrans
           bolConn = True
           Do While Not RsTemp.EOF
              strUpd = ", TMA15=TMA15||'" & ChangeTStringToTDateString(strSrvDate(2)) & "已由" & IIf(strUserNum <> textDB(8), strUserName & "代替", "") & "委查人自行撤回'"
              If m_TMA20 <> "" Then
                 strUpd = strUpd & ", TMA19='" & TMQ_不查 & "' "
              End If
              strSql = "Update tmqappform set TMA14=to_char(sysdate,'yyyymmdd'), TMA34=NULL, TMA35=NULL, TMA36=0, TMA37=0, TMA38=0,TMA13='Y' " & strUpd & " where tma01='" & RsTemp.Fields("tma01") & "' "
              cnnConnection.Execute strSql
              If "" & RsTemp.Fields("tma34") <> "" Then
                  strSql = "delete from tmqcasemap where tqc03='" & RsTemp.Fields("tma01") & "' "
                  cnnConnection.Execute strSql
                  strSql = "delete from casepaperpdf where instr(cpp02,'" & RsTemp.Fields("tma01") & "." & TMQ_查名作業 & ".menu" & "') > 0 and cpp01='" & RsTemp.Fields("tma34") & "' "
                  cnnConnection.Execute strSql
              End If
              '查名單筆數歸零,再重新修正拿單量
              If textDB(10) <> "" Then
                 '更新當日拿單量
                  Call PUB_TMAtoTake("2", textDB(10), "", "0", False)
                 '更新前2日統計量
                  Call PUB_TMAtoTake("2", textDB(10), "", "1", False)
              End If
              strExc(0) = strExc(0) & IIf(strExc(0) <> "", ",", "") & RsTemp.Fields("tma01")
              RsTemp.MoveNext
           Loop
           cnnConnection.CommitTrans
           bolConn = False
           If textDB(10) <> "" Then
              strExc(1) = GetModETitle("，已由" & IIf(strUserNum <> textDB(8), strUserName & "代替", "") & "委查人自行撤回，請不用繼續查名！", strExc(0))
              PUB_SendMail strUserNum, textDB(10), "", strExc(1), strExc(1) & vbCrLf & "若有疑問，可向委查人詢問。"
           End If
           m_TMA13 = "Y"
           MsgBox "查名單已撤回！", vbInformation, "查名單(網中)明細作業"
        End If

        Call cmdExit_Click
               
   Case "A" '資料維護(限電腦中心)
        For Each oObj In textDB
           If oObj.Tag <> oObj.Text Then
              Select Case oObj.Index
                 Case 4, 5, 7
                   '用SQL語法更新:ex.將原本的日期和時間合併為Date型態
                   'select to_char(to_date(tmq13||' '||tmq14||'00','yyyymmdd HH24MISS'),'HH24MISS') ndate,tmq13,tmq14 from trademarkquery where tmq01='HB3060589'
                 Case 9, 11, 12, 14, 31, 32
                    strUpd = strUpd & ", TMA" & Format(oObj.Index, "00") & "=" & CNULL(DBDATE(oObj.Text))
                 Case Else
                    strUpd = strUpd & ", TMA" & Format(oObj.Index, "00") & "=" & CNULL(ChgSQL(oObj.Text))
              End Select
           End If
        Next
        
        '全類檢索TMA21
        If ChkS3(0).Tag <> IIf(ChkS3(0).Value = 1, "Y", "") Then
           strUpd = strUpd & ", TMA21=" & CNULL(IIf(ChkS3(0).Value = 1, "Y", ""))
        End If
        '查詢資料範圍(TMA29)：1-全部, 2-僅查本所代理
        If ChkS3(1).Tag <> IIf(ChkS3(1).Value = 1, "2", "1") Then
           strUpd = strUpd & ", TMA29=" & CNULL(IIf(ChkS3(1).Value = 1, "2", "1"))
        End If
        '是否包含無效或核駁資料(TMA30)
        If ChkS3(2).Tag <> IIf(ChkS3(2).Value = 1, "Y", "") Then
           strUpd = strUpd & ", TMA30=" & CNULL(IIf(ChkS3(2).Value = 1, "Y", ""))
        End If
        '團體標章/證明標章(TMA20)
        If m_TMA20 <> IIf(ChkS2(0).Value = 1, "1", "") Then
           strUpd = strUpd & ", TMA20=" & CNULL(IIf(ChkS2(0).Value = 1, "1", "1"))
        End If
        
        '標章查覆結果
        If Cbo1.Visible = True Then
           strExc(2) = ""
           Call SetCombo("2", Cbo1, strExc(2))
           If Trim(Cbo1.Tag) <> strExc(2) Then
              strUpd = strUpd & ", TMA19=" & CNULL(strExc(2))
           End If
        End If
        '[查覆]是否與本所近似TMA16
        If ChkTMA16.Tag <> IIf(ChkTMA16.Value = 1, "Y", "") Then
           strUpd = strUpd & ", TMA16=" & CNULL(IIf(ChkTMA16.Value = 1, "Y", ""))
        End If
        
        '覆核是否與本所近似TMA67
        strExc(1) = ""
        If ChkTMA67(0).Value = 1 Then strExc(1) = "Y"  'Y=是，需進行協商流程
        If ChkTMA67(1).Value = 1 Then strExc(1) = "N"  'N=否（需確認客戶關係）
        If ChkTMA67(2).Value = 1 Then strExc(1) = "A"  'A=已排除近似
        If ChkTMA67(0).Tag <> strExc(1) Then
           strUpd = strUpd & ", TMA67=" & CNULL(strExc(1))
        End If
        '協商流程結果TMA69
        strExc(1) = ""
        If ChkTMA69(0).Value = 1 Then strExc(1) = "1"  '1.經上級核可代理
        If ChkTMA69(1).Value = 1 Then strExc(1) = "2"  '2.經上級核可先提申再補同意書
        If ChkTMA69(0).Tag <> strExc(1) Then
           strUpd = strUpd & ", TMA69=" & CNULL(strExc(1))
        End If
        If strUpd <> "" Then
           strSql = "Update tmqappform set " & Mid(strUpd, 2) & " where tma01='" & m_TMA01 & "' "
           bolConn = True
           cnnConnection.BeginTrans
              Pub_SeekTbLog strSql
              cnnConnection.Execute strSql
           cnnConnection.CommitTrans
           bolConn = False
        End If
   Case Else
        MsgBox "error code！"
   End Select
   
   Exit Sub

ErrHand:
   If Err.Number <> 0 Then
      Screen.MousePointer = vbDefault
      If bolConn = True Then cnnConnection.RollbackTrans
      MsgBox " 送出失敗！" & vbCrLf & Err.Description
   End If
End Sub

Private Sub SetListScroll(oList As ListBox)
   Dim ii As Integer
   Dim lWnow As Long, lWmax As Long
   
   lWmax = 0
   For ii = 0 To oList.ListCount - 1
      lWnow = TextWidth(oList.List(ii) & " ")
      If lWnow > lWmax Then
         lWmax = lWnow
      End If
   Next
  
   If ScaleMode = vbTwips Then lWmax = lWmax / Screen.TwipsPerPixelX
   SendMessageByNum oList.hWnd, LB_SETHORIZONTALEXTENT, lWmax, 0
End Sub

'*****開啟結果附件*****
Private Sub cmdOpenAtt_Click(Index As Integer)
Dim hLocalFile As Long
Dim stFileName As String
Dim bolIsSelect As Boolean
Dim stF02 As String, stF03 As String

   bolIsSelect = False
   Screen.MousePointer = vbHourglass
   
   stFileName = lstAtt(Index).Text
   
   If stFileName = "" Then
      MsgBox "請選擇欲開啟的附件！"
   Else
      For intJ = 0 To lstAtt(Index).ListCount - 1
         If lstAtt(Index).Selected(intJ) Then
            bolIsSelect = True
            stFileName = lstAtt(Index).List(intJ)
            Call GetTMF0203(stFileName, stF02, stF03)
            If PUB_TMQAppFileGet(m_AttachPath, m_AttachPath & "\" & stFileName, m_TMA01, stF02, stF03) = False Then Exit Sub
            ShellExecute hLocalFile, "open", m_AttachPath & "\" & stFileName, vbNullString, vbNullString, 1
         End If
         
         '不限查覆完畢後,委查人開啟附件更新->已讀
    On Error GoTo ErrOpenAF
         If m_TMA72 = "" And R_type = "Q" And (InStr(stIdList, textDB(8).Text) > 0 Or (bolSpecMan = True And InStr(strSpecCode, textDB(8).Text) > 0)) Then
            cnnConnection.BeginTrans
              strSql = "UPDATE TMQAPPFILE SET TMF08='Y' WHERE TMF01='" & m_TMA01 & "' AND TMF02='" & stF02 & "' AND TMF03='" & stF03 & "' "
              cnnConnection.Execute strSql, intI
               intI = 1
               strExc(0) = "select count(*),count(TMF08) from TMQAPPFILE where TMF01='" & m_TMA01 & "' AND TMF02='" & stF02 & "' "
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  If RsTemp(0) = RsTemp(1) Then
                     strSql = "UPDATE TMQAPPFORM SET TMA72='Y' WHERE TMA01='" & m_TMA01 & "' "
                     cnnConnection.Execute strSql, intI
                     m_TMA72 = "Y"
                  End If
               End If
            cnnConnection.CommitTrans
         End If
         '清除列表中的未讀提示(本次作業中的未讀,實際上由申請者開啟,則狀態不變)
         strExc(0) = lstAtt(Index).List(intJ)
         lstAtt(Index).List(intJ) = Replace(strExc(0), " (未讀)", "")
      Next intJ
      If bolIsSelect = False Then
         MsgBox "請選擇欲開啟的附件！"
      End If
   End If
   
   Screen.MousePointer = vbDefault
   
ErrOpenAF:
   If Err.Number <> 0 Then
      MsgBox Err.Description
   End If
End Sub

'*****全選結果附件*****
Private Sub cmdSelect_Click(Index As Integer)
Dim ii As Integer, oList As ListBox
    
   Set oList = lstAtt(Index)
   If oList.ListCount = 0 Then Exit Sub
   
   For ii = 0 To oList.ListCount - 1
      lstAtt(Index).Selected(ii) = True
   Next
End Sub

'*****下載結果附件*****
Private Sub cmdSaveAtt_Click(Index As Integer)
Dim stFileName As String, stFolderPath As String, stFullName As String
Dim bMultiFile As Boolean
Dim ii As Integer, oList As ListBox
Dim stF02 As String, stF03 As String
Dim pIdx As Integer

   Screen.MousePointer = vbHourglass
   
   Set oList = lstAtt(Index)
   pIdx = -1
   stFileName = ""
   bMultiFile = False
   For ii = 0 To oList.ListCount - 1
      If oList.Selected(ii) Then
         If stFileName <> "" Then
            bMultiFile = True
            Exit For
         Else
            stFileName = oList.List(ii)
            pIdx = ii
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
                  Call GetTMF0203(stFileName, stF02, stF03)
                  stFullName = stFolderPath & stFileName
                  If stFullName <> "" Then
                     If Dir(stFullName) <> "" Then
                        If MsgBox("檔案[ " & stFileName & " ]已存在是否要覆蓋?？", vbYesNo + vbDefaultButton2) = vbNo Then
                           stFullName = ""
                        End If
                     End If
                     If stFullName <> "" Then
                        If PUB_TMQAppFileGet(stFolderPath, stFullName, m_TMA01, stF02, stF03) = False Then
                           MsgBox "無法儲存檔案[ " & stFullName & " ]！"
                           GoTo RunExit
                        'Added by Lydia 2025/08/27 因為智權常用下載PDF來看，所以下載=附件已讀 by 杜協理, 嘉雯
                        Else
                           If m_TMA72 = "" And R_type = "Q" And (InStr(stIdList, textDB(8).Text) > 0 Or (bolSpecMan = True And InStr(strSpecCode, textDB(8).Text) > 0)) Then
                              cnnConnection.BeginTrans
                                strSql = "UPDATE TMQAPPFILE SET TMF08='Y' WHERE TMF01='" & m_TMA01 & "' AND TMF02='" & stF02 & "' AND TMF03='" & stF03 & "' "
                                cnnConnection.Execute strSql, intI
                                 intI = 1
                                 strExc(0) = "select count(*),count(TMF08) from TMQAPPFILE where TMF01='" & m_TMA01 & "' AND TMF02='" & stF02 & "' "
                                 Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                                 If intI = 1 Then
                                    If RsTemp(0) = RsTemp(1) Then
                                       strSql = "UPDATE TMQAPPFORM SET TMA72='Y' WHERE TMA01='" & m_TMA01 & "' "
                                       cnnConnection.Execute strSql, intI
                                       m_TMA72 = "Y"
                                    End If
                                 End If
                              cnnConnection.CommitTrans
                           End If
                          '清除列表中的未讀提示(本次作業中的未讀,實際上由申請者開啟,則狀態不變)
                           strExc(0) = lstAtt(Index).List(ii)
                           lstAtt(Index).List(ii) = Replace(strExc(0), " (未讀)", "")
                        End If
                        'end 2025/08/27
                     End If
                  End If
               End If
            Next
         End If
      Else
         Call GetTMF0203(stFileName, stF02, stF03)
         stFullName = GetSaveName(stFileName)
         If stFullName <> "" Then
           If Dir(stFullName) <> "" Then
             If MsgBox("檔案[ " & stFullName & " ]已存在是否要覆蓋?？", vbYesNo + vbDefaultButton2) = vbNo Then
                stFullName = ""
             End If
           End If
           If stFullName <> "" Then
               If PUB_TMQAppFileGet(Replace(stFullName, "\" & stFileName, ""), stFullName, m_TMA01, stF02, stF03) = False Then
                  MsgBox "無法儲存檔案[ " & stFullName & " ]！"
                  GoTo RunExit
               'Added by Lydia 2025/08/27 因為智權常用下載PDF來看，所以下載=附件已讀 by 杜協理, 嘉雯
               Else
                   If m_TMA72 = "" And R_type = "Q" And (InStr(stIdList, textDB(8).Text) > 0 Or (bolSpecMan = True And InStr(strSpecCode, textDB(8).Text) > 0)) Then
                      cnnConnection.BeginTrans
                        strSql = "UPDATE TMQAPPFILE SET TMF08='Y' WHERE TMF01='" & m_TMA01 & "' AND TMF02='" & stF02 & "' AND TMF03='" & stF03 & "' "
                        cnnConnection.Execute strSql, intI
                         intI = 1
                         strExc(0) = "select count(*),count(TMF08) from TMQAPPFILE where TMF01='" & m_TMA01 & "' AND TMF02='" & stF02 & "' "
                         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                         If intI = 1 Then
                            If RsTemp(0) = RsTemp(1) Then
                               strSql = "UPDATE TMQAPPFORM SET TMA72='Y' WHERE TMA01='" & m_TMA01 & "' "
                               cnnConnection.Execute strSql, intI
                               m_TMA72 = "Y"
                            End If
                         End If
                      cnnConnection.CommitTrans
                   End If
                   '清除列表中的未讀提示(本次作業中的未讀,實際上由申請者開啟,則狀態不變)
                   strExc(0) = lstAtt(Index).List(pIdx)
                   lstAtt(Index).List(pIdx) = Replace(strExc(0), " (未讀)", "")
               'end 2025/08/27
               End If
           End If
         End If
      End If
      
      If stFullName <> "" Then
         MsgBox "下載完成！"
      End If
   End If
RunExit:
   Screen.MousePointer = vbDefault
End Sub

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

'*****取得附件的類別和流水號*****
Private Sub GetTMF0203(ByRef pFileName As String, Optional ByRef pTMF02 As String, Optional ByRef pTMF03 As String)
   
   pTMF02 = "": pTMF03 = ""
   If pFileName = "" Then Exit Sub
            
   '讀取檔案名稱
   If InStr(pFileName, " (") > 0 Then
      pFileName = Left(pFileName, InStr(pFileName, " (") - 1)
   End If
   If InStr(pFileName, "\") > 0 Then
      pFileName = Mid(pFileName, InStrRev(pFileName, "\") - 1)
   End If
   If InStr(pFileName, ".") > 0 Then
      strTmp1 = Mid(pFileName, 1, InStrRev(pFileName, ".") - 1)
   Else
      strTmp1 = pFileName
   End If
   pTMF02 = Mid(strTmp1, Len(strTmp1) - 2, 1)
   pTMF03 = Mid(strTmp1, Len(strTmp1) - 1)
End Sub

'*****新增結果附件*****
Private Sub cmdAddAtt_Click(Index As Integer)
Dim stFileName As String, NowF03 As String
Dim sFile
Dim ii As Integer
Dim fs, f, s
Dim inX As Integer

    If ChkIsTest = True Then Exit Sub 'Added by Lydia 2025/04/15 檢查資料是否可維護
    
    strExc(0) = "select nvl(max(tmf03),0) +1 as mno from tmqappfile where tmf01='" & m_TMA01 & "' and tmf03 >'" & m_STMF03 & "' "
    intI = 1
    Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
    If intI = 1 Then
       NowF03 = Format(RsTemp(0), "00")
    End If
    If Val(NowF03) > fileMax Then
        MsgBox "附件的數量不可超過" & fileMax & "個！", vbCritical
        Exit Sub
    End If
    
On Error GoTo ErrHnd
    stFileName = "*.PDF"
   
    With CommonDialog1
       .CancelError = True
       .FileName = stFileName
       .Filter = "All Files (*.PDF)|*.PDF"
       If GetSetting("TAIE", TMQ_查名作業, UCase(Me.Name) & "Dir", "") <> "" Then
           .InitDir = GetSetting("TAIE", TMQ_查名作業, UCase(Me.Name) & "Dir", "")
       Else
           .InitDir = PUB_Getdesktop
       End If
       .MaxFileSize = 3000
       '允許多選
       .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
       .ShowOpen
       If .FileName <> "" Then
           sFile = Split(.FileName, ChrW$(0))
           '記錄路徑
           SaveSetting "TAIE", TMQ_查名作業, UCase(Me.Name) & "Dir", Left(sFile(0), InStrRev(sFile(0), "\") - 1)
           '多選
           If UBound(sFile) > 1 Then
              inX = 1
           '單選
           Else
              inX = 0
           End If
           For ii = inX To UBound(sFile)
               If InStr(CStr(sFile(ii)), "#") > 0 Then
                  MsgBox CStr(sFile(ii)) & vbCrLf & vbCrLf & "【#】符號為系統保留字，不可使用於檔案命名"
                  Exit Sub
               End If
               If UCase(Right(CStr(sFile(ii)), 4)) <> UCase(".pdf") Then
                  MsgBox "只能新增PDF檔！"
                  Exit Sub
               End If
               
               If InStr(sFile(ii), "\") > 0 Then
                  stFileName = sFile(ii)
               Else
                  stFileName = sFile(0) & "\" & sFile(ii)
               End If
               Set fs = CreateObject("Scripting.FileSystemObject")
               Set f = fs.GetFile(stFileName)
               If f.Size = 0 Then
                  ShowMsg sFile(ii) & MsgText(9221)
                  Exit Sub
               ElseIf f.Size > 5242880 Then
                  If MsgBox("檔案過大（容量超過5MB），確認是否要傳送？", vbYesNo, "警告") = vbNo Then
                     Exit Sub
                  End If
               End If
               If Val(NowF03) > fileMax Then
                   MsgBox "附件的數量不可超過" & fileMax & "個！", vbCritical
                   Exit Sub
               End If
               '允許多選檔案
               If AddListX(lstAtt(Index), NowF03, Format(f.Size, "0")) = True Then
                  If PUB_TMQAppFileSave(False, m_TMA01, m_ATMF02, NowF03, stFileName) = True Then
                     If Pub_StrUserSt03 <> "M51" Then
                        SetAttr stFileName, vbNormal
                        Kill stFileName
                     End If
                     
                     If mbolSend = True And R_type <> "A" Then
                         bolModify = True
                         strMod(0) = GetModETitle
                         strMod(2) = IIf(strMod(2) = "", textDB(8).Text, strMod(2))
                         strMod(1) = strMod(1) & " 結果附件 " & m_TMA01 & m_ATMF02 & NowF03 & ".PDF有所變更" & " ;" & vbCrLf
                     End If
                     
                     NowF03 = Format(Val(NowF03) + 1, "00")
                  End If
               End If
           Next
       Else
           Exit Sub
       End If
   End With
   Exit Sub
ErrHnd:
   If Err.Number <> 0 Then
      MsgBox Err.Description
   End If

End Sub

'*****刪除結果附件*****
Private Sub cmdRemAtt_Click(Index As Integer)
Dim stFileName As String
Dim ii As Integer
Dim stF02 As String, stF03 As String

   If lstAtt(Index).ListCount = 0 Then Exit Sub
   
   If ChkIsTest = True Then Exit Sub 'Added by Lydia 2025/04/15 檢查資料是否可維護
   
   stFileName = ""
   ii = 0
   Screen.MousePointer = vbHourglass
   Do While ii < lstAtt(Index).ListCount
      If lstAtt(Index).Selected(ii) = True Then
         stFileName = lstAtt(Index).List(ii)
         Call GetTMF0203(stFileName, stF02, stF03)
         If MsgBox("確定要刪除" & stFileName & "？", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbNo Then GoTo EXITSUB
         '刪除FTP檔案
         If PUB_TMQAppFileDel(m_TMA01, stF02, stF03) = False Then
         End If
         lstAtt(Index).RemoveItem ii
         SetListScroll lstAtt(Index)
         ii = ii - 1
      End If
      ii = ii + 1
   Loop
   
EXITSUB:
   Screen.MousePointer = vbDefault
End Sub

Private Function AddListX(oList As ListBox, ByVal pF03 As String, Optional pFlen As String, Optional pF08 As String) As Boolean
   Dim idx As Integer, bFound As Boolean, stFileName As String
      
    If oList.ListCount > 0 Then
        For idx = 0 To oList.ListCount - 1
           stFileName = oList.List(idx)
           If stFileName <> "" Then
               If InStr(stFileName, m_TMA01 & m_ATMF02 & pF03) > 0 Then
                   MsgBox "附件 " & stFileName & " 已存在！"
                   AddListX = False
                   bFound = True
                   Exit For
               End If
           End If
        Next
    End If
    
    strExc(6) = GetAFName(pF03, pFlen, pF08)
    idx = oList.ListCount
    If bFound = False And strExc(6) <> "" Then
       oList.AddItem strExc(6), idx
       SetListScroll oList
       AddListX = True
    End If

End Function

'*****取得查名附件檔名*****
Private Function GetAFName(ByVal pF03 As String, Optional pFlen As String, Optional pF08 As String) As String

   If m_TMA01 <> "" And m_ATMF02 <> "" Then
       GetAFName = m_TMA01 & m_ATMF02 & pF03 & ".PDF"
       If pFlen <> "" Then GetAFName = GetAFName & " (" & Round(pFlen / 1024, 2) & " KB)" & IIf(pF08 <> "Y", " (未讀)", "")
   End If
End Function

'*****取得/記錄檔案來源*****
Private Function GetSaveName(ByVal pFileName As String) As String
Dim sFile

On Error GoTo ErrHnd
         
   With CommonDialog1
      .CancelError = True
      .FileName = pFileName
      .Filter = "All Files (*.*)|*.*"
      If GetSetting("TAIE", TMQ_查名作業, UCase(Me.Name) & "Dir", "") <> "" Then
         .InitDir = GetSetting("TAIE", TMQ_查名作業, UCase(Me.Name) & "Dir", "")
      Else
         .InitDir = PUB_Getdesktop
      End If
      .MaxFileSize = 3000
      .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
   End With
    CommonDialog1.ShowSave
    If CommonDialog1.FileName <> "" Then
        sFile = Split(CommonDialog1.FileName, ChrW$(0))
        '記錄路徑
        SaveSetting "TAIE", TMQ_查名作業, UCase(Me.Name) & "Dir", Left(sFile(0), InStrRev(sFile(0), "\") - 1)
        GetSaveName = CommonDialog1.FileName
    End If
   Exit Function
   
ErrHnd:
   If Err.Number <> 0 Then
      MsgBox Err.Description
   End If
End Function

Public Sub PubShowNextData()
Dim i As Integer, j As Integer
Dim mLoad As Boolean
Dim sPath As String
Dim APKind As String
Dim oRunform As Form
   
   Set oRunform = frm090801_New

   If m_TMQApp <> "" Then
      '從接洽單回來
      If mbolCall = True Then
         Unload Me
      End If
   Else
      Me.Enabled = False: mLoad = False
      Screen.MousePointer = vbHourglass
      APKind = m_TMA01
      oRunform.bolExternalCall = True '記錄是外部程式呼叫使用
      oRunform.SetParent Me
      oRunform.Show
      oRunform.Option1(0).Value = True '新案
      oRunform.Text1(6) = "T" '商標案
      Call oRunform.Text1_LostFocus(9)
   
      If m_TMA25 <> "2" Then  '文字、文字+圖形
         oRunform.opt1(0).Value = True
         'oRunform.PicText = textDB(26) 'Mark by Lydia 2024/10/07 商標文字欄位中，勿直接帶入文字，以留空方式讓智權人員填寫---杜協理
      Else
         mLoad = True
         APKind = cmdKD(0).Tag
         sPath = Dir(m_AttachPath & "\" & APKind & "*.*")
         If sPath = "" Then
            mLoad = PUB_TMQAppFileGet(m_AttachPath, m_AttachPath & "\" & APKind & "." & outType, m_TMA01, m_TMA02, m_STMF03)
         Else
            sPath = m_AttachPath & "\" & sPath
         End If
         If mLoad = True Then
            oRunform.opt1(1).Value = True
            oRunform.optColor(0).Value = True
            Call oRunform.PicToObj(sPath)
         End If
      End If
      
      m_TMQApp = m_TMA01
      oRunform.cmdTMQ.Tag = m_TMA01
      oRunform.Combo1(0).Text = "000" & " " & GetPrjNationName("000")
      '設定案件性質
      Call oRunform.Text1_LostFocus(6)
      Call oRunform.QueryTMQ
      'Added by Lydia 2016/07/12 TS案無商標種類
      If oRunform.Text1(6) = "TS" Then
       
      ElseIf oRunform.Text1(6) = "T" Then
         oRunform.Combo6.ListIndex = 0 'Added by Lydia 2016/05/30 接洽單的商標種類
      End If
       
      oRunform.bolExternalCall = False '還原預設值
      Screen.MousePointer = vbDefault
      Me.Enabled = True
      Me.Hide
   End If  '-----If m_TMQApp <> "" Then


End Sub

''*****檢查是否已收文***** '1234
''因為查名人進入畫面時尚未收文,當要上傳附件時可能要檢查
'Private Sub CheckTMA34isExists()
'Dim strRead As String
'Dim intB As Integer
'
'   If FirstCP09 = "" Then
'      strRead = "SELECT TMA01,TMA34,CP01,CP02,CP03,CP04,CP10,CP27,CP14,CP57 FROM TMQAppForm,CASEPROGRESS WHERE TMA01='" & m_TMA01 & "' AND TMA34=CP09(+) "
'      intB = 1
'      Set RsTemp = ClsLawReadRstMsg(intB, strRead)
'      If intB = 1 Then
'         If "" & RsTemp.Fields("TMA34") <> "" Then
'             Text1(1).Text = ChangeWStringToTString("" & RsTemp.Fields("CP27"))
'             '櫃台收文號
'             FirstCP09 = "" & RsTemp.Fields("TMA34")
'             FirstCP(1) = "" & RsTemp.Fields("CP01")
'             FirstCP(2) = "" & RsTemp.Fields("CP02")
'             FirstCP(3) = "" & RsTemp.Fields("CP03")
'             FirstCP(4) = "" & RsTemp.Fields("CP04")
'             FirstCP14 = "" & RsTemp.Fields("CP14")
'             FirstCP57 = "" & RsTemp.Fields("CP57")
'             '目前進度的總收文號
'             ShowCP09 = FirstCP09
'         End If
'      End If
'   End If
'End Sub

'*****ShellExecute程式無法直接開啟PDF檔的問題*****
'在網路上找的資料，應該是.pdf檔副檔名在本機註冊機碼中，預設開啟的程式路徑有問題。
Private Sub OpenDocument(sFile As String, sPath As String)
Dim sResult As String
Dim lSuccess As Long, lPos As Long
sResult = Space$(MAX_PATH)

 lSuccess = FindExecutable(sFile, sPath, sResult)
Select Case lSuccess
    Case ERROR_FILE_NO_ASSOCIATION
        If Right$(sFile, 3) = "pdf" Then
           MsgBox "You must have a PDF viewer such as Acrobat Reader to view pdf files."
        Else
           MsgBox "There is no registered program to open the selected file." & vbCrLf & sFile
        End If
    Case ERROR_FILE_NOT_FOUND: MsgBox "File not found: " & sFile
    Case ERROR_PATH_NOT_FOUND: MsgBox "Path not found: " & sPath
    Case ERROR_BAD_FORMAT: MsgBox "Bad format."
    Case Is >= ERROR_FILE_SUCCESS:
        lPos = InStr(sResult, Chr$(0))
        If lPos Then sResult = Left$(sResult, lPos - 1)
        MsgBox "PDF預設開啟程式:" & sResult
        
        SHELL sResult & " " & sFile, vbMaximizedFocus
End Select

'如何查看某個文件是和誰相關聯呢？例如：.txt是由哪個程式開啟，
'1.查[HKEY_CLASSES_ROOT\.txt]
'取預設值，如本人電腦預設值為 "txtfile"
'2.查[HKEY_CLASSES_ROOT\txtfile\shell\open\command]
'取預設值，如本人電腦預設值為 "C:\WINDOWS\NOTEPAD.EXE %1"
'如此可知.txt 是內定由NotePad.exe所執行。
'註：若step 1.取得的預設值是 "xxxx"，則step 2.便是查
'[HKEY_CLASSES_ROOT\xxxx\shell\open\command] 的預設值
End Sub

'取得人員請假的職代
Private Function GetDutyList(ByVal stIdList As String) As String

Dim inX As Integer
Dim stTmp1 As String
      
    GetDutyList = ""
    If stIdList <> "" Then
        tmpArr = Split(stIdList, ";")
        For inX = 0 To UBound(tmpArr)
            If Trim(tmpArr(inX)) <> "" Then
                stTmp1 = GetCaseDutyAgent(tmpArr(inX), "", False, , True, "A") 'A.指定抓全部職代
                If stTmp1 <> "" Then
                    GetDutyList = GetDutyList & ";" & stTmp1
                End If
            End If
        Next inX
    End If
    If GetDutyList <> "" Then GetDutyList = Mid(GetDutyList, 2)
End Function

'*****圖形路徑*****
Private Sub cmdRoute_Click()

   If Not nfrm090131 Is Nothing Then
      nfrm090131.SetParent IIf(cmdRoute.Caption = "顯示", "Q", "M"), Me, m_TMA01, textDB(28)
      nfrm090131.Show vbModal
   End If
   
End Sub

Public Sub SetData(ByVal pInputVal As String)
   textDB(28).Text = pInputVal
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textDB_Change(Index As Integer)
   Select Case Index
      Case 15, 68
         PUB_RefreshText textDB(Index)
   End Select
End Sub

Private Sub textDB_GotFocus(Index As Integer)
   TextInverse textDB(Index)
End Sub

Private Sub textDB_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   Select Case Index
      Case 1, 6, 8, 10, 17, 28, 40, 42, 65
         KeyAscii = UpperCase(KeyAscii)
      Case 4, 9, 11, 12, 14, 31, 32, 36, 37, 38, 22, 23, 24
         KeyAscii = Pub_NumAscii(KeyAscii)
      Case Else
   End Select
End Sub

Private Sub textDB_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = 2 Then Forms(0).PopupMenu2 textDB(Index)
End Sub

Private Sub textDB_Validate(Index As Integer, Cancel As Boolean)
Dim strTmpA As String

   Select Case Index
      Case 8, 10, 65  '委查人員TMA08、查名人TMA10、覆核人員TMA65
         If textDB(Index).Tag <> textDB(Index).Text Then
            If Index = 65 And InStr(strPreAgree, textDB(Index).Text) = 0 And textDB(Index).Text <> "" Then
               MsgBox "請輸入內商查名覆核人員的員工編號！", vbCritical
               GoTo EXITSUB
            End If
            strTmpA = GetStaffName(textDB(Index), True)
            lblData(Index).Caption = strTmpA
         Else
            lblData(Index).Caption = ""
            If Index = 8 Or Index = 10 Then
               MsgBox IIf(Index = 8, "委查人", "查名人") & "不可空白！", vbCritical
            End If
         End If
      Case 4, 9, 11, 14, 31, 32  '委查日期TMA04,分發日期TMA09、查覆期限TMA11、送出期限TMA12、查覆日期TMA14、查詢區間-起始日期TMA31、查詢區間-終止日期TMA32
         If textDB(Index).Tag = textDB(Index).Text Then Exit Sub
         If textDB(Index).Text <> "" Then
            If CheckIsTaiwanDate(textDB(Index).Text) = False Then
               MsgBox "請輸入民國年月日！", vbCritical
               GoTo EXITSUB
            End If
         End If
      Case 22, 23, 24 '類別TMA22,組群TMA23,3519組群TMA24
         If textDB(Index).Tag = textDB(Index).Text Then Exit Sub
         
         If cmdSend.Tag = "A" Then
            textDB(Index).Text = Replace(textDB(Index).Text, ".", ",") '組群間隔置換為","
            strTmpA = PUB_RepToOneSpace(PUB_StringFilter(textDB(Index).Text))   '清除字串中的enter & 清除連續空白
            textDB(Index).Text = IIf(Right(strTmpA, 1) = ",", Mid(strTmpA, 1, Len(strTmpA) - 1), strTmpA)
            If Pub_ChkTMQCisExist(Me.Name, textDB(Index), IIf(Index = 23, "1", "2"), IIf(m_TMA25 <> "2", "W", "P")) = False Then
               GoTo EXITSUB
            End If
         End If
      Case 18 '客戶名稱TMA18
         If textDB(Index).Text = "" Then
            MsgBox "客戶名稱不可空白！", vbCritical
            GoTo EXITSUB
         End If
      Case 28  '查名路徑TMA28
         If textDB(Index).Tag = textDB(Index).Text Then Exit Sub
         textDB(Index).Text = Replace(textDB(Index).Text, ".", ",")
         strTmpA = PUB_RepToOneSpace(PUB_StringFilter(textDB(Index).Text))   '清除字串中的enter & 清除連續空白
         textDB(Index).Text = IIf(Right(strTmpA, 1) = ",", Mid(strTmpA, 1, Len(strTmpA) - 1), strTmpA)
      Case 17, 40, 42  'TMA17查覆近似本所申請號／審定號,TMA40文字檢索近似本所申請號/審定號(註冊號),TMA42圖形檢索近似本所申請號/審定號(註冊號)
         If textDB(Index).Tag = textDB(Index).Text Or Trim(textDB(Index)) = "" Then Exit Sub
         If InStr("U,M,A", R_type) = 0 Then Exit Sub '排除非維護模式
         
         textDB(Index).Text = Replace(textDB(Index).Text, ".", ",")
         strTmpA = textDB(Index).Text
         If strTmpA <> "" Then
             tmpArr = Split(strTmpA, ",")
             '檢查申請號/審定號
             For intJ = 0 To UBound(tmpArr)
               'Modified by Lydia 2017/08/28 查名若發現與統一公司商標近似情形時，仍列為客戶間利益衝突案件，於申請號/審定號前加上P
               'Memo by Lydia 2022/07/12 排除P開頭: 因為近似商標權人若為統一企業，即使代理人非為本所，也視為本所代理案件。
               If Trim(tmpArr(intJ)) <> "" And Left(Trim(tmpArr(intJ)), 1) <> "P" Then
                 intI = 1
                 'Modified by Lydia 2016/05/05 閉卷有可能復活,拿掉tm29
                 'Modified by Lydia 2017/03/15 閉卷改成彈訊息和增加意見(備註)
                 strExc(0) = "select tm01,tm02,tm03,tm04,tm29,tm57 from trademark where tm10='000' and (tm12='" & Trim(tmpArr(intJ)) & "' or tm15='" & Trim(tmpArr(intJ)) & "') "
                 Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                 If intI = 0 Then
                    MsgBox tmpArr(intJ) & " 查無相關本所案號，請確認！！" & vbCrLf & "統一公司案件，請在申請號/審定號數前加上P", vbExclamation + vbOKOnly
                    'textdb(Index).SetFocus  'Mark by Lydia 2022/07/12 會造成無法跳出
                    Cancel = True
                    Exit Sub
                 Else
                    'Added by Lydia 2017/03/15 閉卷改成彈訊息和增加意見(備註)
                    If Trim("" & RsTemp.Fields("tm57")) <> "" Or Trim("" & RsTemp.Fields("tm29")) <> "" Then
                        If InStr(textDB(Index - 1), tmpArr(intJ)) = 0 Then
                           strExc(1) = "申請號/審定號:" & tmpArr(intJ) & IIf(Trim("" & RsTemp.Fields("tm57")) <> "", " 已銷卷", " 已閉卷")
                           MsgBox strExc(1) & "！", vbExclamation + vbOKOnly
                           textDB(Index - 1).Text = textDB(Index - 1).Text & IIf(textDB(Index - 1).Text <> "", ";", "") & strExc(1)
                        End If
                    End If
                 End If
               End If
             Next intJ
         End If
      Case Else
   End Select
   '檢查長度
   If textDB(Index).MaxLength > 0 Then
      If Not CheckLengthIsOK(textDB(Index), textDB(Index).MaxLength) Then
         GoTo EXITSUB
      End If
   End If
   Exit Sub
   
EXITSUB:
   Cancel = True
   textDB(Index).SetFocus
   textDB_GotFocus Index
End Sub

'*****變更項目的底色*****
Private Sub ChgObjEnabled(ByRef pTarget As Object, ByVal pEnabled As Boolean)
   
   pTarget.Enabled = pEnabled
   If pEnabled = True Then
      pTarget.BackColor = &H8000000F
   Else
      pTarget.BackColor = &H80000005
   End If
End Sub

'*****取得修改查名結果Email的主旨
Private Function GetModETitle(Optional ByVal pSpecTitle As String, Optional ByVal pNoList As String) As String
   GetModETitle = "「" & textDB(18) & "」" & "(" & lblData(25) & ")" & " 查名單: " & IIf(pNoList = "", textDB(1), pNoList) & _
                  IIf(FirstCP09 <> "", "，已收文案件" & FirstCP(1) & "-" & FirstCP(2) & IIf(FirstCP(3) & FirstCP(4) <> "000", "-" & FirstCP(3) & "-" & FirstCP(4), ""), "") & _
                  IIf(pSpecTitle = "", "，查名結果有變更！", pSpecTitle)
End Function

Private Sub txtFM2_Change(Index As Integer)
   PUB_RefreshText txtFM2(Index)
End Sub

Private Sub txtFM2_GotFocus(Index As Integer)
   TextInverse txtFM2(Index)
End Sub

Private Sub txtFM2_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = 2 Then Forms(0).PopupMenu2 txtFM2(Index)
End Sub

'*****AI檢索文字Grid*****
Private Sub MGRD1_Click()
Dim lngColor As Long
   With MGRD1
       If .MouseRow > 0 Then
          lngColor = &H80000005
          '----單選
          GridClick MGRD1, intLastRow, 0, 0, 0, "V", lngColor
          If "" & .TextMatrix(intLastRow, 0) = "V" Then
             For intI = 0 To 3
                txtFM2(intI).Text = "" & .TextMatrix(intLastRow, colWord(intI))
             Next
             txtFM2(0).Tag = "" & .TextMatrix(intLastRow, colORD)
          Else
             For intI = 0 To 3
                txtFM2(intI).Text = ""
             Next
             txtFM2(0).Tag = ""
          End If
       End If
   End With
End Sub

'*****新增AI檢索文字*****
Private Sub cmdAdd_Click()
   If Trim(txtFM2(0)) & Trim(txtFM2(1)) & Trim(txtFM2(2)) & Trim(txtFM2(3)) = "" Then
      MsgBox "請輸入檢索中文、英文、日文或記號！", vbExclamation, "檢索文字檢查"
      txtFM2(0).SetFocus
      txtFM2_GotFocus 0
   Else
      If MGRD1.Rows >= 6 Then
         MsgBox "檢索文字最多5組！", vbExclamation, "檢索文字檢查"
      Else
         strTmp1 = IIf(txtFM2(0).Tag <> "", txtFM2(0).Tag, IIf(Trim("" & MGRD1.TextMatrix(MGRD1.Rows - 1, colORD)) = "", "1", Val("" & MGRD1.TextMatrix(MGRD1.Rows - 1, colORD)) + 1))
         Call SWordForUpdate(strTmp1, "ADD")
      End If
   End If
End Sub

'*****刪除AI檢索文字*****
Private Sub cmdRemove_Click()
   If txtFM2(0).Tag = "" Then
      MsgBox "請勾選檢索文字項目！", vbExclamation, "檢索文字檢查"
   Else
      Call SWordForUpdate(txtFM2(0).Tag, "DEL")
   End If
End Sub

Private Sub cmdDown_Click()
   If txtFM2(0).Tag = "" Then
      MsgBox "請勾選檢索文字項目！", vbExclamation, "檢索文字檢查"
   Else
      If Val(txtFM2(0).Tag) = MGRD1.Rows - 1 Then
         MsgBox "檢索文字已在最後一筆！", vbExclamation, "檢索文字檢查"
      Else
         Call SWordForUpdate(txtFM2(0).Tag, "DOWN")
      End If
   End If
End Sub

Private Sub cmdUp_Click()
   If txtFM2(0).Tag = "" Then
      MsgBox "請勾選檢索文字項目！", vbExclamation, "檢索文字檢查"
   Else
      If Val(txtFM2(0).Tag) = 1 Then
         MsgBox "檢索文字已在第一筆！", vbExclamation, "檢索文字檢查"
      Else
         Call SWordForUpdate(txtFM2(0).Tag, "UP")
      End If
   End If
End Sub

'*****讀取AI檢索文字*****
Private Sub SWordForRead()

   SetGrd True
   strSql = "SELECT '' as V,'1' AS ORD1,TMA45 as NEW1,TMA46 as NEW2,TMA47 as NEW3,TMA48 as NEW4 FROM TMQAPPFORM WHERE TMA01='" & m_TMA01 & "' AND TMA25<>'2' AND length(ltrim(TMA45||TMA46||TMA47||TMA48)) > 0 " & _
         "UNION SELECT '' as V,'2' AS ORD1,TMA49 as NEW1,TMA50 as NEW2,TMA51 as NEW3,TMA52 as NEW4 FROM TMQAPPFORM WHERE TMA01='" & m_TMA01 & "'  AND TMA25<>'2' AND length(ltrim(TMA49||TMA50||TMA51||TMA52)) > 0 " & _
         "UNION SELECT '' as V,'3' AS ORD1,TMA53 as NEW1,TMA54 as NEW2,TMA55 as NEW3,TMA56 as NEW4 FROM TMQAPPFORM WHERE TMA01='" & m_TMA01 & "'  AND TMA25<>'2' AND length(ltrim(TMA53||TMA54||TMA55||TMA56)) > 0 " & _
         "UNION SELECT '' as V,'4' AS ORD1,TMA57 as NEW1,TMA58 as NEW2,TMA59 as NEW3,TMA60 as NEW4 FROM TMQAPPFORM WHERE TMA01='" & m_TMA01 & "'  AND TMA25<>'2' AND length(ltrim(TMA57||TMA58||TMA59||TMA60)) > 0 " & _
         "UNION SELECT '' as V,'5' AS ORD1,TMA61 as NEW1,TMA62 as NEW2,TMA63 as NEW3,TMA64 as NEW4 FROM TMQAPPFORM WHERE TMA01='" & m_TMA01 & "'  AND TMA25<>'2' AND length(ltrim(TMA61||TMA62||TMA63||TMA64)) > 0 "
   strSql = strSql & " ORDER BY ORD1"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   MGRD1.FixedCols = 0
   If intI = 1 Then
      Set MGRD1.Recordset = RsTemp
      SetGrd False
      MGRD1.FixedCols = colORD
   End If
End Sub

'*****更新AI檢索文字*****
Private Sub SWordForUpdate(ByVal pIdx As String, ByVal pType As String)
Dim bolMsg As Boolean

   If pType = "ADD" Then '加入/置換
      strSql = ""
      For intI = 0 To 3
         strSql = strSql & ", TMA" & 45 + (pIdx - 1) * 4 + intI & "=" & CNULL(ChgSQL(txtFM2(intI)))
      Next intI
      strSql = "Update TMQAppForm set " & Mid(strSql, 2) & " Where TMA01='" & m_TMA01 & "' "
      If R_type = "A" Then Pub_SeekTbLog strSql
      cnnConnection.Execute strSql
      Call SWordForRead
   End If
   If pType = "DEL" Then '刪除
      strSql = ""
      For intI = 0 To 3
         strSql = strSql & ", TMA" & 45 + (pIdx - 1) * 4 + intI & "=NULL"
      Next intI
      strSql = "Update TMQAppForm set " & Mid(strSql, 2) & " Where TMA01='" & m_TMA01 & "' "
      If R_type = "A" Then Pub_SeekTbLog strSql
      cnnConnection.Execute strSql
      '重整順序
      If Val(pIdx) < MGRD1.Rows - 1 Then
         strTmp1 = ""
         For intJ = pIdx To MGRD1.Rows - 1
            For intI = 0 To 3
               strTmp1 = strTmp1 & ", TMA" & 45 + (intJ - 1) * 4 + intI & "=TMA" & 45 + (intJ) * 4 + intI
            Next intI
         Next intJ
         strSql = "Update TMQAppForm set " & Mid(strTmp1, 2) & " Where TMA01='" & m_TMA01 & "' "
         If R_type = "A" Then Pub_SeekTbLog strSql
         cnnConnection.Execute strSql
      End If
      Call SWordForRead
   End If
   If pType = "UP" Then '向上置換
      strTmp1 = ""
      For intI = 0 To 3
         strTmp1 = strTmp1 & ", TMA" & 45 + (pIdx - 2) * 4 + intI
      Next intI
      strTmp1 = "SELECT " & Mid(strTmp1, 2) & " FROM TMQAPPFORM WHERE TMA01='" & m_TMA01 & "' "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strTmp1)
      If intI = 1 Then
         strSql = ""
         For intI = 0 To 3
            strSql = strSql & ", TMA" & 45 + (pIdx - 2) * 4 + intI & "=" & CNULL(ChgSQL(txtFM2(intI)))
         Next intI
         strSql = "Update TMQAppForm set " & Mid(strSql, 2) & " Where TMA01='" & m_TMA01 & "' "
         If R_type = "A" Then Pub_SeekTbLog strSql
         cnnConnection.Execute strSql
         strSql = ""
         For intI = 0 To 3
            strSql = strSql & ", TMA" & 45 + (pIdx - 1) * 4 + intI & "=" & CNULL(ChgSQL("" & RsTemp(intI)))
         Next intI
         strSql = "Update TMQAppForm set " & Mid(strSql, 2) & " Where TMA01='" & m_TMA01 & "' "
         If R_type = "A" Then Pub_SeekTbLog strSql
         cnnConnection.Execute strSql
         Call SWordForRead
      End If
   End If

   If pType = "DOWN" Then '向下置換
      strTmp1 = ""
      For intI = 0 To 3
         strTmp1 = strTmp1 & ", TMA" & 45 + (pIdx) * 4 + intI
      Next intI
      strTmp1 = "SELECT " & Mid(strTmp1, 2) & " FROM TMQAPPFORM WHERE TMA01='" & m_TMA01 & "' "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strTmp1)
      If intI = 1 Then
         strSql = ""
         For intI = 0 To 3
            strSql = strSql & ", TMA" & 45 + (pIdx) * 4 + intI & "=" & CNULL(ChgSQL(txtFM2(intI)))
         Next intI
         strSql = "Update TMQAppForm set " & Mid(strSql, 2) & " Where TMA01='" & m_TMA01 & "' "
         If R_type = "A" Then Pub_SeekTbLog strSql
         cnnConnection.Execute strSql
         strSql = ""
         For intI = 0 To 3
            strSql = strSql & ", TMA" & 45 + (pIdx - 1) * 4 + intI & "=" & CNULL(ChgSQL("" & RsTemp(intI)))
         Next intI
         strSql = "Update TMQAppForm set " & Mid(strSql, 2) & " Where TMA01='" & m_TMA01 & "' "
         If R_type = "A" Then Pub_SeekTbLog strSql
         cnnConnection.Execute strSql
         Call SWordForRead
      End If
   End If
   
   If InStr("ADD,DEL,UP,DOWN", pType) > 0 And R_type = "A" Then
      MsgBox "存檔完成,請重新呼叫委查單!", vbInformation
      strSql = GetModETitle("，檢索文字有變更！")
      If Trim(textDB(10)) <> "" Then
         PUB_SendMail strUserNum, Trim(textDB(10)), "", strSql, "同主旨"
      End If
   End If
   
   For Each oObj In txtFM2
      oObj.Text = ""
      oObj.Tag = ""
   Next
End Sub

Private Function TxtValidate() As Boolean
Dim bolCancel As Boolean, bolChk As Boolean

   TxtValidate = False
   
   If m_TMA25 <> "1" And cmdSend.Tag <> "Q" Then
      If cmdKD(2).Tag = "" Then  '目前只開放查名人員更換圖片
         MsgBox lblData(25).Caption & "請上傳圖片！", vbCritical, "資料稽核"
         Exit Function
      End If
      If textDB(28).Text = "" And cmdRoute.Visible = True And cmdRoute.Caption = "輸入" Then
         MsgBox lblData(25).Caption & "請輸入圖形路徑！", vbCritical, "資料稽核"
         Exit Function
      End If
   End If
   If Trim(textDB(31)) <> "" And Trim(textDB(32)) <> "" And Val(textDB(31)) > Val(textDB(32)) Then
      MsgBox "查詢區間起值不可大於迄值！", vbCritical, "資料稽核"
      Exit Function
   End If
   
   Select Case cmdSend.Tag
      Case "U" '〔查名人〕送出(網中),查覆完畢,修改完畢(未收文前查名人員可修改)
          If InStr(cmdSend.Caption, "送出") > 0 Then
             If m_TMA25 <> "2" And Trim("" & MGRD1.TextMatrix(1, colORD)) = "" Then
                If MsgBox("尚未輸入檢索文字，是否繼續送出作業？" & vbCrLf & "選擇""否""回到輸入畫面", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
                   Exit Function
                End If
             End If
          Else   '查覆完畢,修改完畢
JumpToU:
             If Cbo1.Visible = True Then
                If Trim(Cbo1.Text) = "" Then
                   MsgBox "請輸入標章查覆結果！", vbCritical, "資料稽核"
                   Exit Function
                Else
                   Call Cbo1_Validate(bolCancel)
                   If bolCancel = True Then
                      Exit Function
                   End If
                End If
                If Val(textDB(36).Tag) + Val(textDB(37).Tag) + Val(textDB(38).Tag) > 0 And Trim(Cbo1.Text) = "不查" Then
                    If MsgBox("查覆結果為" & Trim(Cbo1.Text) & "，委查筆數會變更為０筆，請確認是否繼續作業？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
                       Exit Function
                    Else
                       textDB(36).Text = "0": textDB(37).Text = "0": textDB(38).Text = "0"
                    End If
                End If
                If Trim(Cbo1.Text) = TMQ_近似T1 Or Trim(Cbo1.Text) = TMQ_近似T2 Then
                   If Trim(textDB(17)) = "" Then
                      MsgBox "請輸入與本所近似的本所案件之申請號或審定號！", vbCritical, "資料稽核"
                      textDB(17).SetFocus
                      textDB_GotFocus 17
                      Exit Function
                   Else
                      Call textDB_Validate(17, bolCancel)
                      If bolCancel = True Then
                         Exit Function
                      End If
                   End If
                Else
                   If Trim(textDB(17).Text) <> "" Then
                      MsgBox "非與本所案件近似，請勿輸入申請號或審定號！", vbCritical
                      textDB(17).SetFocus
                      textDB_GotFocus 17
                      Exit Function
                   End If
                End If
             End If
             '檢查:查覆附件
             strExc(1) = "select count(*) cnt from tmqappfile where tmf01='" & m_TMA01 & "' and tmf02='" & m_ATMF02 & "' " & IIf(m_STMF03 <> "", "and tmf03>'" & m_STMF03 & "' ", "")
             intI = 1
             Set RsTemp = ClsLawReadRstMsg(intI, strExc(1))
             If intI = 1 Then
                If Val("" & RsTemp.Fields("cnt")) = 0 Then
                   '結果為不查或無，不需新增查覆附件
                   If (Cbo1.Visible = True And Trim(Cbo1.Text) <> "無" And Trim(Cbo1.Text) <> "不查") Or Trim(textDB(17).Text) <> "" Then
                       MsgBox "尚未新增查名結果附件，請確認資料的正確性！", vbCritical, "資料稽核"
                       Exit Function
                   End If
                   If m_TMA20 = "" Then  '(網中)查名
                      If MsgBox("尚未新增查名結果附件，是否繼續" & Mid(cmdSend.Caption, 1, 4) & "作業？", vbCritical + vbYesNo + vbDefaultButton2, "資料稽核") = vbNo Then
                         Exit Function
                      End If
                   End If
                End If
             End If  '----檢查:查覆附件
          End If   '----查覆完畢,修改完畢

      Case "M" '〔覆核人〕覆核完畢
JumpToM:
          strExc(1) = "Select tma01,tma19,tma16,tma17 from tmqappform where tma01='" & m_TMA01 & "' and (nvl(tma16,'N')='Y' or nvl(tma19,'9')=" & TMQ_近似1 & " or nvl(tma19,'9')=" & TMQ_近似2 & ")"
          intI = 1
          Set RsTemp = ClsLawReadRstMsg(intI, strExc(1))
          If intI = 0 Then
             MsgBox "查名結果已變更，請重新進入！", vbCritical, "資料稽核"
             Exit Function
          End If
          If ChkTMA67(0).Value = 0 And ChkTMA67(1).Value = 0 And ChkTMA67(2).Value = 0 Then
              MsgBox "請輸入覆核結果！", vbCritical, "資料稽核"
              Exit Function
          Else
              If ChkTMA67(0).Value = 1 Then
                 If ChkTMA69(0).Value = 0 And ChkTMA69(1).Value = 0 Then
                    If MsgBox("尚未輸入協商流程結果，是否需要輸入？", vbExclamation + vbYesNo + vbDefaultButton1, "資料稽核") = vbYes Then
                       Exit Function
                    End If
                 End If
              Else
                 If ChkTMA69(0).Value = 1 Or ChkTMA69(1).Value = 1 Then
                    MsgBox "與本所近似或相同，才可以輸入協商流程結果！", vbCritical, "資料稽核"
                    Exit Function
                 End If
              End If
          End If

      Case "Q" '〔委查人〕撤回
          If MsgBox("確定撤回查名作業嗎？", vbInformation + vbYesNo, "資料稽核") = vbYes Then
             If Left(m_TMA71, 1) = "H" And Left(m_TMA71, 2) <> "HM" Then
                strExc(1) = "select count(*) cnt from tmqappfile where tmf01 in (select tma01 from tmqappform where tma71='" & m_TMA71 & "') "
                intI = 1
                Set RsTemp = ClsLawReadRstMsg(intI, strExc(1))
                If intI = 1 Then
                   If Val("" & RsTemp.Fields("cnt")) > 0 Then
                      MsgBox "查名作業已經有輸入查覆結果或附件，不可撤回！", vbCritical, "資料稽核"
                      Exit Function
                   End If
                End If
             End If
             strExc(1) = "select count(*) cnt from tmqappfile where tmf01='" & m_TMA01 & "' and tmf02='" & m_ATMF02 & "' " & IIf(m_STMF03 <> "", "and tmf03>'" & m_STMF03 & "' ", "")
             intI = 1
             Set RsTemp = ClsLawReadRstMsg(intI, strExc(1))
             If intI = 1 Then
                If Val("" & RsTemp.Fields("cnt")) > 0 Then
                   MsgBox "查名作業已經有輸入查覆結果或附件，不可撤回！", vbCritical, "資料稽核"
                   Exit Function
                End If
             End If
          Else  '不撤回
             Exit Function
          End If
   End Select
   
   If R_type = "A" Then '資料維護(限電腦中心)
      If bolChk = False Then
         bolChk = True
         If Trim(textDB(65)) = "" Then
            GoTo JumpToU
         Else
            GoTo JumpToM
         End If
      End If
      For Each oObj In textDB
         If oObj.Text <> oObj.Tag Then
            Select Case oObj.Index
               Case 4, 5, 7
                    MsgBox "請改用SQL語法更新Date", vbCritical, "資料維護(限電腦中心)"
                    Exit Function
               Case Else
                    Call textDB_Validate(oObj.Index, bolCancel)
                    If bolCancel = True Then
                       Exit Function
                    End If
            End Select
         End If
      Next
      
      '*****雖然限制不可修改，但是修改語法有寫
      '全類檢索TMA21
      If ChkS3(0).Tag <> IIf(ChkS3(0).Value = 1, "Y", "") Then
          MsgBox "是否全類檢索不可變更！", vbCritical, "資料維護(限電腦中心)"
          Exit Function
      End If
      '查詢資料範圍(TMA29)：1-全部, 2-僅查本所代理
      If ChkS3(1).Tag <> IIf(ChkS3(1).Value = 1, "2", "1") Then
         MsgBox "查詢資料範圍不可變更！", vbCritical, "資料維護(限電腦中心)"
         Exit Function
      End If
      '是否包含無效或核駁資料(TMA30)
      If ChkS3(2).Tag <> IIf(ChkS3(2).Value = 1, "Y", "") Then
         MsgBox "是否包含無效或核駁資料不可變更！", vbCritical, "資料維護(限電腦中心)"
         Exit Function
      End If
      '團體標章/證明標章(TMA20)
      If m_TMA20 <> IIf(ChkS2(0).Value = 1, "1", "") Then
         MsgBox "團體標章不可變更！", vbCritical, "資料維護(限電腦中心)"
         Exit Function
      End If
      '***********

   End If
   
   TxtValidate = True

End Function

'******查名人員：送出,查覆完畢,修改完畢(未收文前查名人員可修改)******
Private Function ProcSaveU() As Boolean
Dim strUpd As String, intB As Integer, strB1 As String
Dim rsB1 As New ADODB.Recordset
Dim bolConn As Boolean
Dim strSub As String, strContent As String, strTo As String, strTempCC As String

   strUpd = ""
   For Each oObj In textDB
      Select Case oObj.Index
         Case 28, 36, 37, 38  '圖形路徑TMA28,委查筆數TMA36,TMA37,TMA38
             If oObj.Text <> oObj.Tag Then
                strUpd = strUpd & ", TMA" & Format(oObj.Index, "00") & "=" & CNULL(ChgSQL(oObj.Text))
             End If
      End Select
   Next
   
On Error GoTo ErrHandle
   If InStr(cmdSend.Caption, "送出") > 0 Then
      'Added by Lydia 2025/04/24 要先寫入完整資料，才能送出
      If strUpd <> "" Then
         strSql = "Update tmqappform set " & Mid(strUpd, 2) & " Where TMA01='" & m_TMA01 & "' "
         cnnConnection.Execute strSql
      End If
      '*****啟動(網中)商標查詢系統******
      'Added by Lydia 2025/04/21 開放測試模式
      strB1 = "Y"
      If InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") <> 0 Or Pub_StrUserSt03 = "M51" Then
         If MsgBox("是否啟動台一商標查詢系統", vbYesNo + vbDefaultButton2) = vbNo Then
            strB1 = ""
         End If
      End If
      If strB1 = "Y" Then
         If ProcCallSearch(m_TMA01) = True Then
            strB1 = ""
         End If
      End If
      If strB1 <> "" Then
          iStiu = 0
          FormEnabled
          Exit Function
      Else
      'end 2025/04/21
          strSql = "Update tmqappform set tma05=sysdate Where TMA01='" & m_TMA01 & "' "
          cnnConnection.Execute strSql
      End If
      iStiu = 0
      FormEnabled
      Call cmdNext_Click
   Else
      If textDB(14).Tag <> "" Then bolModify = True
      If bolModify = False Then strUpd = strUpd & ", tma14=to_char(sysdate,'yyyymmdd')"
      
      cnnConnection.BeginTrans
         bolConn = True
         If strUpd <> "" Then
            strSql = "Update tmqappform set " & Mid(strUpd, 2) & " Where TMA01='" & m_TMA01 & "' "
            cnnConnection.Execute strSql
         End If
         
         '在查覆完畢時，針對所有案件進行通知。
         strSql = "select cp01,cp02,cp03,cp04,tqc02,cp06,cp13,cp14,cp122,ep06,count(tqc03) as 已勾 ,sum(decode(TMA14,null,0,1)) as 已完成 " & _
                  "from tmqcasemap,tmqappform,caseprogress,engineerprogress where tqc02 in (select tqc02 from tmqcasemap where nvl(tqc02,'N') <>'N' and tqc03='" & m_TMA01 & "' ) " & _
                  "and tqc03=TMA01(+) and tqc02=cp09(+) and cp09=ep02(+) group by cp01,cp02,cp03,cp04,tqc02,cp06,cp13,cp14,cp122,ep06 "
         intB = 1
         strUpd = ""
         Set rsB1 = ClsLawReadRstMsg(intB, strSql)
         If intB = 1 Then
            rsB1.MoveFirst
            Do While Not rsB1.EOF
               If Val("" & rsB1.Fields("已勾")) = Val("" & rsB1.Fields("已完成")) Then
                  strExc(1) = "update caseprogress set cp143=" & strSrvDate(1) & " where cp09='" & rsB1.Fields("tqc02") & "' " '所有案件的收文號
                  cnnConnection.Execute strExc(1), intB
                  '發信給申請者
                  strUpd = strUpd & "查覆" & ";"
                  If "" & rsB1.Fields("cp14") <> "" And Val("" & rsB1.Fields("ep06")) > 0 Then
                     strExc(0) = PUB_TMdebateCountCP48("" & rsB1.Fields("cp06"), "" & rsB1.Fields("cp122"), "" & rsB1.Fields("ep06"), "" & rsB1.Fields("tqc02"), "" & rsB1.Fields("cp13"))
                     If strExc(0) <> "" Then
                         strExc(1) = "UPDATE CaseProgress SET CP48 = " & strExc(0) & " " & _
                                     "WHERE CP09 = '" & "" & rsB1.Fields("tqc02") & "' "
                         cnnConnection.Execute strExc(1), intB
                     End If
                  End If
                  If strUpd <> "" Then
                      If ChkTMA16.Value = 1 Or Trim(Cbo1.Text) = TMQ_近似T2 Then strUpd = strUpd & "近似本所案;"
                      If Trim(Cbo1.Text) = TMQ_近似T1 Then strUpd = strUpd & "相同本所案;"
                      If InStr(strUpd, "本所案") > 0 Then
                         '通知覆核人員
                         strTo = Pub_GetSpecMan("內商查名覆核通知")
                         strSub = GetModETitle("，查名結果需覆核。")
                         strContent = vbCrLf & vbCrLf & "請進入覆核區進行覆核作業。"
                         If Len(strTo) <= 6 Then '排除職代直接納入收件人
                            strTempCC = GetDutyList(strTo)
                            If strTempCC <> "" Then
                               strContent = "因收件人" & PUB_ReadUserData(strTo) & "請假，請副本收件人處理此郵件；" & vbCrLf & vbCrLf & strContent
                            End If
                         End If
                      Else
                         '通知委查人和承辦人
                         strTo = textDB(8) & IIf("" & rsB1.Fields("cp14") <> "", ";" & rsB1.Fields("cp14"), "")
                         strSub = GetModETitle("，查名作業已完成。")
                         strContent = vbCrLf & vbCrLf & "請進入承辦人系統->智權部->專利商標作業->商標查名／查覆區" & _
                                    "，點選記錄進入查覆明細作業查閱結果和附件。" & vbCrLf & vbCrLf
                         If FirstCP09 <> "" Then strContent = strContent & "已收文案件日後可到共同查詢的卷宗區點選查名結果。"
                      End If
                      'Added by Lydia 2025/04/21 開放測試模式，指定測試收件者
                      If strTestReceiver <> "" Then
                         strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                                  " values( '" & strUserNum & "','" & strTestReceiver & "',to_char(sysdate,'yyyymmdd')" & _
                                  ",to_char(sysdate,'hh24miss'),'" & ChgSQL(strSub) & "','" & "原收件者：" & PUB_ReadUserData(strTo, True) & vbCrLf & IIf(strTempCC <> "", "副本：　" & PUB_ReadUserData(strTempCC, True), "") & vbCrLf & ChgSQL(strContent) & "',null)"
                      Else
                      'end 2025/04/21
                         strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                                  " values( '" & strUserNum & "','" & strTo & "',to_char(sysdate,'yyyymmdd')" & _
                                  ",to_char(sysdate,'hh24miss'),'" & ChgSQL(strSub) & "','" & ChgSQL(strContent) & "'," & CNULL(strTempCC) & ")"
                      End If
                      cnnConnection.Execute strSql
                  End If
               End If

               rsB1.MoveNext
            Loop
         Else '沒有收文
            strUpd = strUpd & "查覆" & ";"
            If ChkTMA16.Value = 1 Or Trim(Cbo1.Text) = TMQ_近似T2 Then strUpd = strUpd & "近似本所案;"
            If Trim(Cbo1.Text) = TMQ_近似T1 Then strUpd = strUpd & "相同本所案;"
            If InStr(strUpd, "本所案") > 0 Then
               '通知覆核人員
               strTo = Pub_GetSpecMan("內商查名覆核通知")
               strSub = GetModETitle("，查名結果需覆核。")
               strContent = vbCrLf & vbCrLf & "請進入覆核區進行覆核作業。"
               strTempCC = GetDutyList(strTo)
               If strTempCC <> "" Then
                  strContent = "因收件人" & PUB_ReadUserData(strTo) & "請假，請副本收件人處理此郵件；" & vbCrLf & vbCrLf & strContent
               End If
            Else
               '通知委查人
               strTo = textDB(8)
               strSub = GetModETitle("，查名作業已完成。")
               strContent = vbCrLf & vbCrLf & "請進入承辦人系統->智權部->專利商標作業->商標查名／查覆區" & _
                          "，點選記錄進入查覆明細作業查閱結果和附件。" & vbCrLf & vbCrLf
               If FirstCP09 <> "" Then strContent = strContent & "已收文案件日後可到共同查詢的卷宗區點選查名結果。"
            End If
            'Added by Lydia 2025/04/21 開放測試模式，指定測試收件者
            If strTestReceiver <> "" Then
               strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                        " values( '" & strUserNum & "','" & strTestReceiver & "',to_char(sysdate,'yyyymmdd')" & _
                        ",to_char(sysdate,'hh24miss'),'" & ChgSQL(strSub) & "','" & "原收件者：" & PUB_ReadUserData(strTo, True) & vbCrLf & IIf(strTempCC <> "", "副本：　" & PUB_ReadUserData(strTempCC, True), "") & vbCrLf & ChgSQL(strContent) & "',null)"
            Else
            'end 2025/04/21
               strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                        " values( '" & strUserNum & "','" & strTo & "',to_char(sysdate,'yyyymmdd')" & _
                        ",to_char(sysdate,'hh24miss'),'" & ChgSQL(strSub) & "','" & ChgSQL(strContent) & "'," & CNULL(strTempCC) & ")"
            End If
            cnnConnection.Execute strSql
         End If
         
         '查名單筆數歸零,再重新修正拿單量
         If textDB(10) <> "" And Val(textDB(36).Text) + Val(textDB(37).Text) + Val(textDB(38).Text) = 0 And Val(textDB(36).Tag) + Val(textDB(37).Tag) + Val(textDB(38).Tag) > 0 Then
            '更新當日拿單量
             Call PUB_TMAtoTake("2", textDB(10), "", "0", False)
            '更新前2日統計量
             Call PUB_TMAtoTake("2", textDB(10), "", "1", False)
         End If
      cnnConnection.CommitTrans
      bolConn = False
      
      iStiu = 0
      FormEnabled
      Call cmdNext_Click
   End If
   ProcSaveU = True
   Set rsB1 = Nothing
   Exit Function
   
ErrHandle:
   If bolConn = True Then
      cnnConnection.RollbackTrans
   End If
   If Err.Number <> 0 Then
      MsgBox Mid(cmdSend.Tag, 1, 2) & "作業失敗：" & Err.Description
   End If
   Set rsB1 = Nothing
End Function

'******覆核人員：覆核完畢,更改
Private Function ProcSaveM() As Boolean
Dim strUpd As String, intB As Integer
Dim rsB1 As New ADODB.Recordset
Dim bolConn As Boolean
Dim strSub As String, strContent As String, strTo As String, strTempCC As String
Dim str相同 As String, str近似 As String
Dim strMailKind As String 'Added by Lydia 2024/10/14 覆核增加確認機制; 依覆核結果區別通知內容:1-經覆核結果為「相同TQD09=4」、「近似TQD09=5」時，通知內容為：覆核意見：（帶入覆核意見）,
                                                                 '2-經覆核結果為「稍近似TQD09=6」或「無TQD09=7」時，通知內容為：覆核意見：（帶入覆核意見）（如有覆核意見,請先列覆核意見，再接續原通知內容，如無則僅列通知內容）
Dim strCP13name As String  '判斷智權人員名稱


On Error GoTo ErrHandle
   strSql = "Update TmqAppForm Set TMA65='" & strUserNum & "', TMA66=to_char(sysdate,'yyyymmdd') where nvl(tma65,'N')='N' and tma01='" & m_TMA01 & "'"
   cnnConnection.Execute strSql

   'Modified by Lydia 2022/06/06 覆核增加確認機制; 當查名結果從近似本所案△改為非近似本所案是查名人員的看法，尚需要智權人員確認結果，所以修改通知email提醒兩方要再次確認(杜經理的需求by 嘉雯)
   If ChkTMA67(0).Value = 1 Then
      'Modified by Lydia 2017/08/28 區分本所案或統一案;
      '查名若發現與統一公司商標近似情形時 , 仍列為客戶間利益衝突案件:
      '1.於審定號/申請號欄位，在號數前加上P時，查名結果可為「相同△、近似△」，並且不檢查是否為本所案；
      '2.遇到P審定號/申請號，系統發通知要求應徵詢的智權同仁為「蘇威廷」。
      strExc(1) = "": strExc(7) = "": strExc(5) = ""
      tmpArr = Empty
      tmpArr = Split(Trim(textDB(17)), ",")
      For intI = 0 To UBound(tmpArr)
         If Trim(tmpArr(intI)) <> "" Then
            If Left(Trim(tmpArr(intI)), 1) = "P" Then
               strExc(7) = strExc(7) & Trim(tmpArr(intI)) & ","
            Else
               strExc(1) = strExc(1) & Trim(tmpArr(intI)) & ","
            End If
         End If
      Next intI
      'Added by Lydia 2017/08/28 統一公司案件
      If strExc(7) <> "" Then
         '指定智權人員
         strExc(0) = " select decode(s1.st04,'1',s1.st01,nvl(a0924,a0908)) sno,decode(s1.st04,'1',s1.st02,getstaffnamelist(nvl(a0924,a0908))) sname " & _
              "from staff s1, acc090,acc090new where s1.st01='A2026' and s1.st03=a0901(+) and s1.st93=a0921(+) "
         intB = 1
         Set rsB1 = ClsLawReadRstMsg(intB, strExc(0))
         If intB = 1 Then
            strExc(5) = strExc(5) & rsB1.Fields("sname") & ","
         End If
      End If
      
      '本所案
      If strExc(1) <> "" Then
         strExc(1) = GetAddStr(strExc(1))
         strExc(0) = "SELECT '1' ORD,TM01,TM02,TM03,TM04,TM15 AS TM1215 FROM TRADEMARK WHERE TM10='000' AND TM15 IN (" & strExc(1) & ") " & _
                     "Union SELECT '2' ORD,TM01,TM02,TM03,TM04,TM12 AS TM1215 FROM TRADEMARK WHERE TM10='000' AND TM12 IN (" & strExc(1) & ") ORDER BY 1,2,3 "

         intB = 1
         Set rsB1 = ClsLawReadRstMsg(intB, strExc(0))
         If intB = 1 Then
            rsB1.MoveFirst
            Do While Not rsB1.EOF
               '若同時有一樣的審定號,只顯示在相同本所案
               'Added by Lydia 2016/04/28 內文代出近似本所案的智權人員
               strExc(3) = GetStaffName(PUB_GetAKindSalesNo(rsB1.Fields("tm01"), rsB1.Fields("tm02"), rsB1.Fields("tm03"), rsB1.Fields("tm04")))
               If strExc(3) <> "" And InStr(strExc(5) & ",", strExc(3)) = 0 Then
                  strExc(5) = strExc(5) & strExc(3) & ","
               End If
               rsB1.MoveNext
            Loop
         End If
      End If
      
      '相同△、近似△之查名結果，覆核完畢後，系統除通知委查人及案件承辦人(已收文)外，增加通知查名人。
      strTo = textDB(8) & IIf(FirstCP14 <> "", ";" & FirstCP14, "") & ";" & textDB(10)
      strSub = GetModETitle("，查名作業已完成覆核，智權人員請確認覆核結果。")
      strContent = vbCrLf & vbCrLf & "請進入承辦人系統->智權部->專利商標作業->商標查名／查覆區" & _
                                    "，點選記錄進入查覆明細作業查閱結果和附件。" & vbCrLf & vbCrLf
      strExc(3) = "近似"
      If Cbo1.Visible = True And Trim(Cbo1.Text) = TMQ_近似T1 Then strExc(3) = "相同"
      strContent = strContent & "已覆核但與本所" & Trim(textDB(17)) & strExc(3) & ",不得申請" & vbCrLf
      'Added by Lydia 2025/04/21 開放測試模式，指定測試收件者
      If strTestReceiver <> "" Then
         strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
            " values( '" & strUserNum & "','" & strTestReceiver & "',to_char(sysdate,'yyyymmdd')" & _
            ",to_char(sysdate,'hh24miss'),'" & ChgSQL(strSub) & "','" & "原收件者：" & PUB_ReadUserData(strTo, True) & vbCrLf & IIf(strTempCC <> "", "副本：　" & PUB_ReadUserData(strTempCC, True), "") & vbCrLf & ChgSQL(strContent) & "',null)"
      Else
      'end 2025/04/21
         strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
            " values( '" & strUserNum & "','" & strTo & "',to_char(sysdate,'yyyymmdd')" & _
            ",to_char(sysdate,'hh24miss'),'" & ChgSQL(strSub) & "','" & ChgSQL(strContent) & "'," & CNULL(strTempCC) & ")"
      End If
      cnnConnection.Execute strSql
   Else
      'mType = mType & ";增加確認;"
      If ChkTMA67(1).Value = 1 Then
         strMailKind = "1"   '1-經覆核結果為「相同」、「近似」
      Else
         strMailKind = "2"   '2-經覆核結果為「稍近似」或「無」
      End If
      
      strExc(2) = "": strExc(3) = "": strExc(4) = "": strExc(5) = "": strExc(6) = ""
      tmpArr = Empty
      tmpArr = Split(Trim(textDB(17)), ",")
      For intI = 0 To UBound(tmpArr)
         If Trim(tmpArr(intI)) <> "" Then
             strExc(0) = "select tm23, nvl(cu04,nvl(cu05,cu06)) as cname,tm01,tm02,tm03,tm04,nvl(cu80,'N') as CU80 " & _
                              "from trademark,customer where (tm12=" & CNULL(IIf(Left(Trim(tmpArr(intI)), 1) = "P", Mid(Trim(tmpArr(intI)), 2), Trim(tmpArr(intI)))) & " or tm15=" & CNULL(IIf(Left(Trim(tmpArr(intI)), 1) = "P", Mid(Trim(tmpArr(intI)), 2), Trim(tmpArr(intI)))) & " ) and substr(tm23,1,8)=cu01(+) and  substr(tm23,9,1)=cu02(+) "
             'Added by Lydai 2022/07/12 區分本所案或統一案/統一公司案件(P審定號/申請號，系統發通知要求應徵詢的智權同仁為「蘇威廷」);
             If Left(Trim(tmpArr(intI)), 1) = "P" Then
                '統一公司非本所案用T-230642為代表
                strExc(0) = strExc(0) & " union select tm23, nvl(cu04,nvl(cu05,cu06)) as cname,'T' as tm01,'999999' as tm02, '0' as tm03,'00' as tm04,nvl(cu80,'N') as CU80 " & _
                                 "from trademark,customer where tm01='T' and tm02='230642' and tm03='0' and tm04='00' and substr(tm23,1,8)=cu01(+) and  substr(tm23,9,1)=cu02(+) "
             End If
             strExc(0) = strExc(0) & " order by tm01,tm02,tm03,tm04"
             'end 2022/07/12
             intB = 1
             Set rsB1 = ClsLawReadRstMsg(intB, strExc(0))
             If intB = 1 Then
               If Left(Trim(tmpArr(intI)), 1) = "P" Then
                   '區分本所案或統一案/統一公司案件(P審定號/申請號，系統發通知要求應徵詢的智權同仁為「蘇威廷」);
                   strExc(2) = "A2026"
               Else
                   strExc(2) = PUB_GetAKindSalesNo(rsB1.Fields("tm01"), rsB1.Fields("tm02"), rsB1.Fields("tm03"), rsB1.Fields("tm04"))
               End If
               '覆核增加確認機制;覆核後為稍近似的案件,若前案為無效客戶時,無須發送通知 ; 依覆核結果區別通知內容
               If InStr("解散,廢止,撤銷,死亡", "" & rsB1.Fields("cu80")) > 0 And strMailKind = "2" Then
                  strExc(2) = ""
               End If
               If strExc(2) <> "" Then
                   '判斷智權人員名稱
                   strCP13name = ""
                   If Left(strExc(2), 4) = "MCTF" Then
                      strExc(0) = Pub_GetSpecMan(strExc(2))
                      If strExc(0) <> "" Then
                         strCP13name = PUB_ReadUserData(strExc(0))
                      End If
                   End If
                   If strCP13name = "" Then
                      strCP13name = GetStaffName(strExc(2), True)
                   End If
                   If InStr(strExc(4) & ",", "(" & rsB1.Fields("tm01") & "-" & rsB1.Fields("tm02")) = 0 Then
                      '統一公司非本所案
                      If Left(Trim(tmpArr(intI)), 1) = "P" And "" & rsB1.Fields("tm02") = "999999" Then
                          strExc(4) = strExc(4) & "、 " & strCP13name & "的" & rsB1.Fields("tm23") & rsB1.Fields("cname") & Mid(Trim(tmpArr(intI)), 2)
                      Else
                          strExc(4) = strExc(4) & "、 " & strCP13name & "的" & rsB1.Fields("tm23") & rsB1.Fields("cname") & "(" & rsB1.Fields("tm01") & "-" & rsB1.Fields("tm02") & IIf(rsB1.Fields("tm03") <> "0", "-" & rsB1.Fields("tm03"), "") & IIf(rsB1.Fields("tm04") <> "00", "-" & rsB1.Fields("tm04"), "") & ")"
                      End If

                      If InStr(";" & strExc(5), strExc(2)) = 0 Then
                          strExc(5) = strExc(5) & ";" & strExc(2) '相關衝突智權(B智權)
                          'Modified by Lydia 2022/09/21 外商日文組的此類案件協調不寄給承辦人，而直接只寄給主管
                          strExc(1) = strExc(2)
                          If rsB1.Fields("tm01") = "FCT" Or rsB1.Fields("tm01") = "CFT" Then
                              strExc(1) = PUB_GetF11ToMan(strExc(2))
                          End If
                          '收件者
                          If InStr(";" & strExc(6), strExc(1)) = 0 Then
                              strExc(6) = strExc(6) & ";" & strExc(1)
                          End If
                          strExc(3) = GetDeptMan(GetST15(strExc(2)))
                          'Modified by Lydia 2022/09/21 排除外商日文組
                          If InStr(";" & strExc(6), strExc(3)) = 0 And strExc(1) = strExc(2) Then
                              strExc(6) = strExc(6) & ";" & strExc(3) '相關衝突智權(B智權)的區主管
                          End If
                      End If
                   End If
               End If
             End If
         End If
      Next intI
      If strExc(5) & strExc(6) <> "" Then 'Added by Lydia 2024/10/09 覆核增加確認機制;覆核後為稍近似的案件,若前案為無效客戶時,無須發送通知
         If strMailKind = "1" Then
            strContent = "覆核意見：" & Trim(textDB(68))
         Else
            strExc(5) = Mid(strExc(5), 2)
            strContent = "因 " & lblData(8) & " 之客戶委查商標與" & Mid(strExc(4), 2) & " 商標相同或近似，" & _
                      "但經覆核後，商標部認為在商標整體上有可辦理機會，但為避免往後客戶間的衝突事件，" & _
                      "請 " & lblData(8) & " 在向客戶說明之前，先與 " & PUB_ReadUserData(strExc(5)) & " 連絡進行協商，" & _
                      "若 " & PUB_ReadUserData(strExc(5)) & " 表示同意，請於收文時於接洽單備註處載明，且於收文後，將相關資訊存入該案之卷宗區中。"
            If Trim(textDB(68)) <> "" Then
               strContent = "覆核意見：" & Trim(textDB(68)) & vbCrLf & vbCrLf & strContent
            End If
         End If
         '系統通知發送：委查人(A智權)、A智權的區主管、相關衝突智權(B智權)、B智權的區主管、杜協理(全所智權部主管)、中所林協理(中所智權部主管)、商標承辦人員(FirstCP14)、商標部覆核人員(strUserNme)、江協理(V2)。
                                   'rTo 傳入=查名結果給近似△的查名人員
         strExc(2) = GetDeptMan(GetST15(textDB(8)))
         If InStr(strExc(6), textDB(8)) = 0 Then
             strExc(6) = ";" & textDB(8) & strExc(6)
         End If
         If InStr(strExc(6), strExc(2)) = 0 Then
             strExc(6) = ";" & strExc(2) & strExc(6)
         End If
      End If 'Added by Lydia 2024/10/09 覆核增加確認機制;覆核後為稍近似的案件,若前案為無效客戶時,無須發送通知
              
      strSub = GetModETitle("，查名作業已完成覆核，智權人員請確認覆核結果。")
      strTo = Mid(strExc(6), 2) & ";" & Pub_GetSpecMan("全所智權部主管") & _
              IIf(FirstCP14 <> "", ";" & FirstCP14, "") & ";" & strUserNum & ";" & Pub_GetSpecMan("V2") & ";" & textDB(10)
      strTo = Replace(strTo, ";;", ";")
   
      'Added by Lydia 2025/04/21 開放測試模式，指定測試收件者
      If strTestReceiver <> "" Then
         strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
            " values( '" & strUserNum & "','" & strTestReceiver & "',to_char(sysdate,'yyyymmdd')" & _
            ",to_char(sysdate,'hh24miss'),'" & ChgSQL(strSub) & "','" & "原收件者：" & PUB_ReadUserData(strTo, True) & vbCrLf & IIf(strTempCC <> "", "副本：　" & PUB_ReadUserData(strTempCC, True), "") & vbCrLf & ChgSQL(strContent) & "', null)"
      Else
      'end 2025/04/21
         strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
            " values( '" & strUserNum & "','" & strTo & "',to_char(sysdate,'yyyymmdd')" & _
            ",to_char(sysdate,'hh24miss'),'" & ChgSQL(strSub) & "','" & ChgSQL(strContent) & "'," & CNULL(strTempCC) & ")"
      End If
      cnnConnection.Execute strSql
   End If

   iStiu = 0
   FormEnabled
   

   ProcSaveM = True
   Set rsB1 = Nothing
   Exit Function
   
ErrHandle:
   If bolConn = True Then
      cnnConnection.RollbackTrans
   End If
   If Err.Number <> 0 Then
      MsgBox Mid(cmdSend.Tag, 1, 2) & "作業失敗：" & Err.Description
   End If
   Set rsB1 = Nothing
End Function

'Added by Lydia 2025/04/15 檢查資料是否可維護
Private Function ChkIsTest() As Boolean

   ChkIsTest = True
   If Left(lblData(35), 1) = "H" Then  '匯入112年查名單
      MsgBox "查名單為匯入舊資料，僅提供瀏覽不提供其他功能！", vbInformation
      Exit Function
   End If
   If m_TMA02 = "2" And Trim(textDB(26)) = "" And m_TMA27 = "" Then
      MsgBox "查名單為網站匯入資料，僅提供瀏覽不提供其他功能！", vbInformation
      Exit Function
   End If
   ChkIsTest = False
End Function

'Added by Lydia 2025/04/21 啟動(網中)商標查詢系統
Private Function ProcCallSearch(ByVal pTMA01 As String) As Boolean
On Error GoTo 0
Dim strB01 As String, intRetry As Integer
Dim oCNHttp As New WinHttp.WinHttpRequest
    
   ProcCallSearch = False
   strB01 = Pub_GetSpecMan("TMSearch送出功能")
   If strB01 <> "" Then
      Set oCNHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
      oCNHttp.Option(WinHttpRequestOption_SslErrorIgnoreFlags) = &H3300
      oCNHttp.SetTimeouts 10000, 10000, 10000, 10000  'Resolve, Connect, Send and Receive
      oCNHttp.Open "GET", strB01 & pTMA01, False
JumpToRetry:
      oCNHttp.Send
      If oCNHttp.Status = 200 Then
         ProcCallSearch = True
      Else
         If intRetry < 3 Then
            Sleep 5000
            intRetry = intRetry + 1
            GoTo JumpToRetry
         Else
            MsgBox "啟動台一商標查詢系統失敗！", vbCritical, "台一商標查詢系統"
         End If
      End If
      Set oCNHttp = Nothing
   Else
      MsgBox "沒有網址！", vbCritical, "台一商標查詢系統"
   End If
   
End Function
