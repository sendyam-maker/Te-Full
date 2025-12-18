VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm140410_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "電子報排程維護"
   ClientHeight    =   6680
   ClientLeft      =   50
   ClientTop       =   340
   ClientWidth     =   9000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6680
   ScaleWidth      =   9000
   Begin VB.Frame FrameCU 
      BackColor       =   &H00FF8080&
      Caption         =   "目前使用於國外電子報"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   580
      Left            =   60
      TabIndex        =   76
      Top             =   5730
      Visible         =   0   'False
      Width           =   4150
      Begin VB.TextBox textCU01_2 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  '沒有框線
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2850
         Locked          =   -1  'True
         TabIndex        =   79
         TabStop         =   0   'False
         Top             =   240
         Width           =   1190
      End
      Begin VB.TextBox textCU01 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   290
         Left            =   630
         TabIndex        =   77
         Top             =   210
         Width           =   2170
      End
      Begin VB.Label Label2 
         Caption         =   "國籍："
         Height          =   230
         Index           =   0
         Left            =   60
         TabIndex        =   78
         Top             =   240
         Width           =   560
      End
   End
   Begin VB.Timer Timer2 
      Left            =   6840
      Top             =   600
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8280
      Top             =   540
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
            Picture         =   "frm140410_1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140410_1.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140410_1.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140410_1.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140410_1.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140410_1.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140410_1.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140410_1.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140410_1.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140410_1.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140410_1.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Left            =   7290
      Top             =   600
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   7785
      Top             =   600
      _ExtentX        =   494
      _ExtentY        =   494
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4965
      Left            =   90
      TabIndex        =   14
      Top             =   720
      Width           =   8775
      _ExtentX        =   15469
      _ExtentY        =   8767
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "基本資料"
      TabPicture(0)   =   "frm140410_1.frx":20F4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label9"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblTestName"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label13"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblCreate"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblUpdate"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdTest"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtToMail"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtNo"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdDetect"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Command1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lstImport(0)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lstImport(1)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmdSend"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "FrameTag"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      TabCaption(1)   =   "多筆查詢"
      TabPicture(1)   =   "frm140410_1.frx":2110
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Line2"
      Tab(1).Control(1)=   "Label1(4)"
      Tab(1).Control(2)=   "txtQueryDate(1)"
      Tab(1).Control(3)=   "txtQueryDate(0)"
      Tab(1).Control(4)=   "cmdQuery(0)"
      Tab(1).Control(5)=   "grdList"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Log"
      TabPicture(2)   =   "frm140410_1.frx":212C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "List1"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "寄發對象說明"
      TabPicture(3)   =   "frm140410_1.frx":2148
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Text1"
      Tab(3).ControlCount=   1
      Begin VB.Frame FrameTag 
         Height          =   600
         Left            =   7150
         TabIndex        =   80
         Top             =   2685
         Visible         =   0   'False
         Width           =   1300
         Begin VB.CheckBox chkTestData 
            BackColor       =   &H00FFFF80&
            Caption         =   "測式資料"
            Height          =   285
            Left            =   20
            TabIndex        =   82
            Top             =   360
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.CheckBox chkNoneTag 
            Caption         =   "不解析Tag"
            Height          =   285
            Left            =   20
            TabIndex        =   81
            Top             =   90
            Width           =   1140
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
         Height          =   4092
         Left            =   -74952
         TabIndex        =   72
         Top             =   816
         Width           =   8580
         _ExtentX        =   15134
         _ExtentY        =   7214
         _Version        =   393216
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
         _Band(0).Cols   =   2
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.5
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4425
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   2  '垂直捲軸
         TabIndex        =   69
         Text            =   "frm140410_1.frx":2164
         Top             =   420
         Width           =   8535
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "寄送"
         Height          =   300
         Left            =   7650
         TabIndex        =   66
         Top             =   2010
         Width           =   825
      End
      Begin VB.ListBox lstImport 
         Height          =   220
         Index           =   1
         ItemData        =   "frm140410_1.frx":2985
         Left            =   4725
         List            =   "frm140410_1.frx":2987
         TabIndex        =   47
         Top             =   1410
         Width           =   2400
      End
      Begin VB.ListBox lstImport 
         Height          =   220
         Index           =   0
         ItemData        =   "frm140410_1.frx":2989
         Left            =   4725
         List            =   "frm140410_1.frx":298B
         TabIndex        =   46
         Top             =   870
         Width           =   2400
      End
      Begin VB.CommandButton Command1 
         Caption         =   "信箱重整(&R)"
         Height          =   345
         Left            =   1845
         TabIndex        =   45
         Top             =   420
         Width           =   1185
      End
      Begin VB.ListBox List1 
         Height          =   2740
         Left            =   -74820
         TabIndex        =   44
         Top             =   420
         Width           =   8430
      End
      Begin VB.CommandButton cmdDetect 
         Height          =   300
         Left            =   5580
         Picture         =   "frm140410_1.frx":298D
         Style           =   1  '圖片外觀
         TabIndex        =   38
         Top             =   4020
         Visible         =   0   'False
         Width           =   350
      End
      Begin VB.TextBox txtNo 
         Height          =   300
         Left            =   1215
         TabIndex        =   0
         Top             =   450
         Width           =   510
      End
      Begin VB.TextBox txtToMail 
         Height          =   300
         Left            =   4680
         TabIndex        =   1
         Top             =   450
         Width           =   2400
      End
      Begin VB.CommandButton cmdTest 
         Caption         =   "測試"
         Height          =   300
         Left            =   7110
         TabIndex        =   29
         Top             =   450
         Width           =   825
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "查詢(&Q)"
         Height          =   400
         Index           =   0
         Left            =   -71940
         TabIndex        =   15
         Top             =   360
         Width           =   912
      End
      Begin VB.TextBox txtQueryDate 
         Height          =   270
         Index           =   0
         Left            =   -74040
         MaxLength       =   7
         TabIndex        =   12
         Top             =   450
         Width           =   945
      End
      Begin VB.TextBox txtQueryDate 
         Height          =   270
         Index           =   1
         Left            =   -72990
         MaxLength       =   7
         TabIndex        =   13
         Top             =   450
         Width           =   945
      End
      Begin VB.Frame Frame1 
         Height          =   3945
         Left            =   135
         TabIndex        =   18
         Top             =   690
         Width           =   8475
         Begin VB.TextBox txtMS26 
            Height          =   264
            Left            =   1584
            MaxLength       =   1
            TabIndex        =   83
            Top             =   1080
            Width           =   300
         End
         Begin VB.CheckBox chkMS26 
            Caption         =   "優先寄發 順位:       (1~9,小的優先)"
            Height          =   285
            Left            =   108
            TabIndex        =   70
            Top             =   1080
            Width           =   2988
         End
         Begin VB.CheckBox chkOutlook 
            Caption         =   "OutLook 範本"
            Height          =   285
            Left            =   6120
            TabIndex        =   65
            Top             =   1320
            Width           =   1335
         End
         Begin VB.CheckBox chkByMailServer 
            Caption         =   "是否用Mail Server發信(backup會收到)"
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   2700
            TabIndex        =   64
            Top             =   1320
            Width           =   3315
         End
         Begin VB.CheckBox chkNoYahoo 
            Caption         =   "Yahoo信箱不寄"
            Height          =   285
            Left            =   1065
            TabIndex        =   63
            Top             =   1335
            Width           =   1560
         End
         Begin VB.CheckBox ChkAtt 
            Caption         =   "有附件"
            Height          =   285
            Left            =   108
            TabIndex        =   3
            Top             =   1332
            Value           =   1  '核取
            Width           =   915
         End
         Begin VB.Frame Frame4 
            Height          =   465
            Left            =   90
            TabIndex        =   53
            Top             =   1530
            Width           =   8250
            Begin VB.CheckBox chkNoneBig5 
               Caption         =   "非 big5 碼"
               Height          =   285
               Left            =   6930
               TabIndex        =   54
               Top             =   150
               Width           =   1140
            End
            Begin MSForms.TextBox txtSubject 
               Height          =   324
               Left            =   984
               TabIndex        =   75
               Top             =   120
               Width           =   5892
               VariousPropertyBits=   679495707
               Size            =   "10393;572"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "主　　旨："
               Height          =   180
               Index           =   1
               Left            =   45
               TabIndex        =   55
               Top             =   180
               Width           =   900
            End
         End
         Begin VB.CommandButton cmdImport 
            Caption         =   "匯入..."
            Height          =   315
            Index           =   0
            Left            =   6975
            TabIndex        =   42
            Top             =   180
            Width           =   825
         End
         Begin VB.ComboBox cboTime 
            Height          =   300
            Left            =   3195
            Style           =   2  '單純下拉式
            TabIndex        =   9
            Top             =   2970
            Width           =   1140
         End
         Begin VB.ComboBox cboDisplayName 
            Height          =   300
            Left            =   1080
            TabIndex        =   5
            Text            =   "cboDisplayName"
            Top             =   2310
            Width           =   5910
         End
         Begin VB.ComboBox cboEmail 
            Height          =   300
            ItemData        =   "frm140410_1.frx":2A8F
            Left            =   1080
            List            =   "frm140410_1.frx":2A91
            TabIndex        =   4
            Text            =   "cboEmail"
            Top             =   2010
            Width           =   5910
         End
         Begin VB.TextBox txtSample 
            Height          =   315
            Left            =   1080
            TabIndex        =   6
            Top             =   2610
            Width           =   5910
         End
         Begin VB.CommandButton cmdOpen 
            Caption         =   "開啟..."
            Height          =   315
            Left            =   6975
            TabIndex        =   7
            Top             =   2610
            Width           =   825
         End
         Begin VB.TextBox txtDate 
            Height          =   300
            Left            =   1620
            MaxLength       =   7
            TabIndex        =   8
            Top             =   2970
            Width           =   870
         End
         Begin VB.TextBox txtCount 
            Enabled         =   0   'False
            Height          =   300
            Left            =   5625
            TabIndex        =   10
            Top             =   2970
            Width           =   555
         End
         Begin VB.CommandButton cmdEstimate 
            Caption         =   "粗估"
            Height          =   315
            Left            =   6165
            TabIndex        =   11
            Top             =   2970
            Visible         =   0   'False
            Width           =   825
         End
         Begin VB.ListBox lstMailChoice 
            BeginProperty Font 
               Name            =   "細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   852
            IntegralHeight  =   0   'False
            ItemData        =   "frm140410_1.frx":2A93
            Left            =   912
            List            =   "frm140410_1.frx":2AAF
            Sorted          =   -1  'True
            Style           =   1  '項目包含核取方塊
            TabIndex        =   2
            Top             =   192
            Width           =   2220
         End
         Begin VB.Frame Frame2 
            BorderStyle     =   0  '沒有框線
            Height          =   495
            Left            =   5805
            TabIndex        =   39
            Top             =   3240
            Visible         =   0   'False
            Width           =   2580
            Begin MSComctlLib.ProgressBar ProgressBar1 
               Height          =   225
               Left            =   45
               TabIndex        =   40
               Top             =   120
               Width           =   2445
               _ExtentX        =   4322
               _ExtentY        =   406
               _Version        =   393216
               Appearance      =   1
            End
            Begin VB.Label lblProgress 
               Alignment       =   2  '置中對齊
               Caption         =   "( 0/0 )"
               BeginProperty Font 
                  Name            =   "新細明體"
                  Size            =   8.5
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Left            =   90
               TabIndex        =   41
               Top             =   360
               Width           =   2355
            End
         End
         Begin VB.Frame Frame3 
            Height          =   825
            Left            =   3150
            TabIndex        =   49
            Top             =   480
            Width           =   5190
            Begin VB.CheckBox chkSpecList 
               Caption         =   "要匯入特殊名單"
               Height          =   285
               Left            =   1440
               TabIndex        =   71
               Top             =   468
               Visible         =   0   'False
               Width           =   1668
            End
            Begin VB.CommandButton cmdImport 
               Caption         =   "匯入..."
               Height          =   315
               Index           =   1
               Left            =   3870
               TabIndex        =   51
               Top             =   150
               Width           =   825
            End
            Begin VB.CheckBox chkMainOnly 
               Caption         =   "只寄代表號"
               Height          =   285
               Left            =   135
               TabIndex        =   50
               Top             =   468
               Visible         =   0   'False
               Width           =   1230
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "指定X,Y,R編號："
               Height          =   180
               Left            =   135
               TabIndex        =   56
               Top             =   210
               Width           =   1350
            End
            Begin VB.Label lblImpCount 
               AutoSize        =   -1  'True
               Caption         =   "000"
               Height          =   180
               Index           =   1
               Left            =   4815
               TabIndex        =   52
               Top             =   150
               Width           =   270
            End
            Begin VB.Label Label5 
               Caption         =   "記事本請關閉自動換行!!!"
               BeginProperty Font 
                  Name            =   "新細明體"
                  Size            =   8.5
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   315
               Left            =   3960
               TabIndex        =   58
               Top             =   480
               Visible         =   0   'False
               Width           =   1155
            End
         End
         Begin VB.Label lblImpCount 
            AutoSize        =   -1  'True
            Caption         =   "000"
            Height          =   180
            Index           =   0
            Left            =   7920
            TabIndex        =   48
            Top             =   180
            Width           =   270
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "排除編號："
            Height          =   180
            Left            =   3285
            TabIndex        =   43
            Top             =   210
            Width           =   900
         End
         Begin VB.Label lblState 
            AutoSize        =   -1  'True
            Caption         =   "寄信中..."
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   180
            Left            =   4545
            TabIndex        =   37
            Top             =   3360
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.Label lblFailCount 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "lblFailCount"
            Height          =   180
            Left            =   3420
            TabIndex        =   36
            Top             =   3690
            Width           =   870
         End
         Begin VB.Label lblActCount 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "lblActCount"
            Height          =   180
            Left            =   1260
            TabIndex        =   35
            Top             =   3690
            Width           =   855
         End
         Begin VB.Label lblActTo 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "lblActTo"
            Height          =   180
            Left            =   3150
            TabIndex        =   34
            Top             =   3360
            Width           =   615
         End
         Begin VB.Label lblActFrom 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "lblActFrom"
            Height          =   180
            Left            =   1125
            TabIndex        =   33
            Top             =   3360
            Width           =   795
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "～"
            Height          =   180
            Left            =   2700
            TabIndex        =   32
            Top             =   3360
            Width           =   180
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "寄發對象："
            Height          =   180
            Index           =   5
            Left            =   60
            TabIndex        =   28
            Top             =   192
            Width           =   900
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "顯示名稱："
            Height          =   180
            Index           =   3
            Left            =   135
            TabIndex        =   27
            Top             =   2340
            Width           =   900
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "樣本檔案："
            Height          =   180
            Index           =   0
            Left            =   135
            TabIndex        =   26
            Top             =   2640
            Width           =   900
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "寄件信箱："
            Height          =   180
            Index           =   2
            Left            =   135
            TabIndex        =   25
            Top             =   2040
            Width           =   900
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "預定發信：日期："
            Height          =   180
            Index           =   7
            Left            =   135
            TabIndex        =   24
            Top             =   3000
            Width           =   1440
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "時間："
            Height          =   180
            Left            =   2655
            TabIndex        =   23
            Top             =   3000
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "實際發信："
            Height          =   180
            Index           =   8
            Left            =   135
            TabIndex        =   22
            Top             =   3360
            Width           =   900
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "實際發信數："
            Height          =   180
            Index           =   9
            Left            =   135
            TabIndex        =   21
            Top             =   3690
            Width           =   1080
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "失敗數："
            Height          =   180
            Index           =   10
            Left            =   2655
            TabIndex        =   20
            Top             =   3690
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "預定發信數："
            Height          =   180
            Index           =   11
            Left            =   4455
            TabIndex        =   19
            Top             =   3000
            Width           =   1080
         End
      End
      Begin MSForms.Label lblUpdate 
         Height          =   228
         Left            =   4632
         TabIndex        =   74
         Top             =   4656
         Width           =   3996
         Caption         =   "Update : "
         Size            =   "7048;402"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblCreate 
         Height          =   228
         Left            =   192
         TabIndex        =   73
         Top             =   4656
         Width           =   3996
         Caption         =   "Create : "
         Size            =   "7048;402"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label13 
         Caption         =   "（測試資料庫）"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   8.5
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   3312
         TabIndex        =   62
         Top             =   252
         Visible         =   0   'False
         Width           =   1296
      End
      Begin VB.Label lblTestName 
         AutoSize        =   -1  'True
         Height          =   180
         Left            =   7965
         TabIndex        =   57
         Top             =   510
         Width           =   45
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "排程代碼："
         Height          =   180
         Index           =   1
         Left            =   270
         TabIndex        =   31
         Top             =   510
         Width           =   900
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "測試信箱："
         Height          =   180
         Left            =   3420
         TabIndex        =   30
         Top             =   510
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "排程日期:"
         Height          =   180
         Index           =   4
         Left            =   -74850
         TabIndex        =   16
         Top             =   480
         Width           =   765
      End
      Begin VB.Line Line2 
         X1              =   -73260
         X2              =   -72840
         Y1              =   570
         Y2              =   570
      End
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   520
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   9000
      _ExtentX        =   15875
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
   Begin VB.Label Label11 
      Caption         =   "4.用Mail Server發信(backup會收到)，只適用於ipdept@taie.com.tw"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   2
      Left            =   720
      TabIndex        =   68
      Top             =   6450
      Width           =   7485
   End
   Begin VB.Label Label11 
      Caption         =   "3.主旨內有 [Our Ref:XXXXXXXXX...]為特殊主旨會置換成編號並勾(backup會收到)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   1
      Left            =   720
      TabIndex        =   67
      Top             =   6225
      Width           =   7485
   End
   Begin VB.Label Label12 
      Caption         =   "注意："
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
      Left            =   30
      TabIndex        =   61
      Top             =   5760
      Width           =   705
   End
   Begin VB.Label Label11 
      Caption         =   "2.以ipdept寄發之信件主旨不要放個人Initial，因為會影響到分信狀況"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   0
      Left            =   720
      TabIndex        =   60
      Top             =   6000
      Width           =   7365
   End
   Begin VB.Label Label10 
      Caption         =   "1.寄發對象為國內電子報,專利雙週報,顧問電子報時,都會同時寄給董事長個人信箱!!!"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   720
      TabIndex        =   59
      Top             =   5760
      Width           =   7410
   End
End
Attribute VB_Name = "frm140410_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/10/16 Form2.0已修改 ; grdList從MSFlexGrid改為MSHFlexGrid
'Memo By Sonia 2021/12/10 Form2.0不用改
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Created by Morgan 2011/12/12

'Modify By Sindy 2019/3/13 [ISD ==> [Our Ref:
' [ISDXXXXXXXXX...] 特殊主旨信件--例如主旨為Meeting at 2018 APAA, New Delhi [ISDXXXXXXXXX.2018 INTA] (EY/wc)
'           --2018/10/19 [ISDY46536010.2018 INTA] (EY/wc) 遇[ISD時切割主旨[前存於ms02欄位,[後9個X置換成Y編號存於msd07
'                                發信時主旨為ms02||msd07,且要勾選是否用Mail Server發信(backup會收到),往來記錄才會有資料
'國外電子報 -- 2018/3/13 改只寄非台灣籍的代理人(Y編號)之代表號、R編號之代表號、R編號下聯絡人 --Widen
'           -- Add By Sindy 2020/12/30 + Y編號下聯絡人是否寄電子報=Y者才要寄發 --Widen
'國內電子報 -- 客戶業務區非國外部(Fxx)或無智權人員但國籍為台灣者(寄代表信箱及其他信箱x3)、國內潛在客戶及國籍為020之國外潛在客戶
'              以上均含聯絡人並另加寄董事長信箱chinchanglin@yahoo.com.tw
'              排除:有異常狀態(倒閉、停業...)、是否不寄電子報設 N
'專利雙週報 -- 以國內電子報為基礎,但加大陸非C類代理人(含聯絡人)
'顧問電子報
'國外部價目表 -- 非C類代理人(含聯絡人),排除台灣大陸國籍者

Option Explicit

Dim ActionEdit As Integer
'執行各項功能的權限
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
Dim m_CurrSel As Integer
'------------

'開啟檔案對話框
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias _
  "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Type OPENFILENAME
  lStructSize As Long
  hwndOwner As Long
  hInstance As Long
  lpstrFilter As String
  lpstrCustomFilter As String
  nMaxCustFilter As Long
  nFilterIndex As Long
  lpstrFile As String
  nMaxFile As Long
  lpstrFileTitle As String
  nMaxFileTitle As Long
  lpstrInitialDir As String
  lpstrTitle As String
  Flags As Long
  nFileOffset As Integer
  nFileExtension As Integer
  lpstrDefExt As String
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type

Const MailBefore$ = "IMCEAEX-_O=TAIE_OU=DOMAIN_CN=RECIPIENTS_CN="
Const MailAfter$ = "@taie.com.tw"

Dim Result$, Sec%

Dim fso As New FileSystemObject
Dim ts As TextStream
Dim m_bolTestOK As Boolean '是否有寄測試
Dim m_DepCode As String '過濾用部門代碼
Dim m_GDPR As Boolean 'Added by Morgan 2018/8/10
Public m_AutoRun As Boolean, m_Schedule1 As Boolean, m_Schedule2 As Boolean 'Added by Morgan 2024/6/6


Private Sub cboDisplayName_KeyPress(KeyAscii As Integer)
   If Pub_StrUserSt03 <> "M51" Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub cboEmail_KeyPress(KeyAscii As Integer)
   If Pub_StrUserSt03 <> "M51" Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub chkByMailServer_Validate(Cancel As Boolean)
   'Add By Sindy 2019/3/19 增加此控管,以免信件未發至backup
   If chkByMailServer.Value = 1 And UCase(cboEmail.Text) <> UCase("ipdept@taie.com.tw") Then
      MsgBox "點選 <用Mail Server發信(backup會收到)>" & vbCrLf & _
             "信箱必須是 <ipdept@taie.com.tw> !!!"
      Cancel = True
      chkByMailServer.Value = 0
      cboEmail.SetFocus
   End If
End Sub

Private Sub chkMS26_Click()
   If chkMS26.Value = vbChecked Then
      txtMS26.Enabled = True
   Else
      txtMS26 = ""
      txtMS26.Enabled = False
   End If
End Sub

Private Sub cmdDetect_Click()
   Timer2.Interval = 10000
   Timer2.Enabled = True
   Frame2.Visible = True
   RefreshBar True
End Sub

Private Sub RefreshBar(Optional pIsInit As Boolean)
   If pIsInit Then
      
      strExc(0) = "select count(*) from mailscheduledetail where msd01=" & txtNo
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         ProgressBar1.max = RsTemp(0)
         ProgressBar1.Min = 0
         ProgressBar1.Value = 0
      Else
         ProgressBar1.max = 0
         ProgressBar1.Min = 0
         ProgressBar1.Value = 0
      End If
      lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
      DoEvents
   End If
   
   If ProgressBar1.max > 0 Then
            
      strExc(0) = "select count(*) from mailscheduledetail where msd01=" & txtNo & " and msd03>0"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         ProgressBar1.Value = RsTemp(0)
         lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
         DoEvents
      End If
   End If
End Sub

Private Sub cmdImPort_Click(Index As Integer)
   Dim stFileName As String
   Dim OpenFile As OPENFILENAME
   Dim lReturn As Long
   Dim sFilter As String
   Dim SNo As String
   Dim stCaption As String
   Dim idx As Integer
   'Addd by Lydia 2018/01/23
   Dim intA As Integer
   Dim rsAD As New ADODB.Recordset
   Dim bChkFA122 As Boolean '是否為促銷信
   Dim varTmp As Variant, ii As Integer
   
On Error GoTo ErrHnd

   'Add By Sindy 2023/8/22 要匯入特殊名單寄信
   If chkSpecList.Value = 1 Then
      strExc(1) = "select TBNP01 from TMBulletinNp where TBNP08='M' and TBNP10 is null order by TBNP01 asc"
      intA = 1
      Set rsAD = ClsLawReadRstMsg(intA, strExc(1))
      If intA = 1 Then
         rsAD.MoveFirst
         lstImport(Index).Clear
         Do While Not rsAD.EOF
            lstImport(Index).AddItem Trim(rsAD.Fields("TBNP01"))
            rsAD.MoveNext
         Loop
      End If
      lblImpCount(Index).Caption = lstImport(Index).ListCount
      MsgBox "匯入完畢！", vbInformation
      Exit Sub
   End If
   '2023/8/22 END
   
   'Added by Lydia 2018/01/23 若指定X/Y編號時，彈訊息
   bChkFA122 = False
   If Index = 1 Then
       If MsgBox("請先詢問發信人員是否為促銷信？", vbYesNo + vbDefaultButton2) = vbYes Then
           bChkFA122 = True
       End If
   End If
   'end 2018/01/23
   
   'Add By Sindy 2020/7/7
   Dim stMS15 As String
   GetChoice stMS15, False
   '2020/7/7 END
   
   OpenFile.lStructSize = Len(OpenFile)
   OpenFile.hwndOwner = Me.hWnd
   OpenFile.hInstance = App.hInstance
   sFilter = "文字檔(*.TXT)" & Chr(0) & "*.txt" & Chr(0) & "所有檔案" & Chr(0) & "*.*" & Chr(0)
   OpenFile.lpstrFilter = sFilter
   OpenFile.nFilterIndex = 1
   OpenFile.lpstrFile = String(257, 0)
   OpenFile.nMaxFile = Len(OpenFile.lpstrFile) - 1
   OpenFile.lpstrFileTitle = OpenFile.lpstrFile
   OpenFile.nMaxFileTitle = OpenFile.nMaxFile
   OpenFile.lpstrInitialDir = PUB_Getdesktop
   stCaption = "匯入清單"
   OpenFile.lpstrTitle = stCaption
   OpenFile.Flags = 0
   lReturn = GetOpenFileName(OpenFile)
   If lReturn <> 0 Then
      stFileName = Trim(OpenFile.lpstrFile)
      lstImport(Index).Clear
      lblImpCount(Index).Caption = ""
      If fso.FileExists(stFileName) Then
         Set ts = fso.OpenTextFile(stFileName)
         Do While Not ts.AtEndOfStream
            SNo = Replace(RTrim(ts.ReadLine), " ", "")
            'Modified by Lydia 2018/01/23 若為促銷信,排除代理人有設定"一定不要寄=N"
            'If SNo <> "" Then lstImport(Index).AddItem SNo
            If SNo <> "" Then
               'Modify By Sindy 2020/7/7 +聯絡人編號,檢查編碼規則
               '其他(指定編號)
               If stMS15 = 2 ^ 8 Then
                  If InStr(SNo, "-") > 0 Then
                     If Len(SNo) < 12 Then
                        varTmp = Split(SNo, "-")
                        SNo = ChangeCustomerL(varTmp(0)) & "-" & varTmp(1)
                     ElseIf Len(SNo) > 12 Then
                        MsgBox SNo & "編碼有誤!"
                        GoTo ErrHnd
                     End If
                  ElseIf Len(SNo) = 11 Then
                     SNo = Left(SNo, 9) & "-" & Right(SNo, 2)
                  ElseIf Len(SNo) <= 9 Then
                     If Len(SNo) < 9 Then
                        SNo = ChangeCustomerL(SNo)
                     End If
                  Else
                     MsgBox SNo & "編碼有誤!"
                     GoTo ErrHnd
                  End If
               End If
               '2020/7/7 END
               If bChkFA122 = True And Left(SNo, 1) = "Y" Then
                    'Modify By Sindy 2020/7/6 加聯絡人編號
                    'strExc(1) = "select fa01 from fagent where fa01||fa02='" & ChangeCustomerL(SNo) & "' and nvl(fa122,'Y')<>'N' "
                    strExc(1) = "select fa01 from fagent where fa01||fa02='" & ChangeCustomerL(Left(SNo, 9)) & "' and nvl(fa122,'Y')<>'N' "
                    '2020/7/6 END
                    intA = 1
                    Set rsAD = ClsLawReadRstMsg(intA, strExc(1))
                    If intA = 1 Then
                        lstImport(Index).AddItem SNo
                    End If
               Else
                    lstImport(Index).AddItem SNo
               End If
            End If
            'end 2018/01/23
         Loop
         ts.Close
      End If
      lblImpCount(Index).Caption = lstImport(Index).ListCount
   End If
   
   If bChkFA122 = True Then Set rsAD = Nothing 'Added by Lydia 2018/01/23
   
ErrHnd:
   If Err.Number <> 0 Then
      MsgBox Err.Description
   End If
End Sub

Private Sub cmdQuery_Click(Index As Integer)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
    
    '若頁籤在 基本資料，就不管
    If SSTab1.Tab = 0 Then Exit Sub
    
    If txtQueryDate(0).Text = "" And txtQueryDate(1).Text = "" Then
        MsgBox "請輸入排程日期範圍!!!", vbExclamation + vbOKOnly
        txtQueryDate(0).SetFocus
        Exit Sub
    End If
    
   Screen.MousePointer = vbHourglass
   Me.GrdList.MousePointer = flexHourglass
   If QueryData() = False Then
       strTit = "查詢資料"
       strMsg = "無資料"
       nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   End If
   Me.GrdList.MousePointer = flexDefault
   Screen.MousePointer = vbDefault
   
End Sub

Private Function QueryData() As Boolean
   Dim nRow As Integer
   Dim stCon As String
   
   QueryData = False
   InitialGridList
   stCon = ""
   
   If Pub_StrUserSt03 <> "M51" Then
      stCon = stCon & " And MS22='" & m_DepCode & "'"
   End If
   
   If txtQueryDate(0).Text <> "" Then
       stCon = stCon & " And MS08>=" & DBDATE(txtQueryDate(0).Text) & " "
   End If
   If txtQueryDate(1).Text <> "" Then
       stCon = stCon & " And MS08<=" & DBDATE(txtQueryDate(1).Text) & " "
   End If
   
   strExc(0) = "select '',ms01,SUBSTRB(SQLDATET(MS08)||' '||SQLTIME6(MS09),1,18)" & _
     ",MS02,SUBSTRB(SQLDATET(MS16)||' '||SQLTIME6(MS17),1,18)" & _
     ",SUBSTRB(SQLDATET(MS11)||' '||SQLTIME6(MS12),1,18)" & _
     " from mailschedule where 1=1 " & stCon & " order by 2 desc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      UpdateGridList RsTemp
      QueryData = True
   End If
End Function

' 初始化列表
Private Sub InitialGridList()
   m_CurrSel = 0
   With GrdList
   .Clear
   .Rows = 1
   .Cols = 6
    
   .row = 0
   .col = 0
   .CellAlignment = flexAlignCenterCenter
   .ColWidth(0) = 300
   .ColAlignment(0) = flexAlignCenterCenter
   
   .col = 1
   .Text = "代碼"
   .CellAlignment = flexAlignCenterCenter
   .ColWidth(1) = 450
   .ColAlignment(1) = flexAlignCenterCenter
    
   .col = 2
   .Text = "預定發信時間"
   .CellAlignment = flexAlignCenterCenter
   .ColWidth(2) = 1500
   .ColAlignment(2) = flexAlignLeftCenter
       
   .col = 3
   .Text = "主旨"
   .CellAlignment = flexAlignLeftCenter
   .ColWidth(3) = 4575
   .ColAlignment(3) = flexAlignLeftCenter
    
   .col = 4
   .Text = "實際發信時間(起)"
   .CellAlignment = flexAlignCenterCenter
   .ColWidth(4) = 1500
   .ColAlignment(4) = flexAlignLeftCenter
    
   .col = 5
   .Text = "實際發信時間(迄)"
   .CellAlignment = flexAlignCenterCenter
   .ColWidth(5) = 1500
   .ColAlignment(5) = flexAlignLeftCenter
   End With
End Sub

'Modified by Morgan 2023/2/23 + or cu80='解除對造'
Private Function GetSQL2(pMS01 As Long) As String
   Dim stVTB As String
   'MailScheduleImport
   
   'Modified by Morgan 2021/1/29 +只抓9碼的編號length(msi02)=9
   stVTB = "" & _
      " select fa16 MSD02,FA01||FA02 MSD06 from MailScheduleImport,fagent where msi01=" & pMS01 & " and length(msi02)=9 and fa01(+)=substr(msi02,1,8) and fa02(+)=substr(msi02,9,1) and fa69 is null and instr(fa16,'@')>0"
   If chkMainOnly.Value = 0 Then
      '代理人其他信箱及聯絡人,編號長度為9
      stVTB = stVTB & " Union" & _
         " select fa80 MSD02,FA01||FA02 MSD06 from MailScheduleImport,fagent where msi01=" & pMS01 & " and length(msi02)=9 and fa01(+)=substr(msi02,1,8) and fa02(+)=substr(msi02,9,1) and fa69 is null" & _
         " and instr(fa80,'@')>0" & _
         " Union" & _
         " select fa81 MSD02,FA01||FA02 MSD06 from MailScheduleImport,fagent where msi01=" & pMS01 & " and length(msi02)=9 and fa01(+)=substr(msi02,1,8) and fa02(+)=substr(msi02,9,1) and fa69 is null" & _
         " and instr(fa81,'@')>0" & _
         " Union" & _
         " select fa82 MSD02,FA01||FA02 MSD06 from MailScheduleImport,fagent where msi01=" & pMS01 & " and length(msi02)=9 and fa01(+)=substr(msi02,1,8) and fa02(+)=substr(msi02,9,1) and fa69 is null" & _
         " and instr(fa82,'@')>0" & _
         " Union" & _
         " select pcc08 MSD02,PCC01||'0-'||PCC02 MSD06 from MailScheduleImport,fagent,potcustcont where msi01=" & pMS01 & " and length(msi02)=9 and fa01(+)=substr(msi02,1,8) and fa02(+)=substr(msi02,9,1) and fa69 is null" & _
         " and pcc01(+)=fa01 and instr(pcc08,'@')>0"
   End If
   
   If stVTB <> "" Then stVTB = stVTB & " UNION "
   
   stVTB = stVTB & _
      " select cu20 MSD02, CU01||CU02 MSD06 from MailScheduleImport,customer where msi01=" & pMS01 & " and length(msi02)=9 and cu01(+)=substr(msi02,1,8) and cu02(+)=substr(msi02,9,1) and (cu80 is null or cu80='業務自行處理' or cu80='解除對造')" & _
      " and instr(cu20,'@')>0"
   
   If chkMainOnly.Value = 0 Then
      '客戶其他信箱及聯絡人,編號長度為9
      stVTB = stVTB & " Union" & _
         " select cu116 MSD02, CU01||CU02 MSD06 from MailScheduleImport,customer where msi01=" & pMS01 & " and length(msi02)=9 and cu01(+)=substr(msi02,1,8) and cu02(+)=substr(msi02,9,1) and (cu80 is null or cu80='業務自行處理' or cu80='解除對造')" & _
         " and instr(cu116,'@')>0" & _
         " Union" & _
         " select cu117 MSD02, CU01||CU02 MSD06 from MailScheduleImport,customer where msi01=" & pMS01 & " and length(msi02)=9 and cu01(+)=substr(msi02,1,8) and cu02(+)=substr(msi02,9,1) and (cu80 is null or cu80='業務自行處理' or cu80='解除對造')" & _
         " and instr(cu117,'@')>0" & _
         " Union" & _
         " select cu118 MSD02, CU01||CU02 MSD06 from MailScheduleImport,customer where msi01=" & pMS01 & " and length(msi02)=9 and cu01(+)=substr(msi02,1,8) and cu02(+)=substr(msi02,9,1) and (cu80 is null or cu80='業務自行處理' or cu80='解除對造')" & _
         " and instr(cu118,'@')>0" & _
         " Union" & _
         " select pcc08 MSD02,PCC01||'0-'||PCC02 MSD06 from MailScheduleImport,customer,potcustcont where msi01=" & pMS01 & " and length(msi02)=9 and cu01(+)=substr(msi02,1,8) and cu02(+)=substr(msi02,9,1) and (cu80 is null or cu80='業務自行處理' or cu80='解除對造')" & _
         " and pcc01(+)=cu01 and instr(pcc08,'@')>0"
   End If
   
   If stVTB <> "" Then stVTB = stVTB & " UNION "
   
   'Add By Sindy 2018/6/20 + R編號
   '國外潛在客戶
   stVTB = stVTB & _
      " select pcu18 MSD02, pcu01||pcu02 MSD06 from MailScheduleImport,potcustomer where msi01=" & pMS01 & " and length(msi02)=9 and pcu01(+)=substr(msi02,1,8) and pcu02(+)=substr(msi02,9,1) and pcu39 is null and instr(pcu18,'@')>0"
   'Add By Sindy 2019/10/1
   If chkMainOnly.Value = 0 Then
      '加聯絡人,編號長度為9
      stVTB = stVTB & " Union" & _
         " select pcc08 MSD02,PCC01||'0-'||PCC02 MSD06 from MailScheduleImport,potcustomer,potcustcont where msi01=" & pMS01 & " and length(msi02)=9 and pcu01(+)=substr(msi02,1,8) and pcu02(+)=substr(msi02,9,1) and pcu39 is null" & _
         " and pcc01(+)=pcu01 and instr(pcc08,'@')>0"
   End If
   '2019/10/1 END
   
   If stVTB <> "" Then stVTB = stVTB & " UNION "
   
   '國內潛在客戶
   stVTB = stVTB & _
      " select poc09 MSD02, poc01||poc02 MSD06 from MailScheduleImport,potcustomer1 where msi01=" & pMS01 & " and length(msi02)=9 and poc01(+)=substr(msi02,1,8) and poc02(+)=substr(msi02,9,1) and poc14 is null and instr(poc09,'@')>0"
   '2018/6/20 END
   'Add By Sindy 2019/10/1
   If chkMainOnly.Value = 0 Then
      '加聯絡人,編號長度為9
      stVTB = stVTB & " Union" & _
         " select pcc08 MSD02,PCC01||'0-'||PCC02 MSD06 from MailScheduleImport,potcustomer1,potcustcont where msi01=" & pMS01 & " and length(msi02)=9 and poc01(+)=substr(msi02,1,8) and poc02(+)=substr(msi02,9,1) and poc14 is null" & _
         " and pcc01(+)=poc01 and instr(pcc08,'@')>0"
   End If
   '2019/10/1 END
   
   'Added by Morgan 2021/1/29
   '抓聯絡人,編號長度為12
   If chkMainOnly.Value = 1 Then
      stVTB = stVTB & " Union" & _
         " select pcc08 MSD02,PCC01||'0-'||PCC02 MSD06 from MailScheduleImport,potcustcont where msi01=" & pMS01 & " and length(msi02)=12 and pcc01(+)=substr(msi02,1,8) and pcc02(+)=substr(msi02,11) and instr(pcc08,'@')>0"
   End If
   'end 2021/1/29
   
   GetSQL2 = "select * from (" & stVTB & ") X"
   
End Function

'Modified by Morgan 2023/2/23 + or cu80='解除對造'
Private Function GetSubSQL(pChoice As Integer) As String
   Dim stVTB As String
   Dim stConFa As String, stConFaE As String '代理人
   Dim stConCu As String, stConCuE As String '客戶
   Dim stConPcu As String, stConPcuE As String '國外潛在客戶
   Dim stConPcu1 As String, stConPcu1E As String '國內潛在客戶
   Dim stConECu As String '國外開拓客戶
   Dim stConPccE As String '聯絡人
   Dim stConFaPccE As String '代理人的聯絡人 Add By Sindy 2020/12/30
   Dim strFA10 As String 'Modify By Sindy 2025/9/25
   
   stVTB = ""
   stConFa = ""
   stConFaE = ""
   stConCu = ""
   stConCuE = ""
   stConPcu = ""
   stConPcuE = ""
   stConPcu1 = ""
   stConPcu1E = ""
   stConPccE = ""
   stConFaPccE = "" 'Add By Sindy 2020/12/30
   stConECu = ""
   
   '顧問電子報 Add By Sindy 2011/3/18
   If pChoice = 3 Then
      '2011/8/8 modify by sonia 顧問案件申請人1為X65299000(與謝律師合作案件)則改抓申請人2
      '客戶信箱
      'Modify By Sindy 2024/1/26 增加抓ACS的112智財顧問
      '改抓顧問專用信箱寄發, 且是否寄發顧問電子報為Y; 不抓聯絡人信箱
      stVTB = "SELECT cu199 MSD02,CU01||CU02 MSD06 From CASEPROGRESS,HIRECASE,CUSTOMER " & _
               "WHERE HC01=CP01(+) AND HC02=CP02(+) AND HC03=CP03(+) AND HC04=CP04(+) " & _
               "AND CP54>" & strSrvDate(1) & " AND CP10='0' AND CP57 IS NULL AND CP27 is null " & _
               "AND (SUBSTR(decode(HC05,'X65299000',HC24,HC05),1,8)=CU01(+) AND SUBSTR(decode(HC05,'X65299000',HC24,HC05),9,1)=CU02(+)) " & _
               "AND cu02='0' AND (cu80 is null or cu80='業務自行處理' or cu80='解除對造') AND instr(cu199,'@')>0 AND cu199 is not null " & _
               "AND CU153='Y' "
      stVTB = stVTB & "Union " & _
               "SELECT cu199 MSD02,CU01||CU02 MSD06 From CASEPROGRESS,LAWCASE,CUSTOMER " & _
               "WHERE LC01=CP01(+) AND LC02=CP02(+) AND LC03=CP03(+) AND LC04=CP04(+) " & _
               "AND CP54>" & strSrvDate(1) & " AND CP01='ACS' AND CP10='112' AND CP57 IS NULL AND CP27 is null " & _
               "AND (SUBSTR(decode(LC11,'X65299000',LC43,LC11),1,8)=CU01(+) AND SUBSTR(decode(LC11,'X65299000',LC43,LC11),9,1)=CU02(+)) " & _
               "AND cu02='0' AND (cu80 is null or cu80='業務自行處理' or cu80='解除對造') AND instr(cu199,'@')>0 AND cu199 is not null " & _
               "AND CU153='Y' "
'      stVTB = "SELECT cu20 MSD02,CU01||CU02 MSD06 From CASEPROGRESS,HIRECASE,CUSTOMER " & _
'                     "WHERE HC01=CP01(+) AND HC02=CP02(+) AND HC03=CP03(+) AND HC04=CP04(+) " & _
'                     "AND CP54>" & strSrvDate(1) & " AND CP10='0' AND CP57 IS NULL AND CP27 is null " & _
'                     "AND (SUBSTR(decode(HC05,'X65299000',HC24,HC05),1,8)=CU01(+) AND SUBSTR(decode(HC05,'X65299000',HC24,HC05),9,1)=CU02(+)) " & _
'                     "AND cu02='0' AND (cu80 is null or cu80='業務自行處理' or cu80='解除對造') AND instr(cu20,'@')>0 " & _
'                     "AND CU153='Y' "
'      stVTB = stVTB & "Union " & _
'                     "SELECT cu116 MSD02,CU01||CU02 MSD06 From CASEPROGRESS,HIRECASE,CUSTOMER " & _
'                     "WHERE HC01=CP01(+) AND HC02=CP02(+) AND HC03=CP03(+) AND HC04=CP04(+) " & _
'                     "AND CP54>" & strSrvDate(1) & " AND CP10='0' AND CP57 IS NULL AND CP27 is null " & _
'                     "AND (SUBSTR(decode(HC05,'X65299000',HC24,HC05),1,8)=CU01(+) AND SUBSTR(decode(HC05,'X65299000',HC24,HC05),9,1)=CU02(+)) " & _
'                     "AND cu02='0' AND (cu80 is null or cu80='業務自行處理' or cu80='解除對造') AND instr(cu116,'@')>0 " & _
'                     "AND CU153='Y' "
'      stVTB = stVTB & "Union " & _
'                     "SELECT cu117 MSD02, CU01||CU02 MSD06 From CASEPROGRESS,HIRECASE,CUSTOMER " & _
'                     "WHERE HC01=CP01(+) AND HC02=CP02(+) AND HC03=CP03(+) AND HC04=CP04(+) " & _
'                     "AND CP54>" & strSrvDate(1) & " AND CP10='0' AND CP57 IS NULL AND CP27 is null " & _
'                     "AND (SUBSTR(decode(HC05,'X65299000',HC24,HC05),1,8)=CU01(+) AND SUBSTR(decode(HC05,'X65299000',HC24,HC05),9,1)=CU02(+)) " & _
'                     "AND cu02='0' AND (cu80 is null or cu80='業務自行處理' or cu80='解除對造') AND instr(cu117,'@')>0 " & _
'                     "AND CU153='Y' "
'      stVTB = stVTB & "Union " & _
'                     "SELECT cu118 MSD02, CU01||CU02 MSD06 From CASEPROGRESS,HIRECASE,CUSTOMER " & _
'                     "WHERE HC01=CP01(+) AND HC02=CP02(+) AND HC03=CP03(+) AND HC04=CP04(+) " & _
'                     "AND CP54>" & strSrvDate(1) & " AND CP10='0' AND CP57 IS NULL AND CP27 is null " & _
'                     "AND (SUBSTR(decode(HC05,'X65299000',HC24,HC05),1,8)=CU01(+) AND SUBSTR(decode(HC05,'X65299000',HC24,HC05),9,1)=CU02(+)) " & _
'                     "AND cu02='0' AND (cu80 is null or cu80='業務自行處理' or cu80='解除對造') AND instr(cu118,'@')>0 " & _
'                     "AND CU153='Y' "
'      '聯絡人信箱
'      stVTB = stVTB & "Union " & _
'                     "SELECT pcc08 MSD02, CU01||CU02 MSD06 From CASEPROGRESS,HIRECASE,CUSTOMER,potcustcont " & _
'                     "WHERE HC01=CP01(+) AND HC02=CP02(+) AND HC03=CP03(+) AND HC04=CP04(+) " & _
'                     "AND CP54>" & strSrvDate(1) & " AND CP10='0' AND CP57 IS NULL AND CP27 is null " & _
'                     "AND (SUBSTR(decode(HC05,'X65299000',HC24,HC05),1,8)=CU01(+) AND SUBSTR(decode(HC05,'X65299000',HC24,HC05),9,1)=CU02(+)) " & _
'                     "AND cu02='0' AND (cu80 is null or cu80='業務自行處理' or cu80='解除對造') " & _
'                     "AND pcc01(+)=cu01 AND instr(pcc08,'@')>0 " & _
'                     "AND pcc23='Y' "
      'Add By Sindy 2015/3/2 同時寄給董事長個人信箱
      stVTB = stVTB & " Union SELECT 'chinchanglin@yahoo.com.tw' MSD02,'None' MSD06 FROM DUAL"
      '2015/3/2 END
      GetSubSQL = stVTB
      Exit Function
   End If
                     
   'Added by Morgan 2018/8/10 --Widen
   'GDPR詢問信:歐洲區Y編號之代表號、歐洲區R編號之代表號以及歐洲區R編號下的聯絡人Email
   '不限制代理人性質或是否寄電子報
   '排除已回覆或已寄送
   If pChoice = 5 Then
      '代理人
      stVTB = "select trim(fa16) MSD02,FA01||FA02 MSD06" & _
         " From Nation, fagent" & _
         " where  na02='C20' and fa10(+)=na01 and fa02='0'" & _
         " and fa69 is null and instr(fa16,'@')>0 and FA123 is null"

      '潛在客戶
      stVTB = stVTB & " Union select trim(pcu18) MSD02,PCU01||PCU02 MSD06" & _
         " From Nation, potcustomer" & _
         " where na02='C20' and pcu09(+)=na01 and pcu02='0'" & _
         " and pcu39 is null and instr(pcu18,'@')>0 and PCU50 is null"
      
      '潛在客戶聯絡人(獨立看待,不管潛在客戶是否有回覆或寄送--Widen)
      stVTB = stVTB & " Union select trim(pcc08) MSD02,PCC01||'0-'||PCC02 MSD06" & _
         " From Nation, potcustomer, potcustcont" & _
         " where na02='C20' and pcu09(+)=na01 and pcu02='0'" & _
         " and pcu39 is null and pcc01(+)=pcu01 and instr(pcc08,'@')>0 and PCC26 is null"
      
      GetSubSQL = stVTB
      Exit Function
   
   'Addedby Morgan 2024/5/28
   '索取CF對帳單
   ElseIf pChoice = 10 Or pChoice = 11 Then
      'Modified by Morgan 2025/6/13 改只抓5年內有帳單的代理人--斯閔
      stVTB = "select nvl(fa105,nvl(fa79,fa16)) msd02,FA01||FA02 MSD06" & _
         " from (select distinct a1503 from acc150 where a1507 is null and a1502>=" & (strSrvDate(2) \ 10000 - 5) & "0101) V1,fagent" & _
         " where fa01(+)=substr(a1503,1,8) and fa02='0'" & _
         " and instr(nvl(fa105,nvl(fa79,fa16)),'@')>0 and fa133 is null"
      If pChoice = 10 Then
         stVTB = stVTB & " and fa10='020'"
      Else
         stVTB = stVTB & " and fa10<>'020'"
      End If
      GetSubSQL = stVTB
      Exit Function
   'end 2024/5/28
   End If
   'end 2018/8/10
   
   Select Case pChoice
   Case 0 '國外電子報
      'Modify by Morgan 2011/3/29 主檔與連絡人也改個別判斷
      'Modified by Morgan 2018/3/13 改只寄非台灣籍的代理人(Y編號)之代表號、R編號之代表號、R編號下聯絡人 --Widen
      'stConCu = stConCu & " and CU10>='011' and CU10<='999'"
      'stConCuE = stConCuE & " and cu132 is null"
      'stConFa = stConFa & " and FA10>='011' AND FA10<='999'"
      'stConFaE = stConFaE & " and fa97 is null"
      'stConPcu = stConPcu & " and PCU09>='011' AND PCU09<='999'"
      'stConPcuE = stConPcuE & " and pcu35 is null"
      'stConECu = stConECu & " and ECD10>='011' AND ECD10<='999' and ECD14 is null"
      'stConPccE = stConPccE & " and pcc10 is null"
      '代表號不寄則聯絡人也不寄
      'Modified by Morgan 2018/9/27 排除 GDPR選項為W或N者
      'Modify By Sindy 2020/12/30
      'stConFa = stConFa & " and FA10>='011' AND FA10<='999' and fa97 is null and fa76 in ('A','B') and NVL(FA123,'Y')='Y'"
      'Modify By Sindy 2025/5/27 增加可以抓取某一國籍資料 + IIf(FrameCU.Visible = True And Trim(textCU01) <> "", " and FA10='" & Trim(textCU01) & "'", "")
      'Modify By Sindy 2025/9/25
      strFA10 = ""
      If FrameCU.Visible = True And Trim(textCU01) <> "" Then
         If InStr(textCU01, ",") = 0 Then '單一國家
            strFA10 = " ='" & Trim(textCU01) & "'"
         Else
            'Run VB:抓多國時
            '           日本(011)、韓國(012)、馬來西亞(018)、泰國(019)、加拿大(102) 墨西哥(104) 、德國(231)籍
            'strFA10 = " in('011','012','018','019','102','104','231')"
            strFA10 = " in('" & Replace(textCU01, ",", "','") & "')"
         End If
      End If
      '2025/9/25 END
      stConFa = stConFa & " and FA10>='011' AND FA10<='999' and fa76 in ('A','B') and NVL(FA123,'Y')='Y'" & _
                IIf(strFA10 <> "", " and FA10" & strFA10, "")
      stConFaE = stConFaE & " and fa97 is null"
      '2020/12/30 END
      stConPcu = stConPcu & " and PCU09>='011' AND PCU09<='999' and pcu35 is null and NVL(PCU50,'Y')='Y'" & _
                IIf(strFA10 <> "", " and PCU09" & strFA10, "")
      'Modified by Morgan 2020/3/27 +排除已離職--Widen
      stConPccE = stConPccE & " and pcc10 is null and NVL(PCC26,'Y')='Y' and pcc20 is null"
      'end 2018/3/13
      'Add By Sindy 2020/12/30 + Y編號下聯絡人是否寄電子報=Y者才要寄發 --Widen
      stConFaPccE = stConFaPccE & " and pcc10='Y' and NVL(PCC26,'Y')='Y' and pcc20 is null"
      '2020/12/30 END
      
   'Add By Sindy 2020/12/3
   Case 6 '國外電子報(日本籍)
      '同國外電子報,只抓日本籍
      'Modify By Sindy 2020/12/30
      'stConFa = stConFa & " and substr(FA10,1,3)>='011' AND substr(FA10,1,3)<='011' and fa97 is null and fa76 in ('A','B') and NVL(FA123,'Y')='Y'"
      stConFa = stConFa & " and substr(FA10,1,3)>='011' AND substr(FA10,1,3)<='011' and fa76 in ('A','B') and NVL(FA123,'Y')='Y'"
      stConFaE = stConFaE & " and fa97 is null"
      '2020/12/30 END
      stConPcu = stConPcu & " and substr(PCU09,1,3)>='011' AND substr(PCU09,1,3)<='011' and pcu35 is null and NVL(PCU50,'Y')='Y'"
      stConPccE = stConPccE & " and pcc10 is null and NVL(PCC26,'Y')='Y' and pcc20 is null"
      'Add By Sindy 2020/12/30 + Y編號下聯絡人是否寄電子報=Y者才要寄發 --Widen
      stConFaPccE = stConFaPccE & " and pcc10='Y' and NVL(PCC26,'Y')='Y' and pcc20 is null"
      '2020/12/30 END
      
   Case 1 '國內電子報
      stConCu = stConCu & " and (SUBSTR(CU12,1,1)<>'F' OR (CU12 IS NULL AND CU10>='001' AND CU10<='008'))"
      stConCuE = stConCuE & " and cu132 is null"
      stConPccE = stConPccE & " and pcc10 is null"
      stConPcu1E = stConPcu1E & " and poc11 is null"
      stConPcu = stConPcu & " and pcu09='020'"
      stConPcuE = stConPcuE & " and pcu35 is null" 'Added by Morgan 2012/2/1
      
   Case 2 '專利雙週報
      'Modified by Morgan 2011/12/30 專利雙週報欄位改放 N(不寄)
      stConCu = stConCu & " and (SUBSTR(CU12,1,1)<>'F' OR (CU12 IS NULL AND CU10>='001' AND CU10<='008'))"
      stConCuE = stConCuE & " and cu145 is null"
      stConPcu1E = stConPcu1E & " and poc28 is null"
      'Add By Sindy 2011/3/11 +專利雙週報
      stConFa = stConFa & " and fa100 is null and fa76<>'C' and fa10='020'"
      '2011/3/11 End
      
      'Added by Morgan 2012/1/6 比照國內電子報
      stConPccE = stConPccE & " and pcc24 is null"
      stConPcu = stConPcu & " and pcu09='020'"
      stConPcuE = stConPcuE & " and pcu48 is null" 'Added by Morgan 2012/2/1
   
   'Added by Morgan 2012/2/15
   Case 4 '國外部價目表
      stConFa = stConFa & " and FA10>='011' AND FA10<>'020' and FA76<>'C'"
      
   End Select

   '國內潛在客戶
   Select Case pChoice
   '國內電子報
   'Modify by Morgan 2011/2/14 +專利雙週報
   Case 1, 2
      If stVTB <> "" Then stVTB = stVTB & " UNION "
      stVTB = stVTB & _
         " select POC09 MSD02,POC01||POC02 MSD06 from potcustomer1 where POC02='0' AND POC14 IS NULL" & stConPcu1 & stConPcu1E & _
         " and instr(POC09,'@')>0"
      'Add By Sindy 2015/3/2 同時寄給董事長個人信箱
      stVTB = stVTB & " Union SELECT 'chinchanglin@yahoo.com.tw' MSD02,'None' MSD06 FROM DUAL"
      '2015/3/2 END
   End Select

   '代理人
   Select Case pChoice
   '國外電子報
   'Modify By Sindy 2011/3/11 +專利雙週報
   'Modified by Morgan 2012/2/15 +國外部價目表
   'Modified by Morgan 2018/3/13 國外電子報改只寄非台灣籍的代理人的代表號 --Widen
   'Case 0, 2, 4
   Case 0, 6 'Modify By Sindy 2020/12/3 + 6.國外電子報(日本籍)
      If stVTB <> "" Then stVTB = stVTB & " UNION "
      stVTB = stVTB & _
         " select fa16 MSD02,FA01||FA02 MSD06 from fagent where fa02='0' and fa69 is null" & stConFa & stConFaE & _
         " and instr(fa16,'@')>0"
      'Add By Sindy 2020/12/30 + Y編號下聯絡人是否寄電子報=Y者才要寄發 --Widen
      stVTB = stVTB & " Union" & _
            " select pcc08 MSD02,PCC01||'0-'||PCC02 MSD06 from fagent,potcustcont where fa02='0' and fa69 is null" & stConFa & _
            " and pcc01(+)=fa01 and instr(pcc08,'@')>0" & stConFaPccE
      '2020/12/30 END
      
   Case 2, 4
   'end 2018/3/13
      If stVTB <> "" Then stVTB = stVTB & " UNION "
      stVTB = stVTB & _
         " select fa16 MSD02,FA01||FA02 MSD06 from fagent where fa02='0' and fa69 is null" & stConFa & stConFaE & _
         " and instr(fa16,'@')>0"

      'Modify by Morgan 2009/3/10 +其他信箱也要寄--秀玲
      '代理人其他信箱及聯絡人
      stVTB = stVTB & " Union" & _
         " select fa80 MSD02,FA01||FA02 MSD06 from fagent where fa02='0' and fa69 is null" & stConFa & stConFaE & _
         " and instr(fa80,'@')>0" & _
         " Union" & _
         " select fa81 MSD02,FA01||FA02 MSD06 from fagent where fa02='0' and fa69 is null" & stConFa & stConFaE & _
         " and instr(fa81,'@')>0" & _
         " Union" & _
         " select fa82 MSD02,FA01||FA02 MSD06 from fagent where fa02='0' and fa69 is null" & stConFa & stConFaE & _
         " and instr(fa82,'@')>0" & _
         " Union" & _
         " select pcc08 MSD02,PCC01||'0-'||PCC02 MSD06 from fagent,potcustcont where fa02='0' and fa69 is null" & stConFa & _
         " and pcc01(+)=fa01 and instr(pcc08,'@')>0" & stConPccE
   End Select

   '國外潛在客戶
   Select Case pChoice
   '國外電子報
   '國內電子報(國籍為020者) Add by Morgan 2010/10/4
   '專利雙週報 Added by Morgan 2012/1/6
   Case 0, 1, 2, 6 'Modify By Sindy 2020/12/3 + 6.國外電子報(日本籍)
      If stVTB <> "" Then stVTB = stVTB & " UNION "
      stVTB = stVTB & _
         " select pcu18 MSD02,PCU01||PCU02 MSD06 from potcustomer where pcu02='0' and pcu39 is null" & stConPcu & stConPcuE & _
         " and instr(pcu18,'@')>0"
         
      stVTB = stVTB & " Union" & _
            " select pcc08 MSD02,PCC01||'0-'||PCC02 MSD06 from potcustomer,potcustcont where pcu02='0' and pcu39 is null" & stConPcu & _
            " and pcc01(+)=pcu01 and instr(pcc08,'@')>0" & stConPccE
   End Select
   
   'Add by Morgan 2009/3/16
   '開拓客戶
   'Modified by Morgan 2018/3/13 國外電子報改只寄非台灣籍的代理人的代表號 --Widen
   'Select Case pChoice
   ''國外電子報
   'Case 0
   '   If stVTB <> "" Then stVTB = stVTB & " UNION "
   '   stVTB = stVTB & _
   '      " select ECD13 MSD02,ecd02||'-'||LPAD(ecd01,6,'0') MSD06 from ExpandCusDetail where 1=1" & stConECu & _
   '      " and instr(ECD13,'@')>0"
   'End Select
   'end 2018/3/13
   
   '客戶,客戶聯絡人
   Select Case pChoice
   'Modified by Morgan 2018/3/13 國外電子報改只寄非台灣籍的代理人的代表號 --Widen
   'Case 0, 1, 2, 3
   Case 1, 2, 3
   'end 2018/3/13
      'Modify by Morgan 2009/3/10 +其他信箱也要寄--秀玲
      If stVTB <> "" Then stVTB = stVTB & " UNION "
      stVTB = stVTB & _
         " select cu20 MSD02, CU01||CU02 MSD06 from customer where cu02='0' and (cu80 is null or cu80='業務自行處理' or cu80='解除對造')" & stConCu & stConCuE & _
         " and instr(cu20,'@')>0"

      '客戶其他信箱及聯絡人
      stVTB = stVTB & " Union" & _
         " select cu116 MSD02, CU01||CU02 MSD06 from customer where cu02='0' and (cu80 is null or cu80='業務自行處理' or cu80='解除對造')" & stConCu & stConCuE & _
         " and instr(cu116,'@')>0" & _
         " Union" & _
         " select cu117 MSD02, CU01||CU02 MSD06 from customer where cu02='0' and (cu80 is null or cu80='業務自行處理' or cu80='解除對造')" & stConCu & stConCuE & _
         " and instr(cu117,'@')>0" & _
         " Union" & _
         " select cu118 MSD02, CU01||CU02 MSD06 from customer where cu02='0' and (cu80 is null or cu80='業務自行處理' or cu80='解除對造')" & stConCu & stConCuE & _
         " and instr(cu118,'@')>0"

      '客戶聯絡人
      'Modified by Morgan 2012//1/6 聯絡人也寄--郭
      'If pChoice <> 2 Then 'Add by Morgan 2011/2/14 +專利雙週報先不寄給聯絡人(有可能客戶也不寄)--郭
         stVTB = stVTB & " Union" & _
            " select pcc08 MSD02,PCC01||'0-'||PCC02 MSD06 from customer,potcustcont where cu02='0' and (cu80 is null or cu80='業務自行處理' or cu80='解除對造')" & stConCu & _
            " and pcc01(+)=cu01 and instr(pcc08,'@')>0" & stConPccE
      'End If
   End Select
   
   GetSubSQL = stVTB
End Function

Private Function GetSql(Optional pEstimate As Boolean = False) As String
   Dim stScript As String
   Dim stVTB As String, stSubSQL As String
   Dim ii As Integer, iChoice As Integer
   
   stVTB = ""
   
   For ii = 0 To lstMailChoice.ListCount - 1
      If lstMailChoice.Selected(ii) = True Then
         iChoice = lstMailChoice.ITEMDATA(ii)
         If iChoice <> 8 Then
            stSubSQL = GetSubSQL(iChoice)
            If stVTB <> "" Then stVTB = stVTB & " UNION "
            stVTB = stVTB & stSubSQL
         End If
      End If
   Next
   
   If stVTB <> "" Then
      If pEstimate = True Then
         stScript = "select count(DISTINCT MSD02) from (" & stVTB & ") X"
      Else
         stScript = "select * from (" & stVTB & ") X"
      End If
   End If

   GetSql = stScript
End Function

Private Sub cmdEstimate_Click()
   Screen.MousePointer = vbHourglass
   txtCount = ""
   If GetChoice() = True Then
      strExc(0) = GetSql(True)
      If strExc(0) <> "" Then
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            txtCount = RsTemp.Fields(0)
         End If
      End If
   Else
      MsgBox "請勾選寄發對象！"
      lstMailChoice.SetFocus
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmdOpen_Click()
   Dim stFileName As String
   Dim OpenFile As OPENFILENAME
   Dim lReturn As Long
   Dim sFilter As String

On Error GoTo ErrHnd
   stFileName = PUB_Getdesktop
   OpenFile.lStructSize = Len(OpenFile)
   OpenFile.hwndOwner = Me.hWnd
   If chkOutlook.Value = vbChecked Then
      OpenFile.hInstance = App.hInstance
      sFilter = "Outlook 範本(*.oft)" & Chr(0) & "*.oft" & Chr(0) & "所有檔案" & Chr(0) & "*.*" & Chr(0)
   Else
      OpenFile.hInstance = App.hInstance
         sFilter = "電子信(*.eml)" & Chr(0) & "*.eml" & Chr(0) & "單一網頁(*.mht)" & Chr(0) & "*.mht" & Chr(0) & "所有檔案" & Chr(0) & "*.*" & Chr(0)
   End If
   OpenFile.lpstrFilter = sFilter
   OpenFile.nFilterIndex = 1
   OpenFile.lpstrFile = String(257, 0)
   OpenFile.nMaxFile = Len(OpenFile.lpstrFile) - 1
   OpenFile.lpstrFileTitle = OpenFile.lpstrFile
   OpenFile.nMaxFileTitle = OpenFile.nMaxFile
   OpenFile.lpstrInitialDir = stFileName
   OpenFile.lpstrTitle = "開啟郵件樣本"
   OpenFile.Flags = 0
   lReturn = GetOpenFileName(OpenFile)
   If lReturn <> 0 Then
      txtSample = Trim(OpenFile.lpstrFile)
   End If

ErrHnd:
   If Err.Number <> 0 Then
      MsgBox Err.Description
   End If
End Sub

Private Function SplitEmail(pMS01 As Long, Optional pUpdateRecNum As Boolean) As Long
   Dim arrToMail
   Dim lngAdd As Long, lngDel As Long
   Dim ii As Integer
   Dim stSQL As String
   Dim stMSD06 As String
   Dim stMSD07 As String 'Added by Morgan 2018/11/29
   Dim stEMails As String
   Dim stMsg As String 'Add by Amy 2025/09/02
   
   'Add by Morgan 2009/3/26 多個信箱放一起的資料
   lngAdd = 0
   'Modified by Morgan 2013/1/22 +跳行
   'Added by Morgan 2014/9/18+逗號
   stSQL = "select msd02,msd06,msd07 from MailScheduleDetail where msd01=" & pMS01 & " and (instr(msd02,';')>0 or instr(msd02,chr(13)||chr(10))>0 or instr(msd02,' ')>0 or instr(msd02,',')>0)"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, stSQL)
   If intI = 1 Then
      With RsTemp
      Do While Not .EOF
         'Modified by Morgan 2013/1/22
         '+跳行視為分隔
         'arrToMail = Split("" & .Fields("msd02"), ";")
         stEMails = Replace("" & .Fields("msd02"), vbCrLf, ";")
         'Added by Morgan 2013/3/13 +空白視為分隔
         stEMails = Replace(stEMails, " ", ";")
         'Added by Morgan 2014/9/18+逗號
         stEMails = Replace(stEMails, ",", ";")
         arrToMail = Split(stEMails, ";")
         'end 2013/1/22
         
         stMSD06 = "" & .Fields("msd06")
         stMSD07 = "" & .Fields("msd07") 'Added by Morgan 2018/11/29
         For ii = 0 To UBound(arrToMail)
            '去除前後的空白和跳行符號
            arrToMail(ii) = Trim(Replace(arrToMail(ii), vbCrLf, ""))
            If InStr(arrToMail(ii), "@") > 0 Then
               stSQL = "select msd02,msd06 from MailScheduleDetail where msd01=" & pMS01 & " and msd02='" & arrToMail(ii) & "'"
               intI = 1
               Set AdoRecordSet3 = ClsLawReadRstMsg(intI, stSQL)
               If intI = 0 Then
                  'Modified by Morgan 2018/11/29 +MSD07
                  stSQL = " insert into MailScheduleDetail(MSD01,MSD02,MSD06,MSD07) values(" & pMS01 & ",'" & arrToMail(ii) & "','" & stMSD06 & "','" & stMSD07 & "')"
                  cnnConnection.Execute stSQL, intI
                  lngAdd = lngAdd + intI
                  
               '若信箱已存在但對應的編號不同時該筆資料的編號要更新為 'None'(表多編號同信箱)
               ElseIf AdoRecordSet3(1) <> stMSD06 And AdoRecordSet3(1) <> "None" Then
                  'Modified by Morgan 2018/6/28 多編號改放最小號+"+"(原放 None)
                  'Modified by Morgan 2018/8/10 改原編號無"+"號則+"+"
                  'Modified by Morgan 2019/8/1 信箱大小寫應視為相同
                  stSQL = "update MailScheduleDetail set msd06=msd06||'+' where msd01=" & pMS01 & " and lower(msd02)='" & LCase(arrToMail(ii)) & "' and instr(msd06,'+')=0"
                  cnnConnection.Execute stSQL, intI
                  
                  'Added by Morgan 2018/9/14 GDPR確認信不寄也要紀錄以便更新回覆結果
                  If m_GDPR Then
                     stSQL = " insert into MailScheduleDetail(MSD01,MSD02,MSD03,MSD06) values(" & pMS01 & ",'" & arrToMail(ii) & "',19221111,'" & stMSD06 & "')"
                     cnnConnection.Execute stSQL, intI
                     lngAdd = lngAdd + intI
                  End If
                  'end 2018/9/14
               End If
               If m_GDPR Then Exit For 'Added by Morgan 2018/9/14 GDPR確認信只要寄第一個信箱--Widen Ex:R14812000
            End If
         Next
         .MoveNext
      Loop
      End With
      'Added by Morgan 2014/9/18+逗號
      stSQL = "delete from MailScheduleDetail where msd01=" & pMS01 & " and (instr(msd02,';')>0 or instr(msd02,chr(13)||chr(10))>0 or instr(msd02,' ')>0 or instr(msd02,',')>0)"
      cnnConnection.Execute stSQL, lngDel
      lngAdd = lngAdd - lngDel
      
   End If 'Added by Morgan 2023/9/4

      'Added by Morgan 2018/5/11
      'Yahoo信箱不寄
      If chkNoYahoo.Value = vbChecked Then
         stSQL = "delete from MailScheduleDetail where msd01=" & pMS01 & " and instr(upper(msd02),upper('@yahoo.com'))>0"
         cnnConnection.Execute stSQL, lngDel
         lngAdd = lngAdd - lngDel
      End If
      'end 2018/5/11
      
      'Add by Amy 2025/09/02 Amy 測 雙週專利電子報,客戶不寄電子報回信連結用
      stMsg = ""
      If UCase(pub_DbTerminalName) <> 正式資料庫電腦名稱 And InStr(txtSubject, "專利電子報") > 0 And chkTestData.Visible = True Then
         If chkTestData.Value = vbChecked Then
            If SetTestMailData(pMS01, stMsg) = False Then
               MsgBox "設定測式資料失敗,請確認" & vbCrLf & _
                              "原因:" & stMsg
            Else
               MsgBox "目前明細筆數:" & Pub_GetField("MailScheduleDetail", "msd01=" & pMS01, "Count(*)")
            End If
         End If
      End If
      'end 2025/09/02
      
      SplitEmail = lngAdd
      
   'End If 'Removed by Morgan 2023/9/4

   If pUpdateRecNum = True Then
      cnnConnection.Execute "update MailSchedule set MS10=(select count(*) from MailScheduleDetail where msd01=ms01) where ms01=" & pMS01, intI
   End If
End Function

Private Function FormDelete() As Boolean
   Dim stSQL As String, bInTrans As Boolean
   
On Error GoTo ErrHandle
   
   cnnConnection.BeginTrans
   bInTrans = True
      
   stSQL = "delete MailScheduleImport where MSI01='" & txtNo & "'"
   cnnConnection.Execute stSQL, intI
   
   PUB_DelFtpFile2 txtNo, , UCase("MAILSCHEDULETEMPLET") 'Add By Sindy 2017/7/3 檔案改放 FTP,必須在DB資料刪除前執行
   stSQL = "delete mailscheduletemplet where mst01='" & txtNo & "'"
   cnnConnection.Execute stSQL, intI
   
   stSQL = "delete mailscheduledetail where msd01='" & txtNo & "'"
   cnnConnection.Execute stSQL, intI
   
   stSQL = "delete mailschedule where ms01='" & txtNo & "'"
   cnnConnection.Execute stSQL, intI
   
   cnnConnection.CommitTrans
   FormDelete = True
   
ErrHandle:
   If Err.Number <> 0 Then
      If bInTrans Then cnnConnection.RollbackTrans
      MsgBox Err.Description
   End If
   
End Function

Private Function FormSave() As Boolean
   Dim iFileNo As Integer
   Dim bytes() As Byte
   Dim lngSize As Long '檔案大小
   Dim stSQL As String
   Dim stScript As String
   Dim stFilePath As String
   Dim stFromMail As String
   Dim lngRec As Long
   Dim lngMS01 As Long
   Dim adoRst As New ADODB.Recordset
   Dim arrToMail
   Dim lngMulti As Long
   Dim ii As Integer

   Dim Numblocks As Integer
   Dim LeftOver As Long
   Dim i As Integer
   Dim bSampleOnly As Boolean '其他(只存樣本)
   Dim stMS15 As String
   Dim bInTrans As Boolean
   Dim strFtpPath As String 'Add By Sindy 2017/7/3
   'Add by Amy 2018/10/19
   Dim bolSetMSD07 As Boolean '設定MSD07欄
   Dim strSubject As String, strValue As String
   Dim strMSD06 As String 'Add By Sindy 2024/2/26
   Dim strMS27 As String 'Add by Amy 2025/09/02
 
   Const BlockSize = 500000
   
On Error GoTo ErrHandle
   
   'Add By Sindy 2019/10/2
   If InStr(txtSubject, "[Our Ref: XXXXXXXXX") > 0 Then txtSubject = Replace(txtSubject, "[Our Ref: XXXXXXXXX", "[Our Ref:XXXXXXXXX")
   If InStr(txtSubject, "[Our Ref : XXXXXXXXX") > 0 Then txtSubject = Replace(txtSubject, "[Our Ref : XXXXXXXXX", "[Our Ref:XXXXXXXXX")
   If InStr(txtSubject, "[Our Ref :XXXXXXXXX") > 0 Then txtSubject = Replace(txtSubject, "[Our Ref :XXXXXXXXX", "[Our Ref:XXXXXXXXX")
   '2019/10/2 END
   'Add by Amy 2018/10/19 特殊主旨寫入MSD07欄位中,發mail時主旨抓此欄
   strSubject = txtSubject
   bolSetMSD07 = False
   'Modify By Sindy 2019/3/13 [ISD ==> [Our Ref:
'   If InStr(UCase(strSubject), UCase("[Our Ref:XXXXXXXXX")) > 0 Or _
'      InStr(UCase(strSubject), UCase("[ISDXXXXXXXXX")) > 0 Then
   If InStr(UCase(strSubject), UCase("[Our Ref:XXXXXXXXX")) > 0 Then
      bolSetMSD07 = True
      'If InStr(UCase(strSubject), UCase("[ISDXXXXXXXXX")) > 0 Then strSubject = Mid(strSubject, 1, InStr(strSubject, "[ISDXXXXXXXXX") - 1)
      If InStr(UCase(strSubject), UCase("[Our Ref:XXXXXXXXX")) > 0 Then strSubject = Mid(strSubject, 1, InStr(strSubject, "[Our Ref:XXXXXXXXX") - 1)
      strValue = Replace(txtSubject, strSubject, "")
      'Add By Sindy 2019/10/2
      If InStr(UCase(txtSubject), UCase("] (")) = 0 Then
         MsgBox "主旨格式有誤, 找不到] (" & vbCrLf & _
                "若有聯絡人編號, 解析主旨時會有問題。"
         Exit Function
      End If
      '2019/10/2 END
   End If
   'end 2018/10/19
   
   GetChoice stMS15, bSampleOnly
   
   stFromMail = cboEmail
   stFilePath = txtSample
   iFileNo = FreeFile
   Open stFilePath For Binary Access Read As #iFileNo
   lngSize = LOF(iFileNo)
   'ReDim bytes(lngSize)
   'Get #iFileNo, , bytes()
   
   'Add by Amy 2025/09/02 +MS27 不解析Tag
   strMS27 = ""
   If chkNoneTag.Value = vbChecked Then
      strMS27 = "Y"
   End If
   'end 2025/09/02
   
   cnnConnection.BeginTrans
   bInTrans = True
   
   If txtNo = "" Then
      stSQL = "select nvl(max(ms01),0)+1 from mailschedule"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, stSQL)
      If intI = 1 Then
         lngMS01 = RsTemp.Fields(0)
      Else
         GoTo ErrHandle
      End If
   Else
      lngMS01 = txtNo
      stSQL = "delete MailScheduleDetail where msd01=" & lngMS01
      cnnConnection.Execute stSQL, lngRec
      
      PUB_DelFtpFile2 CStr(lngMS01), , UCase("MAILSCHEDULETEMPLET") 'Add By Sindy 2017/7/3 檔案改放 FTP,必須在DB資料刪除前執行
      stSQL = "delete MailScheduletemplet where mst01=" & lngMS01
      cnnConnection.Execute stSQL, lngRec
      
      'Added by Morgan 2012/2/8
      stSQL = "delete MailScheduleImport where MSI01=" & lngMS01
      cnnConnection.Execute stSQL, lngRec
   End If
   
   'Added by Morgan 2012/2/8
   For ii = 0 To lstImport(0).ListCount - 1
      stSQL = "insert into MailScheduleImport(MSI01,MSI02,MSI03) select " & lngMS01 & ",'" & ChgSQL(lstImport(0).List(ii)) & "','N' from dual" & _
         " where not exists(select * from MailScheduleImport where MSI01=" & lngMS01 & " and MSI02='" & ChgSQL(lstImport(0).List(ii)) & "')"
      cnnConnection.Execute stSQL, lngRec
   Next
   
   txtCount = 0
   If bSampleOnly = False Then
      'Added by Morgan 2012/3/12 +其他(指定編號)
      If stMS15 = 2 ^ 8 Then
         For ii = 0 To lstImport(1).ListCount - 1
            stSQL = "insert into MailScheduleImport(MSI01,MSI02,MSI03) select " & lngMS01 & ",'" & ChgSQL(lstImport(1).List(ii)) & "','Y' from dual" & _
               " where not exists(select * from MailScheduleImport where MSI01=" & lngMS01 & " and MSI02='" & ChgSQL(lstImport(1).List(ii)) & "')"
            cnnConnection.Execute stSQL, lngRec
         Next
         
         stScript = GetSQL2(lngMS01)
         'Modified by Morgan 2018/6/28 多編號改放最小號+"+"(原放 None)
         'Modified by Morgan 2019/8/1 信箱大小寫應視為相同
         'Modified by Morgan 2024/5/28 是否加'+',用count()或sum()判斷select是正確的但實際insert會不正確,改用max()和min()判斷才會一致
         stSQL = "insert into MailScheduleDetail(MSD01,MSD02,MSD06) SELECT " & lngMS01 & ",lower(MSD02) MSD02,MIN(MSD06)||DECODE(min(MSD06),max(MSD06),'','+') MSD06" & _
            " FROM (" & stScript & ") X,MailScheduleImport A,MailScheduleImport B where A.MSI01(+)=" & lngMS01 & " and A.MSI02(+)=msd06 and A.MSI03(+)='N'" & _
            " and B.MSI01(+)=" & lngMS01 & " and B.MSI02(+)=substrb(msd06,1,decode(sign(instr(msd06,'-')),0,length(msd06),instr(msd06,'-')-1)) and B.MSI03(+)='N'" & _
            " and A.MSI02||B.MSI02 is null" & _
            " GROUP BY lower(MSD02)"
         cnnConnection.Execute stSQL, lngRec
         txtCount = lngRec
         
      'Added by Morgan 2013/1/22 其他(指定信箱)
      ElseIf stMS15 = 2 ^ 7 Then
         For ii = 0 To lstImport(1).ListCount - 1
            'Add By Sindy 2024/2/26
            strMSD06 = "None"
            'If chkSpecList.Value = 1 Then 'Modify By Sindy 2024/12/17 發生代碼1276特殊編號未寫入 MailScheduleDetail 的 msd06
               '抓特殊編號
               stSQL = "SELECT TBNP09 FROM TMBulletinNp WHERE TBNP08='M' AND upper(TBNP01)=upper('" & Trim(ChgSQL(lstImport(1).List(ii))) & "')"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, stSQL)
               If intI = 1 Then
                  strMSD06 = RsTemp.Fields(0)
               End If
            'End If
            '2024/2/26 END
            'Modify By Sindy 2024/2/26 'None' 改為 '" & strMSD06 & "'
            stSQL = "insert into MailScheduleDetail(MSD01,MSD02,MSD06) select " & lngMS01 & ",'" & ChgSQL(lstImport(1).List(ii)) & "','" & strMSD06 & "'" & _
               " from dual where not exists(select * from MailScheduleDetail where msd01=" & lngMS01 & " and msd02='" & ChgSQL(lstImport(1).List(ii)) & "')"
            cnnConnection.Execute stSQL, lngRec
            If lngRec = 0 Then Debug.Print lstImport(1).List(ii)
            txtCount = Val(txtCount) + lngRec
         Next
      
      'Added by Morgan 2018/8/10
      'GDPR詢問信
      ElseIf stMS15 = 2 ^ 5 Then
         m_GDPR = True 'Added by Morgan 2018/9/14
'Modified by Morgan 2021/1/29 後續的詢問信改人工抓 XLS 匯入
'         stScript = GetSql
'         stSQL = "insert into MailScheduleDetail(MSD01,MSD02,MSD06) SELECT " & lngMS01 & ",MSD02,MSD06" & _
'            " FROM (" & stScript & ") X,MailScheduleImport A,MailScheduleImport B where A.MSI01(+)=" & lngMS01 & " and A.MSI02(+)=msd06 and A.MSI03(+)='N'" & _
'            " and B.MSI01(+)=" & lngMS01 & " and B.MSI02(+)=substrb(msd06,1,decode(sign(instr(msd06,'-')),0,length(msd06),instr(msd06,'-')-1)) and B.MSI03(+)='N'" & _
'            " and A.MSI02||B.MSI02 is null"
'         cnnConnection.Execute stSQL, lngRec
'         txtCount = lngRec
'
'         '重複信箱非最小編號發信日期(msd03)上19221111
'         'Modified by Morgan 2019/8/1 信箱大小應視為相同
'         stSQL = "update MailScheduleDetail a set msd03=19221111 where msd01=" & lngMS01 & " and msd03 is null" & _
'            " and exists(select * from MailScheduleDetail b where b.msd01=a.msd01 and lower(b.msd02)=lower(a.msd02) and b.msd06<a.msd06)"
'         cnnConnection.Execute stSQL, lngRec
'         txtCount = txtCount - lngRec
'
'         '重複信箱最小編號msd06附+號
'         'Modified by Morgan 2019/8/1 信箱大小應視為相同
'         stSQL = "update MailScheduleDetail a set msd06=msd06||'+' where msd01=" & lngMS01 & " and msd03 is null" & _
'            " and exists(select * from MailScheduleDetail b where b.msd01=a.msd01 and lower(b.msd02)=lower(a.msd02) and b.msd03>0)"
'         cnnConnection.Execute stSQL, lngRec
         For ii = 0 To lstImport(1).ListCount - 1
            stSQL = "insert into MailScheduleImport(MSI01,MSI02,MSI03) select " & lngMS01 & ",'" & ChgSQL(lstImport(1).List(ii)) & "','Y' from dual" & _
               " where not exists(select * from MailScheduleImport where MSI01=" & lngMS01 & " and MSI02='" & ChgSQL(lstImport(1).List(ii)) & "')"
            cnnConnection.Execute stSQL, lngRec
         Next
         
         stScript = GetSQL2(lngMS01)
         'Modified by Morgan 2024/5/28 是否加'+',用count()或sum()判斷select是正確的但實際insert會不正確,改用max()和min()判斷才會一致
         stSQL = "insert into MailScheduleDetail(MSD01,MSD02,MSD06) SELECT " & lngMS01 & ",lower(MSD02) MSD02,MIN(MSD06)||DECODE(min(MSD06),max(MSD06),'','+') MSD06" & _
            " FROM (" & stScript & ") X,MailScheduleImport A,MailScheduleImport B where A.MSI01(+)=" & lngMS01 & " and A.MSI02(+)=msd06 and A.MSI03(+)='N'" & _
            " and B.MSI01(+)=" & lngMS01 & " and B.MSI02(+)=substrb(msd06,1,decode(sign(instr(msd06,'-')),0,length(msd06),instr(msd06,'-')-1)) and B.MSI03(+)='N'" & _
            " and A.MSI02||B.MSI02 is null" & _
            " GROUP BY lower(MSD02)"
         cnnConnection.Execute stSQL, lngRec
         txtCount = lngRec
         
'end 2021/1/29
         
      Else
         stScript = GetSql
         '新增明細
         'Modified by Morgan 2012/2/9 +MailScheduleImport 例外清單(客戶/代理人例外則聯絡人也例外,聯絡人可單獨例外)
         'stSQL = "insert into MailScheduleDetail(MSD01,MSD02,MSD06) SELECT " & lngMS01 & ",MSD02,DECODE(SUM(1),1,MAX(MSD06),'None') MSD06 FROM (" & stScript & ") GROUP BY MSD02"
         'Modified by Morgan 2018/6/28 多編號改放最小號+"+"(原放 None)
         'Modified by Morgan 2019/8/1 信箱大小應視為相同 Ex:Y48162000(DOCKETING@novozymes.com),Y51883000(Docketing@novozymes.com)
         'Modified by Morgan 2024/5/28 是否加'+',用count()或sum()判斷select是正確的但實際insert會不正確,改用max()和min()判斷才會一致
         stSQL = "insert into MailScheduleDetail(MSD01,MSD02,MSD06) SELECT " & lngMS01 & ",lower(MSD02) MSD02,MIN(MSD06)||DECODE(min(MSD06),max(MSD06),'','+') MSD06" & _
            " FROM (" & stScript & ") X,MailScheduleImport A,MailScheduleImport B where A.MSI01(+)=" & lngMS01 & " and A.MSI02(+)=msd06 and A.MSI03(+)='N'" & _
            " and B.MSI01(+)=" & lngMS01 & " and B.MSI02(+)=substrb(msd06,1,decode(sign(instr(msd06,'-')),0,length(msd06),instr(msd06,'-')-1)) and B.MSI03(+)='N'" & _
            " and A.MSI02||B.MSI02 is null" & _
            " GROUP BY lower(MSD02)"
         cnnConnection.Execute stSQL, lngRec
         
         txtCount = lngRec
      End If
      
      'Add by Morgan 2009/3/26 多個信箱放一起的資料
      lngRec = SplitEmail(lngMS01)
      txtCount = Val(txtCount) + lngRec
      'Add by Amy 2018/10/19 特殊主旨寫入MSD07欄位中
      'Modify By Sindy 2019/9/19
      'If bolSetMSD07 = True And stMS15 = 2 ^ 8 Then 'stMS15 = 2 ^ 8 : 其他(指定編號) : 有編號才能置換 XXXXXXXXX
      If bolSetMSD07 = True Then
      '2019/9/19 END
        '寫一筆操作人員資料 for 組主旨
        stSQL = "insert into MailScheduleDetail(MSD01,MSD02,MSD06,MSD07) Values(" & lngMS01 & ",'" & strUserNum & "@taie.com.tw','M51','" & strValue & "')"
        cnnConnection.Execute stSQL, lngRec
        'Modify By Sindy 2019/3/13 [ISD ==> [Our Ref:
        'stSQL = "Update MailScheduleDetail Set msd07=Replace('" & strValue & "','[ISDXXXXXXXXX','[ISD'||Replace(msd06,'+','')) Where MSD01=" & lngMS01 & " And MSD06<>'M51' "
'        If InStr(UCase(strValue), UCase("[ISDXXXXXXXXX")) > 0 Then
'            stSQL = "Update MailScheduleDetail Set msd07=Replace('" & strValue & "','[ISDXXXXXXXXX','[ISD'||Replace(msd06,'+','')) Where MSD01=" & lngMS01 & " And MSD06<>'M51' "
'            cnnConnection.Execute stSQL, lngRec
'        Else
        If InStr(UCase(strValue), UCase("[Our Ref:XXXXXXXXX")) > 0 Then
            '解析msd06為9碼編號
            stSQL = "Update MailScheduleDetail Set msd07=Replace('" & strValue & "','[Our Ref:XXXXXXXXX','[Our Ref:'||REPLACE(substr(msd06,1,9),'+','')) Where MSD01=" & lngMS01 & " And MSD06<>'M51'"
            cnnConnection.Execute stSQL, lngRec
            'Add By Sindy 2019/10/1 要串聯絡人編號
            stSQL = "Update MailScheduleDetail Set msd07=Replace(msd07,'] (','.'||REPLACE(substr(MSD06,11,length(MSD06)),'+','')||'] (') Where MSD01=" & lngMS01 & " And MSD06<>'M51' and instr(MSD06,'-')>0"
            cnnConnection.Execute stSQL, lngRec
            '2019/10/1 END
        End If
        '2019/3/13 END
        'Add By Sindy 2019/9/19
        If stMS15 = 2 ^ 7 Then '其他(指定信箱)
            MsgBox "要檢查 mailscheduledetail.MSD01=排程編號 的 MSD06=XYR編號 ex:R15510000-01 是否有資料" & vbCrLf & _
                   "並且要自己組合出正確的 MSD07 ex:[Our Ref:R15510000.B51.01] (EY/wc)" & vbCrLf & _
                   "不然寄出去的主旨會有問題！"
        End If
        '2019/9/19 END
      End If
      'end 2018/10/19
   End If
   
   'Modified by Morgan 2018/5/11 +MS24
   'Modified by Morgan 2018/7/6 +MS25
   '新增
   'Modify by Amy 2018/10/19 txtSubject 改抓變數 strSubject
   If txtNo = "" Then
      'Modified by Morgan 2022/3/18 +ms26
      'Modify by Amy 2025/09/02 +ms27
      'Modified by Morgan 2025/9/4 MS26改預設Y但可存數字
      stSQL = "insert into mailschedule(ms01,ms02,ms03,ms04,ms05,ms06,ms07,ms08,ms09,ms10,ms14,ms15,ms22,ms23,ms24,ms25,ms26,ms27)" & _
         " values(" & lngMS01 & ",'" & ChgSQL(strSubject) & "','" & ChgSQL(stFromMail) & "','" & ChgSQL(stScript) & "'" & _
         ",'" & strUserNum & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
         "," & DBDATE(txtDate) & "," & Val(Replace(cboTime, ":", "") & "00") & "," & Val(txtCount) & ",'" & ChgSQL(cboDisplayName) & "'," & stMS15 & _
         ",'" & m_DepCode & "','" & IIf(chkMainOnly.Value = 1, "Y", "") & "','" & IIf(chkNoYahoo.Value = 1, "N", "") & "'" & _
         ",'" & IIf(chkByMailServer.Value = 1, "Y", "") & "','" & IIf(chkMS26.Value = 1, IIf(txtMS26 = "", "Y", txtMS26), "") & "'," & CNULL(ChgSQL(strMS27)) & ")"
      cnnConnection.Execute stSQL, lngRec
   '修改
   Else
      'Modified by Morgan 2022/3/18 +ms26
      'Modify by Amy 2025/09/02 +ms27
      'Modified by Morgan 2025/9/4 MS26改預設Y但可存數字
      stSQL = "update mailschedule set ms02='" & ChgSQL(strSubject) & "'" & _
         ",ms03='" & ChgSQL(stFromMail) & "',ms04='" & ChgSQL(stScript) & "'" & _
         ",ms08=" & DBDATE(txtDate) & ",ms09=" & Val(Replace(cboTime, ":", "") & "00") & _
         ",ms10=" & Val(txtCount) & _
         ",ms14='" & ChgSQL(cboDisplayName) & "',ms15=" & stMS15 & _
         ",ms19='" & strUserNum & "',ms20=to_char(sysdate,'yyyymmdd'),ms21=to_char(sysdate,'hh24miss')" & _
         ",ms23='" & IIf(chkMainOnly.Value = 1, "Y", "") & "'" & _
         ",ms24='" & IIf(chkNoYahoo.Value = 1, "N", "") & "'" & _
         ",ms25='" & IIf(chkByMailServer.Value = 1, "Y", "") & "'" & _
         ",ms26='" & IIf(chkMS26.Value = 1, IIf(txtMS26 = "", "Y", txtMS26), "") & "'" & _
         ",ms27=" & CNULL(ChgSQL(strMS27)) & _
         " where ms01=" & lngMS01
      cnnConnection.Execute stSQL, lngRec
   End If
   'end 2018/10/19
  
   'Modify By Sindy 2017/7/3 改上傳FTP File Server
'   stSQL = "select * from MailScheduleTemplet where rownum<1"
'   If adoRst.State <> adStateClosed Then adoRst.Close
'   With adoRst
'   .CursorLocation = adUseClient
'   .Open stSQL, cnnConnection, adOpenStatic, adLockOptimistic
'   .AddNew
'   .Fields("mst01").Value = lngMS01
'   .Fields("mst02").Value = lngSize
'   'Modify by Morgan 2010/2/6
'   '.Fields("mst03").AppendChunk bytes()
'   Numblocks = lngSize / BlockSize
'   LeftOver = lngSize Mod BlockSize
'   ReDim bytes(LeftOver)
'   Get #iFileNo, , bytes()
'   .Fields("mst03").AppendChunk bytes()
'   ReDim bytes(BlockSize)
'   For i = 1 To Numblocks
'       Get #iFileNo, , bytes()
'       .Fields("mst03").AppendChunk bytes()
'   Next i
'   'end 2010/2/6
   Close #iFileNo
'   .Fields("mst04").Value = IIf(ChkAtt.Value = 1, "Y", Null)
'   .Fields("mst05").Value = IIf(chkNoneBig5.Value = 1, "Y", Null) 'Added by Morgan 2012/8/14
'   .UPDATE
'   End With
   PUB_PutFtpFile stFilePath, CStr(lngMS01), "eml", strFtpPath, "MAILSCHEDULETEMPLET"
   If strFtpPath <> "" Then
      'CNULL(GetFileName(stFilePath)) ==> strFtpPath
      'Modified by Morgan 2018/8/3 +mst03
      strSql = "insert into MailScheduleTemplet(mst01,mst02,mst03,mst04,mst05,mst06) " & _
               "values(" & lngMS01 & "," & lngSize & ",'" & chkOutlook.Value & "'" & _
               ",'" & IIf(chkAtt.Value = 1, "Y", Null) & "','" & IIf(chkNoneBig5.Value = 1, "Y", Null) & "'" & _
               "," & CNULL(strFtpPath) & ")"
      cnnConnection.Execute strSql, lngRec
   End If
   '2017/7/3 END
   
   'Add By Sindy 2025/4/21
   stSQL = "select count(*) from MAILSCHEDULEDETAIL where msd01='" & lngMS01 & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, stSQL)
   If intI = 1 Then
      If RsTemp.Fields(0) = 0 Then
         If MsgBox("注意！此電子報排程(" & lngMS01 & ")尚無明細資料，確定要儲存嗎？", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
            GoTo ErrHandle
         End If
      End If
   End If
   '2025/4/21 END
   
   cnnConnection.CommitTrans
   
   'Add By Sindy 2024/12/17 發生代碼1276特殊編號未寫入 MailScheduleDetail 的 msd06
   '其他(指定信箱)
   If stMS15 = 2 ^ 7 Then
      If chkSpecList.Value = 1 Then
         stSQL = "select * from MAILSCHEDULEDETAIL where msd01='" & lngMS01 & "' and (msd06 is null or msd06='None')"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, stSQL)
         If intI = 1 Then
            MsgBox "匯入特殊名單，發生特殊編號未寫入 MailScheduleDetail 的 msd06 請檢查原因！"
         End If
      End If
   End If
   '2024/12/17 END
   
   txtNo = lngMS01
   FormSave = True
   
   Exit Function

ErrHandle:
   If bInTrans = True Then cnnConnection.RollbackTrans
   If Err.Number <> 0 Then
      MsgBox Err.Description
   End If
   
   If iFileNo > 0 Then Close #iFileNo
End Function

Private Sub cmdSend_Click()
   Dim bolIsTest As Boolean
   Dim stSQL As String
   Dim intR As Integer
   Dim rsQuery As ADODB.Recordset
   Dim bolOK As Boolean
   Dim intTestRecs As Integer '測試筆數
   
On Error GoTo ErrHnd:
   
   If chkOutlook.Value <> vbChecked Then MsgBox "勾選【Outlook 範本】才需在此寄送！": Exit Sub
   
   If MsgBox("勾選【Outlook 範本】寄信時，控制台【郵件】設定的【預設信箱】必須為【" & cboEmail & "】！" & vbCrLf & vbCrLf & "是否確定要繼續？", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
      Exit Sub
   End If
   
   List1.Clear
   
   Screen.MousePointer = vbHourglass
   If txtNo <> "" And txtSample = "" Then
      If GetTemplete(txtNo) = False Then
         MsgBox "樣本檔案讀取失敗！"
         Screen.MousePointer = vbDefault
         Exit Sub
      End If
   End If
   
   'Added by Morgan 2018/11/28
   '是否為測試
   intTestRecs = 10
   If MsgBox("是否為測試？" & vbCrLf & vbCrLf & "若為測試則將會寄送前 " & intTestRecs & " 封信到測試信箱！", vbYesNo + vbQuestion + vbDefaultButton1) = vbYes Then
      If txtToMail = "" Then
         MsgBox "請輸入測試信箱！"
         txtToMail.SetFocus
         Screen.MousePointer = vbDefault
         Exit Sub
      Else
         bolIsTest = True
      End If
   'Removed by Morgan 2018/12/4 Client_Win7無法讀取寄件人
   'ElseIf CheckAccount() = False Then
   '   Screen.MousePointer = vbDefault
   '   Exit Sub
   'end 2018/12/4
   End If
   'end 2018/11/28
   
   If GetChoiceNew() = 5 Then
      m_GDPR = True
      stSQL = "select msd02,msd06,nvl(msd07, '(' || msd06 ||')') Tag from MailScheduleDetail where msd01='" & txtNo & "' and msd03 is null order by 1"
   'Added by Morgan 2018/11/28
   '聖誕卡,賀年卡
   Else
      stSQL = "select msd02,msd06,msd07 Tag from MailScheduleDetail where msd01='" & txtNo & "' and msd03 is null order by 1"
   End If
   
   intR = 1
   Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      If MsgBox("共有 " & rsQuery.RecordCount & " 封通知函待寄送！是否確定要繼續？", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
         With rsQuery
         If bolIsTest = False Then
            stSQL = "update MailSchedule set  ms16=to_char(sysdate,'yyyymmdd'),ms17=to_char(sysdate,'hh24miss') where ms01=" & txtNo & " and ms16 is null"
            cnnConnection.Execute stSQL, intR
         End If
         Do While Not .EOF
            
            'Debug.Print .AbsolutePosition & " -> " & Now
            Sleep 3000 '等1秒
            
            '測試用
            If bolIsTest Then
               If .AbsolutePosition > intTestRecs Then Exit Do
               If SendOutlookMail(txtToMail, txtSubject & " ( " & .Fields("msd06") & " )", "" & .Fields("Tag"), True, m_GDPR) = False Then
                  Exit Do
               End If
            Else
               bolOK = SendOutlookMail(.Fields("msd02"), txtSubject & " ( " & .Fields("msd06") & " )", "" & .Fields("Tag"), , m_GDPR)
               stSQL = "update MailScheduleDetail set msd03=to_char(sysdate,'yyyymmdd'),msd04=to_char(sysdate,'hh24miss')"
               If bolOK Then
                  stSQL = stSQL & ",MSD05=null" '重寄成功要清除
               Else
                  stSQL = stSQL & ",MSD05='1'"
               End If
               stSQL = stSQL & " where msd01='" & txtNo & "' and msd02='" & ChgSQL(.Fields("msd02")) & "' and msd03 is null"
               cnnConnection.Execute stSQL, intR
                        
               If bolOK And m_GDPR Then
                  stSQL = "update fagent set FA123='W'" & _
                     " where (fa01||fa02) in (select replace(msd06,'+','') from MailScheduleDetail where msd01=" & txtNo & _
                     " and msd02='" & ChgSQL(.Fields("msd02")) & "') and FA123 is null"
                  cnnConnection.Execute stSQL, intR
                  
                  stSQL = "update potcustomer set PCU50='W'" & _
                     " where (PCU01||PCU02) in (select replace(msd06,'+','') from MailScheduleDetail where msd01=" & txtNo & _
                     " and msd02='" & ChgSQL(.Fields("msd02")) & "') and PCU50 is null"
                  cnnConnection.Execute stSQL, intR
                  
                  stSQL = "update potcustcont set PCC26='W'" & _
                     " where (PCC01||'0-'||PCC02) in (select replace(msd06,'+','') from MailScheduleDetail where msd01=" & txtNo & _
                     " and msd02='" & ChgSQL(.Fields("msd02")) & "') and PCC26 is null"
                  cnnConnection.Execute stSQL, intR
               End If
            End If
            .MoveNext
         Loop
         
         If bolIsTest = False Then
            stSQL = "update MailSchedule set ms11=to_char(sysdate,'yyyymmdd'),ms12=to_char(sysdate,'hh24miss')" & _
               ",ms13=(select count(*) from mailscheduledetail where msd01=ms01 and msd03>19221111)" & _
               ",ms18=(select count(*) from mailscheduledetail where msd01=ms01 and msd05 is not null) where ms01=" & txtNo
            cnnConnection.Execute stSQL, intR
         End If
         
         End With
         MsgBox "寄送完畢！"
      End If
   Else
      MsgBox "無待寄送資料！"
   End If
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   
   Set rsQuery = Nothing
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmdTest_Click()
   Dim stMime As String, stFromName As String
   
   Screen.MousePointer = vbHourglass
   
   If txtNo <> "" And txtSample = "" Then
      If GetTemplete(txtNo) = False Then
         MsgBox "樣本檔案讀取失敗！"
         Screen.MousePointer = vbDefault
         Exit Sub
      End If
   End If
   
   If TxtValidate(True) = True Then
      'Modified by Morgan 2018/8/1 +以Outlook範本寄信選項
      If chkOutlook.Value Then
         'Added by Morgan 2021/1/29
         If GetChoiceNew() = 5 Then
            m_bolTestOK = SendOutlookMail(txtToMail, txtSubject, , True, True)
         Else
         'end 2021/1/29
         
            m_bolTestOK = SendOutlookMail(txtToMail, txtSubject, lblTestName, True)
            
         End If 'Added by Morgan 2021/1/29
      Else
         m_bolTestOK = SendXMail()
      End If
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Function GetChoice(Optional stVaule As String, Optional pbolIs9 As Boolean) As Boolean
   Dim ii As Integer
   stVaule = 0
   For ii = 0 To lstMailChoice.ListCount - 1
      If lstMailChoice.Selected(ii) = True Then
         GetChoice = True
         If lstMailChoice.ITEMDATA(ii) = 9 Then pbolIs9 = True
         stVaule = stVaule + 2 ^ lstMailChoice.ITEMDATA(ii)
      End If
   Next
End Function

Private Function TxtValidate(Optional bTestMail As Boolean) As Boolean
   
   If bTestMail = False Then
      
      If GetChoice() = False Then
         MsgBox "請勾選寄發對象！"
         lstMailChoice.SetFocus
         Exit Function
      End If
   
      If txtSubject = "" Then
         MsgBox "主旨不可空白！"
         txtSubject.SetFocus
         Exit Function
         
      ElseIf ActionEdit = 0 And txtSubject = txtSubject.Tag And m_GDPR = False Then
         MsgBox "主旨錯誤!!尚未輸入期數!!"
         txtSubject.SetFocus
         Exit Function
         
      End If
   End If
   
   'Add by Amy 2018/10/19
   'Modify By Sindy 2019/3/13 [ISD ==> [Our Ref:
'   If ActionEdit = 0 And m_bolTestOK = True And _
'      (InStr(UCase(txtSubject), UCase("[Our Ref:")) > 0 Or InStr(UCase(txtSubject), UCase("[ISD")) > 0) Then
   If ActionEdit = 0 And m_bolTestOK = True And _
      InStr(UCase(txtSubject), UCase("[Our Ref:")) > 0 Then
        If MsgBox("主旨內有[Our Ref:字樣若為特殊主旨且需勾選「是否用Mail Server發信(backup會收到)」" & vbCrLf & _
            "要再修改為[Our Ref:XXXXXXXXX...]或勾選「是否用Mail Server發信(backup會收到)」？" & vbCrLf & vbCrLf & _
            "[Our Ref:後面需大寫9個X才能置換成指定編號", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbYes Then
            txtSubject.SetFocus
            Exit Function
        End If
   End If
   
   If cboEmail = "" Then
      MsgBox "發送信箱不可空白！"
      cboEmail.SetFocus
      Exit Function
   End If
   
   If cboDisplayName = "" Then
      MsgBox "顯示名稱不可空白！"
      cboDisplayName.SetFocus
      Exit Function
   End If
   
   If txtSample = "" Then
      MsgBox "尚未設定樣本檔案！"
      txtSample.SetFocus
      Exit Function
      
   ElseIf Not fso.FileExists(txtSample) Then
      MsgBox "樣本檔案路徑錯誤！"
      txtSample.SetFocus
      TextInverse txtSample
      Exit Function
   End If

   If bTestMail Then
      If txtToMail = "" Then
         MsgBox "測試信箱不可空白！"
         txtToMail.SetFocus
         Exit Function
      End If
   Else
   
      If txtDate = "" Then
         MsgBox "預定發信日期不可空白！"
         If txtDate.Enabled = True Then txtDate.SetFocus
         Exit Function
      ElseIf ChkDate(txtDate) = False Then
         If txtDate.Enabled = True Then txtDate.SetFocus
         Exit Function
      End If
      If cboTime.ListIndex = -1 Then
         MsgBox "請選擇發信時間！"
         If cboTime.Enabled = True Then cboTime.SetFocus
         Exit Function
      End If
      
      If m_bolTestOK = False Then
         If MsgBox("尚未寄過測試信是否確定要存檔??", vbYesNo + vbDefaultButton2 + vbExclamation) = vbNo Then
            Exit Function
         End If
      End If
      
      If Not m_AutoRun Then 'Added by Morgan 2024/6/11 自動新增除外
         'Added by Morgan 2020/9/18
         '存檔前確認寄件信箱是否設定無誤(退信處理問題).
         If MsgBox("寄件信箱是否正確？" & vbCrLf & vbCrLf & "※要確認退信處理OK!!", vbYesNo + vbExclamation + vbDefaultButton2) = vbNo Then
            Exit Function
         End If
         'end 2020/9/18
      End If
      
      
      'Add by Sindy 2021/7/7 提醒訊息 1.國內電子報 2.專利雙週報
      If (GetChoiceNew() = 1 Or GetChoiceNew() = 2) Then 'And ActionEdit = 0:新增
         MsgBox "寄發『台一專利商標雜誌』和『台一雙週專利電子報』時，" & vbCrLf & vbCrLf & _
                "增加(寄發)前研發(楊監察人)/智權提供之電郵，約4,000餘筆。" & vbCrLf & vbCrLf & _
                "統計時比照台一專利商標雜誌，(例如：共計 10538 + 4571)" & vbCrLf & vbCrLf & _
                "通知 發信人員 和 業拓(陳增廣主任 及 楊雯芳經理)。", vbInformation
      End If
      '2021/7/7 END
      
      If Not m_AutoRun Then 'Added by Morgan 2024/6/11 自動新增除外
         'Add By Sindy 2022/10/21 檢查目前是否已有待寄送的郵件
         strSql = "select ms01 from mailschedule where ms11 is null" & IIf(txtNo.Text <> "", " and ms01<>" & txtNo.Text, "")
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            If MsgBox("目前已有待寄送的郵件排程，" & vbCrLf & vbCrLf & "請考量是否有（優先寄發）的需求，是否繼續？", vbYesNo + vbDefaultButton2 + vbExclamation) = vbNo Then
               Exit Function
            End If
         End If
         '2022/10/21 END
      End If
   End If
   TxtValidate = True
End Function

Private Sub Command1_Click()
   If txtNo = "" Then
      MsgBox "請輸入排程代碼!!"
      txtNo.SetFocus
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   
On Error GoTo ErrHnd
   cnnConnection.BeginTrans
   
   SplitEmail Val(txtNo), True
   cnnConnection.CommitTrans
   ReadSchedule txtNo
   MsgBox "重整結束!!"
   Screen.MousePointer = vbDefault
   Exit Sub
   
ErrHnd:
   If Err.Number <> 0 Then
      cnnConnection.RollbackTrans
      MsgBox Err.Description, vbCritical
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
   'Add By Sindy 2017/12/11 電腦中心人員時,檢查資料庫連線是測試資料庫,則顯示詢問訊息
   If UCase(pub_DbTerminalName) <> 正式資料庫電腦名稱 Then
      If MsgBox("目前是測試資料庫，請確認是否繼續？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
         Unload Me
      Else
         Label13.Visible = True
      End If
   End If
   '2017/12/11 END
   
   'Added by Morgan 2024/6/11
   If m_AutoRun Then
      If AutoAddSchedule Then
         Unload Me
      End If
   End If
   'end 2024/6/11
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   'Add By Sindy 2017/12/11 檢查資料是否有存檔
   If ActionEdit = 0 Or ActionEdit = 1 Then
      If MsgBox(Me.Caption & "資料尚未存檔，請確認是否結束此作業？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
         Cancel = 1
      End If
   End If
   '2017/12/11 END
End Sub

Private Sub lstMailChoice_ItemCheck(Item As Integer)
   If ActionEdit > 1 Then Exit Sub
   Dim ii As Integer
   
   chkNoYahoo.Enabled = True 'Added by Morgan 2023/9/4
   
   If lstMailChoice.Selected(Item) = True Then
      '選其他時剩餘選項的勾要拿掉,反之其他的勾要拿掉
      lstMailChoice.Enabled = False
      'Modified by Morgan 2012/1/3 改都只能選一個
      'If lstMailChoice.ItemData(Item) = 9 Then
         For ii = 0 To lstMailChoice.ListCount - 1
            If ii <> Item Then
               lstMailChoice.Selected(ii) = False
            End If
         Next
      '
      'Else
      '   For ii = 0 To lstMailChoice.ListCount - 1
      '      If lstMailChoice.ItemData(ii) = 9 Then
      '         lstMailChoice.Selected(ii) = False
      '      End If
      '   Next
      'End If
      'end 2012/1/3
      
      'Added by Morgan 2012/3/12
      chkSpecList.Visible = False 'Add By Sindy 2023/8/22
      chkMainOnly.Visible = False 'Add By Sindy 2023/8/22
      If lstMailChoice.ITEMDATA(Item) = 8 Then
         chkMainOnly.Visible = True
         cmdImport(1).Enabled = True
         Label4.Caption = "指定X,Y編號："
      'Added by Morgan 2013/1/22
      ElseIf lstMailChoice.ITEMDATA(Item) = 7 Then
         chkMainOnly.Visible = False
         chkSpecList.Visible = True 'Add By Sindy 2023/8/22
         cmdImport(1).Enabled = True
         Label4.Caption = "指定Email信箱："
      'end 2013/1/22
         
         'Added by Morgan 2023/9/4
         chkNoYahoo.Value = vbUnchecked
         chkNoYahoo.Enabled = False
         'end 2023/9/4
         
      'Added by Morgan 2021/1/29
      ElseIf lstMailChoice.ITEMDATA(Item) = 5 Then
         chkMainOnly.Visible = True
         chkMainOnly.Value = vbChecked
         cmdImport(1).Enabled = True
         Label4.Caption = "指定R,Y編號："
      'end 2021/1/29
      
      Else
         lstImport(1).Clear
         cmdImport(1).Enabled = False
         chkMainOnly.Visible = False
      End If
      'end 2012/3/12
      
      lstMailChoice.Enabled = True
      CheckClick Item
      
   'Added by Morgan 2012/3/12
   Else
      lstImport(1).Clear
      cmdImport(1).Enabled = False
      chkMainOnly.Value = vbUnchecked
      chkMainOnly.Visible = False
      'Add By Sindy 2023/8/24
      chkSpecList.Value = vbUnchecked
      chkSpecList.Visible = False
      '2023/8/24 END
   End If
End Sub

Private Sub CheckClick(Index As Integer)
   chkNoYahoo.Value = vbUnchecked 'Added by Morgan 2018/5/11
   chkMS26.Value = 0 'Add By Sindy 2023/1/12 預設無
   FrameCU.Visible = False 'Add By Sindy 2025/5/27 預設無
   FrameTag.Visible = False 'Add by Amy 2025/09/02
   Select Case lstMailChoice.ITEMDATA(Index)
      Case 0 '國外電子報
         txtSubject = "Tai E Quarterly Issue No. ????"
         txtSubject.Tag = txtSubject
         cboEmail.Text = "newsletter@taie.com.tw"
         cboDisplayName.ListIndex = 0
         txtSubject.SetFocus
         FrameCU.Visible = True 'Add By Sindy 2025/5/27
         
      Case 1 '國內電子報
         txtSubject = "智慧財產權專業新知"
         txtSubject.Tag = ""
         'cboEmail.Text = "lawoffice@taie.com.tw"
         'cboEmail.Text = "news@taie.com.tw" 'Modify By Sindy 2020/4/1
         cboEmail.Text = "office@taie.com.tw" 'Modify By Sindy 2020/6/2
         cboDisplayName.ListIndex = 1
         chkNoYahoo.Value = vbChecked 'Added by Morgan 2018/5/11
         
      Case 2 '專利雙週報
         Frame2.Enabled = False
         'Modify By Sindy 2012/11/29
         'txtSubject = "台一雙週電子報 No.???"
         txtSubject = "台一雙週專利電子報 No.???"
         '2012/11/29 End
         txtSubject.Tag = txtSubject
         'cboEmail.Text = "lawoffice@taie.com.tw"
         'cboEmail.Text = "news@taie.com.tw" 'Modify By Sindy 2020/4/1
         cboEmail.Text = "office@taie.com.tw" 'Modify By Sindy 2020/6/2
         cboDisplayName.ListIndex = 1
         chkNoYahoo.Value = vbChecked 'Added by Morgan 2018/5/11
         'chkMS26.Value = 1 'Add By Sindy 2023/1/12 預設優先寄發(準時寄發)
         'Add by Amy 2025/09/02
         FrameTag.Visible = True
         If strUserNum = "A2004" Then chkTestData.Visible = True
         'end 2025/09/02
         
      Case 3 '顧問電子報
         Frame2.Enabled = False
         txtSubject = "第???期顧問通訊電子報"
         txtSubject.Tag = txtSubject
         'cboEmail.Text = "lawoffice@taie.com.tw"
         'cboEmail.Text = "news@taie.com.tw" 'Modify By Sindy 2020/4/1
         cboEmail.Text = "office@taie.com.tw" 'Modify By Sindy 2020/6/2
         cboDisplayName.ListIndex = 1
         'Modify By Sindy 2024/10/1 mark
         'chkNoYahoo.Value = vbChecked 'Added by Morgan 2018/5/11
      
      Case 5 'GDPR詢問信
         txtSubject = "Confirmation for European Union General Data Protection Regulation (EY/wc)"
         txtSubject.Tag = txtSubject
         cboEmail.Text = "qadept@taie.com.tw"
         cboDisplayName.ListIndex = 0
         chkOutlook.Value = vbChecked
         txtSubject.SetFocus
         
      Case 6 '國外電子報(日本籍) 'Modify By Sindy 2020/12/3
         txtSubject = "" 'WW/al:台一知財情報?載??知??
         txtSubject.Tag = txtSubject
         cboEmail.Text = "newsletter@taie.com.tw"
         cboDisplayName.ListIndex = 0
         txtSubject.SetFocus
      
      Case 10 '索取CF對帳單(中文) Added by Morgan 2024/5/28
         txtSubject = "請求提供對帳單"
         txtSubject.Tag = ""
         cboEmail.Text = "account@taie.com.tw"
         cboDisplayName.ListIndex = 1
         loadMailTemplete lstMailChoice.ITEMDATA(Index)
         
      Case 11 '索取CF對帳單(英文) Added by Morgan 2024/5/28
         txtSubject = "Request for Statement of Account"
         txtSubject.Tag = ""
         cboEmail.Text = "account@taie.com.tw"
         cboDisplayName.ListIndex = 0
         loadMailTemplete lstMailChoice.ITEMDATA(Index)
         
      Case 9
         txtSubject = ""
         cboEmail = ""
         cboDisplayName = ""
         
      Case Else
         
   End Select
   
End Sub

'Add By Sindy 2025/5/27
Private Sub textCU01_GotFocus()
   InverseTextBox textCU01
End Sub
'國籍
Private Sub textCU01_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   textCU01_2 = Empty
   If IsEmptyText(textCU01) = False Then
      textCU01_2 = GetNationName(Left(textCU01, 3), 0)
      If IsEmptyText(textCU01_2) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "代理人國籍不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCU01_GotFocus
      End If
   End If
End Sub
'2025/5/27 END

Private Sub Timer2_Timer()
   RefreshBar
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

'Added by Morgan 2018/7/6
Private Function GetSMTP() As String
   Dim stSQL As String, intR As Integer
   
   If chkByMailServer.Value = 1 Then
      stSQL = "select oMan from setSpecMan where ocode='SMTP_IP_MS'"
   Else
      stSQL = "select oMan from setSpecMan where ocode='SMTP_IP_FW'"
   End If
   intR = 1
   Set RsTemp = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      GetSMTP = RsTemp(0)
   End If
End Function

Private Function SendOutlookMail(pToMail As String, pSubject As String, Optional pTag As String, Optional pIsTest As Boolean, Optional pIsGDPR As Boolean = False) As Boolean
   Dim stTag As String, ii As Integer
   Dim objOutLook As Object
   Dim objMail As Object
   Dim strHTMLBody As String
   Dim objTmp As Object
         
   Set objOutLook = CreateObject("Outlook.Application")
   Set objMail = objOutLook.CreateItemFromTemplate(txtSample)
     
On Error GoTo ErrHnd
   
   objMail.Subject = pSubject
   If pTag <> "" Then
      stTag = pTag
      
   ElseIf pIsTest And Me.ActiveControl = Me.cmdTest Then
      ii = InStr(txtToMail, "@")
      If ii > 0 Then
         stTag = "(" & Left(txtToMail, ii - 1) & ")"
      Else
         stTag = "(" & txtToMail & ")"
      End If
      objMail.Subject = pSubject & " " & stTag 'Added by Morgan 2021/1/29
   End If
   
   strHTMLBody = objMail.HTMLBody
   
   'GDPR確認信
   If pIsGDPR = True Then
      stTag = Replace(stTag, " ", "%20")
      strHTMLBody = Replace(strHTMLBody, "subject=I%20CONSENT%20", "subject=I%20CONSENT%20" & stTag)
      strHTMLBody = Replace(strHTMLBody, "subject=I%20DO%20NOT%20CONSENT%20", "subject=I%20DO%20NOT%20CONSENT%20" & stTag)
   '聖誕卡,賀年卡
   ElseIf Me.txtNo = "528" Or Me.txtNo = "529" Then
      If pTag <> "" Then
         strHTMLBody = Replace(strHTMLBody, "智權人員　夏慧珠", "智權人員　" & pTag)
      Else
         strHTMLBody = Replace(strHTMLBody, "智權人員　夏慧珠", "")
      End If
      
   Else
      If pTag <> "" Then
         strHTMLBody = Replace(strHTMLBody, "智權人員　ＯＯＯ", "智權人員　" & pTag)
      Else
         strHTMLBody = Replace(strHTMLBody, "智權人員　ＯＯＯ", "")
      End If
   End If
   objMail.HTMLBody = strHTMLBody
   
   objMail.To = pToMail
   objMail.DeleteAfterSubmit = True
   'objMail.ReplyRecipients.add cboEmail
   If pIsTest Then
      'objMail.Display
      strExc(0) = ""
      'Removed by Morgan 2018/12/4 Client_Win7無法讀取寄件人
      'If Not objMail.SendUsingAccount Is Nothing Then
      '   strExc(0) = strExc(0) & "From: " & objMail.SendUsingAccount & vbCrLf
      'End If
      'end 2018/12/4
      strExc(0) = strExc(0) & "To: " & objMail.To & vbCrLf
      If objMail.cc <> "" Then
         strExc(0) = strExc(0) & "CC: " & objMail.cc & vbCrLf
      End If
      strExc(0) = strExc(0) & "Subject: " & objMail.Subject & vbCrLf
      If MsgBox(strExc(0) & vbCrLf & "是否確定要寄送？", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
         objMail.Send
         SendOutlookMail = True
         MsgBox "已寄出！", vbInformation
      Else
         objMail.Delete
         MsgBox "已取消！", vbInformation
      End If
   Else
      objMail.Send
      SendOutlookMail = True
   End If
   
ErrHnd:
   If Err.Number <> 0 Then
      If pIsTest Then MsgBox Err.Description, vbCritical
      objMail.Delete
   End If
   Set objMail = Nothing
   Set objOutLook = Nothing
End Function

Private Function SendXMail(Optional iErrCode As Integer) As Boolean

   Dim strFromName As String, strFromMail As String, strToName As String, StrToMail As String, strSubj As String, strMime As String
   Dim bolNoneBig5 As Boolean
   Dim SMTP As String
   Dim strData(0 To 9) As String
   Dim DateNow As String
   Dim iRetry As Integer
   Dim stBas64 As String
   'Add by Amy 2025/02/11
   Dim intCodeType As Integer, stUPMime As String, stOfficeTag As String, stEndMime As String

On Error GoTo ErrHnd
   
   bolNoneBig5 = IIf(chkNoneBig5.Value = 1, True, False)
   
   strFromName = cboDisplayName
   strFromMail = cboEmail
   strToName = txtToMail
   StrToMail = txtToMail
   strSubj = txtSubject
   PUB_ConvUni2UTF8Base64 strSubj 'Modified by Morgan 2024/8/16 2024/8/16 若主旨含UniCode則轉換成UTF-8Base64的編碼
   'Modify by Amy 2025/02/11 +if 主旨有[專利電子報]抓eml檔Tag,將客戶回覆mail之主旨,加上客戶編號,以利將客戶信箱設定不寄電子報
   'Modify by Amy 2025/09/02 +MS27 專利雙週報 No.382 解析Tag 有問題,無法跳過解析,故加「不解析Tag」
   If InStr(txtSubject, "專利電子報") > 0 And chkNoneTag.Value = vbUnchecked Then
      strMime = GetMime(txtSample, IIf(chkAtt.Value = 1, True, False), bolNoneBig5, True, True, intCodeType, stUPMime, stOfficeTag, stEndMime, txtNo, strUserNum)
      '為Big5編碼
      If intCodeType = 1 Then
         'Add by Amy 2025/07/24 +if InStr(stOfficeTag, "錯誤") > 0 專利雙週電子報 No.380 回不需電子報之內文不見,因body tag 被切成bod=(換行)y,不應再寄出信
         If InStr(stOfficeTag, "錯誤") > 0 Then
            '訊息於 GetMime已彈
            Exit Function
         ElseIf stUPMime <> "" And stOfficeTag <> "" And stEndMime <> "" Then
            strMime = stUPMime & stOfficeTag & stEndMime
            'Add by Amy 2025/09/02 測式與寄出Tag比對用-ex:首字為英文句點[.]寄出會被吃掉而導致內容缺.或格式異常
            If strUserNum = "A2004" Then Call SaveMime(stUPMime, stOfficeTag, stEndMime)
         Else
            MsgBox "主旨有[專利電子報]且為Big5編碼,但分割Tag有誤請確認！"
            Exit Function
         End If
      '為Utf8編碼,薛經理:彈訊息看是否能找出為UTF8的字,請User不要使用
      Else
         'Memo Utf8編碼,照原本的寄(因信另存eml為utf8 編碼<body  編碼後的code 可能不同無法觸析,故先不做)
         'ex:1131004 No.359 (PGJvZHk-<body 編碼) / 1131017 No.360 (DQo8Ym9keSB-<body 編碼)
         If MsgBox("此編碼為【UTF8】" & vbCrLf & _
                        "請確認信件內容是否能修改成【Big5】" & vbCrLf & _
                        "因客戶不寄電子報回信無法加入編號辨識" & vbCrLf & _
                        "[已]確認,繼續操作->請按「是」" & vbCrLf & _
                        "回前畫面,與再User確認->請按「否」", vbInformation + vbYesNo + vbDefaultButton1) = vbNo Then
            Exit Function
         End If
      End If
   Else
      strMime = GetMime(txtSample, IIf(chkAtt.Value = 1, True, False), bolNoneBig5, True)
   End If
   
   'Modified by Morgan 2018/7/6 SMTP 改抓特殊設定
   'X400
   If InStr(StrToMail, "@") = 0 Then
      'Modified by Morgan 2013/1/21
      'SMTP = "192.168.1.2" 'exchange
      'strToMail = MailBefore & strToMail & MailAfter
      'SMTP = "192.168.1.10" '台一 spam firewall
      StrToMail = StrToMail & "@taie.com.tw"
   'SMTP
   Else
      'hinet會限制廣告信數量,改用自己的SMTP
      'SMTP = "192.168.1.10" '台一 spam firewall
   End If
   SMTP = GetSMTP()
   'end 2018/7/6
   
   iErrCode = 0
   Result = ""
   DoEvents
   
   strData(1) = "mail from: " & strFromMail & vbCrLf
   strData(2) = "rcpt to: " & StrToMail & vbCrLf

   'Added by Morgan 2013/9/4
   '不要指定編碼否則郵件內容若編碼不同不會自動選擇
   If bolNoneBig5 Then
      strData(3) = "From: """ & strFromName & """ <" & strFromMail & ">" & vbCrLf
      strData(4) = "To: """ & strToName & """ <" & StrToMail & ">" & vbCrLf
      strData(5) = "Subject: " & strSubj & vbCrLf
   Else
   'end 2013/9/4
   
      stBas64 = ConvertToBase64(strFromName, False, False)
      strData(3) = "From: =?Big5?B?" & stBas64 & "?= <" & strFromMail & ">" & vbCrLf
      stBas64 = ConvertToBase64(strToName, False, False)
      strData(4) = "To: =?Big5?B?" & stBas64 & "?= <" & StrToMail & ">" & vbCrLf
      'Added by Morgan 2012/8/10
      If Left(Trim(strSubj), 2) = "=?" Then
         strData(5) = "Subject: " & strSubj & vbCrLf
      Else
      'end 2012/8/10
         stBas64 = ConvertToBase64(strSubj, False, False)
         strData(5) = "Subject: =?Big5?B?" & stBas64 & "?=" & vbCrLf
      End If 'Added by Morgan 2012/8/10
      
   End If 'Added by Morgan 2013/9/4

   DateNow = Format(Date, "Ddd") & ", " & Format(Date, "dd Mmm YYYY") & " " & Format(time, "hh:mm:ss") & "" & " +0800"
   strData(6) = "Date:" + Chr(32) + DateNow + vbCrLf
   strData(0) = strData(3) + strData(4) + strData(5) + strData(6)
   'Added by Morgan 2022/11/22 Gmail會檢查Message-ID
   strData(0) = strData(0) & "Message-ID: <" & GetGUID() & "@Exchange>" & vbCrLf
   'end 2022/11/22
   strData(9) = strMime
   If strData(9) = "" Then
      strData(7) = "MIME-Version: 1.0" & vbCrLf & _
                   "Content-Type: text/plain;" + vbCrLf & _
                   "   charset=""big5""" + vbCrLf
      strData(8) = "testing..." + vbCrLf
      strData(9) = strData(7) & strData(8)
   End If

   strData(0) = strData(0) + strData(9)

RetryPoint:

   If Winsock1.State <> sckClosed Then Winsock1.Close

   Winsock1.LocalPort = 0
   Winsock1.Protocol = sckTCPProtocol
   Winsock1.RemoteHost = SMTP
   Winsock1.RemotePort = 25
   DoEvents

   List1.AddItem Now & " -> SMTP:" & SMTP & "," & strData(4), 0

   Winsock1.Connect
   If Not Response("220") Then
      Winsock1.Close
      iErrCode = 1
      GoTo ERRORMail
   End If

   DoEvents
   Winsock1.SendData ("HELO " & Winsock1.LocalHostName & ".taie.com.tw" & vbCrLf)
   If Not Response("250") Then
      iErrCode = 2
      GoTo ERRORMail
   End If

   DoEvents
   Winsock1.SendData (strData(1))
   If Not Response("250") Then
      iErrCode = 3
      GoTo ERRORMail
   End If

   DoEvents
   Winsock1.SendData (strData(2))
   If Not Response("250") Then
      iErrCode = 4
      GoTo ERRORMail
   End If

   DoEvents
   Winsock1.SendData ("data" + vbCrLf)
   If Not Response("354") Then
      iErrCode = 5
      GoTo ERRORMail
   End If

   DoEvents
   Winsock1.SendData (strData(0) & vbCrLf & "." & vbCrLf)
   If Not Response("250") Then
      iErrCode = 6
      GoTo ERRORMail
   End If

   DoEvents
   Winsock1.SendData ("quit" + vbCrLf)
   If Not Response("221") Then
      iErrCode = 7
      GoTo ERRORMail
   End If
   Winsock1.Close
   SendXMail = True
   
   If Not m_AutoRun Then 'Added by Morgan 2024/6/11 自動新增除外
      MsgBox "已寄出！", vbInformation
   End If
   Exit Function

ERRORMail:
   iRetry = iRetry + 1
   If iRetry < 3 Then
      GoTo RetryPoint
   End If

ErrHnd:

End Function

Private Sub txtMS26_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And (Chr(KeyAscii) < "1" Or Chr(KeyAscii) > "9") Then
      Beep
      KeyAscii = 0
   End If
End Sub

Private Sub txtSample_Change()
   m_bolTestOK = False
End Sub

Private Sub txtSubject_GotFocus()
   Dim iPos As Integer
   iPos = InStr(txtSubject, "?")
   If iPos > 0 Then
      txtSubject.SelStart = iPos - 1
      txtSubject.SelLength = Len(Mid(txtSubject, iPos))
   Else
      TextInverse txtSubject
   End If
End Sub

Private Sub txtSubject_Validate(Cancel As Boolean)
   If InStr(txtSubject, "Tai E Quarterly") > 0 And InStr(txtSubject, "_") > 0 Then
      If MsgBox("主旨含有底線(_)，是否確定要繼續？" & vbCrLf & vbCrLf & "注意：若複製樣本檔名可能會與信件主旨不同!!!", vbYesNo + vbDefaultButton2 + vbExclamation) = vbNo Then
         Cancel = True
      End If
   End If
End Sub

Private Sub txtToMail_Change()
   lblTestName = ""
End Sub

Private Sub txtToMail_KeyPress(KeyAscii As Integer)
   'KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtToMail_Validate(Cancel As Boolean)
   Dim strName As String
   If InStr(txtToMail, "@") = 0 Then
      If ClsPDGetStaffN(txtToMail, strName) = True Then
         lblTestName = strName
      End If
   End If
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Winsock1.GetData Result, vbString
    List1.AddItem Now & " -> " & Result, 0
End Sub
'
Private Function Response(RCode$, Optional IsShow As Boolean = True) As Boolean

   Const TimeOut% = 20
   Sec = 0
   Timer1.Interval = 500
   Timer1.Enabled = True
   Response = True

   Do While Left$(Result, 3) <> RCode
      '收件者被拒504,Unsupport Option 555
      If Left(Result, 3) = "504" Or Left(Result, 3) = "555" Then
         Response = False
         Exit Do
      End If
      DoEvents
      If Sec > TimeOut * 2 Then
         Response = False
         Exit Do
      End If
   Loop
   Result = ""
   Timer1.Enabled = False
End Function

Private Sub Timer1_Timer()
  Sec = Sec + 1
  DoEvents
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   
   SetTimeList
   
   cboDisplayName.Clear
   cboDisplayName.AddItem "Tai E International Patent & Law Office"
   'Modified by Morgan 2020/3/30 事務所名稱改用函數抓
   'cboDisplayName.AddItem "台一國際專利法律事務所"
   cboDisplayName.AddItem CompNameQuery("2")
   'end 2020/3/30
   
   cboEmail.Clear
   
   SSTab1.TabVisible(2) = False
   SSTab1.TabVisible(3) = False 'Add By Sindy 2021/11/30
   Command1.Visible = False
   
   'Added by Morgan 2024/6/6 財務系統
   If InStr(UCase(App.EXEName), "ACCOUNT") > 0 Then
      cboEmail.AddItem "account@taie.com.tw"
      List1.Visible = False
      m_DepCode = "M31"
   'end 2024/6/6
   ElseIf Pub_StrUserSt03 = "M51" Then
      cboEmail.AddItem "ipdept@taie.com.tw"
      cboEmail.AddItem "newsletter@taie.com.tw"
      'cboEmail.AddItem "lawoffice@taie.com.tw"
      'cboEmail.AddItem "news@taie.com.tw" 'Modify By Sindy 2020/4/1
      cboEmail.AddItem "office@taie.com.tw" 'Modify By Sindy 2020/6/2
      cboEmail.AddItem "patent@taie.com.tw" 'Modify By Sindy 2024/5/8
      SSTab1.TabVisible(2) = True
      SSTab1.TabVisible(3) = True 'Add By Sindy 2021/11/30
      Command1.Visible = True
   '專利處
   ElseIf Left(Pub_StrUserSt03, 2) = "P1" Then
      'cboEmail.AddItem "lawoffice@taie.com.tw"
      'cboEmail.AddItem "news@taie.com.tw" 'Modify By Sindy 2020/4/1
      cboEmail.AddItem "office@taie.com.tw" 'Modify By Sindy 2020/6/2
      cboEmail.AddItem "patent@taie.com.tw"
      List1.Visible = False
      m_DepCode = "P1"
   '研發
   ElseIf Left(Pub_StrUserSt03, 1) = "D" Then
      'cboEmail.AddItem "lawoffice@taie.com.tw"
      'cboEmail.AddItem "news@taie.com.tw" 'Modify By Sindy 2020/4/1
      cboEmail.AddItem "office@taie.com.tw" 'Modify By Sindy 2020/6/2
      List1.Visible = False
      m_DepCode = "D"
   '國外開拓
   ElseIf Left(Pub_StrUserSt03, 2) = "F4" Then
      'cboEmail.AddItem "newsletter@taie.com.tw"
      cboEmail.AddItem "office@taie.com.tw" 'Modify By Sindy 2020/6/2
      List1.Visible = False
      m_DepCode = "F4"
   End If
   
   m_bInsert = IsUserHasRightOfFunction("frm140410_1", strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction("frm140410_1", strEdit, False)
   m_bDelete = IsUserHasRightOfFunction("frm140410_1", strDel, False)
   m_bQuery = IsUserHasRightOfFunction("frm140410_1", strFind, False)
   
   SSTab1.Tab = 0
      
   '專利雙週報
   FormReset
   
   ActionEdit = 3
   Action 9 '預設最後一筆
End Sub

Private Sub SetTimeList()
   Dim ii As Integer
   cboTime.Clear
   For ii = 0 To 47
      cboTime.AddItem Format(ii \ 2, "00") & ":" & Format((ii Mod 2) * 30, "00")
   Next
End Sub

Private Sub SetList(Optional ByVal pSelectValue As Integer, Optional pbolAll As Boolean)
   'Modified by Morgan 2024/5/28 增加第10,11選項(每月批次自動新增)
   Dim ii As Integer, jj As Integer
   Dim arrAllMailType(11) As String '所有種類
   Dim arrCanMailType(11) As String '可選種類
   
   lstMailChoice.Visible = False
   
   '2 ^ 1 ~ 2 ^ 9:
   arrAllMailType(0) = "國外電子報"
   arrAllMailType(1) = "國內電子報"
   arrAllMailType(2) = "專利雙週報"
   arrAllMailType(3) = "顧問電子報"
   arrAllMailType(4) = "國外部價目表"
   arrAllMailType(5) = "GDPR詢問信"
   arrAllMailType(6) = "國外電子報(僅日本籍)" 'Add By Sindy 2020/12/3
   arrAllMailType(7) = "其他(指定信箱)"
   arrAllMailType(8) = "其他(指定編號)"
   arrAllMailType(9) = "其他(只存樣本)"
   arrAllMailType(10) = "索取CF對帳單(中文)" 'Added by Morgan 2024/5/28
   arrAllMailType(11) = "索取CF對帳單(英文)" 'Added by Morgan 2024/5/28
   
   '有勾選的放前面
   If pSelectValue > 0 Then
      lstMailChoice.Clear
      
      For ii = UBound(arrAllMailType) To 0 Step -1
         If pSelectValue >= 2 ^ ii Then
            lstMailChoice.AddItem arrAllMailType(ii), 0
            lstMailChoice.ITEMDATA(0) = ii
            lstMailChoice.Selected(0) = True
            pSelectValue = pSelectValue Mod 2 ^ ii
         End If
         If pSelectValue = 0 Then Exit For
      Next
   End If
      
   If pbolAll Then
      
      '財務系統
      If m_DepCode = "M31" Then
         arrCanMailType(10) = arrAllMailType(10)
         arrCanMailType(11) = arrAllMailType(11)
      '專利處
      ElseIf m_DepCode = "P1" Then
         arrCanMailType(2) = "專利雙週報"
         
       '研發
      ElseIf m_DepCode = "D" Then
         arrCanMailType(1) = "國內電子報"
         arrCanMailType(3) = "顧問電子報"
         
      '國外開拓
      ElseIf m_DepCode = "F4" Then
         arrAllMailType(0) = "國外電子報"
         arrAllMailType(6) = "國外電子報(僅日本籍)" 'Add By Sindy 2020/12/3
         
      ElseIf Pub_StrUserSt03 = "M51" Then
         For ii = 0 To UBound(arrAllMailType)
            arrCanMailType(ii) = arrAllMailType(ii)
         Next
      End If
   
      For ii = 0 To UBound(arrCanMailType)
         If arrCanMailType(ii) <> "" Then
            For jj = 0 To lstMailChoice.ListCount - 1
               If lstMailChoice.ITEMDATA(jj) = ii Then
                  Exit For
               End If
            Next
            '沒有勾的選項放後面
            If jj = lstMailChoice.ListCount Then
               intI = lstMailChoice.ListCount
               lstMailChoice.AddItem arrCanMailType(ii), intI
               lstMailChoice.ITEMDATA(intI) = ii
            End If
         End If
      Next
   End If
   
   lstMailChoice.Visible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm140410_1 = Nothing
End Sub

Private Sub FormReset()
   Dim oCheck As CheckBox
   txtNo = ""
   
   lstImport(0).Clear
   lstImport(1).Clear
   lblImpCount(0).Caption = ""
   lblImpCount(1).Caption = ""
   
   cmdImport(1).Enabled = False
   lstMailChoice.Clear
   chkAtt.Value = 1 '0 Modify By Sindy 2024/4/18
   txtSubject = ""
   chkNoneBig5.Value = 0 'Added by Morgan 2013/8/22
   chkNoYahoo.Value = 0 'Added by Morgan 2018/7/6
   chkByMailServer.Value = 0 'Added by Morgan 2018/7/6
   chkOutlook.Value = 0 'Added by Morgan 2020/1/6
   chkMS26.Value = 0 'Added by Morgan 2022/3/18
   
   cboEmail = ""
   cboDisplayName = ""
   
   txtSample = ""
   txtToMail = strUserNum
   txtDate = ""
   cboTime.ListIndex = -1
   txtCount = ""
   
   lblActFrom = ""
   lblActTo = ""
   lblActCount = ""
   lblFailCount = ""
   
   lblCreate.Caption = "Create : "
   lblUpdate.Caption = "Update : "
   
   Command1.Enabled = False
   cmdEstimate.Visible = False
   lblState.Visible = False
   cmdDetect.Visible = False
   Frame2.Visible = False
   Timer2.Enabled = False
   Timer2.Interval = 0
   m_GDPR = False 'Added by Morgan 2018/9/14
End Sub

Private Sub UpdateGridList(ByRef rsTmp As ADODB.Recordset)
Dim iRow As Integer, iCol As Integer
   With rsTmp
   .MoveFirst
   GrdList.Visible = False
   Do While Not .EOF
      GrdList.Rows = GrdList.Rows + 1
      iRow = GrdList.Rows - 1
      For iCol = 0 To .Fields.Count - 1
         GrdList.TextMatrix(iRow, iCol) = "" & .Fields(iCol)
      Next
      .MoveNext
   Loop
   GrdList.FixedRows = 1 'Added by Lydia 2023/10/16
   GrdList.Visible = True
   End With
End Sub

Private Sub grdList_Click()
    grdList_ShowSelection
End Sub


' 將GridList所選取的列反白, 並將未選取的列設成一般顏色
Private Sub grdList_ShowSelection()
Dim nCurrSel As Integer
Dim nCol As Integer
   
    nCurrSel = GrdList.row
    ' 與前一選擇的列位置相同則不處理
    If m_CurrSel = GrdList.row Then
        GoTo EXITSUB
    End If
    ' 將原先選取的列回復到正常的顏色
    If m_CurrSel > 0 And m_CurrSel < GrdList.Rows Then
        GrdList.row = m_CurrSel
        GrdList.col = 1
        If GrdList.CellBackColor <> &H80000005 Then
            For nCol = 1 To GrdList.Cols - 1
                GrdList.col = nCol
                If GrdList.CellBackColor <> &H80000005 Then: GrdList.CellBackColor = &H80000005
                If GrdList.CellForeColor <> &H80000008 Then: GrdList.CellForeColor = &H80000008
            Next nCol
        End If
        GrdList.col = 0
    End If
    ' 設定成所選取的列
    m_CurrSel = nCurrSel
    ' 將所選取的列反白
    If m_CurrSel > 0 And m_CurrSel < GrdList.Rows Then
        GrdList.row = m_CurrSel
        GrdList.col = 1
        For nCol = 1 To GrdList.Cols - 1
            GrdList.col = nCol
            GrdList.CellBackColor = &H8000000D
            GrdList.CellForeColor = &H80000005
        Next nCol
        GrdList.col = 0
    End If
EXITSUB:
End Sub

Private Sub grdList_DblClick()
   SSTab1.Tab = 0
End Sub

Private Sub grdList_SelChange()
   Dim nRow As Integer
    grdList_ShowSelection

    If GrdList.row > 0 And GrdList.row <= GrdList.Rows - 1 Then
        nRow = GrdList.row
        ReadSchedule GrdList.TextMatrix(nRow, 1)
    End If
End Sub

Private Sub ReceiverTypeCheck(ByVal pValue As Single, ByRef arrValue() As Integer)
   Dim ii As Integer
   
   If pValue = 0 Then
      arrValue(0) = 1
   Else
      For ii = 3 To 0 Step -1
         If pValue >= 2 ^ ii Then arrValue(2 ^ ii) = 1: pValue = pValue Mod 2 ^ ii
      Next
   End If
End Sub

Private Function GetTemplete(p_MST01 As String) As Boolean
   
   Dim stAttPath As String
   Dim lngSize As Long
   Dim iFileNo As Integer
   Dim bytes() As Byte
   
   stAttPath = App.path & "\mailschedule.tmp"
   strExc(0) = "select * from mailscheduleTemplet b where mst01=" & p_MST01
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If Dir(stAttPath) <> "" Then Kill stAttPath
      'Add By Sindy 2017/7/3
      If "" & RsTemp.Fields("mst06") <> "" Then
         GetTemplete = PUB_GetFtpFile(RsTemp.Fields("mst06"), stAttPath, UCase("MAILSCHEDULETEMPLET"))
      Else
      '2017/7/3 END
         With RsTemp
         lngSize = Val(.Fields("mst02").Value)
         ReDim bytes(lngSize)
         bytes() = .Fields("mst03").GetChunk(lngSize)
         End With
         iFileNo = FreeFile
         If fso.FileExists(stAttPath) Then
            Kill stAttPath
         End If
         Open stAttPath For Binary Access Write As #iFileNo
         Put #iFileNo, , bytes()
         Close #iFileNo
      End If
      
      txtSample = stAttPath
      GetTemplete = True
   End If
   
End Function

Private Function ReadSchedule(p_MS01 As String) As Boolean
   
   Dim iMS15 As Single

   FormReset
   'Modify By Sindy 2019/3/13 [ISD ==> [Our Ref:
   'Modify by Amy 2018/10/19 +MSD07(主旨內容有[ISD->預帶第一筆電腦中心人員設定)
   strExc(0) = "select a.*,b.*,s1.st02||' '||sqldatet(ms06)||' '||sqltime6(ms07) cCreate" & _
      ",s2.st02||' '||sqldatet(ms20)||' '||sqltime6(ms21) cUpdate" & _
      ",sqldatet(ms16)||' '||sqltime6(ms17) cActFrom" & _
      ",sqldatet(ms11)||' '||sqltime6(ms12) cActTo,MSD07" & _
      " from mailschedule a,mailscheduleTemplet b,staff s1,staff s2,(Select msd01,msd07 From MailScheduleDetail Where msd01=" & p_MS01 & "  And msd06='M51' )" & _
      " where ms01=" & p_MS01 & " and mst01(+)=ms01 and s1.st01(+)=ms05 and s2.st01(+)=ms19 And ms01=msd01(+) "
   'end 2018/10/19
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
         txtNo = "" & .Fields("ms01")
         iMS15 = Val("" & .Fields("ms15"))
         
         SetList iMS15
               
         txtSubject = "" & .Fields("ms02")
         'Add by Amy 2018/10/19 特殊主旨-記錄於msd06='M51' 那筆Record
         If Not IsNull(.Fields("MSD07")) Then
            txtSubject = txtSubject & .Fields("MSD07")
         End If
         cboEmail = "" & .Fields("ms03")
         cboDisplayName = "" & .Fields("ms14")
         
         txtDate = TransDate("" & .Fields("ms08"), 1)
         
         cboTime.ListIndex = 2 * (Val("" & .Fields("ms09")) \ 10000) + IIf((Val("" & .Fields("ms09")) Mod 10000) = 0, 0, 1)
         txtCount = "" & .Fields("ms10")
         
         lblActFrom = Trim("" & .Fields("cActFrom"))
         'Added by Morgan 2018/11/29
         '已寄發不可再重整
         FrameCU.Visible = False 'Add By Sindy 2025/9/25
         If lblActFrom <> "" Then
            Command1.Enabled = False
         Else
            Command1.Enabled = True
            'Add By Sindy 2025/9/25
            If iMS15 = 1 Then
               FrameCU.Visible = True
            End If
            '2025/9/25 END
         End If
         'end 2018/11/29
         lblActCount = "" & .Fields("ms13")
         
         lblActTo = Trim("" & .Fields("cActTo"))
         lblFailCount = "" & .Fields("ms18")
         
         'Added by Morgan 2012/3/12
         If .Fields("ms23").Value = "Y" Then
            chkMainOnly.Value = 1
         Else
            chkMainOnly.Value = 0
         End If
         'Add By Sindy 2023/8/24
         If iMS15 = 2 ^ 7 Then
            chkSpecList.Visible = True
            '2023/8/24 END
         ElseIf iMS15 = 2 ^ 8 Then
            chkMainOnly.Visible = True
            cmdImport(1).Enabled = True
         Else
            chkMainOnly.Visible = False
            chkSpecList.Visible = False 'Add By Sindy 2023/8/24
         End If
         'end 2012/3/12
         
         If .Fields("mst04").Value = "Y" Then
            chkAtt.Value = 1
         Else
            chkAtt.Value = 0
         End If
         
         'Added by Morgan 2018/5/11
         If .Fields("ms24").Value = "N" Then
            chkNoYahoo.Value = 1
         Else
            chkNoYahoo.Value = 0
         End If
         'end 2018/5/11
         
         'Added by Morgan 2018/7/6
         If .Fields("ms25").Value = "Y" Then
            chkByMailServer.Value = 1
         Else
            chkByMailServer.Value = 0
         End If
         'end 2018/7/6
         
         'Added by Morgan 2021/3/18
         If Not IsNull(.Fields("ms26").Value) Then
            chkMS26.Value = 1
            'Added by Morgan 2025/9/4
            If .Fields("ms26") <> "Y" Then
               txtMS26 = .Fields("ms26")
            End If
            'end 2025/9/4
         Else
            chkMS26.Value = 0
         End If
         'end 2021/3/18
         
         'Add by Amy 2025/09/02 +不解析Tag
         FrameTag.Visible = False
         chkNoneTag.Value = 0
         chkTestData.Value = 0
         If InStr(txtSubject, "專利電子報") > 0 Then
            FrameTag.Visible = True
            If "" & .Fields("ms27").Value = "Y" Then
               chkNoneTag.Value = 1
            End If
            If strUserNum = "A2004" Then chkTestData.Visible = True
         End If
         'end 2025/09/02
         
         'Added by Morgan 2012/8/14
         If .Fields("mst05").Value = "Y" Then
            chkNoneBig5.Value = 1
         Else
            chkNoneBig5.Value = 0
         End If
         
         'Added by Morgan 2018/8/3
         cmdSend.Enabled = False
         If .Fields("mst03").Value = "1" Then
            chkOutlook.Value = 1
            cmdSend.Enabled = True
         Else
            chkOutlook.Value = 0
         End If
         
         lblCreate.Caption = lblCreate.Caption & .Fields("cCreate")
         lblUpdate.Caption = lblUpdate.Caption & .Fields("cUpdate")
         
         If lblActFrom <> "" And lblActTo = "" Then
            lblState.Visible = True
            cmdDetect.Visible = True
         End If
         
      End With
      
      strExc(0) = "select MSI02 from MailScheduleImport" & _
         " where MSI01=" & p_MS01 & " AND MSI03='N' order by 1"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         With RsTemp
         Do While Not .EOF
            lstImport(0).AddItem .Fields(0)
            .MoveNext
         Loop
         lblImpCount(0).Caption = lstImport(0).ListCount
         End With
      End If
      
      If iMS15 = 2 ^ 8 Then
         strExc(0) = "select MSI02 from MailScheduleImport" & _
            " where MSI01=" & p_MS01 & " AND MSI03='Y' order by 1"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            With RsTemp
            Do While Not .EOF
               lstImport(1).AddItem .Fields(0)
               .MoveNext
            Loop
            lblImpCount(1).Caption = lstImport(1).ListCount
            End With
         End If
         
      'Added by Morgan 2013/1/22
      ElseIf iMS15 = 2 ^ 7 Then
         strExc(0) = "select MSd02 from MailScheduledetail" & _
            " where MSd01=" & p_MS01
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            With RsTemp
            Do While Not .EOF
               lstImport(1).AddItem .Fields(0)
               .MoveNext
            Loop
            lblImpCount(1).Caption = lstImport(1).ListCount
            End With
         End If
      
      'Added by Morgan 2018/9/14
      ElseIf iMS15 = 2 ^ 5 Then
         strExc(0) = "select MSI02 from MailScheduleImport" & _
            " where MSI01=" & p_MS01 & " AND MSI03='Y' order by 1"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            With RsTemp
            Do While Not .EOF
               lstImport(1).AddItem .Fields(0)
               .MoveNext
            Loop
            lblImpCount(1).Caption = lstImport(1).ListCount
            End With
         End If
         
         m_GDPR = True
         chkMainOnly.Visible = True
         chkMainOnly.Value = vbChecked
         cmdImport(1).Enabled = True
         Label4.Caption = "指定R,Y編號："
      End If
      
      ReadSchedule = True
   End If
End Function

Private Sub SSTab1_Click(PreviousTab As Integer)
    
On Error Resume Next
    Select Case Me.SSTab1.Tab
    Case 0
        txtNo.SetFocus
        txtNo_GotFocus
        cmdQuery(0).Default = False
    Case 1
        txtQueryDate(0).SetFocus
        txtQueryDate_GotFocus 0
        cmdQuery(0).Default = True
    End Select
End Sub

Private Sub txtNo_GotFocus()
   CloseIme
   TextInverse txtNo
End Sub

Private Sub txtQueryDate_GotFocus(Index As Integer)
   CloseIme
   TextInverse txtQueryDate(Index)
End Sub


Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   Screen.MousePointer = vbHourglass
   SSTab1.Tab = 0
   Action Button.Index
   Screen.MousePointer = vbDefault
End Sub

Private Sub RsAction(ByVal Sty As Integer)
 Dim stCon As String
 
 If m_DepCode <> "" Then
   stCon = stCon & " And MS22='" & m_DepCode & "'"
End If
   
On Error GoTo ErrHand
   Screen.MousePointer = vbHourglass
   intI = 1
   Select Case Sty
      Case 0 '第一筆
         strExc(0) = "SELECT nvl(min(ms01),0) FROM mailschedule where 1=1 " & stCon
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp.Fields(0) > 0 Then
               ReadSchedule RsTemp.Fields(0)
            End If
         End If
         
      Case 1 '前一筆
         stCon = stCon & " and ms01<" & Val(txtNo)
         strExc(0) = "SELECT nvl(max(ms01)," & Val(txtNo) & ") FROM mailschedule where 1=1 " & stCon
         
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp.Fields(0) = Val(txtNo) Then
               DataErrorMessage 6
            Else
               ReadSchedule RsTemp.Fields(0)
            End If
         End If
         
      Case 2 '後一筆
         stCon = stCon & " and ms01>" & Val(txtNo)
         strExc(0) = "SELECT nvl(min(ms01)," & Val(txtNo) & ") FROM mailschedule where 1=1 " & stCon
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp.Fields(0) = Val(txtNo) Then
               DataErrorMessage 7
            Else
               ReadSchedule RsTemp.Fields(0)
            End If
         End If
         
      Case 3 '最後筆
         strExc(0) = "SELECT nvl(max(ms01),0) FROM mailschedule where 1=1 " & stCon
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp.Fields(0) > 0 Then
               ReadSchedule RsTemp.Fields(0)
            End If
         End If
   End Select
   Screen.MousePointer = vbDefault
   Exit Sub
   
ErrHand:
   Screen.MousePointer = vbDefault
   MsgBox "錯誤 : " & Err.Description, vbCritical
End Sub

Private Sub TxtLock(ByVal Lt As Integer)
   Select Case Lt
   Case 0 '新增
      SSTab1.TabEnabled(1) = False
      txtNo.Enabled = False
      Frame1.Enabled = True
      CmdSitu False
      cmdEstimate.Visible = True
      
   Case 1 '修改
      SSTab1.TabEnabled(1) = False
      txtNo.Enabled = False
      Frame1.Enabled = True
      CmdSitu False
      cmdEstimate.Visible = True
      
   Case 2 '查詢
      SSTab1.TabEnabled(1) = False
      txtNo.Enabled = True
      txtNo.SetFocus
      Frame1.Enabled = False
      CmdSitu False
      cmdEstimate.Visible = False
      
   Case 3 '瀏覽
      SSTab1.TabEnabled(1) = True
      txtNo.Enabled = False
      Frame1.Enabled = False
      CmdSitu True
      cmdEstimate.Visible = False
      
   End Select
End Sub

Private Sub Action(Index As Integer)

   If TBar1.Buttons(Index).Enabled = False Then Exit Sub

On Error GoTo ErrHand
   
   Select Case Index
      Case 1 '按下新增
         txtNo.Tag = txtNo
         FormReset
         ActionEdit = 0
         SetList 0, True
         txtDate = strSrvDate(2)
         m_bolTestOK = False
         
      Case 2 '按下修改
         If GetTemplete(txtNo) = False Then
            MsgBox "樣本檔案讀取失敗！"
            Exit Sub
         End If
         SetList , True
         txtNo.Tag = txtNo
         ActionEdit = 1
         
      Case 3 '按下刪除
         If MsgBox("是否確定要刪除排程??", vbYesNo + vbDefaultButton2) = vbYes Then
            If FormDelete() = False Then
               MsgBox "刪除失敗!", vbCritical
               Exit Sub
            '刪除後移到最末筆
            Else
               RsAction 3
            End If
         Else
            Exit Sub
         End If
         ActionEdit = 3
      Case 4 '按下查詢
         txtNo.Tag = txtNo
         FormReset
         ActionEdit = 2
      Case 6 '第一筆
         RsAction 0
         ActionEdit = 3
      Case 7 '前一筆
         RsAction 1
         ActionEdit = 3
      Case 8 '後一筆
         RsAction 2
         ActionEdit = 3
      Case 9 '最後筆
         RsAction 3
         ActionEdit = 3
         
      Case 11 '按下確定
         Select Case ActionEdit
            Case 0, 1 '新增,修改
               If TxtValidate = False Then
                  Exit Sub
               Else
                  If FormSave() = False Then
                     MsgBox "存檔失敗!", vbCritical
                     Exit Sub
                  End If
               End If
         End Select
         
         If ReadSchedule(txtNo) = False Then
            MsgBox "排程讀取失敗!", vbCritical
            Exit Sub
         End If
         ActionEdit = 3
      
      Case 12 '按下取消
         txtNo = txtNo.Tag
         If txtNo <> "" Then
            ReadSchedule txtNo
         End If
         ActionEdit = 3
      Case 14 '結束
         Unload Me
         Exit Sub
   End Select
   
   TxtLock ActionEdit
   Exit Sub
   
ErrHand:
   ShowMsg "錯誤 : " & Err.Description
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Screen.MousePointer = vbHourglass
   Select Case KeyCode
      Case vbKeyF2: Action 1 '新增
      Case vbKeyF3: Action 2 '修改
      Case vbKeyF5: Action 3 '刪除
      Case vbKeyF4: Action 4 '查詢
      Case vbKeyHome: Action 6 '第一筆
      Case vbKeyPageUp: Action 7 '前一筆
      Case vbKeyPageDown: Action 8 '後一筆
      Case vbKeyEnd: Action 9 '最後筆
      Case vbKeyF9, vbKeyReturn: Action 11 '確定
      Case vbKeyF10: Action 12 '取消
      Case vbKeyEscape: Action 14 '結束
    End Select
    Screen.MousePointer = vbDefault
End Sub

Private Sub CmdSitu(ByVal TF As Boolean)
   Dim i As Integer, txt As TextBox
   Dim oButton As Button
 
   For i = 1 To 4
      TBar1.Buttons(i).Enabled = False
      TBar1.Buttons(i + 5).Enabled = False
   Next
   TBar1.Buttons(11).Enabled = False
   TBar1.Buttons(12).Enabled = False
   TBar1.Buttons(14).Enabled = False
      
   If TF = True Then
      If m_bInsert Then
          TBar1.Buttons(1).Enabled = True
      End If
      
      If txtNo <> "" Then
         If lblActFrom & lblActTo = "" Or Val(txtCount) = 0 Then
            If m_bUpdate Then
                TBar1.Buttons(2).Enabled = True
            End If
            If m_bDelete Then
                TBar1.Buttons(3).Enabled = True
            End If
         End If
         For i = 1 To 4
            TBar1.Buttons(i + 5).Enabled = True
         Next
         TBar1.Buttons(4).Enabled = True
      End If
      TBar1.Buttons(14).Enabled = True
   Else
      TBar1.Buttons(11).Enabled = True
      TBar1.Buttons(12).Enabled = True
   End If
End Sub
'Added by Morgan 2018/8/10
Private Function GetChoiceNew() As Integer
   Dim ii As Integer
   For ii = 0 To lstMailChoice.ListCount - 1
      If lstMailChoice.Selected(ii) = True Then
         GetChoiceNew = lstMailChoice.ITEMDATA(ii)
         Exit For
      End If
   Next
End Function
'Added by Morgan 2018/11/28
'檢查寄件信箱是否正確
Private Function CheckAccount() As Boolean
   Dim objOutLook As Object
   Dim objMail As Object
   Dim bolOK As Boolean
   
   CheckAccount = False
   Set objOutLook = CreateObject("Outlook.Application")
   Set objMail = objOutLook.CreateItemFromTemplate(txtSample)
   
On Error GoTo ErrHnd

   If Not objMail.SendUsingAccount Is Nothing Then
      If objMail.SendUsingAccount = cboEmail Then
         CheckAccount = True
      End If
   End If
   If CheckAccount = False Then
      MsgBox "寄件信箱錯誤！", vbCritical
   End If
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   
   objMail.Delete
   Set objMail = Nothing
   Set objOutLook = Nothing
End Function

'Added by Morgan 2024/6/6
Private Sub loadMailTemplete(pChoice As Integer)
   Dim stAttPath As String, stFileName As String
   
   txtSample = ""
   stFileName = "mailschedule.tmp"
   stAttPath = App.path & "\" & stFileName
   If Dir(stAttPath) <> "" Then Kill stAttPath
   If pChoice = 10 Then
      Call PUB_GetSampleFile(stFileName, "TOT-000M31-0-04")
   ElseIf pChoice = 11 Then
      Call PUB_GetSampleFile(stFileName, "TOT-000M31-0-05")
   End If
   If Dir(stAttPath) <> "" Then txtSample = stAttPath
End Sub

'Added by Morgan 2024/6/6
Private Function AutoAddSchedule() As Boolean
   Dim strMsg As String
   
On Error GoTo ErrHnd
   
   If m_Schedule1 Then
      Tbar1_ButtonClick TBar1.Buttons(1) '新增
      cboTime = "18:00"
      lstMailChoice.Selected(0) = True '索取CF對帳單(中文)
      cmdTest.Value = True
      Tbar1_ButtonClick TBar1.Buttons(11)  '確定
      
      strMsg = vbCrLf & vbCrLf & "#" & txtNo & ": " & lstMailChoice.List(0) & vbCrLf & "　預定發信日期: " & txtDate & vbCrLf & "　　　　　時間: " & cboTime & vbCrLf & "　　　　發信數: " & txtCount
   End If
   
   If m_Schedule2 Then
      Tbar1_ButtonClick TBar1.Buttons(1) '新增
      cboTime = "18:30"
      lstMailChoice.Selected(1) = True '索取CF對帳單(英文)
      cmdTest.Value = True
      Tbar1_ButtonClick TBar1.Buttons(11)  '確定
      
      strMsg = strMsg & vbCrLf & vbCrLf & "#" & txtNo & ": " & lstMailChoice.List(0) & vbCrLf & "　預定發信日期: " & txtDate & vbCrLf & "　　　　　時間: " & cboTime & vbCrLf & "　　　　發信數: " & txtCount
   End If
   
   AutoAddSchedule = True
   MsgBox "排程已新增如下：" & strMsg & vbCrLf & vbCrLf & "※已自動寄發測試信，請務必確認是否正確！", vbInformation
   
   Exit Function
   
ErrHnd:
   MsgBox Err.Description, vbCritical
End Function


'Add by Amy 2025/09/02 設定測式資料
Private Function SetTestMailData(pMS01, stMsg As String) As Boolean
   Dim ii As Integer, stTest_Fix As String, stCmd As String, lngCmd As Long
   
   SetTestMailData = False: stMsg = ""
   stTest_Fix = "Select min(msd06) From MailScheduleDetail  Where msd01=" & pMS01 & " And Substr(msd06,1,1)='X' "
   For ii = 1 To 3
      If ii = 1 Then
         stCmd = stTest_Fix & "And Instr(msd06,'-')>0 "
      ElseIf ii = 2 Then
         stCmd = Replace(stTest_Fix, "Substr(msd06,1,1)='X'", "Substr(msd06,1,1)='Y'")
      Else
         stCmd = Replace(stTest_Fix, "Substr(msd06,1,1)='X'", "Substr(msd06,1,1)='R'")
      End If
      stCmd = "Update MailScheduleDetail Set msd06='*'||msd06 Where msd01=" & pMS01 & " And msd06 in (" & stCmd & ")"
      cnnConnection.Execute stCmd
   Next ii
   
   stCmd = "Delete From MailScheduleDetail Where msd01=" & pMS01 & " And SubStr(msd06,1,1)<>'*' "
   cnnConnection.Execute stCmd, lngCmd
   If lngCmd = 0 Then
      stMsg = "刪除測式資料失敗"
      Exit Function
   End If
               
   '更新mail Address
   stCmd = "Update MailScheduleDetail Set msd02='A2004@taie.com.tw',msd06=SubStr(msd06,2) " & _
                  "Where msd01=" & pMS01 & " "
   cnnConnection.Execute stCmd, lngCmd
   If lngCmd = 0 Then
      stMsg = "更新測式Mail Addr失敗"
      Exit Function
   End If
   
   SetTestMailData = True
End Function

'Add by Amy 2025/09/02 測式與寄出Tag比對用-ex:首字為英文句點[.]寄出會被吃掉而導致內容缺.或格式異常
Private Sub SaveMime(stUPMime As String, stOfficeTag As String, stEndMime As String)
   Dim F1 As Integer
   Dim FileFullPath As String, Mime_file As String
   
   Mime_file = App.path & "\" & "解析完Tag.txt"

   If F1 > 0 Then Close #F1
   F1 = FreeFile
   Open Mime_file For Output As F1
   
   Print #F1, stUPMime & stOfficeTag & stEndMime
   Close F1
End Sub

