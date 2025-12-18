VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm140113 
   BorderStyle     =   1  '單線固定
   Caption         =   "教育訓練登錄作業"
   ClientHeight    =   5736
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9108
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5736
   ScaleWidth      =   9108
   Begin TabDlg.SSTab SSTab1 
      Height          =   5055
      Left            =   90
      TabIndex        =   23
      Top             =   660
      Width           =   8925
      _ExtentX        =   15748
      _ExtentY        =   8911
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "基本資料"
      TabPicture(0)   =   "frm140113.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(9)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblCreateData"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label7"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblReceiver"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Text1(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Option1(1)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Option1(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Command2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtReceiver"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "登記"
      TabPicture(1)   =   "frm140113.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label4"
      Tab(1).Control(1)=   "lblMemo"
      Tab(1).Control(2)=   "lblTitle"
      Tab(1).Control(3)=   "MSHFlexGrid2"
      Tab(1).Control(4)=   "cmdBooinSave"
      Tab(1).Control(5)=   "Timer1"
      Tab(1).Control(6)=   "cmdPrint"
      Tab(1).Control(7)=   "CboChoose"
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "教育訓練資料"
      TabPicture(2)   =   "frm140113.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label1(8)"
      Tab(2).Control(1)=   "cmdOpenAtt(0)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "lstAtt"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "cmdSaveAtt(0)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "cmdAddAtt(0)"
      Tab(2).Control(5)=   "cmdRemAtt(0)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Frame2"
      Tab(2).Control(7)=   "cmdSelect(0)"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).ControlCount=   8
      TabCaption(3)   =   "多筆查詢"
      TabPicture(3)   =   "frm140113.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Line2"
      Tab(3).Control(1)=   "Label1(11)"
      Tab(3).Control(2)=   "Label1(13)"
      Tab(3).Control(3)=   "Label1(14)"
      Tab(3).Control(4)=   "Label1(15)"
      Tab(3).Control(5)=   "Label1(16)"
      Tab(3).Control(6)=   "Label6"
      Tab(3).Control(7)=   "lblName"
      Tab(3).Control(8)=   "Label1(17)"
      Tab(3).Control(9)=   "Line1"
      Tab(3).Control(10)=   "Label9"
      Tab(3).Control(11)=   "txtSpeaker"
      Tab(3).Control(12)=   "txtAttender"
      Tab(3).Control(13)=   "txtSubject"
      Tab(3).Control(14)=   "grdList"
      Tab(3).Control(15)=   "txtQueryDate(1)"
      Tab(3).Control(16)=   "txtQueryDate(0)"
      Tab(3).Control(17)=   "cmdQuery(0)"
      Tab(3).Control(18)=   "Frame3"
      Tab(3).Control(19)=   "txtDept(0)"
      Tab(3).Control(20)=   "txtDept(1)"
      Tab(3).Control(21)=   "cmdQuery(1)"
      Tab(3).ControlCount=   22
      Begin VB.ComboBox CboChoose 
         Height          =   276
         ItemData        =   "frm140113.frx":0070
         Left            =   -67650
         List            =   "frm140113.frx":0072
         Style           =   2  '單純下拉式
         TabIndex        =   99
         Top             =   420
         Width           =   1500
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "參加記錄(未失效)"
         Height          =   400
         Index           =   1
         Left            =   -69120
         TabIndex        =   93
         Top             =   510
         Width           =   1600
      End
      Begin VB.TextBox txtReceiver 
         Height          =   285
         Left            =   6915
         MaxLength       =   6
         TabIndex        =   91
         Top             =   550
         Width           =   885
      End
      Begin VB.CommandButton Command2 
         Caption         =   "測試信"
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
         Left            =   6120
         TabIndex        =   90
         TabStop         =   0   'False
         Top             =   550
         Width           =   770
      End
      Begin VB.TextBox txtDept 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   -72720
         MaxLength       =   3
         TabIndex        =   51
         Top             =   1750
         Width           =   700
      End
      Begin VB.TextBox txtDept 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   -73590
         MaxLength       =   3
         TabIndex        =   50
         Top             =   1750
         Width           =   700
      End
      Begin VB.OptionButton Option1 
         Caption         =   "公開"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   1680
         TabIndex        =   80
         Top             =   600
         Value           =   -1  'True
         Width           =   800
      End
      Begin VB.OptionButton Option1 
         Caption         =   "不公開"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   2520
         TabIndex        =   79
         Top             =   600
         Width           =   1000
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "列印點名單"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -68820
         TabIndex        =   76
         Top             =   390
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   30
         Left            =   -72300
         TabIndex        =   75
         Top             =   1110
         Width           =   15
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "全選"
         Height          =   345
         Index           =   0
         Left            =   -67125
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   900
         Width           =   735
      End
      Begin VB.Frame Frame2 
         Height          =   2295
         Left            =   -74910
         TabIndex        =   60
         Top             =   2610
         Width           =   8610
         Begin VB.CommandButton cmdSelect 
            Caption         =   "全選"
            Height          =   345
            Index           =   1
            Left            =   7785
            TabIndex        =   67
            TabStop         =   0   'False
            Top             =   540
            Width           =   735
         End
         Begin VB.CommandButton cmdOpenAtt 
            Caption         =   "開啟"
            Height          =   345
            Index           =   1
            Left            =   7785
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   180
            Width           =   735
         End
         Begin VB.CommandButton cmdSaveAtt 
            Caption         =   "另存"
            Height          =   345
            Index           =   1
            Left            =   7785
            TabIndex        =   64
            TabStop         =   0   'False
            Top             =   900
            Width           =   735
         End
         Begin VB.ListBox lstAtt1 
            Height          =   1848
            ItemData        =   "frm140113.frx":0074
            Left            =   720
            List            =   "frm140113.frx":0076
            MultiSelect     =   2  '進階多重選取
            Sorted          =   -1  'True
            TabIndex        =   63
            TabStop         =   0   'False
            Top             =   180
            Width           =   7080
         End
         Begin VB.CommandButton cmdRemAtt 
            Caption         =   "-> 移除"
            Height          =   345
            Index           =   1
            Left            =   7785
            TabIndex        =   62
            TabStop         =   0   'False
            Top             =   1620
            Width           =   735
         End
         Begin VB.CommandButton cmdAddAtt 
            Caption         =   "<- 新增"
            Height          =   345
            Index           =   1
            Left            =   7785
            TabIndex        =   61
            Top             =   1260
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "原始　附件："
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   12
            Left            =   90
            TabIndex        =   66
            Top             =   210
            Width           =   615
            WordWrap        =   -1  'True
         End
      End
      Begin VB.CommandButton cmdRemAtt 
         Caption         =   "-> 移除"
         Height          =   345
         Index           =   0
         Left            =   -67125
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   1980
         Width           =   735
      End
      Begin VB.CommandButton cmdAddAtt 
         Caption         =   "<- 新增"
         Height          =   345
         Index           =   0
         Left            =   -67125
         TabIndex        =   57
         Top             =   1620
         Width           =   735
      End
      Begin VB.CommandButton cmdSaveAtt 
         Caption         =   "另存"
         Height          =   345
         Index           =   0
         Left            =   -67125
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   1260
         Width           =   735
      End
      Begin VB.ListBox lstAtt 
         Height          =   1848
         ItemData        =   "frm140113.frx":0078
         Left            =   -74190
         List            =   "frm140113.frx":007A
         MultiSelect     =   2  '進階多重選取
         Sorted          =   -1  'True
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   540
         Width           =   7080
      End
      Begin VB.CommandButton cmdOpenAtt 
         Caption         =   "開啟"
         Height          =   345
         Index           =   0
         Left            =   -67125
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   540
         Width           =   735
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "查詢(&Q)"
         Height          =   400
         Index           =   0
         Left            =   -70140
         TabIndex        =   52
         Top             =   510
         Width           =   912
      End
      Begin VB.TextBox txtQueryDate 
         Height          =   315
         Index           =   0
         Left            =   -73590
         MaxLength       =   7
         TabIndex        =   45
         Top             =   540
         Width           =   945
      End
      Begin VB.TextBox txtQueryDate 
         Height          =   315
         Index           =   1
         Left            =   -72540
         MaxLength       =   7
         TabIndex        =   46
         Top             =   540
         Width           =   945
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   -68745
         Top             =   4470
      End
      Begin VB.CommandButton cmdBooinSave 
         Caption         =   "確認登記"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   -67620
         Style           =   1  '圖片外觀
         TabIndex        =   41
         Top             =   4440
         Width           =   1320
      End
      Begin VB.Frame Frame1 
         Height          =   4185
         Left            =   90
         TabIndex        =   24
         Top             =   750
         Width           =   8745
         Begin VB.ComboBox cboTime 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.4
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   0
            Left            =   3210
            TabIndex        =   9
            Text            =   "cboTime"
            Top             =   1830
            Width           =   1100
         End
         Begin VB.ComboBox cboTime 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.4
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   1
            Left            =   4560
            TabIndex        =   10
            Text            =   "cboTime"
            Top             =   1830
            Width           =   1100
         End
         Begin VB.ComboBox cboRoom 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   690
            TabIndex        =   7
            Text            =   "cboRoom"
            Top             =   1500
            Width           =   2300
         End
         Begin VB.CheckBox Check2 
            Caption         =   "不需登記"
            Height          =   225
            Index           =   1
            Left            =   7200
            TabIndex        =   6
            Top             =   1400
            Width           =   1300
         End
         Begin VB.CheckBox Check2 
            Caption         =   "部門全體人員"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   8.4
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   7200
            TabIndex        =   5
            Top             =   1150
            Width           =   1300
         End
         Begin VB.CommandButton CmdMeeting 
            BackColor       =   &H00C0FFC0&
            Caption         =   "會議室預約"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5685
            Style           =   1  '圖片外觀
            TabIndex        =   11
            Top             =   1865
            Width           =   1160
         End
         Begin VB.CommandButton CmdAdd 
            Caption         =   "<- 新增"
            Height          =   285
            Index           =   1
            Left            =   7150
            TabIndex        =   87
            Top             =   510
            Width           =   735
         End
         Begin VB.CommandButton CmdDel 
            Caption         =   "刪除 ->"
            Height          =   285
            Index           =   1
            Left            =   7150
            TabIndex        =   86
            Top             =   840
            Width           =   735
         End
         Begin VB.CommandButton CmdDel 
            Caption         =   "刪除 ->"
            Height          =   285
            Index           =   0
            Left            =   2200
            TabIndex        =   84
            Top             =   840
            Width           =   735
         End
         Begin VB.CommandButton CmdAdd 
            Caption         =   "<- 新增"
            Height          =   285
            Index           =   0
            Left            =   2200
            TabIndex        =   12
            Top             =   510
            Width           =   735
         End
         Begin VB.CommandButton Command1 
            Caption         =   "人員確認"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   7650
            TabIndex        =   77
            Top             =   3810
            Width           =   960
         End
         Begin VB.CommandButton Command1 
            Caption         =   "刪除"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   5
            Left            =   7830
            TabIndex        =   16
            Top             =   2780
            Width           =   780
         End
         Begin VB.CommandButton Command1 
            Caption         =   "修改"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   7830
            TabIndex        =   15
            Top             =   2470
            Width           =   780
         End
         Begin VB.CommandButton Command1 
            Caption         =   "新增"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   7830
            TabIndex        =   14
            Top             =   2180
            Width           =   780
         End
         Begin VB.CommandButton Command1 
            Caption         =   "＞"
            BeginProperty Font 
               Name            =   "@新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   375
            TabIndex        =   17
            Top             =   2460
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.CommandButton Command1 
            Caption         =   "＜"
            BeginProperty Font 
               Name            =   "@標楷體"
               Size            =   9.6
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   4
            Left            =   375
            TabIndex        =   18
            Top             =   2730
            Visible         =   0   'False
            Width           =   285
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   45
            Top             =   2580
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   315
            Index           =   3
            Left            =   1155
            TabIndex        =   20
            Top             =   3810
            Width           =   1365
            _ExtentX        =   2413
            _ExtentY        =   550
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   185139201
            CurrentDate     =   40942
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   315
            Index           =   4
            Left            =   2775
            TabIndex        =   21
            Top             =   3810
            Width           =   1365
            _ExtentX        =   2413
            _ExtentY        =   550
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   185139201
            CurrentDate     =   40942
         End
         Begin VB.CheckBox Check1 
            Caption         =   "存檔後寄發通知信"
            Height          =   225
            Left            =   5775
            TabIndex        =   13
            Top             =   3840
            Width           =   1815
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   870
            Left            =   690
            TabIndex        =   44
            Top             =   2180
            Width           =   7050
            _ExtentX        =   12425
            _ExtentY        =   1545
            _Version        =   393216
            BackColor       =   -2147483624
            Rows            =   1
            Cols            =   1
            FixedRows       =   0
            FixedCols       =   0
            WordWrap        =   -1  'True
            HighLight       =   2
            GridLinesFixed  =   1
            SelectionMode   =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "細明體-ExtB"
               Size            =   9.6
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "細明體"
               Size            =   11.4
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   1
            _Band(0).GridLinesBand=   0
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
         Begin MSMask.MaskEdBox MaskEdBox1 
            Height          =   345
            Left            =   690
            TabIndex        =   98
            Top             =   1830
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   593
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "新細明體"
               Size            =   10.8
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSForms.ListBox lstCC 
            Height          =   600
            Left            =   4960
            TabIndex        =   89
            TabStop         =   0   'False
            Top             =   510
            Width           =   2150
            VariousPropertyBits=   746586139
            ScrollBars      =   2
            DisplayStyle    =   2
            Size            =   "3792;1058"
            MatchEntry      =   0
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.ListBox lstUsers 
            Height          =   600
            Left            =   690
            TabIndex        =   82
            TabStop         =   0   'False
            Top             =   510
            Width           =   1500
            VariousPropertyBits=   746586139
            ScrollBars      =   2
            DisplayStyle    =   2
            Size            =   "2646;1058"
            MatchEntry      =   0
            MultiSelect     =   2
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.ComboBox cboEmp 
            Height          =   300
            Left            =   2950
            TabIndex        =   2
            Top             =   510
            Width           =   1300
            VariousPropertyBits=   679495707
            DisplayStyle    =   3
            Size            =   "2293;529"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox Text1 
            Height          =   285
            Index           =   23
            Left            =   7710
            TabIndex        =   97
            Top             =   1860
            Width           =   250
            VariousPropertyBits=   679493659
            MaxLength       =   50
            Size            =   "441;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox Text1 
            Height          =   345
            Index           =   3
            Left            =   690
            TabIndex        =   4
            Top             =   1150
            Width           =   6420
            VariousPropertyBits=   679493659
            MaxLength       =   50
            Size            =   "11324;609"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox Text1 
            Height          =   705
            Index           =   9
            Left            =   690
            TabIndex        =   19
            Top             =   3090
            Width           =   7920
            VariousPropertyBits=   -1467989989
            MaxLength       =   500
            ScrollBars      =   2
            Size            =   "13970;1244"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox Text1 
            Height          =   345
            Index           =   18
            Left            =   690
            TabIndex        =   1
            Top             =   150
            Width           =   7920
            VariousPropertyBits=   679493659
            MaxLength       =   50
            Size            =   "13970;609"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox Text1 
            Height          =   345
            Index           =   4
            Left            =   7910
            TabIndex        =   3
            Top             =   480
            Width           =   700
            VariousPropertyBits=   679493659
            MaxLength       =   50
            Size            =   "1235;609"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox Text1 
            Height          =   345
            Index           =   8
            Left            =   2970
            TabIndex        =   8
            Top             =   1500
            Width           =   4140
            VariousPropertyBits=   679493659
            MaxLength       =   50
            Size            =   "7302;609"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Y:不檢查"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   19
            Left            =   7980
            TabIndex        =   100
            Top             =   1920
            Width           =   705
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "不檢查 ："
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   18
            Left            =   6930
            TabIndex        =   96
            Top             =   1920
            Width           =   825
         End
         Begin VB.Label Label8 
            Caption         =   "(限智權部使用)"
            ForeColor       =   &H00FF0000&
            Height          =   225
            Left            =   7380
            TabIndex        =   94
            Top             =   1620
            Width           =   1200
         End
         Begin MSForms.Label lblCC 
            Height          =   270
            Left            =   7920
            TabIndex        =   88
            Top             =   840
            Width           =   690
            BackColor       =   -2147483634
            VariousPropertyBits=   27
            Size            =   "1217;466"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   195
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.Label lblChairMan 
            Height          =   270
            Left            =   3000
            TabIndex        =   83
            Top             =   840
            Width           =   1005
            BackColor       =   -2147483634
            VariousPropertyBits=   27
            Size            =   "1773;466"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   195
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label5 
            Caption         =   "標題："
            Height          =   255
            Left            =   100
            TabIndex        =   42
            Top             =   210
            Width           =   555
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   " ～"
            Height          =   180
            Left            =   2505
            TabIndex        =   35
            Top             =   3870
            Width           =   270
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "可登記日期："
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   10
            Left            =   60
            TabIndex        =   34
            Top             =   3840
            Width           =   1170
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "說明："
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   7
            Left            =   60
            TabIndex        =   33
            Top             =   3090
            Width           =   585
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "議題："
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   6
            Left            =   100
            TabIndex        =   32
            Top             =   2220
            Width           =   585
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "出席："
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   100
            TabIndex        =   26
            Top             =   1185
            Width           =   585
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "主席："
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   100
            TabIndex        =   25
            Top             =   555
            Width           =   585
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "副本："
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   4360
            TabIndex        =   27
            Top             =   555
            Width           =   585
         End
         Begin VB.Label lblWeek 
            AutoSize        =   -1  'True
            Caption         =   "星期一"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1920
            TabIndex        =   36
            Top             =   1905
            Width           =   585
         End
         Begin VB.Label Label2 
            Caption         =   "～"
            Height          =   180
            Left            =   4350
            TabIndex        =   30
            Top             =   1920
            Width           =   300
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "時間："
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   4
            Left            =   2685
            TabIndex        =   29
            Top             =   1905
            Width           =   585
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "日期："
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   105
            TabIndex        =   28
            Top             =   1900
            Width           =   585
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "地點："
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   5
            Left            =   105
            TabIndex        =   31
            Top             =   1575
            Width           =   585
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   3555
         Left            =   -74865
         TabIndex        =   37
         Top             =   810
         Width           =   8565
         _ExtentX        =   15113
         _ExtentY        =   6265
         _Version        =   393216
         BackColor       =   -2147483624
         WordWrap        =   -1  'True
         AllowBigSelection=   0   'False
         HighLight       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體-ExtB"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
         Height          =   2844
         Left            =   -74928
         TabIndex        =   101
         Top             =   2112
         Width           =   8652
         _ExtentX        =   15261
         _ExtentY        =   5017
         _Version        =   393216
         AllowUserResizing=   3
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
      Begin MSForms.TextBox txtSubject 
         Height          =   315
         Left            =   -73590
         TabIndex        =   47
         Top             =   840
         Width           =   1995
         VariousPropertyBits=   679493659
         Size            =   "3519;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtAttender 
         Height          =   315
         Left            =   -73590
         TabIndex        =   49
         Top             =   1440
         Width           =   1095
         VariousPropertyBits=   679493659
         MaxLength       =   6
         Size            =   "1931;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtSpeaker 
         Height          =   315
         Left            =   -73590
         TabIndex        =   48
         Top             =   1140
         Width           =   1095
         VariousPropertyBits=   679493659
         Size            =   "1931;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   345
         Index           =   1
         Left            =   795
         TabIndex        =   0
         Top             =   390
         Width           =   735
         VariousPropertyBits=   679493659
         Size            =   "1296;609"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label9 
         Caption         =   "＊未登記"
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   -66960
         TabIndex        =   95
         Top             =   1860
         Width           =   795
      End
      Begin MSForms.Label lblReceiver 
         Height          =   264
         Left            =   7875
         TabIndex        =   92
         Top             =   580
         Width           =   840
         BackColor       =   -2147483634
         VariousPropertyBits=   27
         Size            =   "1482;466"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Line Line1 
         X1              =   -72960
         X2              =   -72540
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "部門:"
         Height          =   180
         Index           =   17
         Left            =   -74760
         TabIndex        =   85
         Top             =   1800
         Width           =   405
      End
      Begin VB.Label Label7 
         Caption         =   "公開副本發送固定人員"
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   1680
         TabIndex        =   81
         Top             =   360
         Width           =   1900
      End
      Begin MSForms.Label lblCreateData 
         Height          =   195
         Left            =   3645
         TabIndex        =   78
         Top             =   360
         Width           =   5000
         BackColor       =   -2147483634
         VariousPropertyBits=   27
         Size            =   "8819;344"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   165
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblName 
         Height          =   270
         Left            =   -72435
         TabIndex        =   74
         Top             =   1440
         Width           =   1485
         BackColor       =   16777152
         VariousPropertyBits=   27
         Size            =   "2619;466"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   195
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "( 可輸入員工編號或名稱)"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -72390
         TabIndex        =   73
         Top             =   1200
         Width           =   2115
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "( 模糊比對 )"
         Height          =   180
         Index           =   16
         Left            =   -71490
         TabIndex        =   72
         Top             =   930
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "參加人員工號:"
         Height          =   180
         Index           =   15
         Left            =   -74760
         TabIndex        =   71
         Top             =   1500
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "主講人:"
         Height          =   180
         Index           =   14
         Left            =   -74760
         TabIndex        =   70
         Top             =   1207
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "議題:"
         Height          =   180
         Index           =   13
         Left            =   -74760
         TabIndex        =   69
         Top             =   907
         Width           =   405
      End
      Begin VB.Label Label1 
         Caption         =   "開放　附件："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   8
         Left            =   -74820
         TabIndex        =   59
         Top             =   540
         Width           =   615
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "研討會日期:"
         Height          =   180
         Index           =   11
         Left            =   -74760
         TabIndex        =   53
         Top             =   607
         Width           =   945
      End
      Begin VB.Line Line2 
         X1              =   -72810
         X2              =   -72390
         Y1              =   660
         Y2              =   660
      End
      Begin VB.Label lblTitle 
         Caption         =   "標題"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -74865
         TabIndex        =   43
         Top             =   420
         Width           =   6000
      End
      Begin VB.Label lblMemo 
         AutoSize        =   -1  'True
         Caption         =   "* 請於 ??? 下班前完成「V」註。"
         Height          =   180
         Left            =   -74775
         TabIndex        =   40
         Top             =   4680
         Width           =   2550
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "* 煩請「V」 註(雙擊滑鼠左鍵)欲參加之課程。"
         Height          =   180
         Left            =   -74775
         TabIndex        =   39
         Top             =   4470
         Width           =   3660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "編號："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   9
         Left            =   225
         TabIndex        =   38
         Top             =   450
         Width           =   585
      End
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
            Picture         =   "frm140113.frx":007C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140113.frx":0398
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140113.frx":06B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140113.frx":0890
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140113.frx":0BAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140113.frx":0EC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140113.frx":11E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140113.frx":1500
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140113.frx":181C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140113.frx":1B38
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140113.frx":1E54
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   576
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   9108
      _ExtentX        =   16066
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
End
Attribute VB_Name = "frm140113"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/10/16 Form2.0已修改 ; grdList從MSFlexGrid改為MSHFlexGrid
'Memo by Amy 2022/01/05 Form2.0已修改 lblCreate/lblReceiver/lblChairMan/lblName/lblCC/text1()/cboEmp/lstUsers/lstCC/txtSpeaker/txtAttender/MSHFlexGrid1/MSHFlexGrid2/grdList
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Created by Morgan 2012/2/13
Option Explicit

Dim ActionEdit As Integer '0:新增/1:修改/2:查詢/3:取消
'執行各項功能的權限
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
Dim m_bOpen As Boolean '開啟原始附件權限(有設定列印權限者)
Dim m_bPrint As Boolean

Dim m_CurrSel As Integer

Dim m_iRow As Integer
Dim m_FilesRemoved() As String
Dim m_selRow As Integer, m_selCol As Integer
Dim m_IsOpen As Boolean
Dim m_UpdateCount As Integer
Dim m_AttachPath As String

Private Declare Function SendMessageByNum Lib "user32" _
  Alias "SendMessageA" (ByVal hWnd As Long, ByVal _
  wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Const LB_SETHORIZONTALEXTENT = &H194

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

Dim m_arrBookInList() As String
Public m_stNumList As String
'Add by Amy 2018/09/18
'Modify by Amy 2022/12/30 69009(楊監察人)退休-->B1015(林岱嫻特助)
'Modify by Amy 2023/07/10 76012(桂所長)退休-->拿掉
Const strCC_Fix = "63001;81040;94007;B1015" '固定副本收受者人員(不可刪)
Dim i As Integer
Dim Del_arrBookinList() As String
Dim strDeptNo As String, strDeptName As String, strSN12 As String '部門/部門名稱/建立者
Dim bolJoin As Boolean '登入者為參加人員
Dim m_arrBookPeo() As String '非同部門需勾選人員
'Add by Amy 2019/01/24
Dim strSN13 As String '修改日
Dim strOldRR(1 To 12) As String '最原始會議室預約記錄(for進會議室預約存檔後,回此畫面取消)
Dim strDefDeptMail As String '部門全體人員Email(Outlook Mail)
'Add by Amy 2019/11/12
Dim strSeminar As String '固定抓建立者同部門、主席、副本人員或參加人員才可查詢語法
Dim m_FirstKEY As String, m_LastKEY As String ' 第一筆/最後一筆
Dim strDeptSql As String 'Add by Amy 2020/11/27 固定抓部門語法
Dim intMaxTB As String, intMaxCol As Integer, iTables As Integer, iNowCol As Integer 'Add by Amy 2021/01/25 從run Word 搬過來(iNowCol:目前欄)
Dim m_MeTrackMode  As String 'Add by Amy 2022/01/05 Form2.0 記錄鍵盤傳入順序

Public Sub SetBookInList(pIndex As Integer, pBookInList As String)
   'Add by Amy 2018/09/18
   Dim arrTmp
   Dim ii As Integer, strTemp As String
   'Add by Amy 2020/12/28
   Dim strBookPST15 As String, strBookPST03 As String '登記人員部門
   Dim bolSameDept As Boolean '是同部門
   
   If UBound(m_arrBookInList) < MSHFlexGrid1.Rows Then
      ReDim Preserve m_arrBookInList(MSHFlexGrid1.Rows) As String
      ReDim Preserve m_arrBookPeo(MSHFlexGrid1.Rows) As String 'Add by Amy 2018/09/18
   End If
   m_arrBookInList(pIndex + 1) = pBookInList
   'Add by Amy 2018/09/18 修改時登記人員與新增人員不同部門時需設為已登記(F21不同組視為不同部門)
   If ActionEdit = 1 Then
         arrTmp = Split(pBookInList, ",")
         For ii = 0 To UBound(arrTmp)
            If arrTmp(ii) <> MsgText(601) Then
                'Modify by Amy 2020/12/28 Bug-雅娟建立參加人員時已加杜經理74018,修改時增加非專利處之人員時,也會將杜經理給勾選
'                If (strDeptNo = "F21" And PUB_GetStaffST16(strUserNum) <> PUB_GetStaffST16("" & arrTmp(ii))) Or _
'                (Left(strDeptNo, 2) = "F2" And strDeptNo <> "F21" And strDeptNo <> GetST15("" & arrTmp(ii))) Or _
'                (Left(strDeptNo, 2) <> "F2" And Left(strDeptNo, 1) <> "S" And Left(strDeptNo, 2) <> Left(GetST15("" & arrTmp(ii)), 2)) Or _
'                (Left(strDeptNo, 1) = "S" And Left(GetST15("" & arrTmp(ii)), 1) <> "S") Then
'                    strTemp = strTemp & "," & arrTmp(ii)
'                End If
                bolSameDept = False
                strBookPST15 = GetST15("" & arrTmp(ii))
                strBookPST03 = PUB_GetST03("" & arrTmp(ii))
                'S部門
                If Left(strDeptNo, 1) = "S" And (Left(strBookPST15, 1) = "S" Or Left(strBookPST03, 1) = "S") Then
                    bolSameDept = True
                '外專工程師同組
                ElseIf strDeptNo = "F21" And (strBookPST15 = "F21" Or strBookPST03 = "F21") And PUB_GetStaffST16(strUserNum) = PUB_GetStaffST16("" & arrTmp(ii)) Then
                    bolSameDept = True
                '其他-部門前2碼相同
                ElseIf (Left(strDeptNo, 2) = Left(strBookPST15, 2) Or Left(strDeptNo, 2) = Left(strBookPST03, 2)) And strDeptNo <> "F21" Then
                    bolSameDept = True
                End If
                If bolSameDept = False Then
                    strTemp = strTemp & "," & arrTmp(ii)
                End If
                'end 2020/12/28
            End If
        Next ii
        If strTemp <> MsgText(601) Then
            m_arrBookPeo(pIndex + 1) = Mid(strTemp, 2)
        End If
   End If
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

Private Function GetSaveName(ByVal pFileName As String) As String
   
On Error GoTo ErrHnd

   With CommonDialog1
      .CancelError = True
      .FileName = pFileName
      .Filter = "All Files (*.*)|*.*"
      .InitDir = PUB_Getdesktop
      .MaxFileSize = 3000
      .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
      .ShowSave
      If .FileName <> "" Then
         GetSaveName = .FileName
      End If
   End With
   
   Exit Function
   
ErrHnd:
   If Err.Number <> 32755 Then
      MsgBox Err.Description
   End If
End Function

'Modify by Amy 2022/01/05 原As ListBox->object
Private Sub SetListScroll(oList As Object)
   Dim ii As Integer
   Dim lWnow As Long, lWmax As Long
   
   lWmax = 0
   For ii = 0 To oList.ListCount - 1
      lWnow = TextWidth(oList.List(ii) & " ")
      If lWnow > lWmax Then
         lWmax = lWnow
      End If
   Next
  
   If ScaleMode = vbTwips Then lWmax = lWmax / Screen.TwipsPerPixelX  ' if twips change to pixels
   SendMessageByNum oList.hWnd, LB_SETHORIZONTALEXTENT, lWmax, 0
End Sub

'Modify by Amy 2022/01/05 原:As ListBox->object
Private Function AddListX(oList As Object, stNewItem As String, oList1 As Object, Index As Integer) As Boolean
Dim idx As Integer, bFound As Boolean, stFileName As String
Dim fs, f

'   If InStr(stNewItem, ",") > 0 Then
'      MsgBox "逗號[,]為系統保留字，請重新命名！", vbExclamation
'      cmdAddAtt.SetFocus
'      Exit Function
'   End If
   
   Set fs = CreateObject("Scripting.FileSystemObject")
   Set f = fs.GetFile(stNewItem)
   'Add By Sindy 2017/5/23 檔案大小為 0 KB 有誤
   If f.Size = 0 Then
      ShowMsg stNewItem & MsgText(9221)
      Exit Function
   End If
   '2017/5/23 END
   
   If stNewItem <> "" Then
      For idx = 0 To oList.ListCount - 1
         stFileName = GetFileName(oList.List(idx))
         If GetFileName(stNewItem) = stFileName Then
            MsgBox "附件 " & stFileName & " 已存在！"
            AddListX = False
            bFound = True
            Exit For
         End If
      Next
      
      If bFound = False Then
         For idx = 0 To oList1.ListCount - 1
            stFileName = GetFileName(oList1.List(idx))
            If GetFileName(stNewItem) = stFileName Then
               MsgBox "附件 " & stFileName & " 已存在！"
               AddListX = False
               bFound = True
               Exit For
            End If
         Next
      End If
      If bFound = False Then
         oList.AddItem stNewItem & " (" & Round(f.Size / 1024, 2) & " KB)", 0
         SetListScroll oList
         AddListX = True
      End If
   End If
   
   Set fs = Nothing
   Set f = Nothing
End Function

Private Function GetFileName(ByVal FullPath As String) As String
   Dim stItem As String, iPos As Integer
   
   stItem = FullPath
   iPos = InStr(stItem, "\")
   Do While iPos > 0
      stItem = Mid(stItem, iPos + 1)
      iPos = InStr(stItem, "\")
   Loop
   
   If InStrRev(stItem, " (") > 0 And Right(stItem, 1) = ")" Then
      stItem = Left(stItem, InStrRev(stItem, " (") - 1)
   End If
   
   GetFileName = stItem
End Function

'Add by Amy 2019/01/24
'Modify by Amy 2022/01/05 原: Integer
Private Sub CboEmp_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If cboEmp.ListIndex <> -1 Then Exit Sub
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub CboEmp_LostFocus() '不寫於Validate,因取消會先 run Vaildate,會先檢查並彈訊息
    If cboEmp = MsgText(601) And Not (ActionEdit = 0 Or ActionEdit = 1) Then Exit Sub
    If cboEmp.ListIndex <> -1 Then Exit Sub
    
    If CheckKeyIn(cboEmp) = -1 Then
        Exit Sub
    End If
End Sub

Private Sub cboTime_KeyPress(Index As Integer, KeyAscii As Integer)
    If Not (ActionEdit = 0 Or ActionEdit = 1) Then Exit Sub
    '不可用輸的,因會議室預約只允許每半小時為單位,導致會議室預約畫面看得到無法點兩下看明細
    KeyAscii = 0
End Sub

Private Sub cboTime_LostFocus(Index As Integer) '不寫於Validate,因取消會先 run Vaildate,會先檢查並彈訊息
    If Not (ActionEdit = 0 Or ActionEdit = 1) Then Exit Sub
   
    If CheckKeyIn(cboTime(Index)) = -1 Then
        Exit Sub
    End If
End Sub
'end 2019/01/24

'Add by Amy 2018/09/18
Private Sub Check1_Click()
    If Check1.Value = 1 Then
        '避免沒議題勾選發信造成Error
        If MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 1) = "" Then
            MsgBox "沒有議題不可發信"
            Check1.Value = 0
        End If
    End If
End Sub

Private Sub Check2_Click(Index As Integer)
    If ActionEdit <> 0 And ActionEdit <> 1 Then Exit Sub
    
    'Add by Amy 2020/12/22 「部門全體人員」信箱未未設置彈訊息
    If Check2(0).Value = 1 And strDefDeptMail = MsgText(601) Then
        Check2(0).Value = 0
        MsgBox "「部門全體人員」信箱未設定,請通知電腦中心設置！"
        Exit Sub
    End If
    
    '文雄(S1)勾選「部門全體人員」則「不需登記」需一起勾選
    'Moidfy by Amy 2019/11/12 原:strUserNum = "A4023"
    'Modify by Amy 2020/12/22 開放杜經理操作智權部 原:Left(strDeptNo, 2) = "S1"
    If Left(strDeptNo, 1) = "S" Then
         Check2(1).Value = Check2(0).Value
    End If
    
    If Check2(Index).Value = 1 And Check2(0).Value = 0 Then
        If MsgBox("勾選「不需登記」必需勾選「部門全體人員」", vbYesNo, "確定勾選「部門體人員」") = vbNo Then
            Check2(1).Value = 0
        Else
            Check2(0).Value = 1
        End If
    End If
    
    '未勾選「不需登記」
    If Check2(1).Value = 0 Then
        If InStr(Text1(9), "◎" & strDeptName & "經副理以下人員請至系統內登記是否參加。" & vbCrLf) = 0 Then
            Text1(9) = Text1(9) & "◎" & strDeptName & "經副理以下人員請至系統內登記是否參加。" & vbCrLf
        End If
    '勾選「不需登記」
    Else
        If InStr(Text1(9), "◎" & strDeptName & "經副理以下人員請至系統內登記是否參加。" & vbCrLf) > 0 Then
            Text1(9) = Replace(Text1(9), "◎" & strDeptName & "經副理以下人員請至系統內登記是否參加。" & vbCrLf, "")
        End If
    End If
   
End Sub

Private Sub cmdAdd_Click(Index As Integer)
    Dim strMsg As String, strData As String
    
    If Not (ActionEdit = 0 Or ActionEdit = 1) Then Exit Sub
    '主席
    If Index = 0 Then
        'Modify by Amy 2019/01/24 主席輸入欄改下拉選單
        strData = lblChairMan & "(" & cboEmp & ")"
        If InStr(cboEmp, "Patent") > 0 Then
            strData = "Patent"
        ElseIf InStr(cboEmp, "(") > 0 Then
            strData = cboEmp
        End If
        If AddList(lstUsers, strData, strMsg) = False Then
            MsgBox strMsg, , "警告"
            cboEmp.SetFocus
            Exit Sub
        End If
        cboEmp = "": lblChairMan = ""
        If strMsg <> MsgText(601) Then
            MsgBox strMsg, , "警告"
            lstUsers.SetFocus
            Exit Sub
        End If
        'end 2019/01/24
    '副本
    Else
        If lblCC = MsgText(601) Then
            MsgBox "新增之副本資料有誤請確認！", , "警告"
            cboEmp.SetFocus
            Exit Sub
        End If
        strData = lblCC & "(" & Text1(4) & ")"
        If AddList(lstCC, strData, strMsg) = False Then
            MsgBox strMsg, , "警告"
            Text1(4).SetFocus
            Exit Sub
        End If
        Text1(4) = "": lblCC = ""
        If strMsg <> MsgText(601) Then
            MsgBox strMsg, , "警告"
            lstCC.SetFocus
            Exit Sub
        End If
    End If
End Sub

Private Sub cmdDel_Click(Index As Integer)
    Dim strMsg As String
    
    If Not (ActionEdit = 0 Or ActionEdit = 1) Then Exit Sub
    '主席
    If Index = 0 Then
        Call DelList(lstUsers)
    End If
    '副本
    If Index = 1 Then
        Call DelList(lstCC, strMsg)
        If strMsg <> MsgText(601) Then
            MsgBox strMsg, , "警告"
        End If
    End If
End Sub

'Modify by Amy 2022/01/05 原:As ListBox->object
Private Function AddList(oList As Object, strData As String, ByRef strMsg As String) As Boolean
    Dim bFound As Boolean, strNo As String
    Dim strList As String 'Add by Amy 2020/12/28
    
    If strData = MsgText(601) Then AddList = True: Exit Function
    
    'Modify by Amy 2020/12/28 bug-因加了職稱若抓名稱(員編),導致所有資料已有仍會加入
    If InStr(strData, "Patent") > 0 Then
        strNo = "Patent"
    ElseIf InStr(strData, "(") > 0 Then
        strNo = Mid(strData, Val(InStr(strData, "(")) + 1)
        strNo = Mid(strNo, 1, Val(InStr(strNo, ")")) - 1)
    End If
    For i = 0 To oList.ListCount - 1
        '舊資料
        If InStr(oList.List(i), "(") = 0 And InStr(oList.List(i), "Patent") = 0 Then
            strMsg = strMsg & ";" & oList.List(i) & " 為舊資料請刪除,重新選擇人員"
            bFound = True
        '新資料
        Else
            strList = Mid(oList.List(i), Val(InStr(oList.List(i), "(")) + 1)
            strList = Mid(strList, 1, Val(InStr(strList, ")")) - 1)
            If strList = strNo Then
                strMsg = strMsg & ";" & strData & " 資料已存在"
                bFound = True
            End If
        End If
    Next i
    'end 2020/12/28
    If bFound = False Then
        If InStr(strData, "Patent") > 0 Then
            oList.AddItem strData, 0
        Else
            oList.AddItem GetJobTitle(strNo, 1) & "(" & strNo & ")", 0
        End If
        AddList = True
    End If
    If strMsg <> MsgText(601) Then
        strMsg = Replace(Mid(strMsg, 2), ";", Chr(13) & Chr(10))
    End If
End Function

'Modify by Amy 2022/01/05 原:As ListBox->object
Private Sub DelList(oList As Object, Optional ByRef strMsg As String)
    Dim ii As Integer, bolCanDel As Boolean
    Dim strData As String
    
    If oList.ListCount > 0 Then
        ii = 0
        Do While ii < oList.ListCount
            If oList.Selected(ii) = True Then
                bolCanDel = True
                If UCase(oList.Name) = "LSTCC" Then
                    strData = oList.List(ii)
                    If InStr(strData, "(") > 0 Then
                        If InStr(strCC_Fix, Replace(Mid(strData, InStr(strData, "(") + 1), ")", "")) > 0 Then
                            strMsg = strMsg & Mid(strData, 1, InStr(strData, "(") - 1) & " 為固定收受者不可刪"
                            bolCanDel = False
                        End If
                    End If
                End If
                If bolCanDel = True Then
                    oList.RemoveItem ii
                End If
            End If
            ii = ii + 1
        Loop
    End If
End Sub
'end 2018/09/18

Private Sub cmdAddAtt_Click(Index As Integer)
   Dim stFileName As String
   Dim sFile
   Dim ii As Integer
   
On Error GoTo ErrHnd
   'Modify by Amy 2019/01/24
   stFileName = "*.PDF"
   If Index = 1 Then stFileName = "*.*"
   'end 2019/01/24
   With CommonDialog1
      .CancelError = True
      .FileName = stFileName
      'Modify by Amy 2019/01/24 開放附件只允許pdf
      If Index = 0 Then
        .Filter = "PDF Files (*.PDF)|*.PDF"
      Else
        .Filter = "All Files (*.*)|*.*"
      End If
      .InitDir = PUB_Getdesktop
      .MaxFileSize = 3000
      .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
      .ShowOpen
      If .FileName <> "" Then
         If InStr(.FileName, ChrW$(0)) > 0 Then
            sFile = Split(.FileName, ChrW$(0))
            For ii = 1 To UBound(sFile)
               If InStr(sFile(ii), "\") > 0 Then
                  stFileName = sFile(ii)
               Else
                  stFileName = sFile(0) & "\" & sFile(ii)
               End If
               If Index = 0 Then
                  AddListX lstAtt, stFileName, lstAtt1, Index
               Else
                  AddListX lstAtt1, stFileName, lstAtt, Index
               End If
            Next
         Else
            stFileName = .FileName
            If Index = 0 Then
               AddListX lstAtt, stFileName, lstAtt1, Index
            Else
               AddListX lstAtt1, stFileName, lstAtt, Index
            End If
         End If
      End If
   End With
   Exit Sub
   
ErrHnd:
   If Err.Number <> 0 Then
      MsgBox Err.Description
   End If
End Sub

Private Sub cmdBooinSave_Click()
   MSHFlexGrid2.Visible = False
   CloseTimer
   BooinSave
   MSHFlexGrid2.Visible = True
End Sub

Private Sub BooinSave()
   Dim iCol As Integer, bolInTran As Boolean
   
On Error GoTo ErrHnd
   
   With MSHFlexGrid2
   For iCol = 1 To .Cols - 1
      .row = 0
      .col = iCol
      If .CellForeColor = vbBlue Then
         If .TextMatrix(.Rows - 1, iCol) <> .TextMatrix(.Rows - 2, iCol) Then
            cnnConnection.BeginTrans
            bolInTran = True
            'Modified by Morgan 2012/5/15 改設定參加人員
            'strSql = "delete seminarbookin where sb01=" & Val(Text1(1)) & " and sb02='" & .TextMatrix(.Rows - 3, iCol) & "'"
            'cnnConnection.Execute strSql, intI
            'If .TextMatrix(.Rows - 1, iCol) <> "X" Then
            strSql = "update seminarbookin set sb03='" & .TextMatrix(.Rows - 1, iCol) & "',sb04='" & strUserNum & "',sb05=" & strSrvDate(1) & ",sb06=to_char(sysdate,'HH24MISS')" & _
               " where sb01=" & Val(Text1(1)) & " and sb02='" & .TextMatrix(.Rows - 3, iCol) & "'"
            cnnConnection.Execute strSql, intI
            If intI = 0 Then
               strSql = "insert into seminarbookin(sb01,sb02,sb03,sb04,sb05,sb06)" & _
                  " values(" & Val(Text1(1)) & ",'" & .TextMatrix(.Rows - 3, iCol) & "','" & .TextMatrix(.Rows - 1, iCol) & "'" & _
                  ",'" & strUserNum & "'," & strSrvDate(1) & ",to_char(sysdate,'HH24MISS'))"
               cnnConnection.Execute strSql, intI
            End If
            cnnConnection.CommitTrans
            bolInTran = False
            .TextMatrix(.Rows - 2, iCol) = .TextMatrix(.Rows - 1, iCol)
         End If
         .CellForeColor = .ForeColorFixed
         m_UpdateCount = m_UpdateCount - 1
      End If
   Next
   End With
   Exit Sub
   
ErrHnd:
   If bolInTran Then cnnConnection.RollbackTrans
   MsgBox Err.Description, vbCritical
   
End Sub

Private Function GetAttachFile(ByRef pFileName As String, Optional pSavePath As String) As Boolean
   
   Dim stAttPath As String
   Dim lngSize As Long
   Dim iFileNo As Integer
   Dim bytes() As Byte
   
On Error GoTo ErrHnd
   
   If pSavePath = "" Then
      If Dir(m_AttachPath, vbDirectory) = "" Then
         MkDir m_AttachPath
      End If
      stAttPath = m_AttachPath & "\" & pFileName
      '檔案已存在時不必重新下載
      If Dir(stAttPath) <> "" Then
         'Kill stAttPath
         pFileName = stAttPath
         GetAttachFile = True
         Exit Function
      End If
   Else
      stAttPath = pSavePath
   End If
      
   strExc(0) = "select * from seminarattachment b where sa01=" & Text1(1) & " and sa02='" & ChgSQL(pFileName) & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If Dir(stAttPath) <> "" Then Kill stAttPath
      'Add By Sindy 2017/5/22
      If "" & RsTemp.Fields("sa07") <> "" Then
         GetAttachFile = PUB_GetFtpFile(RsTemp.Fields("sa07"), stAttPath, UCase("seminarattachment"))
      Else
      '2017/5/22 END
         With RsTemp
         lngSize = Val(.Fields("sa03").Value)
         ReDim bytes(lngSize)
         If lngSize > 0 Then bytes() = .Fields("sa04").GetChunk(lngSize)
         End With
         iFileNo = FreeFile
         Open stAttPath For Binary Access Write As #iFileNo
         If lngSize > 0 Then Put #iFileNo, , bytes()
         Close #iFileNo
      End If
      pFileName = stAttPath
      GetAttachFile = True
   End If
   Exit Function
   
ErrHnd:
   MsgBox Err.Description, vbCritical
   If iFileNo > 0 Then Close #iFileNo
End Function

Private Sub CmdMeeting_Click()
    'Add by Amy 2019/12/09
    Dim bolHasRR20 As Boolean '已預約會議室
    Dim stRoom As String, stDate As String, stTimeS As String, stTimeE As String '畫面上資料
    Dim stDBRR(1 To 4) As String '目前DB資料
    Dim stSQL As String, stTmp1 As String
        
    stRoom = GetMeetingRoom(cboRoom, False)
    If Val(stRoom) >= 3 Then Exit Sub
    
    'Moidify by Amy 2020/01/14 從下面搬上來, for 都只進預約第一畫面
    stDate = DBDATE(MaskEdBox1): stTimeS = Format(cboTime(0), "HHmm"): stTimeE = Format(cboTime(1), "HHmm")
    bolHasRR20 = ChkHasRR20(Val(Text1(1)), stDBRR(1), stDBRR(2), stDBRR(3), stDBRR(4), True)
    'end 2020/01/14
    
    '新增/修改
    If ActionEdit = 0 Or ActionEdit = 1 Then
        Text1(23) = "" 'Add by Amy 2019/12/09 按過「會議室預約」鈕都檢查,資料一定要一致,否則不可存檔或取消
        '地點/日期
        'Modify by Amy 2019/11/12 原:DTPicker1(0)
        If CheckKeyIn(MaskEdBox1) = -1 Then
             Exit Sub
        End If
        '時間
        If CheckKeyIn(cboTime(0)) = -1 Then
             Exit Sub
        End If
        If CheckKeyIn(cboTime(1)) = -1 Then
             Exit Sub
        End If
        'Add by Amy 2019/12/09
        '5F/9F中型 有衝突預約,帶預約選取畫面自選時間
        stTmp1 = "Y"
        If ChkReservation(stRoom, stDate, Val(stTimeS), Val(stTimeE), , , , , stTmp1, Text1(1)) = False Then
            MsgBox "該時段會議室已被佔用,自選(刪改)預約範圍！"
            Call Show140112(stRoom, stDate, stTimeS, stTimeE, False)
            Exit Sub
        '沒衝突
        Else
            stSQL = ""
            '未有新增過預約
            If bolHasRR20 = False Then
                stSQL = "Insert into RoomReservation (rr01,rr02,rr03,rr04,rr05,rr07,rr08,rr10,rr11,rr12,rr20) Values " & _
                            "(" & Val(stRoom) & "," & stDate & "," & stTimeS & "," & stTimeE & ",'N'," & _
                            "'" & strUserNum & "'," & CNULL(ChgSQL(Text1(18) & "-" & GetDepartmentName(strDeptNo))) & ",'" & strUserNum & "'," & _
                            strSrvDate(1) & ",to_char(sysdate,'hh24miss')," & Text1(1) & ")"
            '時間與目前DB資料不同,直接存檔後 show 選取畫面
            ElseIf stDBRR(1) & stDBRR(2) & stDBRR(3) & stDBRR(4) <> stRoom & stDate & stTimeS & stTimeE Then
                If stDBRR(1) <> stRoom Then stSQL = stSQL & ",rr01=" & stRoom
                If stDBRR(2) <> stDate Then stSQL = stSQL & ",rr02=" & stDate
                If stDBRR(3) <> stTimeS Then stSQL = stSQL & ",rr03=" & stTimeS
                If stDBRR(4) <> stTimeE Then stSQL = stSQL & ",rr04=" & stTimeE
                stSQL = "Update RoomReservation Set rr13='" & strUserNum & "',rr14=" & strSrvDate(1) & ",rr15=to_char(sysdate,'hh24miss')" & _
                            stSQL & " Where rr20=" & Text1(1)
            End If
            If stSQL <> MsgText(601) Then
                cnnConnection.Execute stSQL
                Call Show140112(stRoom, stDate, stTimeS, stTimeE, True)
                Exit Sub
            End If
        End If
        'end 2019/12/09
    End If
    'Modify by Amy 2020/01/14 都跳預約第一畫面
    If bolHasRR20 = False Then
        MsgBox "無會議室預約資料！"
    Else
        Call Show140112(stRoom, stDate, stTimeS, stTimeE, True)
    End If
End Sub

'開啟附件
Private Sub cmdOpenAtt_Click(Index As Integer)
   Dim hLocalFile As Long
   Dim stFileName As String
   Dim strAtt As String, strType As String
   
   Screen.MousePointer = vbHourglass
   
   If Index = 0 Then
      strAtt = lstAtt.Text
   Else
      strAtt = lstAtt1.Text
   End If
   
   If strAtt = "" Then
      MsgBox "請選擇欲開啟的附件！"
   Else
      stFileName = strAtt
      If InStrRev(stFileName, " (") > 0 Then
         stFileName = Left(stFileName, InStrRev(stFileName, " (") - 1)
      End If
      
      If InStr(stFileName, "\") = 0 Then
         If GetAttachFile(stFileName) = False Then
            Exit Sub
         End If
      End If
      
      ShellExecute hLocalFile, "open", stFileName, vbNullString, vbNullString, 1
   End If
   
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmdPrint_Click()
    Screen.MousePointer = vbHourglass
    runWord
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdQuery_Click(Index As Integer)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
'
'   If txtQueryDate(0).Text = "" And txtQueryDate(1).Text = "" Then
'       MsgBox "請輸入研討會日期範圍!!!", vbExclamation + vbOKOnly
'       txtQueryDate(0).SetFocus
'       Exit Sub
'   End If
     
   Screen.MousePointer = vbHourglass
   Me.grdList.MousePointer = flexHourglass
   'Modify by Amy 2019/01/24 +參加未失效查詢
   If QueryData(Index) = False Then
       strTit = "查詢資料"
       strMsg = "無資料"
       nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   End If
   Me.grdList.MousePointer = flexDefault
   Screen.MousePointer = vbDefault
End Sub

'Modify by Amy 2019/01/24 +intIdx:0-查詢/1-參加記錄(未失效)/2-參加記錄(未失效且未登記,排除無議題資料From_Load用)/99-for FormUnolad
Private Function QueryData(ByVal intIdx As Integer, Optional ByRef stNo As String) As Boolean
   Dim nRow As Integer
   Dim stCon As String, stConSS As String, stConSB As String
   Dim stQ As String
   
   QueryData = False
   Label9.Visible = False 'Add byAmy 2019/01/24
   InitialGridList
   stCon = ""
   stConSS = ""
   stConSB = ""
   
   '多筆查詢
   If intIdx = 0 Then
        If txtQueryDate(0).Text <> "" Then
            stCon = stCon & " And sn05>=" & DBDATE(txtQueryDate(0).Text) & " "
        End If
        If txtQueryDate(1).Text <> "" Then
            stCon = stCon & " And sn05<=" & DBDATE(txtQueryDate(1).Text) & " "
        End If
        
        'Modify by Amy  2019/11/12 原部門抓st03
        'Modify by Amy 2020/11/27 +st03 也需判斷 ex:杜燕文經理為專利處及智權部身份
        If txtDept(0) <> "" And txtDept(1) <> "" Then
             stCon = stCon & " And ( (st15>='" & txtDept(0) & "' And st15<='" & txtDept(1) & "') "
             stCon = stCon & " Or (st03>='" & txtDept(0) & "' And st03<='" & txtDept(1) & "') )"
        ElseIf txtDept(0) <> "" Then
            stCon = stCon & " And (st15>='" & txtDept(0) & "' Or st03>='" & txtDept(0) & "') "
        ElseIf txtDept(1) <> "" Then
             stCon = stCon & " And (st15<='" & txtDept(1) & "' Or st03<='" & txtDept(1) & "') "
        End If
        
        If txtSubject <> "" Then
           stConSS = stConSS & " and instr(upper(ss03),upper('" & ChgSQL(txtSubject) & "'))>0"
        End If
        
        If txtSpeaker <> "" Then
           stConSS = stConSS & " And ss06='" & ChgSQL(txtSpeaker) & "'"
        End If
        
        If txtAttender <> "" Then
           stConSB = stConSB & " and sb02='" & txtAttender & "'"
        End If
        
        If stConSS <> "" Then
           If stConSB <> "" Then
              stCon = stCon & " And exists(select * from SeminarSubject,SeminarBookin where ss01=sn01 and sb01(+)=ss01 and instr(','||sb03||',',','||ss02||',')>0 " & stConSS & stConSB & ")"
           Else
              stCon = stCon & " And exists(select * from SeminarSubject where ss01=sn01 " & stConSS & ")"
           End If
        ElseIf stConSB <> "" Then
           stCon = stCon & " and exists(select * from SeminarBookin where sb01=sn01 and sb03 is not null" & stConSB & ")"
        End If
        stQ = "select '',sn01,sn18,SQLDATET(sn05)" & _
                        " from seminar,Staff where 1=1 And sn12=st01(+)" & stCon & " order by 2 asc"
    '參加記錄(未失效)
    'Modify by Amy 2019/11/12 +intIdx=2(參加記錄-未失效且未登記,排除無議題資料From_Load用)
    ElseIf intIdx = 1 Or intIdx = 2 Then
        Label9.Visible = True
        'Modify by Amy 2020/11/27 ex:杜燕文經理為專利處及智權部身份
        strExc(0) = strDeptSql
        strExc(1) = ""
        
        '抓取個人登記資料
        stQ = "Select '',sn01,sn18,SQLDATET(sn05) From Seminar,SeminarBookin " & _
                 "Where sn01=sb01(+) And sn05>=" & strSrvDate(1) & _
                 " And sb02='" & strUserNum & "' " & IIf(intIdx = 2, "And sb03 is null ", "")
                 
        '參加記錄(未失效)
        If intIdx = 1 Then
            '抓取依建立者部門抓無議題資料(無議題不會有參加人員)
            stQ = stQ & "Union Select '',sn01,sn18,SQLDATET(sn05) From Seminar,SeminarSubject,Staff " & _
                     "Where sn01=ss01(+) And sn05>=" & strSrvDate(1) & " And sn12=st01(+) " & strExc(0)
        '專利處建立之資料,專利處經副理級以下不需登記
        ElseIf intIdx = 2 And (Left(strDeptNo, 2) = "P1" Or Left(Pub_StrUserSt03, 2) = "P1") Then
            strExc(1) = " And sb01 Not In(" & _
                "Select sb01 From Seminar,SeminarBookin,Staff c,Staff a " & _
                 "Where sn01=sb01(+) And sn05>=" & strSrvDate(1) & " And sb02='" & strUserNum & "' And sb03 is null And sb01 is not null " & _
                 "And sn12=c.st01(+) And (SubStr(c.St15,1,2)='P1' or SubStr(c.St03,1,2)='P1'  ) " & _
                 "And sb02=a.st01(+) And (SubStr(a.St15,1,2)='P1' or SubStr(a.St03,1,2)='P1'  ) And a.st20<=44 " & _
                 ")"
        End If
        stQ = stQ & strExc(1) & " Order by sn01 Desc "
        'end 2020/11/27
    'intIdx=99
    Else
        stQ = "Select * From Seminar Where SN01=" & Val(Text1(1))
    End If
   
    intI = 1
    Set RsTemp = ClsLawReadRstMsg(intI, stQ)
    If intI = 1 Then
       If intIdx = 99 Then
            Me.Tag = "" & RsTemp.Fields("SN12") '建立人員
       ElseIf intIdx = 2 Then
            If RsTemp.RecordCount = 1 Then stNo = "" & RsTemp.Fields("SN01")
       Else
            UpdateGridList RsTemp, True
       End If
       QueryData = True
    End If
End Function
'end 2019/01/24

'Modify by Amy 2019/01/24 +bolNotExpired 未失效
Private Sub UpdateGridList(ByRef rsTmp As ADODB.Recordset, Optional ByVal bolNotExpired As Boolean = False)
Dim iRow As Integer, iCol As Integer
Dim RsQ As New ADODB.Recordset, strQ As String, intQ As Integer 'Add by Amy 2019/01/24

   With rsTmp
   .MoveFirst
   grdList.Visible = False
   Do While Not .EOF
      grdList.Rows = grdList.Rows + 1
      iRow = grdList.Rows - 1
      For iCol = 0 To .Fields.Count - 1
         grdList.TextMatrix(iRow, iCol) = "" & .Fields(iCol)
         'Moidfy by Amy 2019/01/24 未失效且未登記顯示＊
         If bolNotExpired = True And iCol = .Fields.Count - 1 Then
            strQ = "Select SB03 From SeminarBookin Where SB01=" & grdList.TextMatrix(iRow, 1) & " And SB02='" & strUserNum & "' And SB03 is null "
            intQ = 1
            Set RsQ = ClsLawReadRstMsg(intQ, strQ)
            If intQ = 1 Then
                grdList.TextMatrix(iRow, iCol) = grdList.TextMatrix(iRow, iCol) & " ＊"
            End If
         End If
      Next
      .MoveNext
   Loop
   grdList.FixedRows = 1 'Added by Lydia 2023/10/16
   grdList.Visible = True
   End With
End Sub
' 初始化列表
Private Sub InitialGridList()
   m_CurrSel = 0
   With grdList
   .Clear
   .Rows = 1
   .Cols = 4
    
   .row = 0
   .col = 0
   .CellAlignment = flexAlignCenterCenter
   .ColWidth(0) = 300
   .ColAlignment(0) = flexAlignCenterCenter
   
   .col = 1
   .Text = "編號"
   .CellAlignment = flexAlignCenterCenter
   .ColWidth(1) = 800
   .ColAlignment(1) = flexAlignCenterCenter
    
   .col = 2
   .Text = "研討會標題"
   .CellAlignment = flexAlignCenterCenter
   .ColWidth(2) = 6045
   .ColAlignment(2) = flexAlignLeftCenter
       
   .col = 3
   .Text = "日期"
   .CellAlignment = flexAlignLeftCenter
   .ColWidth(3) = 1200
   .ColAlignment(3) = flexAlignLeftCenter
   End With
End Sub

'Modify by Amy 2022/01/05 原:As ListBox
Private Function RemoveList(oList As Object) As Boolean
   Dim ii As Integer
   If oList.ListCount > 0 Then
      ii = 0
      Do While ii < oList.ListCount
         If oList.Selected(ii) = True Then
            If oList.ITEMDATA(ii) > 0 Then
'               'Add By Sindy 2017/5/23
'               If m_upFileServer = True Then
                  If MsgBox("確定要永久刪除" & oList.List(ii) & "電子檔？", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbNo Then
                     Screen.MousePointer = vbDefault
                     Exit Function
                  End If
                  '直接從資料庫刪除檔案
                  If PUB_DelFtpFile2(Text1(1), " and sa02='" & GetFileName(oList.List(ii)) & "'", UCase("seminarattachment")) = True Then '檔案改放FTP,必須在DB資料刪除前執行
                     strSql = "delete from seminarattachment where sa01='" & Text1(1) & "' and sa02='" & GetFileName(oList.List(ii)) & "'"
                     cnnConnection.Execute strSql
                     'Call ReadAttachFile
                  End If
'               Else
'               '2017/5/23 END
'                  intI = UBound(m_FilesRemoved) + 1
'                  ReDim Preserve m_FilesRemoved(intI) As String
'                  m_FilesRemoved(intI) = GetFileName(oList.List(ii))
'               End If
            End If
            oList.RemoveItem ii
            SetListScroll oList
            RemoveList = True
            ii = ii - 1
         End If
         ii = ii + 1
      Loop
   End If
End Function

Private Sub cmdRemAtt_Click(Index As Integer)
   If Index = 0 Then
      RemoveList lstAtt
   Else
      RemoveList lstAtt1
   End If
End Sub

Private Sub cmdSaveAtt_Click(Index As Integer)
   
   Dim stFileName As String, stFolderPath As String, stFullName As String
   Dim bMultiFile As Boolean
   Dim ii As Integer, oList 'Modify by Amy 2022/01/05 原:As ListBox
   
   Screen.MousePointer = vbHourglass
   
   If Index = 0 Then
      Set oList = lstAtt
   Else
      Set oList = lstAtt1
   End If
   
   stFileName = ""
   bMultiFile = False
   For ii = 0 To oList.ListCount - 1
      If oList.Selected(ii) Then
         If stFileName <> "" Then
            bMultiFile = True
            Exit For
         Else
            stFileName = oList.Text
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
                  If InStrRev(stFileName, " (") > 0 Then
                     stFileName = Left(stFileName, InStrRev(stFileName, " (") - 1)
                  End If
                  stFullName = stFolderPath & stFileName
                  If stFullName <> "" Then
                     If Dir(stFullName) <> "" Then
                        If MsgBox("檔案[ " & stFileName & " ]已存在是否要覆蓋??", vbYesNo + vbDefaultButton2) = vbNo Then
                           stFullName = ""
                        End If
                     End If
                     If stFullName <> "" Then
                        If GetAttachFile(stFileName, stFullName) = False Then
                           MsgBox "無法儲存檔案[ " & stFileName & " ]！"
                        End If
                     End If
                  End If
               End If
            Next
         End If
      
      Else
         stFileName = Left(stFileName, InStrRev(stFileName, " (") - 1)
         stFullName = GetSaveName(stFileName)
         If stFullName <> "" Then
            If Dir(stFullName) <> "" Then
               If MsgBox("檔案[ " & stFileName & " ]已存在是否要覆蓋??", vbYesNo + vbDefaultButton2) = vbNo Then
                  stFullName = ""
               End If
            End If
            If stFullName <> "" Then
               If GetAttachFile(stFileName, stFullName) = False Then
                  MsgBox "無法儲存檔案[ " & stFileName & " ]！"
               End If
            End If
         End If
      End If
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSelect_Click(Index As Integer)
   Dim ii As Integer, oList 'Modify by Amy 2022/01/05 原:As ListBox
   If Index = 0 Then
      Set oList = lstAtt
   Else
      Set oList = lstAtt1
   End If
   
   For ii = 0 To oList.ListCount - 1
      oList.Selected(ii) = True
   Next
End Sub

Private Sub Command1_Click(Index As Integer)
   Dim aSwap() As String, iCurRow As Integer, ii As Integer
   Dim strSB02 As String, strSubject As String 'Add by Amy 2018/09/18 登記人員/議題
   
   Select Case Index
      Case 0 '人員確認
         Set frm140113_2.fmParent = Me
         frm140113_2.m_stNumList = m_stNumList
         frm140113_2.Caption = frm140113_2.Caption & "(收信人員)"
         'Add by Amy 2018/09/18
         frm140113_2.bolPublic = Option1(0).Value
         'Modify by Amy 2020/11/27 不顯示議題人員list,改顯示寄件者資訊,改寬度
         frm140113_2.Width = 9230
         'Call frm140113_2.SetJoinList2(m_arrBookInList)
         Call frm140113_2.SetMailInfo
         'end 2020/11/27
         frm140113_2.lblMemo = GetMailMemo 'Modify by Amy 2018/09/18 改抓Function
         'Add by Amy 2022/01/07 因改From2.0 且解析每個user不一致,Scrollbar 拉到最下面,也可能無法顯示最後一筆資料,怕無法刪除,故加刪除鈕
         frm140113_2.Command1(6).Visible = False
         frm140113_2.Show vbModal
         
      Case 1 '新增
         With frm140113_1
         .lstSpeaker.Clear 'Add by Amy 2020/12/28
         .m_strMode = "A"
         .m_DeptNo = strDeptNo 'Add byAmy 2019/11/12 目前登入者部門
         .bolPublic = Option1(0).Value 'Add by Amy 2018/09/18
         .m_curRow = MSHFlexGrid1.Rows - 1
         If MSHFlexGrid1.TextMatrix(.m_curRow, 2) <> "" Then
            .DTPicker1(1) = CDate(Format(MSHFlexGrid1.TextMatrix(.m_curRow, 3), "00:00"))
            .DTPicker1(2) = CDate(Format(MSHFlexGrid1.TextMatrix(.m_curRow, 3), "00:00"))
         End If
         .Show vbModal
         .bolPublic = False  'Add by Amy 2018/09/18
         End With
      Case 2 '修改
         'Add by Amy 2021/01/06 bug MSHFlexGrid1.Rows =0,按修改鈕會錯
         If MSHFlexGrid1.Rows = 1 And MSHFlexGrid1.TextMatrix(0, 1) = "" Then Exit Sub
         
         With frm140113_1
         .m_strMode = "E"
         .m_DeptNo = strDeptNo 'Add byAmy 2019/11/12 目前登入者部門
         .bolPublic = Option1(0).Value 'Add by Amy 2018/09/18
         .m_curRow = MSHFlexGrid1.row
         If MSHFlexGrid1.TextMatrix(.m_curRow, 0) <> "" Then
            .Text1 = MSHFlexGrid1.TextMatrix(.m_curRow, 1)
            .DTPicker1(1) = CDate(Format(MSHFlexGrid1.TextMatrix(.m_curRow, 2), "00:00"))
            .DTPicker1(2) = CDate(Format(MSHFlexGrid1.TextMatrix(.m_curRow, 3), "00:00"))
            'Modify by Amy 2020/12/28 主講者改多筆
            '.Text2 = MSHFlexGrid1.TextMatrix(.m_curRow, 4)
            .SetSS06List (SetSS06(MSHFlexGrid1.TextMatrix(.m_curRow, 5), True))
         End If
         .m_stNumList = m_arrBookInList(.m_curRow + 1)
         .Show vbModal
         .bolPublic = False  'Add by Amy 2018/09/18
         End With
      'Memo by Amy 2018/09/18 隱藏上移,避免資料更新錯誤-Morgan
      Case 3 '上移
         iCurRow = MSHFlexGrid1.row
         If iCurRow > 0 Then
            'Added by Morgan 2012/5/3
            strExc(1) = m_arrBookInList(iCurRow + 1)
            m_arrBookInList(iCurRow + 1) = m_arrBookInList(iCurRow)
            m_arrBookInList(iCurRow) = strExc(1)
            'end 2012/5/3
            
            ReDim aSwap(MSHFlexGrid1.Cols - 1) As String
            For ii = 1 To MSHFlexGrid1.Cols - 1
               aSwap(ii) = MSHFlexGrid1.TextMatrix(iCurRow - 1, ii)
            Next
            
            For ii = 1 To MSHFlexGrid1.Cols - 1
               MSHFlexGrid1.TextMatrix(iCurRow - 1, ii) = MSHFlexGrid1.TextMatrix(iCurRow, ii)
            Next
            GridRefresh False, iCurRow - 1
            
            For ii = 1 To MSHFlexGrid1.Cols - 1
               MSHFlexGrid1.TextMatrix(iCurRow, ii) = aSwap(ii)
            Next
            GridRefresh False, iCurRow
            
         End If
      'Memo 2018/09/18 by Amy 隱藏下移,避免資料更新錯誤-Morgan
      Case 4 '下移
         iCurRow = MSHFlexGrid1.row
         If iCurRow < MSHFlexGrid1.Rows - 1 Then
            'Added by Morgan 2012/5/3
            strExc(1) = m_arrBookInList(iCurRow + 1)
            m_arrBookInList(iCurRow + 1) = m_arrBookInList(iCurRow + 2)
            m_arrBookInList(iCurRow + 2) = strExc(1)
            'end 2012/5/3
            ReDim aSwap(MSHFlexGrid1.Cols - 1) As String
            For ii = 1 To MSHFlexGrid1.Cols - 1
               aSwap(ii) = MSHFlexGrid1.TextMatrix(iCurRow + 1, ii)
            Next
            
            For ii = 1 To MSHFlexGrid1.Cols - 1
               MSHFlexGrid1.TextMatrix(iCurRow + 1, ii) = MSHFlexGrid1.TextMatrix(iCurRow, ii)
            Next
            GridRefresh False, iCurRow + 1
            
            For ii = 1 To MSHFlexGrid1.Cols - 1
               MSHFlexGrid1.TextMatrix(iCurRow, ii) = aSwap(ii)
            Next
            GridRefresh False, iCurRow
         End If
      Case 5 '刪除
         'Add by Amy 2020/11/27 避免不小心按到,先彈訊息確認
         If MsgBox("要刪除此議題", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbNo Then
            Exit Sub
         End If
         'Add by Amy 2018/09/18 議題已有人員登記,彈訊息通知
         strSubject = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 0)
         strSB02 = GetSB02(Mid(strSubject, 1, InStr(strSubject, ".") - 1))
         If strSB02 <> MsgText(601) Then
            MsgBox MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 1) & "  議題已有登記人員" & vbCrLf & _
                          "請自行發Mail通知參加人員此議題已刪除"
         End If
         '議題已有人員登記需更新SB03
         If strSB02 <> MsgText(601) Then
            If Del_arrBookinList(UBound(Del_arrBookinList)) <> MsgText(601) Then
               ReDim Preserve Del_arrBookinList(UBound(Del_arrBookinList) + 1) As String
            End If
            '記錄原議題及人員
            Del_arrBookinList(UBound(Del_arrBookinList)) = Mid(strSubject, 1, InStr(strSubject, ".")) & m_arrBookInList(MSHFlexGrid1.row + 1)
         End If
         'end 2018/09/18
         'Added by Morgan 2012/5/3
         'm_arrBookInList(MSHFlexGrid1.row + 1) = ""
         For ii = MSHFlexGrid1.row + 1 To MSHFlexGrid1.Rows - 1
            m_arrBookInList(ii) = m_arrBookInList(ii + 1)
            m_arrBookPeo(ii) = m_arrBookPeo(ii + 1) 'Add by Amy 2018/09/18 刪除時登記人員與新增人員不同部門之人員也需更新
         Next
         'end 2012/5/3
         
         If MSHFlexGrid1.Rows = 1 Then
            MSHFlexGrid1.Clear
         Else
            MSHFlexGrid1.RemoveItem MSHFlexGrid1.row
         End If
         
         ReDim Preserve m_arrBookInList(MSHFlexGrid1.Rows) As String
         ReDim Preserve m_arrBookPeo(MSHFlexGrid1.Rows) As String 'Add by Amy 2018/09/18
   End Select
   
End Sub

'測試信鈕
Private Sub Command2_Click()
   'Add by Amy 2018/09/18 沒有議題發信會有錯
   If MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 1) = "" Then
        MsgBox "沒有議題不可發信"
        Exit Sub
   End If
   'end 2018/09/18
   'Add by Amy 2021/01/22 主席/副本不可有舊資料
   If ChkOldListData(lstUsers, True) = True Then
        lstUsers.SetFocus
        Exit Sub
   End If
   If ChkOldListData(lstCC, True) = True Then
        lstCC.SetFocus
        Exit Sub
   End If
   'end 2021/01/22
   If txtReceiver <> "" Then
      SendInformMail txtReceiver
   End If
End Sub

Private Sub DTPicker1_Change(Index As Integer)
   Dim dDate As Date
   If Index = 0 Then
      'Mark by Amy 2019/11/12 改元件
'      lblWeek = GetWeekDay(DTPicker1(0))
'      'Added by Morgan 2015/3/16 改抓2個工作天--郭
'      'dDate = DateAdd("D", -2, DTPicker1(0))
'      dDate = CDate(Format(CompWorkDay(2, CompDate(2, -1, DBDATE(Format(DTPicker1(0), "YYYYMMDD"))), 1), "@@@@/@@/@@"))
'      'end 2015/3/16
'      If dDate > DTPicker1(3) Then
'         DTPicker1(4) = dDate
'      End If
'      UpdateMemo
   ElseIf Index = 4 Then
      UpdateMemo
   End If
   
End Sub

'Mark by Amy 2019/11/12 改元件
''Memo by Amy 2019/01/24 不寫於Validate,直接輸入時間起時先觸發Validate 但值可能還沒改變會是舊值
''DTPicker1物件無法即時觸發Change,導致輸完日期立即按「確定」資料會是舊的日期,以TBar1_MouseDown使其值變新的
'Private Sub DTPicker1_LostFocus(Index As Integer)
'    Dim stRoom As String
'    If Not (ActionEdit = 0 Or ActionEdit = 1) Then Exit Sub
'
'    If CheckKeyIn(DTPicker1(0)) = -1 Then
'        DTPicker1(0).SetFocus
'        Exit Sub
'    End If
'End Sub

Private Sub UpdateMemo()
   Dim iPos1 As String, iPos2 As String
   Dim stLeft As String, stRight As String, stDate As String
   '2.經副理以下人員請於 3 月27日至系統內登記是否參加。
   iPos1 = InStr(Text1(9), "2.經副理以下人員請")
   iPos2 = InStr(Text1(9), "至系統內登記是否參加。")
   If iPos1 > 0 And iPos2 > iPos1 Then
      stLeft = Left(Text1(9), iPos1 + 9)
      stRight = Mid(Text1(9), iPos2)
      If Left(DTPicker1(4), 4) = Left(strSrvDate(1), 4) Then
         stDate = Format(DTPicker1(4), "M月D日")
      Else
         stDate = Format(DTPicker1(4), "E年M月D日")
      End If
      Text1(9).Text = stLeft & "於" & stDate & "下班前" & stRight
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   'Modify by Amy 2022/01/05 原程式搬到Form_KeyUp,加記錄鍵盤傳入順序
   Call PUB_SaveMeTrackMode(m_MeTrackMode, 1, KeyCode)
   
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    'Add by Amy 2022/01/05 Form2.0 記錄鍵盤傳入順序
    Call PUB_SaveMeTrackMode(m_MeTrackMode, 1, KeyCode)
    'Add by Amy 2022/01/05 從Form_KeyDown搬來
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
        'Case vbKeyF9: Action 11 '確定 'Mark by Amy 2022/01/05取消以ENTER控制為換行的功能
        Case vbKeyF10: Action 12 '取消
        Case vbKeyEscape: Action 14 '結束
    End Select
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Dim stNo As String 'Add byAmy 2019/11/12
    
   MoveFormToCenter Me
   'Modify By Sindy 2021/5/19
   'm_AttachPath = App.path & "\SeminarAttach"
   m_AttachPath = App.path & "\SeminarAttach\" & strUserNum
   
   m_bInsert = IsUserHasRightOfFunction("frm140113", strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction("frm140113", strEdit, False)
   m_bDelete = IsUserHasRightOfFunction("frm140113", strDel, False)
   m_bQuery = IsUserHasRightOfFunction("frm140113", strFind, False)
   m_bOpen = IsUserHasRightOfFunction("frm140113", strPrint, False)
   
   GridInitial
   lblCreateData.BackColor = Me.BackColor
   lblReceiver.BackColor = Me.BackColor
   
   ReDim m_FilesRemoved(0)
   
   ActionEdit = 3
   'Add by Amy 2019/01/24
   SetCboRoom '地點拆成會議室編號及文字
   Label9.Visible = False 'for 未失效鈕用
   SetCombo '時間 原DTPicker元件改與會議室預約元件相同,因若輸非半小時為一單位,導致會議室預約畫面看得到無法點兩下看明細
   'end 2019/01/24
   
   'Modify  by Amy 2019/11/12 原:Pub_StrUserSt03
   strDeptNo = Pub_StrUserSt15
   'Add by Amy 2019/01/24 +部門全體人員及不登記(只有文雄及電腦中心用)
   Check2(0).Visible = False
   Check2(1).Visible = False
   Label8.Visible = False
   'Modify by Amy 2020/12/22 開放杜經理74018操作智權部 原:Left(strDeptNo, 2) = "S1"
   If Left(strDeptNo, 1) = "S" Or strDeptNo = "M51" Then
        'Add by Amy 2019/01/24 部門全體人員(只有文雄負責建立智權部及電腦中心用)
        Check2(0).Visible = True
        If strDeptNo = "M51" Then
            '顯示「不需登記」
            Check2(1).Visible = True
            Label8.Visible = True
        End If
   End If
   'end 2019/11/12
   'Modify by Amy 2019/11/12
   Call FormReset(True)
   Call RsAction(99)
   If m_FirstKEY & m_LastKEY <> MsgText(601) Then
        '登入者若有一筆未失效且未登記之記錄切至登記畫面,若有多筆執行顯示未失效資料
        If QueryData(2, stNo) = True Then
             If RsTemp.RecordCount > 1 Then
                 cmdQuery_Click (1)
             Else
                 Call ReadData(stNo)
                 bolJoin = True
             End If
             TxtLock ActionEdit
             Call SetLimit 'Add by Amy 2019/09/18 原始附件控管權限改至 SetLimit
             CmdMeeting.Enabled = True
        Else
             Action 9 '預設最後一筆
        End If
   '無任何資料
   Else
        TxtLock ActionEdit
        Call SetLimit 'Add by Amy 2018/09/18 原始附件控管權限改至 SetLimit
        CmdMeeting.Enabled = True
   End If
   'end 2019/11/12
   
    'Modify by Amy 2018/09/18 非參加人員登入不需切至「登記」頁籤
    If bolJoin = True Then
        SSTab1.Tab = 1
    Else
        SSTab1.Tab = 0
        MSHFlexGrid2.LeftCol = 1
    End If
    'end 2018/09/18
   CloseTimer
   SetTxtDeptX 'Modify by Amy 2018/09/18 原程式改至SetLimit,加多筆查詢預設部門
   SetCboChoose 'Add by Amy 2020/11/27 單名單選項
   lblName.BackColor = &H8000000F
   'Add by Amy 2022/01/05一開始將ListBox拉到需要的大小,字型會自動放大；所以畫面預設為一列高度,Form_Load才放大到需要的大小
   lstUsers.Height = 600
   lstUsers.Width = 1500
   lstCC.Height = 600
   lstCC.Width = 2150
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
   KillAttach
   Set frm140113 = Nothing
End Sub

Private Sub KillAttach()
On Error Resume Next
   If Dir(m_AttachPath & "\.") <> "" Then
      Kill m_AttachPath & "\*.*"
   End If
End Sub

Private Sub GridInitial()
   Dim ii As Integer
   With MSHFlexGrid1
      .Clear
      .Cols = 6 'Modify by Amy 2020/12/28 加主講人多筆 原:5
      .ColWidth(0) = .Width - 360
      .ColAlignment(0) = flexAlignLeftCenter
      For ii = 1 To .Cols - 1
         .ColWidth(ii) = 0
      Next
   End With
   With MSHFlexGrid2
      .Clear
      .Cols = 2
      .ColWidth(0) = 2000
      .ColAlignmentFixed = flexAlignLeftCenter
      .ColWidth(1) = 0
      .RowHeight(0) = 800
      .FontHeader(0).Size = 11
      .RowHeight(1) = 0 'Added by Morgan 2012/5/3 參加人員各議題分開,皆不參加選項已無存在需要,但只隱藏不刪除以免改變行數影響程式
   End With
End Sub

Private Sub grdList_Click()
   grdList_ShowSelection
End Sub

' 將GridList所選取的列反白, 並將未選取的列設成一般顏色
Private Sub grdList_ShowSelection()
Dim nCurrSel As Integer
Dim nCol As Integer
   
    nCurrSel = grdList.row
    ' 與前一選擇的列位置相同則不處理
    If m_CurrSel = grdList.row Then
        GoTo EXITSUB
    End If
    ' 將原先選取的列回復到正常的顏色
    If m_CurrSel > 0 And m_CurrSel < grdList.Rows Then
        grdList.row = m_CurrSel
        grdList.col = 1
        If grdList.CellBackColor <> &H80000005 Then
            For nCol = 1 To grdList.Cols - 1
                grdList.col = nCol
                If grdList.CellBackColor <> &H80000005 Then: grdList.CellBackColor = &H80000005
                If grdList.CellForeColor <> &H80000008 Then: grdList.CellForeColor = &H80000008
            Next nCol
        End If
        grdList.col = 0
    End If
    ' 設定成所選取的列
    m_CurrSel = nCurrSel
    ' 將所選取的列反白
    If m_CurrSel > 0 And m_CurrSel < grdList.Rows Then
        grdList.row = m_CurrSel
        grdList.col = 1
        For nCol = 1 To grdList.Cols - 1
            grdList.col = nCol
            grdList.CellBackColor = &H8000000D
            grdList.CellForeColor = &H80000005
        Next nCol
        grdList.col = 0
    End If
EXITSUB:
End Sub

Private Sub grdList_DblClick()
   'Add by Amy 2019/11/12 +權限控制
   Dim nRow As Integer
   
   If grdList.row > 0 And grdList.row <= grdList.Rows - 1 Then
        nRow = grdList.row
        If ChkSeminarLimit(grdList.TextMatrix(nRow, 1)) = False Then
            Exit Sub
        End If
   End If
   'end 2019/11/12
   SSTab1.Tab = 0
End Sub

Private Sub grdList_SelChange()
   Dim nRow As Integer
    grdList_ShowSelection

    If grdList.row > 0 And grdList.row <= grdList.Rows - 1 Then
        nRow = grdList.row
        'Add by Amy 2019/11/12 +權限控制
        If ChkSeminarLimit(grdList.TextMatrix(nRow, 1)) = False Then
            MsgBox "無權限查詢！"
            Exit Sub
        End If
        ReadData grdList.TextMatrix(nRow, 1)
    End If
End Sub

Private Sub lstAtt_DblClick()
   If cmdOpenAtt(0).Enabled = True Then
      cmdOpenAtt(0).Value = True
   End If
End Sub

'Add by Amy 2019/11/12
Private Sub MaskEdBox1_Validate(Cancel As Boolean)
    Dim dDate As Date
    If (MaskEdBox1 = MsgText(601) Or MaskEdBox1.Text = "____/__/__") And Not (ActionEdit = 0 Or ActionEdit = 1) Then Exit Sub
    
    If DateCheck(MaskEdBox1.Text) = MsgText(603) Then
        MsgBox "日期" & MsgText(63), , MsgText(5)
        Cancel = True
        MaskEdBox1.SetFocus
        Exit Sub
    End If
    lblWeek = GetWeekDay(CDate(MaskEdBox1))
     
    dDate = CDate(Format(CompWorkDay(2, CompDate(2, -1, DBDATE(MaskEdBox1)), 1), "@@@@/@@/@@"))
    If dDate > DTPicker1(3) Then
        DTPicker1(4) = dDate
    End If
    UpdateMemo
End Sub
'end 2019/11/12

Private Sub MSHFlexGrid1_Click()
   Dim ii As Integer, iNowRow As Integer
   
   With MSHFlexGrid1
   .Visible = False
   iNowRow = .row
   For ii = 0 To .Rows - 1
      .row = ii
      If ii = iNowRow Then
         .CellBackColor = .BackColorSel
         .CellForeColor = vbWhite
      Else
         .CellBackColor = .BackColor
         .CellForeColor = .ForeColor
      End If
   Next
   .row = iNowRow
   .Visible = True
   End With
End Sub

Private Sub MSHFlexGrid2_DblClick()
   Dim ii As Integer, strSB03 As String
   If m_selCol > 1 Then
      If MSHFlexGrid2.CellBackColor <> MSHFlexGrid2.BackColor Then Exit Sub 'Added by Morgan 2012/5/3
      
      If m_bUpdate Or (m_IsOpen And strUserNum = MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 3, m_selCol)) Then
         '全選
         If m_selRow = 0 Then
            MSHFlexGrid2.TextMatrix(1, m_selCol) = ""
            For ii = 2 To MSHFlexGrid2.Rows - 4
               MSHFlexGrid2.TextMatrix(ii, m_selCol) = "V"
               strSB03 = strSB03 & "," & MSHFlexGrid2.TextMatrix(ii, 1)
            Next
            strSB03 = Mid(strSB03, 2)
         '單選
         Else
            If MSHFlexGrid2.TextMatrix(m_selRow, m_selCol) = "V" Then
               MSHFlexGrid2.TextMatrix(m_selRow, m_selCol) = ""
               
               For ii = 2 To MSHFlexGrid2.Rows - 4
                  If MSHFlexGrid2.TextMatrix(ii, m_selCol) = "V" Then
                     strSB03 = strSB03 & "," & MSHFlexGrid2.TextMatrix(ii, 1)
                  End If
               Next
               strSB03 = Mid(strSB03, 2)

               'If strSB03 = "" Then strSB03 = "X"
            Else
               MSHFlexGrid2.TextMatrix(m_selRow, m_selCol) = "V"
               '皆不參加
               If m_selRow = 1 Then
                  For ii = 2 To MSHFlexGrid2.Rows - 4
                     MSHFlexGrid2.TextMatrix(ii, m_selCol) = ""
                  Next
                  strSB03 = "0"
               Else
                  MSHFlexGrid2.TextMatrix(1, m_selCol) = ""
                  For ii = 2 To MSHFlexGrid2.Rows - 4
                     If MSHFlexGrid2.TextMatrix(ii, m_selCol) = "V" Then
                        strSB03 = strSB03 & "," & MSHFlexGrid2.TextMatrix(ii, 1)
                     End If
                  Next
                  strSB03 = Mid(strSB03, 2)
               End If
               
            End If
         End If
         MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 1, m_selCol) = strSB03
         
         '有修改則姓名變藍色
         If strSB03 <> MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 2, m_selCol) Then
            MSHFlexGrid2.row = 0
            MSHFlexGrid2.col = m_selCol
            If MSHFlexGrid2.CellForeColor <> vbBlue Then
               MSHFlexGrid2.CellForeColor = vbBlue
               m_UpdateCount = m_UpdateCount + 1
            End If
         Else
            MSHFlexGrid2.row = 0
            MSHFlexGrid2.col = m_selCol
            If MSHFlexGrid2.CellForeColor = vbBlue Then
               MSHFlexGrid2.CellForeColor = MSHFlexGrid2.ForeColorFixed
               m_UpdateCount = m_UpdateCount - 1
            End If
         End If
         If m_UpdateCount > 0 Then
            If Timer1.Enabled = False Then
               Timer1.Enabled = True
            End If
         Else
            CloseTimer
         End If
      End If
   End If
End Sub

Private Sub MSHFlexGrid2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   m_selRow = 0
   m_selCol = 0
   If Button = 1 And x > MSHFlexGrid2.ColWidth(0) + MSHFlexGrid2.ColWidth(1) Then
      If y < MSHFlexGrid2.RowHeight(0) Then
         m_selRow = 0
         m_selCol = MSHFlexGrid2.col
      Else
         m_selRow = MSHFlexGrid2.row
         m_selCol = MSHFlexGrid2.col
      End If
   End If
End Sub

'Add by Amy 2018/09/18
Private Sub Option1_Click(Index As Integer)
    If ActionEdit > 1 Then Exit Sub
    '新增才預設資料
    If ActionEdit = 0 Then
        FormReset (True)
        FormSet
    End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   If SSTab1.Tab = 1 Then
      If m_bUpdate Or m_IsOpen Then
         cmdBooinSave.Visible = True
      Else
         cmdBooinSave.Visible = False
      End If
   End If
End Sub

Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   Screen.MousePointer = vbHourglass
   Action Button.Index
   Screen.MousePointer = vbDefault
End Sub

'Mark by Amy 2019/11/12 改元件
''Add by Amy 2019/01/24 DTPicker1物件無法即時觸發Change,導致輸完日期立即按「確定」資料會是舊的日期
'Private Sub TBar1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'
'    If Me.ActiveControl.Name = "DTPicker1" Then
'        Text1(18).SetFocus
'    End If
'End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
   If Index = 1 Then
      CloseIme
   Else
      OpenIme
   End If
End Sub

Private Sub Action(Index As Integer)
    
   If TBar1.Buttons(Index).Enabled = False Then Exit Sub
   'Add by Amy 2022/01/05 Form2.0記錄鍵盤傳入順序，判斷是否可執行
   If PUB_ChkMeTrackMode(m_MeTrackMode) = False Then
        Exit Sub
   End If

On Error GoTo ErrHand
   lstUsers.Tag = "" 'Add by Amy 2020/01/14
   
   Select Case Index
      Case 1 '按下新增
         SSTab1.Tab = 0
         ActionEdit = 0
         FormReset
         'Modify by Amy 2018/09/18 預設改至function
'         'Modified by Morgan 2012/3/20 預設當月份--郭雅娟
'         'If Mid(strSrvDate(1), 5, 2) = "12" Then
'         '   Text1(18) = (Val(Left(strSrvDate(1), 4)) + 1) & "年1月份工程師研討會"
'         'Else
'            Text1(18) = Left(strSrvDate(1), 4) & "年" & (Val(Mid(strSrvDate(1), 5, 2))) & "月份工程師研討會"
'         'End If
'
'         Text1(2) = "王副總"                              '2015/1/14 modify by sonia 原為'江總經理'
'         Text1(3) = "專利處工程師、薛經理"                '2015/8/27 modify by sonia 取消何副總
'         Text1(4) = "董事長、桂所長、閻副所長、林副所長"  '2015/1/14 modify by sonia 原為'董事長、所長、副所長'
'         Text1(8) = "台北所五樓會議室及台中、台南、高雄會議室(同步)"
'         Text1(9) = "1.本次研討會書面資料請自行至系統內參看附件。" & vbCrLf & "2.經副理以下人員請至系統內登記是否參加。"
         Option1(1).Value = True
         FormSet
         Text1(1) = Val(GetSaveAutoNo) 'Modify by Amy 2020/01/14 原:ClsPDGetAutoNumber("SS", strNo, True, False)
         'end 2018/09/18
         
      Case 2 '按下修改
         If SSTab1.Tab <> 0 And SSTab1.Tab <> 2 Then
            SSTab1.Tab = 0
         End If
         ActionEdit = 1
         'Add by Amy 2019/01/24
         'Modify by Amy 2019/11/12 改元件
         If Val(strSN13) < Val(教育訓練登錄啟用日) And Val(DBDATE(MaskEdBox1)) < Val(教育訓練登錄啟用日) Then
             CmdMeeting.Enabled = False
         End If
         lstUsers.Tag = ""
         
      Case 3 '按下刪除
         SSTab1.Tab = 0
         If MsgBox("是否確定要刪除??", vbYesNo + vbDefaultButton2) = vbYes Then
            If FormDelete() = False Then
               MsgBox "刪除失敗!", vbCritical
               Exit Sub
            '刪除後移到最末筆
            Else
               Call RsAction(99) 'Add by Amy 2019/11/12 重抓 m_FirstKey/m_LastKey
               FormReset
               RsAction 3
            End If
         Else
            Exit Sub
         End If
         
      Case 4 '按下查詢
         SSTab1.Tab = 0
         FormReset
         ActionEdit = 2
         
      Case 6 '第一筆
         RsAction 0
      Case 7 '前一筆
         RsAction 1
      Case 8 '後一筆
         RsAction 2
      Case 9 '最後筆
         RsAction 3

      Case 11 '按下確定
         Select Case ActionEdit
            Case 0, 1 '新增,修改
               If TxtValidate = False Then
                  Exit Sub
               Else
                  If FormSave() = False Then
                     MsgBox "存檔失敗!", vbCritical
                     Exit Sub
                  Else
                     If ActionEdit = 0 Then Call RsAction(99) 'Add by Amy 2019/11/12 重抓 m_FirstKey/m_LastKey
                     If Check1.Value = 1 Then
                        SendInformMail
                     End If
                  End If
               End If
         End Select
         
         If Text1(1) <> "" Then
            ReadData Text1(1)
         End If
         
      Case 12 '按下取消
        Call ChkRRAndReCover 'Modify by Amy 2020/02/13 原程式改至function
        ActionEdit = 3
        Text1(1) = Text1(1).Tag
        If Text1(1) <> "" Then
            ReadData Text1(1)
        Else
            FormReset
        End If
         
      Case 14 '結束
         If ActionEdit = 0 Or ActionEdit = 1 Then
            If MsgBox("尚未存檔，是否確定要結束??", vbYesNo + vbDefaultButton2) = vbNo Then
               Exit Sub
            'Add by Amy 2020/01/21 還原會議室預約紀錄(避免會議室預約返回後取消)
            Else
                Call ChkRRAndReCover 'Modify by Amy 2020/02/13 原程式改至function
            'end 2020/01/14
            End If
         End If
         Unload Me
         Exit Sub
   End Select
   
   TxtLock ActionEdit
   If ActionEdit <> 0 Then Call SetLimit 'Add by Amy 2018/09/18
   If ActionEdit <> 1 Then CmdMeeting.Enabled = True  'Add by Amy 2019/01/24
   Exit Sub
   
ErrHand:
   ShowMsg "錯誤 : " & Err.Description
End Sub

'Add by Amy 2019/11/12
Private Sub RsAction(ByVal pCmd As Integer)
    Dim rsTmp As New ADODB.Recordset
    Dim strQ As String
    
On Error GoTo ErrHand

    Screen.MousePointer = vbHourglass
    
    'Memo strDeptSql/strSeminar 有修改,要看QueryData是否有誤
    
    '建立者同部門、主席、副本人員或參加人員才可查詢
    'Modify by Amy 2020/11/27 +st03 因杜經理部門兩個部門別
    'Modify by Amy 2020/12/23 開放杜經理可操作智權部,導致杜經理建的P1可看,故改為st15與st03不同才多加st03條件
    If Pub_StrUserSt15 <> Pub_StrUserSt03 Then
        If Left(Pub_StrUserSt03, 1) = "S" Then
            strDeptSql = "Or SubStr(st15,1,1)='" & Left(Pub_StrUserSt03, 1) & "' Or SubStr(st03,1,1)='" & Left(Pub_StrUserSt03, 1) & "' Or SubStr(st03,1,1)='" & Left(strDeptNo, 1) & "' "
        ElseIf Left(strDeptNo, 2) = "F2" Then
            strDeptSql = "Or st03='" & strDeptNo & "' "
        Else
            strDeptSql = "Or SubStr(st15,1,2)='" & Left(Pub_StrUserSt03, 2) & "' Or SubStr(st03,1,2)='" & Left(Pub_StrUserSt03, 2) & "' Or SubStr(st03,1,2)='" & Left(strDeptNo, 2) & "' "
        End If
    End If
    '智權部建的所有智權人員都要可看
    If Left(strDeptNo, 1) = "S" Then
        strDeptSql = " And (SubStr(st15,1,1)='" & Left(strDeptNo, 1) & "'  " & strDeptSql & " ) "
    Else
        strDeptSql = " And (SubStr(st15,1,2)='" & Left(strDeptNo, 2) & "' " & strDeptSql & " ) "
    End If
    If Left(strDeptNo, 2) = "F2" Then
        strDeptSql = " And (st15='" & strDeptNo & "' " & strDeptSql & " ) "
        '外專工程師需抓組別
        If strDeptNo = "F21" Then
            strDeptSql = strDeptSql & " And st16='" & PUB_GetStaffST16(strUserNum) & "' "
        End If
    End If
    'end 2020/12/23
    'end 2020/11/27
    strSeminar = "Select SN01 From Seminar,Staff Where sn12=st01(+) " & strDeptSql & _
      "Union Select SN01 From Seminar Where InStr(sn20,'" & strUserNum & "')>0 Or InStr(sn21,'" & strUserNum & "')>0 " & _
      "Union Select SN01 From Seminar,SeminarBookin Where sn01=sb01(+) And sb02='" & strUserNum & "' "
 
    Select Case pCmd
        Case 0 '第一筆
            strQ = "Select SN01 From Seminar Where SN01=" & m_FirstKEY
        Case 1 '前一筆
            strQ = "Select Nvl(Max(SN01)," & Val(Text1(1)) & ") as SN01 From (" & strSeminar & ") Where SN01<" & Text1(1)
            If strDeptNo = "M51" Then
                strQ = "Select Nvl(Max(SN01)," & Val(Text1(1)) & ") as SN01 From Seminar Where SN01<" & Text1(1)
            End If
        Case 2 '下一筆
            strQ = "Select Nvl(Min(SN01)," & Val(Text1(1)) & ") as SN01 From (" & strSeminar & ") Where SN01>" & Text1(1)
            If strDeptNo = "M51" Then
                strQ = "Select Nvl(Min(SN01)," & Val(Text1(1)) & ") as SN01 From Seminar Where SN01>" & Text1(1)
             End If
        Case 3 '最後筆
            strQ = "Select SN01 From Seminar Where SN01=" & m_LastKEY
        Case 99 '取m_FirstKEY/m_LastKEY
            '取m_FirstKEY
            strQ = "Select Nvl(Min(SN01),'') as SN01 From (" & strSeminar & ") "
            If strDeptNo = "M51" Then
                strQ = "Select Nvl(Min(SN01),'') as SN01 From Seminar "
            End If
            If rsTmp.State <> adStateClosed Then rsTmp.Close
            rsTmp.CursorLocation = adUseClient
            rsTmp.Open strQ, cnnConnection, adOpenStatic, adLockReadOnly
            If rsTmp.RecordCount > 0 Then
                If IsNull(rsTmp.Fields("SN01")) = False Then: m_FirstKEY = rsTmp.Fields("SN01")
            End If
            rsTmp.Close
            
            '取m_LastKEY
            strQ = "Select Nvl(Max(SN01),'') as Sn01 From (" & strSeminar & ") "
            If strDeptNo = "M51" Then
                strQ = "Select Nvl(Max(SN01),'') as Sn01 From Seminar "
            End If
            If rsTmp.State <> adStateClosed Then rsTmp.Close
            rsTmp.CursorLocation = adUseClient
            rsTmp.Open strQ, cnnConnection, adOpenStatic, adLockReadOnly
            If rsTmp.RecordCount > 0 Then
                If IsNull(rsTmp.Fields("SN01")) = False Then: m_LastKEY = rsTmp.Fields("SN01")
            End If
            rsTmp.Close
            Screen.MousePointer = vbDefault
            Exit Sub
    End Select
    
    intI = 1
    If rsTmp.State <> adStateClosed Then rsTmp.Close
    Set rsTmp = ClsLawReadRstMsg(intI, strQ)
    If intI = 1 Then
        If rsTmp.RecordCount > 0 Then
            If (pCmd = 1 Or pCmd = 2) And Val("" & rsTmp.Fields("SN01")) = Val(Text1(1)) Then
                If pCmd = 1 Then
                    DataErrorMessage 6
                Else
                    DataErrorMessage 7
                End If
            Else
                ReadData rsTmp.Fields(0)
            End If
        End If
    End If
   
    Set rsTmp = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub
   
ErrHand:
    Screen.MousePointer = vbDefault
    MsgBox "錯誤 : " & Err.Description, vbCritical
End Sub

Private Sub SetFrame(ByVal pMode As Integer)
   Dim oText, oDTPicker As DTPicker, oCommand As CommandButton  'Modify by Amy 2022/01/05 原:As TextBox
   Dim bValue As Boolean
   Dim oCbo  'Add by Amy 2019/01/24
   
   Select Case pMode
   Case 0, 1 '新增,修改
      bValue = True
   Case 2, 3 '查詢,瀏覽
      bValue = False
   End Select
   
   'Add by Amy 2018/09/18
   Option1(0).Enabled = bValue
   Option1(1).Enabled = bValue
   If pMode = 1 Then
    Option1(0).Enabled = False
    Option1(1).Enabled = False
   End If
   'end 2018/09/18
   'Add by Amy 2019/01/24
   cboEmp.Enabled = bValue
   cboRoom.Enabled = bValue
   'end 2019/01/24
   For Each oText In Text1
      oText.Locked = Not bValue
   Next
   
   For Each oDTPicker In DTPicker1
      oDTPicker.Enabled = bValue
   Next
   
   MaskEdBox1.Enabled = bValue 'Add by Amy 2019/11/12 日期改元件
   
   'Add by Amy 2019/01/24 時間改元件
   For Each oCbo In cboTime
      oCbo.Enabled = bValue
   Next
   
   For Each oCommand In Command1
      oCommand.Enabled = bValue
   Next
   
   Check1.Visible = bValue
   'Add by Amy 2019/01/24
   Check2(0).Enabled = bValue
   Check2(1).Enabled = bValue
   'end 2019/01/24
   lblReceiver.Visible = bValue
   txtReceiver.Visible = bValue
   Command2.Visible = bValue
   
   cmdAddAtt(0).Enabled = bValue
   cmdAddAtt(1).Enabled = bValue
   
   cmdRemAtt(0).Enabled = bValue
   cmdRemAtt(1).Enabled = bValue
   
   cmdOpenAtt(0).Enabled = Not bValue
   cmdSaveAtt(0).Enabled = Not bValue
   
   If m_bOpen Then
      cmdOpenAtt(1).Enabled = Not bValue
      cmdSaveAtt(1).Enabled = Not bValue
   Else
      cmdOpenAtt(1).Enabled = False
      cmdSaveAtt(1).Enabled = False
   End If
   
End Sub

Private Sub TxtLock(ByVal pMode As Integer)
   
   SetFrame pMode
   
   Select Case pMode
   Case 0, 1 '新增,修改
      SSTab1.TabEnabled(1) = False
      SSTab1.TabEnabled(3) = False
      
      Text1(1).Locked = True
      Text1(18).SetFocus
      CmdSitu False
   
   Case 2 '查詢
      SSTab1.TabEnabled(1) = False
      SSTab1.TabEnabled(3) = False
      Text1(1).Locked = False
      Text1(1).SetFocus
      CmdSitu False
     
    Case 3 '瀏覽
      SSTab1.TabEnabled(1) = True
      SSTab1.TabEnabled(3) = True
      CmdSitu True
   End Select
      
End Sub

Private Sub CmdSitu(ByVal TF As Boolean)
   Dim ii As Integer, txt 'Modify by Amy 原:As TextBox
   Dim oButton As Button
 
   For ii = 1 To 4
      TBar1.Buttons(ii).Enabled = False
      TBar1.Buttons(ii + 5).Enabled = False
   Next
   TBar1.Buttons(11).Enabled = False
   TBar1.Buttons(12).Enabled = False
      
   If TF = True Then
      If m_bInsert Then
          TBar1.Buttons(1).Enabled = True
      End If
      If Text1(1) <> "" Then
        'Modify by Amy 2018/09/18 +建立者部門判斷
'         If m_bUpdate Then
'             TBar1.Buttons(2).Enabled = True
'         End If
'         If m_bDelete Then
'             TBar1.Buttons(3).Enabled = True
'         End If
         'Modify by by Amy 2020/12/15 修改刪除專利改回109/11/11教育訓上線前規則其他部門建立者才可操作,原建立者同部門Mark
'        'Modify by Amy 2019/11/12 部門原抓PUB_GetST03
'        'F21外專工程師同組,其他外專人員同部門
'        If (strDeptNo = "F21" And PUB_GetStaffST16(strUserNum) = PUB_GetStaffST16(strSN12)) Or _
'              (Left(strDeptNo, 2) = "F2" And strDeptNo <> "F21" And strDeptNo = GetST15(strSN12)) Then
'            TBar1.Buttons(2).Enabled = True
'            TBar1.Buttons(3).Enabled = True
'        '建立者同部門
'        ElseIf Left(strDeptNo, 2) <> "F2" And (Left(strDeptNo, 2) = Left(GetST15(strSN12), 2) Or strDeptNo = "M51") Then
'            TBar1.Buttons(2).Enabled = True
'            TBar1.Buttons(3).Enabled = True
'        End If
         'end 2018/09/18
         'end 2019/11/12
         If (m_bUpdate = True And Left(strDeptNo, 2) = "P1" And strDeptNo = GetST15(strSN12)) _
           Or (m_bUpdate = True And strSN12 = strUserNum) Or strDeptNo = "M51" Then
            TBar1.Buttons(2).Enabled = True
         End If
         If (m_bDelete = True And Left(strDeptNo, 2) = "P1" And strDeptNo = GetST15(strSN12)) _
           Or (m_bDelete = True And strSN12 = strUserNum) Or strDeptNo = "M51" Then
            TBar1.Buttons(3).Enabled = True
         End If
         'end 2020/12/15
         
         For ii = 1 To 4
            TBar1.Buttons(ii + 5).Enabled = True
         Next
         TBar1.Buttons(4).Enabled = True
      End If
   Else
      TBar1.Buttons(11).Enabled = True
      TBar1.Buttons(12).Enabled = True
   End If
   TBar1.Buttons(14).Enabled = True
End Sub

Private Function ReadData(pKey As String) As Boolean
   Dim iRow As Integer, ii As Integer
   'Add by Amy 2019/01/24
   Dim strSN20 As String, strDefSN20 As String, strSN21 As String, strSN22 As String
   Dim strSpeaker As String 'Add by Amy 2020/12/28
    
On Error GoTo ErrHnd

   KillAttach
   MSHFlexGrid1.Visible = False
   MSHFlexGrid2.Visible = False
   
   m_IsOpen = False
   m_UpdateCount = 0
   CloseTimer
   
   FormReset
   FormSet (True) 'Add 2019/01/24 主席改下拉
   
GoBack:
   strExc(0) = "select A.*,S1.st02 C1,sqldatet(sn13) C2,sn14 C3,S2.st02 C4,sqldatet(sn16) C5,sn17 C6 ,'台北所 '||mr03||' '||Replace(mr02,'五樓','') as mr02 " & _
                    "From Seminar A,staff S1,staff S2,MeetingRoom " & _
                    "Where sn01=" & pKey & " and S1.st01(+)=sn12 and S2.st01(+)=sn15 And sn22=mr01(+)"

   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      '查詢且非電腦中心需判斷是否有權限
      If ActionEdit = 2 And strDeptNo <> "M51" Then
        If ChkSeminarLimit(pKey) = False Then
            MsgBox "無權限查詢！"
            pKey = Text1(1).Tag
            GoTo GoBack
         End If
      End If
      
      With RsTemp
      Text1(1) = .Fields("sn01")
      'Modify by Amy 2019/01/24 教育訓練上線後 sn02(不使用) sn04(改存部門全體人員之Mail Addr ),但上線前仍抓舊欄位
      'Text1(4) = "" & .Fields("sn04")
      strSN13 = "" & .Fields("sn13")
      If Val(strSN13) >= Val(教育訓練登錄啟用日) Then
            strSN20 = "" & .Fields("sn20") '主席
            strSN20 = GetSortNo(strSN20, False)
            Call SetList(lstUsers, strSN20)
            strSN21 = "" & .Fields("sn21") '副本
            strSN21 = GetSortNo(strSN21, True)
            Call SetList(lstCC, strSN21)
            strDefDeptMail = "" & .Fields("sn04")
            If strDefDeptMail <> MsgText(601) Then
                Check2(0).Value = 1 '部門全體人員
            End If
            If ChkSB02All = True Then
                Check2(1).Value = 1 '不需登記
            End If
      Else
            strSN20 = "" & .Fields("sn02") '主席 原:Text1(2)
            Call SetList(lstUsers, strSN20, True)
             strSN21 = "" & .Fields("sn04") '副本
            Call SetList(lstCC, strSN21, True)
      End If
      Text1(3) = "" & .Fields("sn03") '出席
      
      'Modify by Amy 2019/11/12 改元件
      'DTPicker1(0) = CDate(Format("" & .Fields("sn05"), "####/##/##"))
      'lblWeek = GetWeekDay(DTPicker1(0))
      MaskEdBox1.Mask = MsgText(601)
      MaskEdBox1 = Format("" & .Fields("sn05"), "####/##/##")
      MaskEdBox1.Mask = ADFormat
      lblWeek = GetWeekDay(CDate(MaskEdBox1))
      'Modify by Amy 2019/01/24 改元件 原:DTPicker1(1)和(2)
      'DTPicker1(1)= CDate(Format("" & .Fields("sn06"), "00:00"))
      cboTime(0) = Format("" & .Fields("sn06"), "00:00")
      cboTime(1) = Format("" & .Fields("sn07"), "00:00")
      '會議室
      strSN22 = "" & .Fields("mr02")
      cboRoom.Tag = "" & .Fields("sn22")
      cboRoom = strSN22
      '不檢查會議室預約
      Text1(23) = "" & .Fields("sn23")
      'end 2019/01/24
      
      Text1(8) = "" & .Fields("sn08") '地點
      Text1(9) = "" & .Fields("sn09")
      DTPicker1(3) = CDate(Format("" & .Fields("sn10"), "####/##/##"))
      DTPicker1(4) = CDate(Format("" & .Fields("sn11"), "####/##/##"))
      
      Text1(18) = "" & .Fields("sn18")
      lblTitle = Text1(18)
      If strSrvDate(1) >= .Fields("sn10") And strSrvDate(1) <= .Fields("sn11") Then
         m_IsOpen = True
      End If
     
      strSN12 = "" & .Fields("sn12")  '建立者
      If IsNull(.Fields("sn19")) Then
        Option1(1).Value = True '不公開
      Else
        Option1(0).Value = True '公開
      End If
      'end 2018/09/18
      If Left(DTPicker1(4), 4) = Left(strSrvDate(1), 4) Then
         lblMemo = "* 請於 " & Mid(DTPicker1(4), 6) & " 下班前完成「V」註。"
      Else
         lblMemo = "* 請於 " & DTPicker1(4) & " 下班前完成「V」註。"
      End If
     
      lblCreateData = "CREATE : " & .Fields("C1") & " " & _
        " " & .Fields("C2") & " " & _
        " " & Format(.Fields("C3"), "00:00:00") & String(2, " ")  'Modify by Amy 2019/01/24 原 vbCrLf
      If Not IsNull(.Fields("C4")) Then
         lblCreateData = lblCreateData & _
           "UPDATE : " & .Fields("C4") & " " & _
           " " & .Fields("C5") & " " & _
           " " & Format("" & .Fields("C6"), "00:00:00")
      End If
         
      Text1(1).Tag = Text1(1)
      Call SetRoomReservation 'Add by Amy 2019/01/24
      ActionEdit = 3
      End With
   Else
      MsgBox "查無資料！"
      Exit Function
   End If
   
   MSHFlexGrid2.Cols = 2
   MSHFlexGrid2.Rows = 2
   MSHFlexGrid2.TextMatrix(1, 0) = "以下皆不參加"
   
   strExc(0) = "select * from SeminarSubject where ss01=" & pKey & " order by ss02"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      MSHFlexGrid1.Rows = RsTemp.RecordCount
      MSHFlexGrid2.Rows = RsTemp.RecordCount + 2
      With RsTemp
      Do While Not .EOF
         '議題
         iRow = RsTemp.AbsolutePosition - 1
         MSHFlexGrid1.TextMatrix(iRow, 1) = "" & .Fields("ss03")
         MSHFlexGrid1.TextMatrix(iRow, 2) = "" & .Fields("ss04")
         MSHFlexGrid1.TextMatrix(iRow, 3) = "" & .Fields("ss05")
         'Modify by Amy 2020/12/28 主講者多筆
         strSpeaker = "" & .Fields("ss06")
         MSHFlexGrid1.TextMatrix(iRow, 4) = Replace(SetSS06(strSpeaker, False), ";", ",")
         MSHFlexGrid1.TextMatrix(iRow, 5) = strSpeaker
         'end 2020/12/28
         
         ReDim Preserve m_arrBookInList(MSHFlexGrid1.Rows) As String
         ReDim Preserve m_arrBookPeo(MSHFlexGrid1.Rows) As String 'Add by Amy 2018/09/18
         m_arrBookInList(iRow + 1) = GetNumList(.Fields("ss02"))
         
         '登記
         iRow = .AbsolutePosition + 1
         MSHFlexGrid2.TextMatrix(iRow, 1) = "" & .Fields("ss02")
         SetGrid2Subject iRow, "" & .Fields("ss03"), "(" & Format(.Fields("ss04"), "00:00") & "～" & Format(.Fields("ss05"), "00:00") & ")"
         .MoveNext
      Loop
      End With
      GridRefresh True
      
   End If
   MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 4
   MSHFlexGrid2.RowHeight(MSHFlexGrid2.Rows - 1) = 0 '新登記內容
   MSHFlexGrid2.RowHeight(MSHFlexGrid2.Rows - 2) = 0 '原登記內容
   MSHFlexGrid2.RowHeight(MSHFlexGrid2.Rows - 3) = 0 '員工編號
   MSHFlexGrid2.RowHeight(MSHFlexGrid2.Rows - 4) = 0 '所別
   
   Call ReadAttachFile(pKey) 'Add By Sindy 2017/5/23
'   strExc(0) = "select sa02,sa03,sa05 from SeminarAttachment where sa01=" & pKey & " order by 1"
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      With RsTemp
'      Do While Not .EOF
'         If .Fields("sa05") = "1" Then
'            lstAtt1.AddItem .Fields("sa02") & " (" & Round(.Fields("sa03") / 1024, 2) & " KB)", 0
'            lstAtt1.ITEMDATA(0) = 1
'         Else
'            lstAtt.AddItem .Fields("sa02") & " (" & Round(.Fields("sa03") / 1024, 2) & " KB)", 0
'            lstAtt.ITEMDATA(0) = 1
'         End If
'         .MoveNext
'      Loop
'      End With
'   End If
'   If lstAtt.ListCount > 0 Then SetListScroll lstAtt
'   If lstAtt1.ListCount > 0 Then SetListScroll lstAtt1
   
   'Modify by Amy 2020/11/27 原登記人員資料改至CheckinList函數中
   Call CheckInList(pKey)
   MSHFlexGrid1.Visible = True
   MSHFlexGrid2.Visible = True
   MSHFlexGrid2.col = 0
   MSHFlexGrid2.row = 0
   strDeptName = GetSC03 'Add by Amy 2018/09/18
   ReadData = True
   Exit Function
ErrHnd:
   MsgBox Err.Description
End Function

Private Sub ReadAttachFile(pKey As String)
'   m_upFileServer = True 'Add By Sindy 2017/5/23
   'Add By Sindy 2017/5/19 + ,sa07
   strExc(0) = "select sa02,sa03,sa05,sa07 from SeminarAttachment where sa01=" & pKey & " order by 1"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      'Add By Sindy 2017/5/23
      RsTemp.MoveFirst
'      If "" & .Fields("sa07") = "" Then m_upFileServer = False
      '2017/5/23 END
      Do While Not .EOF
         If .Fields("sa05") = "1" Then
            lstAtt1.AddItem .Fields("sa02") & " (" & Round(.Fields("sa03") / 1024, 2) & " KB)", 0
            lstAtt1.ITEMDATA(0) = 1
         Else
            lstAtt.AddItem .Fields("sa02") & " (" & Round(.Fields("sa03") / 1024, 2) & " KB)", 0
            lstAtt.ITEMDATA(0) = 1
         End If
         .MoveNext
      Loop
      End With
   End If
   If lstAtt.ListCount > 0 Then SetListScroll lstAtt
   If lstAtt1.ListCount > 0 Then SetListScroll lstAtt1
End Sub

Private Sub SetGrid2Subject(pRow As Integer, pData1 As String, pData2 As String)
   Dim strData As String, strLeft As String, strRight As String, iLen As Integer
   Dim iRows As Integer
   Dim iMax As Single
   Dim iPos As Integer

   iMax = 20
   
   strData = ""
   strRight = pData1
   iRows = 1
   Do While GetTextLength(strRight) > iMax Or InStr(strRight, vbCrLf) > 0
      strLeft = PUB_StrToStr(strRight, iMax)
      '考慮跳行符號 Added by Morgan 2012/4/23
      iPos = InStr(strLeft, vbCrLf)
      If iPos > 0 Then
         strLeft = Left(strLeft, iPos - 1)
         iLen = Len(strLeft & vbCrLf)
      Else
         iLen = Len(strLeft)
      End If
      strRight = Mid(strRight, iLen + 1)
      strData = strData & strLeft & vbCrLf
      iRows = iRows + 1
   Loop
   strData = strData & strRight & vbCrLf & pData2
   iRows = iRows + 1
   
   MSHFlexGrid2.row = pRow
   MSHFlexGrid2.col = 0
   MSHFlexGrid2.CellFontSize = 9
   MSHFlexGrid2.TextMatrix(pRow, 0) = strData
   MSHFlexGrid2.RowHeight(pRow) = 210 * iRows
   
End Sub

Public Function GridRefresh(Optional pAll As Boolean = True, Optional pRow As Integer) As String
   Dim strData As String, strLeft As String, strRight As String, iLen As Integer
   Dim iRows As Integer, iRow As Integer
   Dim iMax As Single
   Dim ii As Integer, iRowStart As Integer, iRowEnd As Integer
   Dim iPos As Integer

   iMax = 62

   With MSHFlexGrid1
      If pAll = True Then
         iRowStart = 0
         iRowEnd = .Rows - 1
      Else
         iRowStart = pRow
         iRowEnd = pRow
      End If
      
      For ii = iRowStart To iRowEnd
         strData = ""
         strRight = ii + 1 & ".  " & .TextMatrix(ii, 1)
         iRows = 1
         Do While GetTextLength(strRight) > iMax Or InStr(strRight, vbCrLf) > 0
            strLeft = PUB_StrToStr(strRight, iMax)
            '考慮跳行符號 Added by Morgan 2012/4/23
            iPos = InStr(strLeft, vbCrLf)
            If iPos > 0 Then
               strLeft = Left(strLeft, iPos - 1)
               iLen = Len(strLeft & vbCrLf)
            Else
               iLen = Len(strLeft)
            End If
            strRight = "    " & Mid(strRight, iLen + 1)
            strData = strData & strLeft & vbCrLf
            iRows = iRows + 1
         Loop
         
         strLeft = "(" & Format(.TextMatrix(ii, 2), "00:00") & "～" & Format(.TextMatrix(ii, 3), "00:00") & ")"
         If .TextMatrix(ii, 4) <> "" Then
            strLeft = strLeft & "-----" & .TextMatrix(ii, 4)
         End If
         If GetTextLength(strRight & strLeft) > iMax Then
            'Modify by Amy 2020/12/28 +if 因GetTextLength(strLeft) > iMax 會Error
            If GetTextLength(strLeft) > iMax Then
                strData = strData & strRight & vbCrLf & strLeft
            Else
                strData = strData & strRight & vbCrLf & String(iMax - GetTextLength(strLeft), " ") & strLeft
            End If
            'end 2020/12/28
            iRows = iRows + 1
         Else
            strData = strData & strRight & String(iMax - GetTextLength(strRight & strLeft), " ") & strLeft
         End If
         .TextMatrix(ii, 0) = strData
         .RowHeight(ii) = 230 * iRows + 30
      Next
   End With

End Function

Private Function FormDelete() As Boolean
   Dim stSQL As String, bInTrans As Boolean
   
On Error GoTo ErrHandle
   
   cnnConnection.BeginTrans
   bInTrans = True
      
   stSQL = "delete SeminarBookin where SB01='" & Text1(1) & "'"
   cnnConnection.Execute stSQL, intI
   
   stSQL = "delete SeminarSubject where SS01='" & Text1(1) & "'"
   cnnConnection.Execute stSQL, intI
   
   stSQL = "delete Seminar where SN01='" & Text1(1) & "'"
   cnnConnection.Execute stSQL, intI
   
   'Add by Amy 2019/01/24 刪除會議室預約紀錄
   Call ReCoverRR(0)
      
   PUB_DelFtpFile2 Text1(1), , UCase("SEMINARATTACHMENT") 'Add By Sindy 2017/5/23 檔案改放 FTP,必須在DB資料刪除前執行
   stSQL = "delete SeminarAttachment where SA01='" & Text1(1) & "'"
   cnnConnection.Execute stSQL, intI
   
   cnnConnection.CommitTrans
   FormDelete = True
   Text1(1).Text = ""
   Text1(1).Tag = Text1(1)
   
   Exit Function
   
ErrHandle:
   If Err.Number <> 0 Then
      If bInTrans Then cnnConnection.RollbackTrans
      MsgBox Err.Description
   End If
End Function

Private Function FormSave() As Boolean
   Dim bInTrans As Boolean
   Dim stKEY As String
   Dim SN(23) As String 'Modify by Amy 2019/01/24 原:18
   Dim ii As Integer, jj As Integer
   
   Dim adoRst As New ADODB.Recordset
   Dim stFilePath As String
   Dim iFileNo As Integer
   Dim bytes() As Byte
   Dim lngSize As Long '檔案大小
   Dim Numblocks As Integer
   Dim LeftOver As Long
   Const BlockSize = 500000
   Dim arrSB02
   Dim stReName As String
   Dim strFtpPath As String
   'Add by Amy 2018/09/18
   Dim strSS02 As String
   Dim arrSB02C '需勾選
   Dim strTp(1 To 4) As String 'Add by Amy 2019/01/24
   Dim strSS06 As String 'Add by Amy 2020/12/28 演講者多筆
   
On Error GoTo ErrHandle

   cnnConnection.BeginTrans
   bInTrans = True
   
   'Add by Amy 2019/01/24 編號改於按新增時抓自動編號檔
   SN(1) = Val(Text1(1))
   'Modify by Amy 2018/09/18 主席改存SN20,副本改存SN21(原欄位存文字,改存員編)
   'SN(2) = Text1(2) 主席
   SN(3) = Text1(3)
   'SN(4) = Text1(4) 副本
   'Add by Amy 2019/01/24 SN04改為部門全體人員 存Mail Addr
   SN(4) = ""
   If Check2(0).Value = 1 Then
        SN(4) = strDefDeptMail
   End If
   'Modify by Amy 209/11/12 改元件
   'SN(5) = Format(DTPicker1(0), "YYYYMMDD")
   SN(5) = DBDATE(MaskEdBox1)
   'Modify by Amy 2019/01/24 改元件 原:.DTPicker1(1)和(2)
   SN(6) = Format(cboTime(0), "HHmm")
   SN(7) = Format(cboTime(1), "HHmm")
   'end 2019/01/24
   SN(8) = Text1(8)
   SN(9) = Text1(9)
   SN(10) = Format(DTPicker1(3), "YYYYMMDD")
   SN(11) = Format(DTPicker1(4), "YYYYMMDD")
   'Modify by Amy 2019/01/24由下往上搬
   SN(15) = strUserNum
   SN(16) = strSrvDate(1)
   SN(17) = "TO_CHAR(SYSDATE,'hh24miss')"
   'end 2019/01/24
   SN(18) = Text1(18)
   SN(19) = IIf(Option1(0).Value = True, "Y", "")
   SN(20) = "Y"
   Call ChkOldListData(lstUsers, False, SN(20)) '主席(員編)
   SN(21) = "Y"
   Call ChkOldListData(lstCC, False, SN(21)) '副本(員編)
   SN(22) = GetMeetingRoom(cboRoom, False) '會議室編號(會議室預約紀錄沒有的 9F小會議室)
   SN(23) = Text1(23) '不檢查會議室預約,因週期性預約會先於預約會議室Booking,之後才輸此,無法判斷是否已預約
      
   '新增
   'Modify by Amy 2019/01/24 編號改抓自動編號檔
   'If Text1(1) = "" Then
   If ActionEdit = 0 Then
      SN(12) = strUserNum
      SN(13) = strSrvDate(1)
      SN(14) = "TO_CHAR(SYSDATE,'hh24miss')"
      
     'Modify by Amy 2018/09/18 主席改存SN20(原SN02),副本改存SN21(原SN04),加SN19/SN23
     'Moidfy by Amy 2019/01/24 SN01改抓自動編號檔,SN04改存,加SN22
'     strSql = "insert into Seminar(SN01,SN20,SN03,SN21,SN04,SN05,SN06,SN07,SN08,SN09,SN10,SN11,SN12,SN13,SN14,SN18,SN19) " & _
'         " select nvl(max(SN01),0)+1,'" & ChgSQL(SN(20)) & "','" & ChgSQL(SN(3)) & "'" & _
'         " ,'" & ChgSQL(SN(21)) & "' ," & CNULL(ChgSQL(SN(4))) & "," & SN(5) & "," & SN(6) & "," & SN(7) & _
'         " ,'" & ChgSQL(SN(8)) & "','" & SN(9) & "'," & SN(10) & "," & SN(11) & ",'" & SN(12) & "'," & SN(13) & "," & SN(14) & _
'         " ,'" & ChgSQL(SN(18)) & "'," & CNULL(ChgSQL(SN(19))) & " From Seminar"
'      strExc(0) = "select max(SN01) from seminar"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         SN(1) = RsTemp(0)
'      Else
'         GoTo ErrHandle
'      End If
      
      strSql = "insert into Seminar(SN01,SN20,SN03,SN21,SN04,SN05,SN06,SN07,SN08,SN09,SN10,SN11,SN12,SN13,SN14,SN18,SN19,SN22,SN23) " & _
         "Values( " & SN(1) & ",'" & ChgSQL(SN(20)) & "','" & ChgSQL(SN(3)) & "'" & _
         " ,'" & ChgSQL(SN(21)) & "' ," & CNULL(ChgSQL(SN(4))) & "," & SN(5) & "," & SN(6) & "," & SN(7) & _
         " ,'" & ChgSQL(SN(8)) & "','" & SN(9) & "'," & SN(10) & "," & SN(11) & ",'" & SN(12) & "'," & SN(13) & "," & SN(14) & _
         " ,'" & ChgSQL(SN(18)) & "'," & CNULL(ChgSQL(SN(19))) & "," & Val(SN(22)) & "," & CNULL(ChgSQL(SN(23))) & ")"
      cnnConnection.Execute strSql, intI
   '修改
   Else
      'Mark by Amy 2019/01/24 往上搬 SN(1) /SN(15~18)
      
      'Modify by Amy 2018/09/18 主席改存SN20(原SN02),副本改存SN21(原SN04),加SN19
      'Moidfy by Amy 2019/01/24 +SN22 會議室編號
      strSql = "update Seminar set SN20='" & ChgSQL(SN(20)) & "',SN03='" & ChgSQL(SN(3)) & "',SN04=" & CNULL(ChgSQL(SN(4))) & "" & _
         ",SN21='" & ChgSQL(SN(21)) & "',SN05=" & SN(5) & ",SN06=" & SN(6) & ",SN07=" & SN(7) & _
         ",SN08='" & ChgSQL(SN(8)) & "',SN09='" & ChgSQL(SN(9)) & "',SN10=" & SN(10) & ",SN11=" & SN(11) & _
         ",SN15='" & SN(15) & "',SN16=" & SN(16) & ",SN17=" & SN(17) & ",SN18='" & ChgSQL(SN(18)) & "'" & _
         ",SN19=" & CNULL(ChgSQL(SN(19))) & ",SN22=" & SN(22) & ",SN23=" & CNULL(ChgSQL(SN(23))) & _
         " where SN01=" & SN(1)
      cnnConnection.Execute strSql, intI
   End If
   '原預約5F或9F換小會議室需刪除5F或9F之預約(跳會議室預約畫面會先存檔)
   '設定不預約5F改9F或9F改5F
   If ChkHasRR20(Val(Text1(1)), strTp(1), strTp(2), strTp(3), strTp(4), True) = True Then
    If (Val(SN(22)) >= 3 And Val(strTp(1)) < 3) _
       Or (Text1(23) = "Y" And Val(SN(22)) < 3 And Val(strTp(1)) < 3 And SN(22) & SN(5) & SN(6) & SN(7) <> strTp(1) & strTp(2) & strTp(3) & strTp(4)) Then
        Call ReCoverRR(0)
    End If
   End If
   
   'Add by Amy 2019/01/24
   '勾「不登記」,需於SeminarBookin寫入一筆記錄
   If Check2(1).Value = 1 And ChkSB02All = False Then
        strSql = "Insert Into SeminarBookin (SB01,SB02,SB03,SB04,SB05,SB06,SB07) Values(" & SN(1) & ",'ALL999'" & ",'ALL','" & SN(15) & "'," & SN(16) & "," & SN(17) & ",'ALL')"
        cnnConnection.Execute strSql, intI
   '取消勾選「不登記」,有記錄要刪除
   ElseIf Check2(1).Value = 0 And ChkSB02All = True Then
        strSql = "Delete SeminarBookin Where SB01=" & SN(1) & " And SB02='ALL999' "
        cnnConnection.Execute strSql, intI
   End If
   'Add by Amy 2018/09/18 刪除已有登記人員之議題,將議題編號(SB03)取代為空
   For ii = LBound(Del_arrBookinList) To UBound(Del_arrBookinList)
        If Del_arrBookinList(ii) <> MsgText(601) Then
            strSS02 = Mid(Del_arrBookinList(ii), 1, InStr(Del_arrBookinList(ii), ".") - 1) '議題編號
            strSql = Replace(Del_arrBookinList(ii), strSS02 & ".", "")
            strSql = Replace(Mid(strSql, 1, IIf(Right(strSql, 1) = ",", Len(strSql) - 1, Len(strSql))), ",", "','")
            strSql = "Update seminarbookin Set SB03=SubStr(Replace(','||sb03,'," & strSS02 & "',''),2) " & _
                        "Where sb01=" & SN(1) & " And sb02 in ('" & strSql & "')"
            cnnConnection.Execute strSql, intI
        End If
   Next ii
   'end 2018/09/18
   strSql = "delete SeminarSubject where ss01=" & SN(1)
   cnnConnection.Execute strSql, intI
   
   With MSHFlexGrid1
   For ii = 0 To .Rows - 1
      If .TextMatrix(ii, 0) <> "" Then
        'Modify by Amy 2018/09/18 將已變更議題編號之已登記人員之議題編號(SB03)取代新議題編號 ex:刪除議題2,議題3改為2 後,原SB03=3需改為2
        strSS02 = Mid(.TextMatrix(ii, 0), 1, InStr(.TextMatrix(ii, 0), ".") - 1)
        If ii <> Val(strSS02) Then
            strSql = Replace(Mid(m_arrBookInList(Val(ii + 1)), 1, IIf(Right(m_arrBookInList(Val(ii + 1)), 1) = ",", Len(m_arrBookInList(Val(ii + 1))) - 1, Len(m_arrBookInList(Val(ii + 1))))), ",", "','")
            strSql = "Update seminarbookin Set SB03=SubStr(Replace(','||sb03,'," & strSS02 & "','," & ii + 1 & "' ),2) " & _
                        "Where sb01=" & SN(1) & " And sb02 in ('" & strSql & "')"
            cnnConnection.Execute strSql, intI
        End If
        'end 2018/09/18
        'Modify by Amy 2020/12/28 ss06改存多筆,員工存員編 原:.TextMatrix(ii, 4)
         strSS06 = .TextMatrix(ii, 5)
         strSql = "insert into SeminarSubject(ss01,ss02,ss03,ss04,ss05,ss06) values" & _
          " (" & SN(1) & "," & (ii + 1) & ",'" & ChgSQL(.TextMatrix(ii, 1)) & "'," & Val(.TextMatrix(ii, 2)) & _
          "," & Val(.TextMatrix(ii, 3)) & ",'" & ChgSQL(strSS06) & "')"
          'end 2020/12/28
         cnnConnection.Execute strSql, intI
      End If
   Next
   End With
   
   For ii = 1 To UBound(m_FilesRemoved)
      strSql = "delete SeminarAttachment where sa01=" & SN(1) & " and sa02='" & ChgSQL(m_FilesRemoved(ii)) & "'"
      cnnConnection.Execute strSql, intI
   Next
   
   '修改設定登記清單
   'Modify by Amy 2019/01/24 編號改於按新增時抓自動編號檔
   'If Text1(1) <> "" Then
   If ActionEdit = 1 Then
      '刪除無登記的資料
      strSql = "delete seminarbookin where sb01=" & Text1(1) & " and sb03 is null"
      cnnConnection.Execute strSql, intI
      '清除已登記的可登記議題欄位(sb07)
      strSql = "update seminarbookin set sb07=null where sb01=" & Text1(1) & " And sb07<>'ALL'"
      cnnConnection.Execute strSql, intI
   End If
   For jj = 1 To UBound(m_arrBookInList)
      If m_arrBookInList(jj) <> "" Then
         arrSB02 = Split(m_arrBookInList(jj), ",")
         For ii = LBound(arrSB02) To UBound(arrSB02)
            If arrSB02(ii) <> "" Then
               strSql = "update seminarbookin set sb07=sb07||decode(sb07,null,'',',')||'" & jj & "' where sb01=" & SN(1) & " and sb02='" & arrSB02(ii) & "'"
               cnnConnection.Execute strSql, intI
               If intI = 0 Then
                  strSql = "insert into seminarbookin(sb01,sb02,sb07) values (" & SN(1) & ",'" & arrSB02(ii) & "','" & jj & "')"
                  cnnConnection.Execute strSql, intI
               End If
            End If
         Next ii
         'Add by Amy 2018/09/18 +修改時且登記人員與新增人員不同部門時直接設為已登記
         arrSB02C = Split(m_arrBookPeo(jj), ",")
         For ii = LBound(arrSB02C) To UBound(arrSB02C)
            If arrSB02C(ii) <> MsgText(601) Then
                strSql = "Update seminarbookin set sb03=sb03||decode(sb03,null,'',',')||'" & jj & "' Where sb01=" & SN(1) & " and sb02='" & arrSB02C(ii) & "'"
                cnnConnection.Execute strSql, intI
            End If
         Next ii
         'end 2018/09/18
      End If
   Next
   'Add by Amy 2018/09/18 刪除已刪除議題之登記資料
   If Text1(1) <> "" Then
      strSql = "Delete seminarbookin where sb01=" & Text1(1) & " And sb07 is null"
      cnnConnection.Execute strSql, intI
   End If
   'end 2018/09/18
   
   '開放附件
   For ii = 0 To lstAtt.ListCount - 1
      If lstAtt.ITEMDATA(ii) = 0 Then
         stFilePath = lstAtt.List(ii)
         stFilePath = Left(stFilePath, InStrRev(stFilePath, " (") - 1)
         'Modified by Morgan 2017/6/3 跑執行檔時附件會被鎖住無法刪除,改寫法
         'If iFileNo > 0 Then Close #iFileNo
         'iFileNo = FreeFile
         'Open stFilePath For Binary Access Read As #iFileNo
         'lngSize = LOF(iFileNo)
         'Close #iFileNo
         lngSize = FileLen(stFilePath)
         'end 2017/6/3
         
         'Add By Sindy 2017/5/23
         '上傳FTP File Server
'         If m_upFileServer = True Then
            stReName = SN(1) & ".2." & lngSize & "." & GetFileName(stFilePath)
            PUB_PutFtpFile stFilePath, SN(1), stReName, strFtpPath, "SEMINARATTACHMENT"
            If strFtpPath <> "" Then
               strSql = "insert into seminarattachment(sa01,sa02,sa03,sa05,sa07) " & _
                        "values(" & CNULL(SN(1)) & "," & CNULL(GetFileName(stFilePath)) & _
                        "," & lngSize & ",'2'," & CNULL(strFtpPath) & ")"
               cnnConnection.Execute strSql
            End If
            'Call PUB_DelPCOrgFile(stFilePath) '一併將PC上的實體檔案刪除
'         Else
'         '2017/5/23 END
'            With adoRst
'            If adoRst.State = adStateClosed Then
'               strExc(0) = "select * from SeminarAttachment where rownum<1"
'               .CursorLocation = adUseClient
'               .Open strExc(0), cnnConnection, adOpenStatic, adLockOptimistic
'            End If
'
'            .AddNew
'            .Fields("sa01").Value = SN(1)
'            .Fields("sa02").Value = GetFileName(stFilePath)
'            .Fields("sa05").Value = "2"
'            .Fields("sa03").Value = lngSize
'            Numblocks = lngSize / BlockSize
'            LeftOver = lngSize Mod BlockSize
'
'            ReDim bytes(LeftOver)
'            Get #iFileNo, , bytes()
'            .Fields("sa04").AppendChunk bytes()
'
'            ReDim bytes(BlockSize)
'            For jj = 1 To Numblocks
'                Get #iFileNo, , bytes()
'                .Fields("sa04").AppendChunk bytes()
'            Next jj
'            Close #iFileNo
'            .UPDATE
'            End With
'         End If
      End If
   Next ii
   '原始附件
   For ii = 0 To lstAtt1.ListCount - 1
      If lstAtt1.ITEMDATA(ii) = 0 Then
         stFilePath = lstAtt1.List(ii)
         stFilePath = Left(stFilePath, InStrRev(stFilePath, " (") - 1)
         
         'Modified by Morgan 2017/6/3 跑執行檔時附件會被鎖住無法刪除,改寫法
         'If iFileNo > 0 Then Close #iFileNo
         'iFileNo = FreeFile
         'Open stFilePath For Binary Access Read As #iFileNo
         'lngSize = LOF(iFileNo)
         'Close #iFileNo
         lngSize = FileLen(stFilePath)
         'end 2017/6/3
         
         'Add By Sindy 2017/5/23
         '上傳FTP File Server
'         If m_upFileServer = True Then
            stReName = SN(1) & ".1." & lngSize & "." & GetFileName(stFilePath)
            PUB_PutFtpFile stFilePath, SN(1), stReName, strFtpPath, "SEMINARATTACHMENT"
            If strFtpPath <> "" Then
               strSql = "insert into seminarattachment(sa01,sa02,sa03,sa05,sa07) " & _
                        "values(" & CNULL(SN(1)) & "," & CNULL(GetFileName(stFilePath)) & _
                        "," & lngSize & ",'1'," & CNULL(strFtpPath) & ")"
               cnnConnection.Execute strSql
            End If
            'Call PUB_DelPCOrgFile(stFilePath) '一併將PC上的實體檔案刪除
'         Else
'         '2017/5/23 END
'            With adoRst
'            If adoRst.State = adStateClosed Then
'               strExc(0) = "select * from SeminarAttachment where rownum<1"
'               .CursorLocation = adUseClient
'               .Open strExc(0), cnnConnection, adOpenStatic, adLockOptimistic
'            End If
'
'            .AddNew
'            .Fields("sa01").Value = SN(1)
'            .Fields("sa02").Value = GetFileName(stFilePath)
'            .Fields("sa05").Value = "1"
'            .Fields("sa03").Value = lngSize
'            Numblocks = lngSize / BlockSize
'            LeftOver = lngSize Mod BlockSize
'
'            ReDim bytes(LeftOver)
'            Get #iFileNo, , bytes()
'            .Fields("sa04").AppendChunk bytes()
'
'            ReDim bytes(BlockSize)
'            For jj = 1 To Numblocks
'                Get #iFileNo, , bytes()
'                .Fields("sa04").AppendChunk bytes()
'            Next jj
'            Close #iFileNo
'            .UPDATE
'            End With
'         End If
      End If
   Next ii
   If adoRst.State <> adStateClosed Then adoRst.Close
   
   cnnConnection.CommitTrans
   FormSave = True
   Text1(1) = SN(1)
   Set adoRst = Nothing
   Exit Function

ErrHandle:
   If bInTrans = True Then cnnConnection.RollbackTrans
   If Err.Number <> 0 Then MsgBox Err.Description
   Set adoRst = Nothing
End Function

Private Function TxtValidate() As Boolean
   Dim stTP(0 To 4) As String, stMsg As String 'Add by Amy 2019/01/24
   Dim bolHasRR20 As Boolean 'Add by Amy 2019/12/11
   'Add by Amy 2020/01/14
   Dim m_ChooseSC04 As Variant, stDeptArea As String 'Msgbox訊息回應/部門主管+(編號)
    
    'Add by Amy 2022/01/05檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
    If PUB_ChkUniText(Me, , True) = False Then
        Exit Function
    End If

   'Modify by Amy 2018/09/18 +沒議題不可勾收信 會error,主席、副本不可有舊資料
   If Check1.Value = 1 Then
      If MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 1) = "" Then
        MsgBox "沒有議題不可發信！", vbInformation
        Exit Function
      End If
      
      'Modify by Amy 2020/12/22 開放杜經理操作智權部 原:Left(strDeptNo, 2) = "S1"
      If UBound(m_arrBookInList) = 1 And Left(strDeptNo, 1) = "S" And m_arrBookInList(1) = MsgText(601) And Check2(0).Value = 1 Then
        '有議題沒參加人員且為智權部且勾選「部門全體人員」
        stTP(0) = strDefDeptMail
      Else
        For i = LBound(m_arrBookInList) To UBound(m_arrBookInList)
            stTP(0) = stTP(0) & m_arrBookInList(i)
        Next i
      End If
      If stTP(0) = MsgText(601) Then
          MsgBox "請選擇收信人員！", vbInformation
          Exit Function
      End If
     
   End If
   '主席
   If lstUsers.ListCount = 0 Then
        MsgBox "主席不可為空", vbInformation
        cboEmp.SetFocus 'Modify by Amy 2019/01/24 改為下拉選單
        Exit Function
   End If
   If ChkOldListData(lstUsers, True) = True Then
        lstUsers.SetFocus
        Exit Function
   End If
   '副本
   If ChkOldListData(lstCC, True) = True Then
        lstCC.SetFocus
        Exit Function
   End If
   '日期
   'Modify by Amy 2019/11/12 改元件
   If CheckKeyIn(MaskEdBox1) = -1 Then
        Exit Function
   End If
   '時間
   If CheckKeyIn(cboTime(0)) = -1 Then
        Exit Function
   End If
   If CheckKeyIn(cboTime(1)) = -1 Then
        Exit Function
   End If
   
   '已有會議室預約(5F/9F)需檢查資料是否一致,有可能週期性先預約(不需按「會議室預約」鈕),若不檢查為Y則不檢查
   stTP(0) = GetMeetingRoom(cboRoom, False)
   If Val(stTP(0)) < 3 Then
        'Modify by Amy 2019/12/09
        bolHasRR20 = ChkHasRR20(Val(Text1(1)), stTP(1), stTP(2), stTP(3), stTP(4), True)
        If bolHasRR20 = True Then
            If Val(stTP(1)) & Val(stTP(2)) & Val(stTP(3)) & Val(stTP(4)) <> stTP(0) & Val(DBDATE(MaskEdBox1)) & Val(Replace(cboTime(0), ":", "")) & Val(Replace(cboTime(1), ":", "")) Then
                stMsg = "原預約" & GetMeetingRoom(stTP(1), True) & vbCrLf & _
                            "日期：" & ChangeWStringToWDateString(stTP(2)) & vbCrLf & _
                            "時間：" & stTP(3) & "~" & stTP(4) & vbCrLf & vbCrLf & _
                            "與會議室預約資料不一致請確認！"
                MsgBox stMsg
                Exit Function
             End If
        '按過會議室鈕且刪除了預約資料
        ElseIf Text1(23) = MsgText(601) And bolHasRR20 = False Then
            stMsg = "未預約會議室,不可存檔！"
            MsgBox stMsg
            Exit Function
        '無會議室預約
        ElseIf Text1(23) = "Y" Then
            stMsg = "未預約會議室,請另行處理！"
            MsgBox stMsg
        End If
   End If
   'Add by Amy 2020/01/14 W部門主管未列於主席無法Mail通知(放於最後判斷)
   If Left(strDeptNo, 1) = "W" And (ActionEdit = 0 Or (ActionEdit = 1 And DBDATE(MaskEdBox1) >= strSrvDate(1))) Then
        If ChkDefSC04Exists() = False Then
            stMsg = ""
            m_ChooseSC04 = MsgBox("部門主管不在名單內，是否加入副本？ " & vbCrLf & _
                                                "(是:加入，否:不加入存檔，取消:不存檔)", vbInformation + vbYesNoCancel)
            '是-加入
            If m_ChooseSC04 = 6 Then
                stDeptArea = GetDeptMan(strDeptNo)
                stDeptArea = StaffQuery(stDeptArea) & "(" & stDeptArea & ")"
                Call AddList(lstCC, stDeptArea, stMsg)
            '否-不加入
            ElseIf m_ChooseSC04 = 7 Then
            '不存檔
            Else
                Exit Function
            End If
        End If
   End If
   
   TxtValidate = True
End Function

Private Sub FormReset(Optional ByVal bolKeyKeep As Boolean = False)
    Dim oText, oDTPicker As DTPicker  'Modify by Amy 2021/01/05 原As TextBox
    Dim stTP As String
    
    For Each oText In Text1
        'Modify by Amy 2019/01/24 保留key不清空(點選公開/不公開時)
        If bolKeyKeep = True And oText.Index = 1 Then
        Else
            oText.Text = ""
        End If
    Next
    
    'Add by Amy 2019/11/12 日期預帶
    MaskEdBox1.Mask = MsgText(601)
    MaskEdBox1 = Format(strSrvDate(1), "####/##/##")
    MaskEdBox1.Mask = ADFormat
    'end 2019/11/12
    For Each oDTPicker In DTPicker1
        oDTPicker = Now
    Next
    'Modify by Amy 2019/01/24 改元件 原:DTPicker1(1)和(2)
'    DTPicker1(1).Hour = 10
'    DTPicker1(1).Minute = 0
'    DTPicker1(2).Hour = 17
'    DTPicker1(2).Minute = 30
     'Moidfy by Amy 2019/11/12 原:"10:00" 改抓目前時間
     stTP = Format(Now, "HH:mm")
     If Right(stTP, 2) >= 30 Then
        stTP = Val(Mid(stTP, 1, InStr(stTP, ":") - 1)) + 1 & ":00"
    Else
        stTP = Mid(stTP, 1, InStr(stTP, ":") - 1) & ":30"
    End If
     cboTime(0) = stTP
     'end2019/11/12
     cboTime(1) = "17:30"
     'end 2019/01/24
   
    MSHFlexGrid1.Clear
    MSHFlexGrid1.Rows = 1
    
    lstAtt.Clear
    lstAtt1.Clear
    
    Check1.Value = 0
    
    MSHFlexGrid2.Clear
    MSHFlexGrid2.Rows = 2
    
    lblCreateData = ""
    lblTitle = ""
    lblReceiver = ""
   
    Erase m_FilesRemoved
    ReDim m_FilesRemoved(0) As String
    Erase m_arrBookInList
    ReDim m_arrBookInList(0) As String
    m_stNumList = ""
    'Add by Amy 2018/09/18
    Erase m_arrBookPeo
    ReDim m_arrBookPeo(0) As String
    Erase Del_arrBookinList
    ReDim Del_arrBookinList(0) As String
    'end 2018/09/18
    lstUsers.Clear
    strSN12 = ""
    lblChairMan = ""
    lstCC.Clear
    lblCC = ""
    cboEmp.Clear
    Check2(0).Value = 0
    Check2(1).Value = 0
    cboRoom.Tag = ""
End Sub

'Add by Amy 2018/09/18
'Modify by Amy 2021/01/05 原:Integer
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
    If Not (Index = 4 Or Index = 23) And ActionEdit <> 0 And ActionEdit <> 1 Then Exit Sub
    
    KeyAscii = UpperCase(KeyAscii)
    If Index <> 23 Then Exit Sub
    If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
        KeyAscii = 0
        Beep
    End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    If Index <> 4 Then Exit Sub
    If Text1(Index) = MsgText(601) And Not (ActionEdit = 0 Or ActionEdit = 1) Then Exit Sub
    
    If CheckKeyIn(Text1(Index)) = -1 Then
        Text1(Index).SetFocus
        Text1_GotFocus (Index)
        Exit Sub
    End If
End Sub
'end 2018/09/18

'Added by Morgan 2022/5/25
Private Sub Text1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = 2 Then Forms(0).PopupMenu2 Text1(Index)
End Sub

Private Sub Timer1_Timer()
   If cmdBooinSave.FontSize < 11 Then
      cmdBooinSave.FontSize = 11
   Else
      cmdBooinSave.FontSize = 9
   End If
   
   If cmdBooinSave.BackColor = Me.BackColor Then
      cmdBooinSave.BackColor = RGB(&HFF, &HD7, &H0)
   Else
      cmdBooinSave.BackColor = Me.BackColor
   End If
End Sub

Private Sub CloseTimer()
   Timer1.Enabled = False
   cmdBooinSave.FontSize = 9
   cmdBooinSave.BackColor = Me.BackColor
End Sub

Private Sub txtAttender_Change()
   lblName = ""
   If Left(txtAttender, 1) < "z" And (Len(txtAttender) = 5 Or Len(txtAttender) = 6) Then
      strExc(0) = "select st02 from staff where st01='" & UCase(txtAttender) & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         lblName = "" & RsTemp.Fields(0)
      End If
   End If
End Sub

'Add by Amy 2018/09/18
Private Sub txtDept_GotFocus(Index As Integer)
    CloseIme
    TextInverse txtDept(Index)
End Sub

Private Sub txtDept_KeyPress(Index As Integer, KeyAscii As Integer)
     KeyAscii = UpperCase(KeyAscii)
End Sub
'end 2018/09/18

Private Sub txtQueryDate_GotFocus(Index As Integer)
   CloseIme
   TextInverse txtQueryDate(Index)
End Sub

Private Function GetMailContent(Optional pStaffId As String) As String
   Dim strContent As String, stText As String, ii As Integer, bolStar As Boolean
   
   strContent = "<DIV dir=ltr style=""MARGIN-RIGHT: 0px"" align=left><FONT face=標楷體>"
   
   stText = String(11, " ") & Text1(18)
   strContent = strContent & "<FONT Size=5>" & Replace(stText, " ", "&nbsp;") & "</FONT>" & vbCrLf
   'Modify by Amy 2018/09/18 部門改抓變數,主席、副本 改抓ListBox
   'stText = String(80, " ") & "專利處 " & Left(strSrvDate(1), 4) & " 年 " & Val(Mid(strSrvDate(1), 5, 2)) & " 月 " & Val(Mid(strSrvDate(1), 7)) & " 日"
   'Modify by Amy 2019/12/12 拿掉日期前 部門
   stText = String(80, " ") & "　　　" & " " & Left(strSrvDate(1), 4) & " 年 " & Val(Mid(strSrvDate(1), 5, 2)) & " 月 " & Val(Mid(strSrvDate(1), 7)) & " 日"
   strContent = strContent & "<FONT Size=1>" & Replace(stText, " ", "&nbsp;") & "</FONT>" & vbCrLf
   strContent = strContent & "<FONT Size=3>"
'   strContent = strContent & "主席：" & Replace(Text1(2), " ", "&nbsp;") & vbCrLf
'   strContent = strContent & "副本：" & Replace(Text1(4), " ", "&nbsp;") & vbCrLf
   strContent = strContent & "主席：" & Replace(GetMailTxt(lstUsers), " ", "&nbsp;") & vbCrLf
   strContent = strContent & "副本：" & Replace(GetMailTxt(lstCC), " ", "&nbsp;") & vbCrLf
   'end 2021/01/22
   'end 2018/09/18
   strContent = strContent & "出席：" & Replace(Text1(3), " ", "&nbsp;") & vbCrLf
   'Modify by Amy 2019/11/12 改元件
   strContent = strContent & "日期：" & Replace(Format(MaskEdBox1, "m 月 d 日 ") & lblWeek, " ", "&nbsp;") & vbCrLf
   'Modify by Amy 2019/01/24 改元件 原:DTPicker1(1)和(2)
   'strContent = strContent & "時間：" & Replace(Format(DTPicker1(1), "hh:mm") & "～" & Format(DTPicker1(2), "hh:mm"), " ", "&nbsp;") & vbCrLf
   'Modify by Amy 2021/10/22  +非實際起迄時間
   strContent = strContent & "時間：" & Replace(cboTime(0) & "～" & cboTime(1), " ", "&nbsp;") & "(非實際起迄時間)" & vbCrLf
   strContent = strContent & "地點：" & Replace(cboRoom & Text1(8), " ", "&nbsp;") & vbCrLf
   strContent = strContent & "議題：" & vbCrLf
   stText = ""
   
   For ii = 0 To MSHFlexGrid1.Rows - 1
        If pStaffId <> "" And InStr("," & m_arrBookInList(ii + 1) & ",", "," & pStaffId & ",") > 0 Then
            'bolStar = True 'Mark by Amy 2021/07/30 bug未改到,發測式信仍有*說明
            'Modify by Amy 2021/01/22 不需*
            'stText = stText & "    ＊"
            stText = stText & "      "
        Else
            stText = stText & "      "
        End If
        stText = stText & Replace(MSHFlexGrid1.TextMatrix(ii, 0), vbCrLf, vbCrLf & "      ") & vbCrLf
   Next
   stText = stText & "</FONT>"
   'Modify by Amy 2018/09/18 程式往下搬否則內文會出現<FONT size=3> (*表示指定參加之議題)
'   If bolStar Then
'      stText = stText & vbCrLf & "<FONT size=3>(*表示指定參加之議題)</FONT>" & vbCrLf & vbCrLf
'   End If
   strContent = strContent & Replace(stText, " ", "&nbsp;")
   If bolStar Then
      strContent = strContent & vbCrLf & "<FONT size=3>(*表示指定參加之議題)</FONT>" & vbCrLf & vbCrLf
   End If
   'end 2018/09/18
   strContent = strContent & "<FONT Size=3>說明：</FONT>" & vbCrLf
   strContent = strContent & "<FONT Size=2>" & Replace(Text1(9), " ", "&nbsp;") & "</FONT>"
   
   strContent = strContent & "</FONT></DIV>"
   
   strContent = Replace(strContent, vbCrLf, "<BR>" & vbCrLf)
   GetMailContent = strContent
End Function

'Add by Amy 2018/09/18 SendInformMail_Old修改
'原程式未使用因不好用-雅娟(都用測式信轉寄)
'pTo:測式信人員
'StMailList:只show 收件人
'stMailCC:只show副本
'intSendJoinOlny:0-只單獨寄參加者(一人一封)/1-寄信/2-只show 參加人 'Add by Amy 2020/11/27
'stMailSpeaker:只show 演講者 'Add by Amy 2020/12/28
Private Sub SendInformMail(Optional pTo As String, Optional ByRef stMailList As String = "", Optional ByRef stMailCC As String = "", Optional ByVal intSendJoinOlny As Integer = 2, Optional ByRef stMailSpeaker As String = "")
    Dim strContent As String, ii As Integer, jj As Integer, stText As String
    Dim strTemp(1) As String
    Dim strTo As String, strToList As String  '個別發信人員/收件人
    Dim bolOnlyMailList As Boolean, bolOnlyMailCC As Boolean '只show 收件人/副本
    Dim arr_List
    'Add by Amy 2020/11/27
    Dim strOldSort As String, strChairMan(1) As String, strCC(1) As String, strToJoin(1) As String '排序/主席/副本/參加者不個別發信List
    Dim intCountR As Integer, intMaxCol As Integer, intMaxRow As Integer '目前列/最大顯示收信人員數/最大顯示列數
    Dim bolOnlyMailSpeaker As Boolean, strAllSpeaker(1) As String, strChkList As String 'Add by Amy 2020/12/28 只show演講者/所有演講者/判斷重覆資料
    
    intMaxCol = 6: intMaxRow = 11 'Add by Amy 2020/11/27 一列顯示收信人員數
    bolOnlyMailList = False: bolOnlyMailCC = False
    'Add by Amy 2020/12/28
    bolOnlyMailSpeaker = False
    If stMailSpeaker <> MsgText(601) Then
        bolOnlyMailSpeaker = True
        stMailSpeaker = ""
    End If
    'end 2020/12/28
    If stMailList <> MsgText(601) Then
        bolOnlyMailList = True
        stMailList = ""
    End If
    If stMailCC <> MsgText(601) Then
        bolOnlyMailCC = True
        stMailCC = ""
    End If
    
    '測式信
    If pTo <> "" Then
        strTo = pTo
        strContent = GetMailContent(strTo)
        PUB_SendMail strUserNum, strTo, "", cboRoom & Text1(18) & "開會通知", strContent, , , True, , , , , , , True
        Exit Sub
    End If
    
    '有設定要參加的人單獨寄(信件內容不同)
    'Modify by Amy 2020/11/27 改判斷intSendJoinOlny=0,原參加人員單獨發信,因參加議題顯示*,目前不使用
    'If bolOnlyMailList = False And bolOnlyMailCC = False Then
    If intSendJoinOlny = 0 Then
        strExc(0) = "Select sb02,sb07 From Seminarbookin,Staff " & _
                          "Where st04='1' And SB02=ST01(+) And SB01=" & Text1(1) & " And SB02<>'ALL999' "
        intI = 1
        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
        If intI = 1 Then
            RsTemp.MoveFirst
            Do While Not RsTemp.EOF
                strTo = RsTemp.Fields("sb02")
                strContent = GetMailContent(RsTemp.Fields("sb02"))
                'Modify by Amy 2020/11/27 bol63001CanSend參數設True
                PUB_SendMail strUserNum, strTo, "", cboRoom & Text1(18) & "開會通知", strContent, , , True, , , , , , , True, , , , , , , , , True
                RsTemp.MoveNext
            Loop
        End If
    End If
    
    '發信人員(主席,副本,選擇參加人員)
    'Modify by Amy 2020/11/27 主席與參加者及與副本重覆,以收件者(主席->參加者->副本)優先(寄一次和顯示於收件者)
    strChairMan(0) = "Y"
    Call ChkOldListData(lstUsers, False, strChairMan(0)) '主席(員編)
    strCC(0) = "Y"
    Call ChkOldListData(lstCC, False, strCC(0)) '副本(員編)
    strAllSpeaker(0) = GetSendSS06 'Add by Amy 2020/12/28 演講者
    
    '過濾主席與參加人員重覆
    'Modify by Amy 2019/11/12 原抓m_stNumList,可能為空值
    If intSendJoinOlny >= 1 Then
        For i = LBound(m_arrBookInList) To UBound(m_arrBookInList)
            If m_arrBookInList(i) <> MsgText(601) Then
                arr_List = Split(m_arrBookInList(i), ",")
                For ii = LBound(arr_List) To UBound(arr_List)
                    If InStr(strChairMan(0), arr_List(ii)) = 0 Then
                        strToJoin(0) = strToJoin(0) & ";" & arr_List(ii)
                    End If
                Next ii
            End If
        Next i
        If strToJoin(0) <> MsgText(601) Then strToJoin(0) = Mid(strToJoin(0), 2)
    End If
    'Add by Amy 2020/12/28 過濾主席、參加人員與演講者重覆
    If strAllSpeaker(0) <> MsgText(601) Then
        arr_List = Split(strAllSpeaker(0), ";")
        strAllSpeaker(0) = ""
        strChkList = strChairMan(0)
        If strToJoin(0) <> MsgText(601) Then strChkList = strChkList & ";" & strToJoin(0)
        For i = LBound(arr_List) To UBound(arr_List)
            If InStr(strChkList, arr_List(i)) = 0 Then
                strAllSpeaker(0) = strAllSpeaker(0) & ";" & arr_List(i)
            End If
        Next i
        If strAllSpeaker(0) <> MsgText(601) Then strAllSpeaker(0) = Mid(strAllSpeaker(0), 2)
    End If
    '過濾主席、參加人員與副本重覆
    'Modify by Amy 2020/12/28 +演講者
    If strCC(0) <> MsgText(601) Then
        arr_List = Split(strCC(0), ";")
        strChkList = strChairMan(0)
        If strToJoin(0) <> MsgText(601) Then strChkList = strChkList & ";" & strToJoin(0)
        If strAllSpeaker(0) <> MsgText(601) Then strChkList = strChkList & ";" & strAllSpeaker(0)
        For i = LBound(arr_List) To UBound(arr_List)
            If InStr(strChkList, arr_List(i)) = 0 Then
                strCC(1) = strCC(1) & ";" & arr_List(i)
            End If
        Next i
    End If
    'end 2020/12/28
    '過濾固定副本人員與副本重覆
    If strCC(1) <> MsgText(601) Then
        arr_List = Split(strCC(1), ";")
        strCC(1) = ""
        For i = LBound(arr_List) To UBound(arr_List)
            If InStr(strCC_Fix, arr_List(i)) = 0 Then
                strCC(1) = strCC(1) & ";" & arr_List(i)
            End If
        Next i
    End If
    
    strExc(0) = "":  ii = 0: jj = 0
     '主席
    If strChairMan(0) <> MsgText(601) Then
        strTemp(0) = Replace(strChairMan(0), ";", "','")
        strExc(0) = strExc(0) & "Select st01,st02,'100'||st01 Sort From Staff Where st04='1' And ST01 In('" & strTemp(0) & "') "
    End If
    '副本
    If strExc(0) <> MsgText(601) Then strExc(0) = strExc(0) & " Union "
    '固定副本 以職稱->員編排
    strExc(0) = strExc(0) & "Select st01,st02,'2'||ac02||st01 Sort From Staff,AllCode Where st04='1' And ST01 In('" & Replace(strCC_Fix, ";", "','") & "') And ac02(+)=st20 And ac01(+)='01' "
    '其他副本 以員編排
    If strCC(1) <> MsgText(601) Then
        strTemp(1) = Replace(Mid(strCC(1), 2), ";", "','")
        strExc(0) = strExc(0) & " Union Select st01,st02,'300'||st01 Sort From Staff Where st04='1' And ST01 In('" & strTemp(1) & "') "
    End If
    'Add by Amy 2020/12/28 +演講者
    If strAllSpeaker(0) <> MsgText(601) Then
        strExc(0) = strExc(0) & " Union Select st01,st02,'400'||st01 Sort From Staff Where st04='1' And ST01 In('" & Replace(strAllSpeaker(0), ";", "','") & "') "
    End If
    If strExc(0) <> MsgText(601) Then strExc(0) = strExc(0) & " Order by Sort "
    intI = 1
    Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
    If intI = 1 Then
        strTemp(0) = "": strTemp(1) = ""
        RsTemp.MoveFirst
        Do While Not RsTemp.EOF
            If (Left("" & RsTemp.Fields("Sort"), 1) = "3" Or Left("" & RsTemp.Fields("Sort"), 1) = "4") And Left("" & RsTemp.Fields("Sort"), 1) <> Left(strOldSort, 1) Then
                jj = 0
            End If
            '主席
            If Left("" & RsTemp.Fields("Sort"), 1) = "1" Then
                strChairMan(1) = strChairMan(1) & ";" & RsTemp.Fields(0)
                If bolOnlyMailList = True Then
                    strChairMan(1) = strChairMan(1) & "(" & RsTemp.Fields("st02") & ")"
                    ii = ii + 1
                    If ii Mod intMaxCol = 0 Then strChairMan(1) = strChairMan(1) & vbCrLf & "@@": intCountR = intCountR + 1
                End If
            '固定副本
            ElseIf Left("" & RsTemp.Fields("Sort"), 1) = "2" Then
                strTemp(0) = strTemp(0) & ";" & RsTemp.Fields(0)
                If bolOnlyMailCC = True Then
                    strTemp(0) = strTemp(0) & "(" & RsTemp.Fields("st02") & ")"
                    jj = jj + 1
                    If jj Mod intMaxCol = 0 Then strTemp(0) = strTemp(0) & vbCrLf & "@@": intCountR = intCountR + 1
                End If
            '其他副本
            ElseIf Left("" & RsTemp.Fields("Sort"), 1) = "3" Then
                strTemp(1) = strTemp(1) & ";" & RsTemp.Fields(0)
                If bolOnlyMailCC = True Then
                    strTemp(1) = strTemp(1) & "(" & RsTemp.Fields("st02") & ")"
                    jj = jj + 1
                    If jj Mod intMaxCol = 0 Then strTemp(1) = strTemp(1) & vbCrLf & "@@": intCountR = intCountR + 1
                End If
            '演講者
            Else
                strAllSpeaker(1) = strAllSpeaker(1) & ";" & RsTemp.Fields(0)
                If bolOnlyMailCC = True Then
                    strAllSpeaker(1) = strAllSpeaker(1) & "(" & RsTemp.Fields("st02") & ")"
                    jj = jj + 1
                    If jj Mod intMaxCol = 0 Then strAllSpeaker(1) = strAllSpeaker(1) & vbCrLf & "@@": intCountR = intCountR + 1
                End If
            End If
            strOldSort = "" & RsTemp.Fields("Sort")
            RsTemp.MoveNext
        Loop
    End If
    '參加人員 以所別,職稱,部門,員編排(此處改看登記頁籤登記人員是否也需改)
    If strToJoin(0) <> MsgText(601) Then
        ii = 0
        strToJoin(0) = Replace(strToJoin(0), ";", "','")
        strExc(0) = "Select st01,st02,Decode(st06,'1','北','2','中','3','南','4','高') as st06N,st06||Nvl(st20,'99')||st15||st01 Sort From Staff " & _
                            "Where st04='1' And ST01 In('" & strToJoin(0) & "') Order by st06||st20||st15||st01"
        intI = 1
        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
        If intI = 1 Then
            strOldSort = ""
            RsTemp.MoveFirst
            Do While Not RsTemp.EOF
                'Modify by Amy 2021/01/26 +bolOnlyMailList = True And bolOnlyMailCC = True ,勾選發信存檔會錯
                If bolOnlyMailList = True And bolOnlyMailCC = True And Left("" & RsTemp.Fields("Sort"), 1) <> Left(strOldSort, 1) And strOldSort <> MsgText(601) And intCountR < intMaxRow Then
                    '所別不同換行
                    strToJoin(1) = strToJoin(1) & vbCrLf & "[" & RsTemp.Fields("st06N") & "-：": intCountR = intCountR + 1
                    ii = 0
                End If
                strToJoin(1) = strToJoin(1) & ";" & RsTemp.Fields(0)
                '只show 參加人員
                If bolOnlyMailList = True And bolOnlyMailCC = True Then
                    If intCountR <= intMaxRow Then
                        If intCountR = intMaxRow And InStr(strToJoin(1), "...") = 0 Then
                            strToJoin(1) = Replace(strToJoin(1), ";" & RsTemp.Fields(0), ";...")
                        Else
                            strToJoin(1) = strToJoin(1) & "(" & RsTemp.Fields("st02") & ")"
                            ii = ii + 1
                            If ii Mod intMaxCol = 0 Then strToJoin(1) = strToJoin(1) & vbCrLf & "@@": intCountR = intCountR + 1
                        End If
                    End If
                End If
                strOldSort = "" & RsTemp.Fields("Sort")
                RsTemp.MoveNext
            Loop
        End If
    End If
    
    strCC(1) = "": strToList = strChairMan(1)
    '主席排序無資料,抓strChairMan(0)變數 ex:Patent 員工檔不會有資料
    If InStr(strChairMan(0), "Patent") > 0 Or InStr(strChairMan(0), "patent@taie.com.tw") > 0 Then
        strToList = strToList & ";patent@taie.com.tw"
    End If
    '勾選「部門全體人員」加發mail
    If Check2(0).Value = 1 Then strToList = strToList & ";" & strDefDeptMail
    '其他副本
    If InStr(strTemp(1), "Patent") Or InStr(strTemp(1), "patent@taie.com.tw") > 0 Then
        strCC(1) = strCC(1) & ";patent@taie.com.tw"
    End If
    '只show Mail
    If bolOnlyMailList = True Or bolOnlyMailCC = True Then
        If strToList = MsgText(601) Then
            '避免主席為參加人員(主席為空),文字被「取消」鈕檔住,故加換行 ex:1090016
            strToList = vbCrLf & vbCrLf
        Else
            strToList = "收件者（主　　席）：" & vbCrLf & "@@" & strToList & vbCrLf & vbCrLf
        End If
        'Add by Amy 2020/12/28
        If strAllSpeaker(1) <> MsgText(601) Then
            strToList = strToList & "收件者（演 講 者）：" & vbCrLf & "@@" & strAllSpeaker(1) & vbCrLf & vbCrLf
        End If
        If strToJoin(1) <> MsgText(601) Then
            If InStr(strToJoin(1), "...") > 0 Then strToJoin(1) = Mid(strToJoin(1), 1, Val(InStr(strToJoin(1), "...")) + 3)
            strToList = strToList & "收件者（參加人員）：" & vbCrLf & "@@" & strToJoin(1) & vbCrLf & vbCrLf
        End If
        If strTemp(0) <> MsgText(601) Then strCC(1) = "副本（固定收件者）：" & vbCrLf & "@@" & strTemp(0) & vbCrLf
        If strTemp(1) <> MsgText(601) Then strCC(1) = strCC(1) & vbCrLf & "副本（其他收件者）：" & vbCrLf & "@@" & strTemp(1) & vbCrLf
        '@@;>>取代為換行+全型+半型空白/[>>取代為換行/：;>>所別取代空字串/[>>所別換行/vbCrLf+@@>>空字串
        stMailList = Replace(Replace(Replace(Replace(strToList, vbCrLf & "@@;", vbCrLf & "　 "), "：;", ""), "[", vbCrLf), "@@" & vbCrLf, "")
        stMailCC = Replace(Replace(Replace(strCC(1), vbCrLf & "@@;", vbCrLf & "　 "), "：;", ""), "@@" & vbCrLf, "")
        Exit Sub
    '寄信
    Else
        If strToList <> MsgText(601) Then strToList = Mid(strToList, 2)
        If strAllSpeaker(1) <> MsgText(601) Then strToList = strToList & strAllSpeaker(1) 'Add by Amy 2020/12/28 +演講者(列於收件者)
        If strToJoin(1) <> MsgText(601) Then strToList = strToList & strToJoin(1)
        If strTemp(0) <> MsgText(601) Or strTemp(1) <> MsgText(601) Then strCC(1) = Mid(strTemp(0) & strTemp(1), 2)
    End If
      
    '收件者與副本若有重覆寄信時Outlook會自動排除重覆只寄一封
    strContent = GetMailContent()
    'bol63001CanSend參數設True,才會寄董事長
    PUB_SendMail strUserNum, strToList, "", cboRoom & Text1(18) & "開會通知", strContent, , , True, , , strCC(1), , , , True, , , , , , , , , True
    'end 2020/11/27
End Sub

'Mark by Amy 2018/09/18 因不好用,都沒在用,都用測式信轉發-雅娟
'Modified by Morgan 2012/5/7 因信件內容要標示指定參加的議題改有設定要參加的人單獨寄
Private Sub SendInformMail_Old(Optional pTo As String)
'   Dim strContent As String, ii As Integer, stText As String
'   Dim strTo As String, strCC As String, strToList As String
'
'   If pTo <> "" Then
'      strTo = pTo
'      strContent = GetMailContent(strTo)
'      PUB_SendMail strUserNum, strTo, "", Text1(18) & "開會通知", strContent, , , True, , , , , , , True
'   Else
'
'      '收件者：總經理,國外部副總,專利處工程師,74001,79075
'      strToList = ""
'      'strExc(0) = "select st20,st01,sb07 from staff,seminarbookin where st04='1' and (st03 in ('P10','P11') or st01 in ('68001','68009','74001','79075')) and st01<'F' and sb01(+)=" & Text1(1) & " order by 1,2"
'      'modify by sonia 2014/9/9 改68001為94007
'      'modify by sonia 2015/6/30 改68009為81040
'      strExc(0) = "select st20,st01,sb07 from staff,seminarbookin where st04='1' and instr(','||'" & m_stNumList & "'||',94007,81040,71011,73022,74001,79075,',','||ST01||',')>0  and sb01(+)=" & Text1(1) & " AND SB02(+)=ST01 order by 1,2"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         With RsTemp
'         Do While Not .EOF
'            If IsNull(.Fields("sb07")) Then
'               strToList = strToList & .Fields("st01") & ";"
'            Else
'               strTo = .Fields("st01")
'               strContent = GetMailContent(.Fields("st01"))
'               PUB_SendMail strUserNum, strTo, "", Text1(18) & "開會通知", strContent, , , True, , , , , , , True
'            End If
'            .MoveNext
'         Loop
'         End With
'      End If
'
'      '副本：所長,副所長
'      strCC = ""
'      strExc(0) = "select st20,st01 from staff where st04='1' and st20 in ('11','12') and st01<'F' order by 1,2"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         With RsTemp
'         Do While Not .EOF
'            strCC = strCC & .Fields("st01") & ";"
'            .MoveNext
'         Loop
'         End With
'      End If
'
'      If strToList <> "" Then
'         strContent = GetMailContent()
'         PUB_SendMail strUserNum, strToList, "", Text1(18) & "開會通知", strContent, , , True, , , strCC, , , , True
'      End If
'   End If
End Sub

Private Sub txtReceiver_Change()
   If Left(txtReceiver, 1) < "z" And (Len(txtReceiver) = 5 Or Len(txtReceiver) = 6) Then
      strExc(0) = "select st02 from staff where st01='" & UCase(txtReceiver) & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         lblReceiver = "" & RsTemp.Fields(0)
      End If
   End If
End Sub

Private Sub txtReceiver_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtSpeaker_Change()
   If Left(txtSpeaker, 1) < "z" And (Len(txtSpeaker) = 5 Or Len(txtSpeaker) = 6) Then
      strExc(0) = "select st02 from staff where st01='" & UCase(txtSpeaker) & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         txtSpeaker = "" & RsTemp.Fields(0)
      End If
   End If
End Sub

'Mark by Amy 2020/11/27
'列印點名單
'Modified by Morgan 2012/5/3 改各議題表格分開(因參加人員可不同)
Private Sub runWord_Old()
'   Dim strFontSize As String
'   Dim stTmp As String
'   Dim iRow As Integer, iCol As Integer, iCols As Integer, ii As Integer, jj As Integer
'   Dim stOffice As String
'   Dim iResumeCnt As Integer
'   Dim iRowHeight As Integer
'   Dim iLine As Integer, iLines As Integer
'   Dim bolMyFlag As Boolean, iTables As Integer
'   Dim dblColW As Double
'
'On Error GoTo ErrHnd
'
'   If g_WordAp Is Nothing Then Set g_WordAp = New Word.Application
'   g_WordAp.Documents.add
'   With g_WordAp
'      .Visible = True
'      '.Visible = False
'      .Selection.Font.Name = "標楷體"
'
'      With .Options
'        .DefaultBorderLineStyle = wdLineStyleSingle
'        .DefaultBorderLineWidth = wdLineWidth050pt
'        '.DefaultBorderColor = wdColorBlack 'Word97 沒有這個屬性及常數(Word2007 有)
'      End With
'
'      .Selection.PageSetup.PaperSize = wdPaperA4
'
'      .Selection.PageSetup.Orientation = wdOrientLandscape
'      strFontSize = 14
'
'      .Selection.Orientation = wdTextOrientationHorizontal
'      .Selection.Font.Size = strFontSize
'      '邊界
'      .Selection.PageSetup.LeftMargin = .CentimetersToPoints(1.6)
'      .Selection.PageSetup.RightMargin = .CentimetersToPoints(1.4)
'      .Selection.PageSetup.TopMargin = .CentimetersToPoints(2)
'      .Selection.PageSetup.BottomMargin = .CentimetersToPoints(2)
'
'      stOffice = ""
'      For jj = 2 To MSHFlexGrid2.Cols - 1
'         If MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 4, jj) <> stOffice Then
'            stOffice = MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 4, jj)
'            iTables = 0
'            For iRow = 2 To MSHFlexGrid2.Rows - 5
'
'               '依議題計算人數(可登記的才算)
'               iCols = 0
'               For ii = jj To MSHFlexGrid2.Cols - 1
'                  MSHFlexGrid2.row = iRow
'                  MSHFlexGrid2.col = ii
'                  If stOffice = MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 4, ii) Then
'                     If MSHFlexGrid2.CellBackColor = MSHFlexGrid2.BackColor Then
'                        iCols = iCols + 1
'                     End If
'                  Else
'                     Exit For
'                  End If
'               Next
'
'               If iCols > 0 Then
'                  iTables = iTables + 1
'                  '兩個議題印一張
'                  If iTables Mod 2 = 1 Then
'                     '依所別跳頁
'                     If jj > 2 Or iTables > 1 Then
'                        .Selection.MoveRight Unit:=wdCharacter, Count:=2
'                        .Selection.InsertBreak Type:=wdPageBreak
'                     End If
'                     '印表頭
'                     .Selection.ParagraphFormat.DisableLineHeightGrid = True
'                     'Modify by Amy 2019/11/12 改元件
'                     stTmp = lblTitle & " ( " & MaskEdBox1 & " )"
'                     .Selection.Font.Size = 18
'                     .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
'                     .Selection.TypeText Text:=stTmp
'
'                     .Selection.TypeParagraph
'                     .Selection.Font.Size = strFontSize
'                     .Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
'
'                     Select Case stOffice
'                        Case "1": stTmp = "北所"
'                        Case "2": stTmp = "中所"
'                        Case "3": stTmp = "南所"
'                        Case "4": stTmp = "高所"
'                     End Select
'                     .Selection.Font.Size = 16
'                     .Selection.TypeText Text:=stTmp
'                     .Selection.Font.Size = strFontSize
'                     .Selection.TypeParagraph
'
'                  End If
'
'                  If iTables Mod 2 = 0 Then
'                     .Selection.SelectRow
'                     .Selection.MoveRight Unit:=wdCharacter, Count:=2
'                     .Selection.TypeParagraph
'                  End If
'
'                  If iCols <= 22 Then
'                     dblColW = 0.95
'                  Else
'                     dblColW = 0.8
'                  End If
'
'                  '插入表格
'                  'Modified by Morgan 2013/6/20 修正2007格線不顯示問題,表格設定太多格式會當掉問題
'                  '.Selection.Tables.Add Range:=.Selection.Range, NumRows:=2, NumColumns:=(iCols + 1)
'                  .Selection.Tables.add Range:=.Selection.Range, NumRows:=3, NumColumns:=(iCols + 1)
'
'                  '設框線,高寬
'                  'Modified by Morgan 2013/6/20
'                  '.Selection.SelectRow
'                  .Selection.Tables(1).Select
'
'                  With .Selection.Borders(wdBorderTop)
'                      .LineStyle = g_WordAp.Options.DefaultBorderLineStyle
'                      .LineWidth = g_WordAp.Options.DefaultBorderLineWidth
'                      '.Color = g_WordAp.Options.DefaultBorderColor'Word97 沒有這個屬性及常數(Word2007 有)
'                  End With
'                  With .Selection.Borders(wdBorderLeft)
'                      .LineStyle = g_WordAp.Options.DefaultBorderLineStyle
'                      .LineWidth = g_WordAp.Options.DefaultBorderLineWidth
'                      '.Color = g_WordAp.Options.DefaultBorderColor'Word97 沒有這個屬性及常數(Word2007 有)
'                  End With
'                  With .Selection.Borders(wdBorderBottom)
'                      .LineStyle = g_WordAp.Options.DefaultBorderLineStyle
'                      .LineWidth = g_WordAp.Options.DefaultBorderLineWidth
'                      '.Color = g_WordAp.Options.DefaultBorderColor'Word97 沒有這個屬性及常數(Word2007 有)
'                  End With
'                  With .Selection.Borders(wdBorderRight)
'                      .LineStyle = g_WordAp.Options.DefaultBorderLineStyle
'                      .LineWidth = g_WordAp.Options.DefaultBorderLineWidth
'                      ''Word97 沒有這個屬性及常數(Word2007 有).Color = g_WordAp.Options.DefaultBorderColor
'                  End With
'                  With .Selection.Borders(wdBorderHorizontal)
'                      .LineStyle = g_WordAp.Options.DefaultBorderLineStyle
'                      '.LineWidth = g_WordAp.Options.DefaultBorderLineWidth'Word巨集正常但vb跑會有錯
'                      '.Color = g_WordAp.Options.DefaultBorderColor'Word97 沒有這個屬性及常數(Word2007 有)
'                  End With
'                  With .Selection.Borders(wdBorderVertical)
'                      .LineStyle = g_WordAp.Options.DefaultBorderLineStyle
'                      .LineWidth = g_WordAp.Options.DefaultBorderLineWidth
'                      '.Color = g_WordAp.Options.DefaultBorderColor'Word97 沒有這個屬性及常數(Word2007 有)
'                  End With
'
'                  '設定表格高度
'                  .Selection.MoveLeft Unit:=wdCharacter, Count:=1 'Added by Morgan 2013/6/20
'                  .Selection.SelectRow
'                  '.Selection.Font.Bold = wdToggle
'                  .Selection.Cells.SetHeight RowHeight:=56, HeightRule:=wdRowHeightExactly
'
'                  .Selection.MoveLeft Unit:=wdCharacter, Count:=1
'                  .Selection.SelectColumn
'                  .Selection.Cells.SetWidth ColumnWidth:=.CentimetersToPoints(5.5), RulerStyle:=wdAdjustProportional
'                  .Selection.SelectRow
'                  .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
'                  .Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
'                  .Selection.MoveLeft Unit:=wdCharacter, Count:=1
'
'                  .Selection.MoveDown Unit:=wdLine, Count:=1
'
'                  'Added by Morgan 2013/6/20
'                  .Selection.SelectRow
'                  .Selection.Cells.SetHeight RowHeight:=20, HeightRule:=wdRowHeightExactly
'                  'end 2013/6/20
'
'                  stTmp = MSHFlexGrid2.TextMatrix(iRow, 0)
'                  iLines = 1
'                  intI = InStr(1, stTmp, vbCrLf)
'                  Do While intI > 0
'                     iLines = iLines + 1
'                     intI = InStr(intI + 1, stTmp, vbCrLf)
'                  Loop
'                  iRowHeight = iLines * 20
'                  If iRowHeight < 80 Then iRowHeight = 80
'
'                  'Added by Morgan 2013/6/20
'                  .Selection.MoveLeft Unit:=wdCharacter, Count:=1
'                  .Selection.MoveDown Unit:=wdLine, Count:=1
'                  .Selection.SelectRow
'                  'end 2013/6/20
'
'                  .Selection.Cells.SetHeight RowHeight:=(iRowHeight - 20), HeightRule:=wdRowHeightExactly
'
'                  'Added by Morgan 2013/6/20
'                  .Selection.MoveLeft Unit:=wdCharacter, Count:=1
'                  .Selection.MoveUp Unit:=wdLine, Count:=1, Extend:=wdExtend
'                  .Selection.Cells.Merge
'                  'end2013/6/20
'
'                  .Selection.TypeText Text:=stTmp
'                  bolMyFlag = True
'                  For iCol = jj To MSHFlexGrid2.Cols - 1
'                     If MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 4, iCol) <> stOffice Then Exit For
'
'                     MSHFlexGrid2.row = iRow
'                     MSHFlexGrid2.col = iCol
'                     If MSHFlexGrid2.CellBackColor = MSHFlexGrid2.BackColor Then
'                        .Selection.MoveRight Unit:=wdCell, Count:=1
'                        .Selection.MoveUp Unit:=wdLine, Count:=1
'
'                        .Selection.Cells.SetWidth ColumnWidth:=.CentimetersToPoints(dblColW), RulerStyle:=wdAdjustProportional
'                        .Selection.TypeText Text:=MSHFlexGrid2.TextMatrix(0, iCol)
'                        .Selection.MoveDown Unit:=wdLine, Count:=1
'
'                        'Removed by Morgan 2013/6/20
'                        '.Selection.Cells.Split NumRows:=2, NumColumns:=1, MergeBeforeSplit:=False
'                        'If bolMyFlag Then
'                        '   .Selection.Cells(1).SetHeight RowHeight:=20, HeightRule:=wdRowHeightExactly
'                        '   .Selection.Cells(2).SetHeight RowHeight:=(iRowHeight - 20), HeightRule:=wdRowHeightExactly
'                        '   bolMyFlag = False
'                        'End If
'                        '.Selection.MoveUp Unit:=wdLine, Count:=1
'                        '.Selection.MoveDown Unit:=wdLine, Count:=1
'                        'end 2013/6/20
'
'                        .Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
'                        .Selection.TypeText Text:=MSHFlexGrid2.TextMatrix(iRow, iCol)
'
'                        '.Selection.MoveDown Unit:=wdLine, Count:=1 'Removed by Morgan 2013/6/20
'                     End If
'                  Next iCol
'                  .Selection.MoveDown Unit:=wdLine, Count:=1 'Added by Morgan 2013/6/20
'               End If
'            Next iRow
'         End If
'      Next jj
'      .Selection.WholeStory
'      .Selection.Font.Name = "Times New Roman"
'      .Selection.MoveRight Unit:=wdCharacter, Count:=1
'      '插入頁碼
'      If .ActiveWindow.View.SplitSpecial = wdPaneNone Then
'         .ActiveWindow.ActivePane.View.Type = wdPageView
'      Else
'         .ActiveWindow.View.Type = wdPageView
'      End If
'      .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
'      .Selection.HeaderFooter.PageNumbers.add PageNumberAlignment:=wdAlignPageNumberCenter, FirstPage:=True
'      .ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
'
'      .Visible = True
'      .Activate
'   End With
'   Exit Sub
'
'ErrHnd:
'   If Err.Number <> 0 Then
'      If iResumeCnt > 3 Then
'         MsgBox "錯誤 : " & Err.Description, vbCritical
'      Else
'         iResumeCnt = iResumeCnt + 1
'         Select Case Err.Number
'            Case 91:
'               g_WordAp.Documents.add
'               Resume Next
'            Case 462:
'               Set g_WordAp = New Word.Application
'               Resume
'            Case Else:
'               MsgBox "錯誤" & iLine & " : " & Err.Description, vbCritical
'         End Select
'      End If
'   End If
End Sub

Private Function GetNumList(pSS02 As String) As String
   Dim strSB01 As String
   Dim adoRst As ADODB.Recordset
   
   strSB01 = Text1(1)
   If strSB01 <> "" Then
      strExc(0) = "select st01" & _
         " from seminarbookin,staff,allcode" & _
         " where sb01=" & strSB01 & " and instr(','||sb07||',',','||'" & pSS02 & "'||',')>0 and st01(+)=sb02 and ac02(+)=st20 and ac01(+)='01' order by st02 desc"
      intI = 1
      Set adoRst = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         GetNumList = adoRst.GetString(, , , ",")
      End If
   End If
   Set adoRst = Nothing
End Function

'Add by Amy 2018/09/18 表單預設值
Private Sub FormSet(Optional ByVal bolSN20Only As Boolean)
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String
    Dim intQ As Integer
    Dim strTmp(1) As String
    Dim strManagerNo As String, strTitleN As String, strAttend As String '部門主管員編/標題/出席
    Dim strSC05 As String   '副本
    Dim bolOpen As Boolean 'Add by Amy 2020/11/27

    MaskEdBox1.Mask = MsgText(601)
    MaskEdBox1 = Format(strSrvDate(1), "####/##/##")
    MaskEdBox1.Mask = ADFormat
    lblWeek = GetWeekDay(CDate(MaskEdBox1))
    
    If Left(strDeptNo, 2) <> "F2" Then strDeptName = ""
    'Modify by Amy 2020/11/27
    bolOpen = Option1(0).Value
    strQ = GetSeminarContactSql(1, strDeptNo, bolOpen)
    'end 2020/11/27
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        If Left(strDeptNo, 2) <> "F2" Then strDeptName = "" & RsQ.Fields("SC03")  '部門名稱
        strManagerNo = "" & RsQ.Fields("SC04") '知會主管
        strSC05 = "" & RsQ.Fields("SC05") '副本收受者(不含固定預設)
        strAttend = "" & RsQ.Fields("SC06") '出席
        '公開且為外專抓部門名稱
        If Option1(0).Value And Left(strDeptNo, 2) = "F2" Then
            strAttend = strDeptName
        End If
        strTitleN = "" & RsQ.Fields("SC07") '標題
        strDefDeptMail = "" & RsQ.Fields("SC08") '部門全體人員mail
    End If
    
    'Add by Amy 2019/01/24 +if
    If bolSN20Only = False Then
        '副本收受者為部門之取代為通知主管
        If strSC05 <> MsgText(601) Then
            strQ = Replace(strSC05, ";", "','")
            strSC05 = ";" & strSC05
            strQ = GetSeminarContactSql(3, strDeptNo, bolOpen, " And SC01 In ('" & strQ & "') ") 'Modify by Amy 2020/11/27
            intQ = 1
            Set RsQ = ClsLawReadRstMsg(intQ, strQ)
            If intQ = 1 Then
                RsQ.MoveFirst
                Do While RsQ.EOF = False
                    strSC05 = Replace(strSC05, ";" & RsQ.Fields("SC01"), ";" & RsQ.Fields("SC04"))
                    RsQ.MoveNext
                Loop
            End If
            strSC05 = Replace(Mid(strSC05, 2), strUserNum & ";", "") '排建立者自己(ex:FCP小組建立開公時,副本不需帶自己)
        End If
        '副本 排序(以員工編號排序)
        strSC05 = GetSortNo(strSC05, True)
    End If
    
    '主席List
    If Option1(0).Value = True And Left(strDeptNo, 2) = "F2" Then
        strManagerNo = strUserNum
    Else
        strManagerNo = GetSortNo(strManagerNo, False)
    End If
    'Modify by Amy 2019/01/24 多個主席不預帶,放於下拉選單
    Call SetCboEmp(strManagerNo)
    If InStr(strManagerNo, ";") = 0 Then
        Call SetList(lstUsers, strManagerNo)
    End If
    If bolSN20Only = True Then Exit Sub
    'end 2019/01/24
    
    '副本List
    Call SetList(lstCC, strSC05)
    '出席
    Text1(3) = strAttend
    '標題
    If Option1(0).Value = True Then
        strTitleN = Left(strSrvDate(1), 4) & "年" & (Val(Mid(strSrvDate(1), 5, 2))) & "月份" & strDeptName & "研討會"
    End If
    Text1(18) = strTitleN
   
    '地點
    'Modify by Amy 2019/01/24 地點拆成會議室及文字
    cboRoom = GetMeetingRoom("1", True) '預設5F
    Text1(8) = "及台中、台南、高雄會議室(同步)"
    '說明
    Text1(9) = "◎本次研討會書面資料請自行至系統內參看附件。" & vbCrLf & _
                     "◎" & strDeptName & "經副理以下人員請至系統內登記是否參加。" & vbCrLf
    '勾選公開說明增加3.
    If Option1(0).Value = True Then
       Text1(9) = Text1(9) & _
                        "◎非" & strDeptName & "請回覆寄件者貴部門欲參加人員。"
    End If
    Text1(23) = "Y" 'Add by Amy 2019/12/09 未按過「會議室預約」鈕都不檢查
    Frame2.Visible = True
End Sub

'設定原始附件及按鈕顯示
Private Sub SetLimit()
   'Mark by Amy 2018/09/18 從Form_Load 搬來修改
'   'Added by Morgan 2012/5/1
'   If m_bOpen Then
'      Frame2.Visible = True
'   Else
'      Frame2.Visible = False
'   End If
'   If m_bInsert Then
'      cmdPrint.Visible = True
'   Else
'      cmdPrint.Visible = False
'   End If
   
   Frame2.Visible = False '原始附件區
   cmdPrint.Visible = False '點名單按鈕
   CboChoose.Visible = False 'Add by Amy 2020/11/27 點名單列印選擇
   
   'Modify by Amy 2020/12/15 發現專利處109/11/11教育訓上線後建立者同部門可操作原始檔區
   '                                             與薛經理確認後專利處改回109/11/11教育訓上線前之判斷,
   '                                             其他部門只有建立者(員工權限維護檔)可以操作
   'Moidfy by Amy 2018/09/18 從Form_Load 搬來修改
   'Modify by Amy 2019/11/12 部門原抓PUB_GetST03
'   'F21外專工程師同組,其他外專人員同部門
'   If (strDeptNo = "F21" And PUB_GetStaffST16(strUserNum) = PUB_GetStaffST16(strSN12)) Or _
'        (Left(strDeptNo, 2) = "F2" And strDeptNo <> "F21" And strDeptNo = GetST15(strSN12)) Then
'        Frame2.Visible = True
'        cmdPrint.Visible = True
'    '建立者同部門
'    ElseIf Left(strDeptNo, 2) <> "F2" And (Left(strDeptNo, 2) = Left(GetST15(strSN12), 2) Or strDeptNo = "M51") Then
'        Frame2.Visible = True
'        cmdPrint.Visible = True
'    End If
'   'end 2019/11/12
   'end 2018/09/18
   If (m_bOpen = True And Left(strDeptNo, 2) = "P1") Or (m_bOpen = True And strSN12 = strUserNum) Then
        Frame2.Visible = True
   End If
   If (m_bInsert = True And Left(strDeptNo, 2) = "P1") Or (m_bInsert = True And strSN12 = strUserNum) Then
        cmdPrint.Visible = True
        CboChoose.Visible = True
   End If
   'end 2020/12/15
End Sub

'設定主席 lstUsers/副本 lstCC
'Modify by Amy 2022/01/05 原:As ListBox->Object
Private Sub SetList(oList As Object, ByVal stData As String, Optional ByVal bolOld As Boolean = False)
    Dim arrTmp
    Dim stSign As String '替換符號
    
    oList.Clear
    If stData = MsgText(601) Then Exit Sub
    '舊資料抓SN02(存文字)
    If bolOld = True Then
        stSign = "、"
        If InStr(stData, vbCrLf) > 0 Then
            arrTmp = Split(stData, vbCrLf)
        Else
            arrTmp = Split(stData, stSign)
        End If
        For i = UBound(arrTmp) To LBound(arrTmp) Step -1
            oList.AddItem arrTmp(i), 0
        Next i
    '新資料(存員編)
    Else
        arrTmp = Split(stData, ";")
        For i = LBound(arrTmp) To UBound(arrTmp)
            'Modify by Amy 2021/01/26 原取顯示 GetJobTitle(arrTmp(i), 1) & "(" & arrTmp(i) & ")" 改至GetSortNo做
            'patent 信箱顯示王副總
            If arrTmp(i) = "patent@taie.com.tw" Then
                arrTmp(i) = "Patent"
            Else
                arrTmp(i) = arrTmp(i)
            End If
            oList.AddItem arrTmp(i), 0
        Next i
    End If
End Sub

'設定多筆查詢部門
Private Sub SetTxtDeptX()
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String, strStartD As String, strEndD As String
    Dim intQ As Integer
    
    strQ = "Select 1 as C0,Min(a0901) From Acc090 Where a0911='" & strDeptNo & "' " & _
    "Union Select 2 as C0,Max(a0901) From Acc090 Where a0911='" & strDeptNo & "' " & _
              "Order by C0"
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        Do While RsQ.EOF = False
            If RsQ.Fields("C0") = 1 Then
                strStartD = "" & RsQ.Fields(1)
            Else
                strEndD = "" & RsQ.Fields(1)
            End If
            RsQ.MoveNext
        Loop
    End If
    If strStartD = MsgText(601) Then
        If Left(strDeptNo, 1) = "S" Then
            strStartD = "S00"
        Else
            strStartD = strDeptNo
        End If
    End If
    If strEndD = MsgText(601) Then
        If Left(strDeptNo, 1) = "S" Then
            strEndD = "S49"
        Else
            strEndD = strDeptNo
        End If
    End If
    txtDept(0) = strStartD
    txtDept(1) = strEndD
End Sub

'判斷是否有舊資料,strGetPList不為空回傳人員資料
'Modify by Amy 2022/01/05 原As ListBox->Object
Private Function ChkOldListData(oList As Object, ByVal bolShowMsg As Boolean, Optional ByRef strGetPList As String = "") As Boolean
    Dim bolBackList As Boolean
    Dim strData As String, strMsg As String, strNo As String
    
    If strGetPList <> MsgText(601) Then
        bolBackList = True
        strGetPList = ""
    End If
    
    For i = 0 To oList.ListCount - 1
         strData = oList.List(i)
        '舊資料
        If InStr(strData, "(") = 0 And InStr(strData, "Patent") = 0 Then
            strMsg = strMsg & ";" & strData
            ChkOldListData = True
        ElseIf InStr(strData, "Patent") > 0 Then
            strGetPList = strGetPList & ";" & "Patent"
        Else
            strNo = Mid(oList.List(i), InStr(oList.List(i), "(") + 1)
            strGetPList = strGetPList & ";" & Mid(strNo, 1, InStr(strNo, ")") - 1)
        End If
    Next i
    If ChkOldListData = True Then
        If bolShowMsg = True Then
            strMsg = Replace(Mid(strMsg, 2), ";", Chr(13) & Chr(10)) & vbCrLf & _
                        "為舊資料請刪除,重新選擇人員"
            MsgBox strMsg, , "警告"
        End If
    End If
    If strGetPList <> MsgText(601) Then strGetPList = Replace(Mid(strGetPList, 2), "Patent", "patent@taie.com.tw")
End Function

'Add by Amy 2021/01/26 調整顯示ListBox 的內容 for 寄信
'Modidify by Amy 2022/01/05  原:As ListBox->object
Private Function GetMailTxt(oList As Object) As String
    Dim strData As String, strDataNew, intCount As Integer
    Dim arrTmp
    
    GetMailTxt = ""
    Call ChkOldListData(oList, False, strData)
    If strData <> MsgText(601) Then
        strDataNew = GetSortNo(strData, IIf(UCase(oList.Name) = "LSTCC", True, False), 1, False)
    End If
    strData = ""
    If InStr(strDataNew, ";") > 0 Then
        arrTmp = Split(strDataNew, ";")
        For i = LBound(arrTmp) To UBound(arrTmp)
            intCount = intCount + 1
            strData = Trim(arrTmp(i)) & "、"
            '顯示5個人換行
            If intCount = 5 Then
                strData = Mid(strData, 1, Val(Len(strData)) - 1) & "<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
                intCount = 0
            End If
            GetMailTxt = GetMailTxt & strData
        Next i
    Else
        GetMailTxt = strDataNew
    End If
    If Right(GetMailTxt, 1) = "、" Then GetMailTxt = Mid(GetMailTxt, 1, Val(Len(GetMailTxt)) - 1)
End Function

'顯示ListBox 的內容 for 寄信
'Mark by Amy 2021/01/26 因未存檔就發測式信,副本順序仍需排且避免取字有錯,改新版本
Private Function GetMailTxt_Old(oList As ListBox) As String
'    Dim strData As String, intCount As Integer
'
'    For i = 0 To oList.ListCount - 1
'        strData = oList.List(i)
'        intCount = intCount + 1
'        If InStr(strData, "(") > 0 Then
'            GetMailTxt = GetMailTxt & IIf(intCount = 1, "", "、") & Trim(Mid(strData, 1, InStr(strData, "(") - 1))
'        Else
'            GetMailTxt = GetMailTxt & IIf(intCount = 1, "", "、") & Trim(strData)
'        End If
'        '顯示5個人換行
'        If intCount = 5 Then
'            GetMailTxt = GetMailTxt & "<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
'            intCount = 0
'        End If
'    Next i
'    If GetMailTxt <> MsgText(601) And oList.ListCount > 1 Then GetMailTxt = Mid(GetMailTxt, 2)
End Function

'員編排序
'stSortNo:員工編號/IsCC:是副本欄位
'intChoose:0-傳回姓名稱+職稱+(編號)/1-回傳姓名+職稱 'Add by Amy 2021/01/26
'bolDesc:反序排
Private Function GetSortNo(stSortNo As String, IsCC As Boolean, Optional ByVal intChoose As Integer = 0, Optional bolDesc As Boolean = True) As String
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String, stTmp(1) As String
    Dim intQ As Integer
    Dim arrTp
    Dim stJobTitle As String, stName As String, stNotFixSort As String 'Add by Amy 2021/01/26
    
    'Modify by Amy 2021/01/26 +名稱及職稱,非固定副本也用職稱、員編排
    GetSortNo = ""
    stNotFixSort = ";" & stSortNo
    stTmp(1) = stNotFixSort
    '主席
    'Modify by Amy 2024/04/19 職稱第一個字為[代] 拿掉 ex:李柏翰 代經理
    If IsCC = False Then
        strQ = "Select ST01,ST02,Decode(SubStr(AC03,1,1),'代',SubStr(AC03,2,length(AC03)),AC03) as AC03,'1'||AC02||ST01 as Sort From Staff,AllCode " & _
                    "Where ST01 In ('" & Replace(stSortNo, ";", "','") & "') And ac02(+)=st20 And ac01(+)='01' Order by Sort " & IIf(bolDesc = True, "Desc", "")
    '副本
    Else
        '過濾出非副本收受者
        arrTp = Split(strCC_Fix, ";")
        For i = LBound(arrTp) To UBound(arrTp)
            stTmp(1) = Replace(stTmp(1), ";" & arrTp(i), "")
        Next i
        If Left(stTmp(1), 1) = ";" Then stTmp(1) = Mid(stTmp(1), 2)
     
        'Modify by Amy 2020/11/27 固定收受者改職稱,員編排;其他員編排
        If stTmp(1) <> MsgText(601) Then
            strQ = "Union Select ST01,ST02,Decode(SubStr(AC03,1,1),'代',SubStr(AC03,2,length(AC03)),AC03) as AC03,'2'||AC02||ST01 as Sort " & _
                        "From Staff,AllCode Where ST01 In ('" & Replace(stTmp(1), ";", "','") & "') And ac02(+)=st20 And ac01(+)='01' "
        End If
        '固定收受者
         strQ = "Select ST01,ST02,Decode(SubStr(AC03,1,1),'代',SubStr(AC03,2,length(AC03)),AC03) as AC03,'1'||AC02||ST01 as Sort From Staff,AllCode " & _
                    "Where ST01 In ('" & Replace(strCC_Fix, ";", "','") & "') And ac02(+)=st20 And ac01(+)='01' " & strQ & _
                    "Order by Sort " & IIf(bolDesc = True, "Desc", "")
        'end 2020/11/27
    End If
    'end 2024/04/19
    
    stTmp(0) = "": stTmp(1) = ""
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        RsQ.MoveFirst
        Do While RsQ.EOF = False
            stTmp(1) = "" & RsQ.Fields("ST01")
            stName = "" & RsQ.Fields("ST02")
            stJobTitle = Replace(Replace("" & RsQ.Fields("AC03"), "副總經理", "副總"), "主任祕書", "主秘")
            '固定副本,只取姓+職稱,但董事長不需+姓
            If InStr(strCC_Fix, stTmp(1)) > 0 Or stJobTitle = "副總" Then
                stName = Left(stName, 1)
            End If
            If stJobTitle = "董事長" Then
                GetSortNo = GetSortNo & ";" & stJobTitle
            Else
                GetSortNo = GetSortNo & ";" & stName & stJobTitle
            End If
            If intChoose = 0 Then GetSortNo = GetSortNo & "(" & stTmp(1) & ")"
            '員工編號查不到 ex:patent@taie.com.tw
            stNotFixSort = Replace(stNotFixSort, ";" & RsQ.Fields("ST01"), "")
            RsQ.MoveNext
        Loop
    End If
    RsQ.Close
    GetSortNo = GetSortNo & ";" & IIf(Left(stNotFixSort, 1) = ";", Mid(stNotFixSort, 2), stNotFixSort)
    '去除前後;
    If Left(GetSortNo, 1) = ";" Then GetSortNo = Mid(GetSortNo, 2)
    If Right(GetSortNo, 1) = ";" Then GetSortNo = Mid(GetSortNo, 1, Val(Len(GetSortNo) - 1))
    'end 2021/01/26
End Function

'取得議題已有登記人員名單
Private Function GetSB02(stSB03 As String) As String
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String, intQ As Integer
    
    strQ = "Select ST02||' ('||SB02||')' From SeminarBookin,Staff " & _
                "Where SB01=" & Text1(1) & " And InStr(','||sb03||',',','||'" & stSB03 & "'||',')>0 And SB02=ST01(+)"
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        RsQ.MoveFirst
        Do While RsQ.EOF = False
            GetSB02 = GetSB02 & ";" & RsQ.Fields(0)
            RsQ.MoveNext
        Loop
    End If
    If GetSB02 <> MsgText(601) Then GetSB02 = Mid(GetSB02, 2)
    RsQ.Close
End Function

Private Function GetSC03() As String
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String, intQ As Integer
    Dim bolOpen As Boolean 'Add by Amy 2020/11/27
    
    'Modify by Amy 2020/11/27
    bolOpen = Option1(0).Value
    strQ = GetSeminarContactSql(2, strDeptNo, bolOpen)
    'end 2020/11/27
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        GetSC03 = "" & RsQ.Fields("SC03")
    End If
    RsQ.Close
    Set RsQ = Nothing
End Function

Private Function GetMailMemo() As String
    'Memo 發信時固定收件人改至 副本-與雅娟討論後
    '原:"1.總經理、薛經理為固定收件人，所長、副所長為固定副本" & vbCrLf & "　收受人，皆不另行顯示。" & vbCrLf & _
          "2.專利處王副總、游經理、郭副理不論是否選擇也都會寄發。"
    Dim stMailTo As String, stMailCC As String
    Dim stMailSpeaker As String 'Add by Amy 2020/12/28
    
    stMailTo = "Y": stMailCC = "Y"
    'Modify by Amy 2020/12/28
    stMailSpeaker = "Y"
    Call SendInformMail("", stMailTo, stMailCC, , stMailSpeaker)
    'GetMailMemo = "" & _
            "收件者：" & stMailTo & vbCrLf & _
            "副　本：" & stMailCC & vbCrLf
    GetMailMemo = stMailTo & stMailCC & stMailSpeaker
    'end 2020/12/28
End Function
'end 2018/09/18

'Add by Amy 2019/01/24
Private Function ChkSB02All() As Boolean
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String, intQ As Integer
    
    ChkSB02All = False
    If ActionEdit = 0 Then Exit Function
    strQ = "Select * From SeminarBookin Where sb01=" & Text1(1) & " And sb02='ALL999' "
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        ChkSB02All = True
    End If
    RsQ.Close
End Function

Private Sub SetCboEmp(ByVal strManagerNo As String)
    Dim arrTmp
    
    cboEmp.Clear
    arrTmp = Split(strManagerNo, ";")
    For i = LBound(arrTmp) To UBound(arrTmp)
        'Mark by Amy 2021/01/26 職稱改GetSortNop 做
'        'patent 信箱顯示王副總
'        If arrTmp(i) = "patent@taie.com.tw" Then
'            arrTmp(i) = GetJobTitle("71011", 1) & "(Patent)"
'        Else
'            arrTmp(i) = GetJobTitle(arrTmp(i), 1) & "(" & arrTmp(i) & ")"
'        End If
        cboEmp.AddItem arrTmp(i), i
    Next i
    
End Sub

Private Sub SetCboRoom()
    cboRoom.Clear
    strExc(0) = "Select mr03||' '||Replace(mr02,'五樓','') as mr02,mr01 From MeetingRoom Where mr04='1' Order by mr01 Desc"
 
    intI = 1
    Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
    If intI = 1 Then
        With RsTemp
            Do While Not .EOF
                cboRoom.AddItem "台北所" & " " & .Fields("mr02"), 0
                cboRoom.ITEMDATA(0) = .Fields("mr01")
                .MoveNext
            Loop
        End With
    End If
    cboRoom.AddItem "", 0
End Sub

Private Function ChkOtherRoom(stSN22 As String) As Boolean
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String, stMsg As String
    Dim intQ As Integer
   
    ChkOtherRoom = False
    If ActionEdit = 1 Then strQ = " And SN01<>" & Val(Text1(1))
    'Modify by Amy 2019/11/12 改元件
    strQ = "Select SN05,SN06,SN07 From Seminar,Meetingroom Where SN22=" & Val(stSN22) & strQ & _
                " And SN05=" & Val(DBDATE(MaskEdBox1)) & " And SN01=MR01(+)" & _
                " And " & Val(Format(cboTime(0), "HHmm")) & ">=SN06 And " & Val(Format(cboTime(0), "HHmm")) & "<SN07 " & _
    "Union Select SN05,SN06,SN07 From Seminar,Meetingroom Where SN22=" & Val(stSN22) & strQ & _
                " And SN05=" & Val(DBDATE(MaskEdBox1)) & " And SN01=MR01(+)" & _
                " And " & Val(Format(cboTime(1), "HHmm")) & ">SN06 And " & Val(Format(cboTime(1), "HHmm")) & "<=SN07 " & _
    "Union Select SN05,SN06,SN07 From Seminar,Meetingroom Where SN22=" & Val(stSN22) & strQ & _
                " And SN05=" & Val(DBDATE(MaskEdBox1)) & " And SN01=MR01(+)" & _
                " And " & Val(Format(cboTime(0), "HHmm")) & "<SN06 And " & Val(Format(cboTime(1), "HHmm")) & ">=SN07 "

    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        Do While RsQ.EOF = False
            stMsg = stMsg & ";" & RsQ.Fields("SN05") & "　" & Format(RsQ.Fields("SN06"), "00:00") & _
                        "　" & Format(RsQ.Fields("SN07"), "00:00")
            RsQ.MoveNext
        Loop
        If stMsg <> MsgText(601) Then
            ChkOtherRoom = True
            stMsg = Replace(Mid(stMsg, 2), ";", vbCrLf)
            stMsg = "　日期　　起始 　結束" & vbCrLf & stMsg
            MsgBox stMsg, , cboRoom & "已預約清單"
        End If
    End If
    RsQ.Close
End Function

Private Function CheckKeyIn(obj As Object) As Integer
    Dim stRoom As String, stTmp1 As String, stTmp2 As String
    
    CheckKeyIn = 0
    Select Case UCase(obj.Name)
        Case "CBOEMP", "TEXT1"
            '可輸入員編或姓名(主席/副本)
            If ByInputGetST01or02(obj.Text, stTmp1, stTmp2) = False Then
                CheckKeyIn = -1
                obj.SetFocus
                Exit Function
            End If
            obj.Text = stTmp1
            If UCase(obj.Name) = "TEXT1" Then
                lblCC = stTmp2
            Else
                lblChairMan = stTmp2
            End If
        '*** Memo by Amy 開始-下面範圍有改,請確認是否frm140112_1 TxtValidate是否也要改
        'Modify by Amy 2019/11/12 DTPICKER1(0)改為TxtDate
        Case "MASKEDBOX1"
            If Not (ActionEdit = 0 Or ActionEdit = 1) Then Exit Function
            If cboRoom = MsgText(601) Then
                CheckKeyIn = -1
                MsgBox "地點不可為空值！", vbCritical
                cboRoom.SetFocus
                Exit Function
            End If
            '日期
            If MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = "____/__/__" Then
                CheckKeyIn = -1
                MsgBox "請輸入預約日期！", vbExclamation
                MaskEdBox1.SetFocus
                Exit Function
            End If
            stTmp1 = Val(DBDATE(MaskEdBox1))
            If ChkDate(stTmp1) = False Then
                CheckKeyIn = -1
                Exit Function
            End If
            '新增
            If ActionEdit = 0 Then
                If DBDATE(stTmp1) < strSrvDate(1) Then
                    CheckKeyIn = -1
                    MsgBox "預約日期已經過了！", vbCritical
                    Exit Function
                End If
                '改4個月(原2個月)--經理(同frm140112_1 TxtValidate)
                stTmp2 = CompDate(1, 4, strSrvDate(1))
                If DBDATE(stTmp1) > stTmp2 Then
                    CheckKeyIn = -1
                    MsgBox "預約日期請輸入 4 個月以內的日期！"
                    Exit Function
                End If
            End If
            'end 2019/11/12
        Case "CBOTIME"
            If Not (ActionEdit = 0 Or ActionEdit = 1) Then Exit Function
            'Modify by Amy 2020/02/06 拿掉Val() 避免後面時間判斷有問題 ex:早上Val(0800)
            stTmp1 = Format(cboTime(obj.Index), "HHmm") & "00"
            If Val(stTmp1) = 0 Then
                CheckKeyIn = -1
                MsgBox "請輸入" & IIf(obj.Index = 0, "開始", "結束") & "時間！", vbExclamation
                Exit Function
            End If
            '新增
            If ActionEdit = 0 Then
                If Val(DBDATE(MaskEdBox1) & stTmp1) < Val(strSrvDate(1) & Format(ServerTime, "000000")) Then
                    CheckKeyIn = -1
                    MsgBox IIf(obj.Index = 0, "開始", "結束") & "時間已過請重新輸入！", vbCritical
                    Exit Function
                End If
            End If
            '結束時間
            If obj.Index = 1 Then
                stTmp1 = Format(cboTime(0), "HHmm") & "00"
                stTmp2 = Format(cboTime(1), "HHmm") & "00"
                If stTmp1 >= stTmp2 Then
                    CheckKeyIn = -1
                    MsgBox "結束時間必須晚於開始時間！", vbCritical
                    Exit Function
                End If
            End If
        '*** End 結束-上面範圍有改,請確認是否frm140112_1 TxtValidate是否也要改
    End Select
    
    If UCase(obj.Name) <> "CBOTIME" Then Exit Function
    '結束時間
    If obj.Index = 1 Then
        '檢查預約是否有重疊
        stTmp1 = "Y"
        stRoom = GetMeetingRoom(cboRoom, False)
        '北所小會議室
        If Val(stRoom) >= 3 Then
            If ChkOtherRoom(stRoom) = True Then
                CheckKeyIn = -1
                'Modify by Amy 2019/11/12 改元件
                MaskEdBox1.SetFocus '跳至日期欄(避免一直彈已預約訊息無法跳離)
                Exit Function
            End If
        End If
    End If
    
End Function

'刪除或還原會議室預約記錄
Private Sub ReCoverRR(ByVal intCmd As Integer)
    Dim stSQL As String, stField As String
    
    Select Case intCmd
        Case 0 '刪除
            stSQL = "Delete RoomReservation Where RR20=" & Val(Text1(1))
        Case 1 '新增
            For i = LBound(strOldRR) To UBound(strOldRR)
                If i >= 5 And i <= 10 And i <> 6 Then
                   stField = stField & ",RR" & Format(i, "00")
                   stSQL = stSQL & "," & CNULL(ChgSQL(strOldRR(i)))
                ElseIf i <> 6 Then
                    stField = stField & ",RR" & Format(i, "00")
                    stSQL = stSQL & "," & strOldRR(i)
                End If
            Next i
            stSQL = "Insert into RoomReservation (RR20" & stField & ") Values(" & Val(Text1(1)) & stSQL & ")"
        Case 2 '修改
            For i = LBound(strOldRR) To UBound(strOldRR)
                If i >= 5 And i <= 10 And i <> 6 Then
                   stSQL = stSQL & ",RR" & Format(i, "00") & "=" & CNULL(ChgSQL(strOldRR(i)))
                ElseIf i <> 6 Then
                    stSQL = stSQL & ",RR" & Format(i, "00") & "=" & strOldRR(i)
                End If
            Next i
            stSQL = "Update RoomReservation set  " & Mid(stSQL, 2) & " Where RR20=" & Val(Text1(1))
    End Select
  
    If stSQL <> MsgText(601) Then cnnConnection.Execute stSQL
End Sub

Private Sub SetRoomReservation()
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String, intQ As Integer
    
    For i = LBound(strOldRR) To UBound(strOldRR)
        strOldRR(i) = ""
    Next i
        
    strQ = "Select * From RoomReservation Where RR20=" & Val(Text1(1))
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        For i = LBound(strOldRR) To UBound(strOldRR)
            strOldRR(i) = "" & RsQ.Fields(i - 1)
        Next i
    End If
    RsQ.Close
End Sub

Private Sub SetCombo()
    Dim ii As Integer
    For ii = 0 To 23
        cboTime(0).AddItem Format(ii, "00") & ":" & "00"
        cboTime(1).AddItem Format(ii, "00") & ":" & "30"
        cboTime(0).AddItem Format(ii, "00") & ":" & "30"
        cboTime(1).AddItem Format(ii + 1, "00") & ":" & "00"
    Next
End Sub

'stDate:編號或名稱 /bolNo:傳入為編號
'傳入為編號回名稱,傳入為名稱回傳編號
Public Function GetMeetingRoom(ByVal stData As String, bolNo As Boolean) As String
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String, intQ As Integer
    
    GetMeetingRoom = ""
    strQ = "Where MR01=" & Val(stData)
    If bolNo = False Then
        strQ = "Where MR02='" & Replace(Replace(Replace(stData, "台北所 ", ""), "9F ", ""), "5F ", "五樓") & "' "
    End If
    strQ = "Select * From MeetingRoom " & strQ
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        If bolNo = True Then
            GetMeetingRoom = "台北所 " & RsQ.Fields("MR03") & " " & Replace(RsQ.Fields("MR02"), "五樓", "")
        Else
            GetMeetingRoom = "" & RsQ.Fields("MR01")
        End If
    End If
    RsQ.Close
End Function

'Mark by Amy 2019/11/12 限制同部門主席、副本收授者、參加者可以查
Private Sub RsAction_Old(ByVal pCmd As Integer)
'On Error GoTo ErrHand
'   Screen.MousePointer = vbHourglass
'   intI = 1
'   Select Case pCmd
'      Case 0 '第一筆
'         strExc(0) = "SELECT nvl(min(SN01),0) FROM Seminar"
'         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'         If intI = 1 Then
'            If RsTemp.Fields(0) > 0 Then
'               ReadData RsTemp.Fields(0)
'            End If
'         End If
'
'      Case 1 '前一筆
'         strExc(0) = "SELECT nvl(max(SN01)," & Val(Text1(1)) & ") FROM Seminar where SN01<" & Val(Text1(1))
'         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'         If intI = 1 Then
'            If RsTemp.Fields(0) = Val(Text1(1)) Then
'               DataErrorMessage 6
'            Else
'               ReadData RsTemp.Fields(0)
'            End If
'         End If
'
'      Case 2 '後一筆
'         strExc(0) = "SELECT nvl(min(SN01)," & Val(Text1(1)) & ") FROM Seminar where SN01>" & Val(Text1(1))
'         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'         If intI = 1 Then
'            If RsTemp.Fields(0) = Val(Text1(1)) Then
'               DataErrorMessage 7
'            Else
'               ReadData RsTemp.Fields(0)
'            End If
'         End If
'
'      Case 3 '最後筆
'         strExc(0) = "SELECT nvl(max(SN01),0) FROM Seminar"
'         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'         If intI = 1 Then
'            If RsTemp.Fields(0) > 0 Then
'               ReadData RsTemp.Fields(0)
'            End If
'         End If
'   End Select
'   Screen.MousePointer = vbDefault
'   Exit Sub
'
'ErrHand:
'   Screen.MousePointer = vbDefault
'   MsgBox "錯誤 : " & Err.Description, vbCritical
End Sub

'判斷傳入之編號是建立者同部門、主席、副本人員或參加人員,才可以查詢
Private Function ChkSeminarLimit(ByVal stSN01 As String) As Boolean
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String, intQ As Integer
    
    ChkSeminarLimit = False
    strQ = "Select * From Seminar Where SN01 in (" & strSeminar & ") And SN01=" & Val(stSN01)
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 0 Then
        '不在有權限之編號資料中
        RsQ.Close
        Exit Function
    End If
    ChkSeminarLimit = True
    RsQ.Close
End Function

'傳入畫面會議室/日期/起迄日期
Private Sub Show140112(stRoom As String, stDate As String, stTimeS As String, stTimeE As String, bolShowB As Boolean)
    'Add by Amy 2020/01/14
    Dim bolShowHoliday As Boolean
    
    If Right(lblWeek, 1) = "六" Or Right(lblWeek, 1) = "日" Then
        '新增/修改
        If ActionEdit <= 1 Then
            If MsgBox("日期為假日,繼續操作？", vbYesNo, MsgText(5)) = vbNo Then
                Exit Sub
            End If
        End If
        bolShowHoliday = True
    End If
    'end 2020/01/14
    
    Me.Enabled = False
    If fnSaveParentForm(Me) = False Then
        Me.Enabled = True
        Exit Sub
    End If
    frm140112.m_SN01 = Me.Text1(1)
    frm140112.m_Title = Me.Text1(18)
    frm140112.stOldRR1 = stRoom
    frm140112.stOldRR2 = stDate
    frm140112.stOldRR3 = stTimeS
    frm140112.stOldRR4 = stTimeE
    frm140112.bolReadOnly = IIf(ActionEdit >= 2, True, False)
    frm140112.m_SN12 = strSN12 '建立者
    frm140112.Tag = strDeptNo
    'Add by Amy 2020/01/14 假日需勾選才會顯示
    If bolShowHoliday = True Then
        frm140112.Check1 = 1
        frm140112.cmdMove_Click (2)
    End If
    'end 2020/01/14
    frm140112.Show
    Me.Enabled = True
End Sub

'Add by Amy 2020/01/14 取得自動編號(上線前舊資料不變,上線後 年+4碼流水號)
Private Function GetSaveAutoNo() As String
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String, intQ As Integer
    Dim strUpd As String, strAU02 As String, strAU03 As String
    
    GetSaveAutoNo = ""
    strQ = "Select * From AutoNumber Where au01='SS' "
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        strAU02 = "" & RsQ.Fields("au02")
        strAU03 = "" & RsQ.Fields("au03")
        If strAU03 = "9999" Then
            MsgBox "流水號已達9999請洽電腦中心！"
        Else
            If strAU02 = Mid(strSrvDate(1), 1, 4) Then
                strAU03 = Val(strAU03) + 1
            '新一年重編流水號
            Else
                strAU02 = Mid(strSrvDate(1), 1, 4)
                strAU03 = "1"
                strUpd = ",au02=" & strAU02
            End If
            strUpd = "Update AutoNumber Set au03=" & strAU03 & strUpd & " Where au01='SS' "
            adoTaie.Execute strUpd
            GetSaveAutoNo = Val(strAU02) - 1911 & String(4 - Len(strAU03), "0") & strAU03
        End If
    End If
    RsQ.Close
End Function

'確認預設主席是否被拿掉,W部門預設主席為區主管(只設於W1001/W2001),若主席修改,其區主管看不到此筆資料
Private Function ChkDefSC04Exists() As Boolean
    Dim RsQ As New ADODB.Recordset
    Dim strSql As String, stNowSC04 As String
    Dim intQ As Integer, i As Integer
    Dim bolOpen As Boolean 'Add by Amy 2020/11/27

    Call ChkOldListData(lstUsers, False, stNowSC04) '主席(員編)
    'Modify by Amy 2020/11/27
    bolOpen = Option1(0)
    strSql = GetSeminarContactSql(1, strDeptNo, bolOpen)
    'end 2020/11/27
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strSql)
    If intQ = 1 Then
        If "" & RsQ.Fields("SC04") <> MsgText(601) Then
            '知會主管=stNowSC04
            If InStr(stNowSC04, RsQ.Fields("SC04")) > 0 Then
                ChkDefSC04Exists = True
            End If
        End If
    End If
    RsQ.Close
End Function

'Add by Amy 2020/01/14 取消與結束 預約會議室 資料還原(避免會議室預約返回後取消)
Private Sub ChkRRAndReCover()
    Dim strTp(3) As String
    Dim bolHasRR As Boolean '已預約會議室
    
    If Not (ActionEdit = 0 Or ActionEdit = 1) Then Exit Sub
    
    bolHasRR = ChkHasRR20(Val(Text1(1)), strTp(0), strTp(1), strTp(2), strTp(3), True)
    '新增
    If ActionEdit = 0 Then
        '有預約資料,刪除預約
        If bolHasRR = True Then
            Call ReCoverRR(0)
        End If
    '修改
    ElseIf strOldRR(1) & strOldRR(2) & strOldRR(3) & strOldRR(4) <> strTp(0) & strTp(1) & strTp(2) & strTp(3) Then
        '目前有預約
        If bolHasRR = True Then
            '原 未預約-刪除目前預約
            If strOldRR(1) = MsgText(601) Then
                Call ReCoverRR(0)
            '原 有預約-還原舊預約
            Else
                Call ReCoverRR(2)
            End If
        '目前無預約且有舊預約-還原舊預約
        ElseIf strOldRR(1) <> MsgText(601) Then
            Call ReCoverRR(1)
        End If
    End If
End Sub

'Add by Amy 2020/11/27 從ReadData搬過來登記頁籤-登記人員資料
Private Sub CheckInList(ByVal pKey As String)
    Dim ii As Integer, iMyCol As Integer, iCol As Integer
    
    '語法改至GetCheckInSql修改
    strExc(0) = GetCheckInSql(pKey, False)
    intI = 1
    Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
    If intI = 1 Then
        With RsTemp
        Do While Not .EOF
            MSHFlexGrid2.Cols = MSHFlexGrid2.Cols + 1
            iCol = MSHFlexGrid2.Cols - 1
            MSHFlexGrid2.ColAlignment(iCol) = flexAlignCenterCenter
            MSHFlexGrid2.TextMatrix(0, iCol) = "" & .Fields("st02")
            'Add by Amy 2021/01/28 非同部門設員工姓名設顏色
            If Left(strDeptNo, 1) = "S" Then
                If Left("" & .Fields("st15"), 1) <> "S" Then
                    If Left(PUB_GetST03("" & .Fields("st01")), 1) <> "S" Then
                        MSHFlexGrid2.row = 0
                        MSHFlexGrid2.col = iCol
                        MSHFlexGrid2.CellBackColor = &HFFFF80
                    End If
                End If
            Else
                If Left("" & .Fields("st15"), 2) <> Left(strDeptNo, 2) Then
                    If Left(PUB_GetST03("" & .Fields("st01")), 2) <> Left(strDeptNo, 2) Then
                        MSHFlexGrid2.row = 0
                        MSHFlexGrid2.col = iCol
                        MSHFlexGrid2.CellBackColor = &HFFFF80
                    End If
                End If
            End If
            'end 2021/01/28
            
            If Not IsNull(.Fields("sb02")) Then '有登記紀錄的
                MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 2, iCol) = "" & .Fields("sb03")
                If Not IsNull(.Fields("sb03")) Then
                    If .Fields("sb03") = "0" Then
                        MSHFlexGrid2.TextMatrix(1, iCol) = "V"
                    Else
                        For ii = 2 To MSHFlexGrid2.Rows - 5
                             If InStr("," & .Fields("sb03") & ",", "," & MSHFlexGrid2.TextMatrix(ii, 1) & ",") > 0 Then
                                 MSHFlexGrid2.TextMatrix(ii, iCol) = "V"
                            'Added by Morgan 2012/5/3
                            '沒有設定可登記的議題變色
                             Else
                                 If InStr("," & .Fields("sb07") & ",", "," & MSHFlexGrid2.TextMatrix(ii, 1) & ",") = 0 Then
                                     MSHFlexGrid2.row = ii
                                     MSHFlexGrid2.col = iCol
                                     MSHFlexGrid2.CellBackColor = MSHFlexGrid2.BackColorFixed
                                 End If
                            'end 2012/5/3
                             End If
                        Next
                    End If
                'Added by Morgan 2012/5/3
                '沒有設定可登記的議題變色
                Else
                    For ii = 2 To MSHFlexGrid2.Rows - 5
                        If InStr("," & .Fields("sb07") & ",", "," & MSHFlexGrid2.TextMatrix(ii, 1) & ",") = 0 Then
                            MSHFlexGrid2.row = ii
                            MSHFlexGrid2.col = iCol
                            MSHFlexGrid2.CellBackColor = MSHFlexGrid2.BackColorFixed
                        End If
                    Next
                'end 2012/5/3
                End If
                MSHFlexGrid2.ColWidth(iCol) = 350
           'Added by Morgan 2012/5/15
           '其他要列印人員(經副理人員放在最後)
           Else
                MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 2, iCol) = "X"
                'MSHFlexGrid2.ColWidth(iCol) = 0 'Mark by Amy 2020/11/27 經副理需顯示並放於最後
           End If
           MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 1, iCol) = MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 2, iCol)
           MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 3, iCol) = .Fields("st01")
           MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 4, iCol) = .Fields("st06")
           If .Fields("st01") = strUserNum Then
                iMyCol = iCol
                'If MSHFlexGrid2.ColWidth(iCol) > 0 Then bolJoin = True 'Add by Amy 2018/09/18
           End If
           .MoveNext
        Loop
        End With
    End If
    If iMyCol > 0 Then
        MSHFlexGrid2.LeftCol = iMyCol
    End If
    '**** End 登記人員 ***
End Sub

Private Function GetCheckInSql(ByVal pKey As String, Optional ByVal bolRollCallList As Boolean = False, Optional ByVal strWhere As String = "")
    Dim stTB As String
    Dim strWhere2(1) As String 'Add by Amy 2021/01/28
    
    '**** 登記人員 ***
    '順序:自己,已登記者,其他專利處工程師
    'Modified by Morgan 2012/5/1
    'strExc(0) = "select ST01,ST02,min(C1) X1,max(sb02) sb02,max(sb03) sb03 from (" & _
        " select st01,st02,2 C1,sb02,sb03 from SeminarBookin A,staff where sb01=" & pKey & " and st01(+)=sb02"
    'If m_bUpdate Or m_IsOpen Then
    '   strExc(0) = strExc(0) & " Union All select st01,st02,1 C1,'','' from staff where st01='" & strUserNum & "'"
    'End If
    'If m_bUpdate Then
    '   strExc(0) = strExc(0) & " Union All select st01,st02,3 C1,'','' from staff where st03='P11' and st04='1'"
    'End If
    'strExc(0) = strExc(0) & ") group by st01,st02 order by X1,st02"
     
    'Mark by 2020/11/27 因杜燕文經理為專利處及智權部身份,名單未出現,改至下面判斷
    'Modify by Amy 2019/01/24 排除登記人員為 ALL999,開放其他部門只有P1才加st20條件
'    strWhere = " And st15 like '" & Left(strDeptNo, IIf(Left(strDeptNo, 1) = "S", 1, 2)) & "%' "
'    If Left(strDeptNo, 2) = "P1" Then
'        strWhere = strWhere & " and st20<='44' "
'    End If
    'end 2020/11/27
     
    'Modify by Amy 2021/01/28 後加入的不同部門(ex:94012)且為副理會出現於雅娟部門經副理list中間,改為其他部門以部門職稱員編排最後
    If Left(strDeptNo, 1) = "S" Then
        strWhere2(0) = strWhere2(0) & "And (Substr(st15,1,1)='S' Or Substr(st03,1,1)='S') "
        strWhere2(1) = strWhere2(1) & "And (Substr(st15,1,1)<>'S' And Substr(st03,1,1)<>'S') "
    Else
        strWhere2(0) = strWhere2(0) & "And (Substr(st15,1,2)='" & Left(strDeptNo, 2) & "' Or Substr(st03,1,2)='" & Left(strDeptNo, 2) & "') "
        strWhere2(1) = strWhere2(1) & "And (Substr(st15,1,2)<>'" & Left(strDeptNo, 2) & "' And Substr(st03,1,2)<>'" & Left(strDeptNo, 2) & "') "
    End If
    'Modify by Amy 2020/11/27 有選參加人員,「登記」會出現,點名單也會出現但列於最後
    'Modify by Amy 2019/11/12 部門原抓st03
    '自已部門非經副理級人員
    GetCheckInSql = "Select st01,st02,sb02,sb03,st06,sb07,st15,'1'||st06||'" & strDeptNo & "'||st01 C1 From SeminarBookin A,staff" & stTB & _
                                " Where sb01=" & pKey & " and st01(+)=sb02 And substr(st01,1,3)<>'999' And Nvl(st20,'99')>'44' " & strWhere2(0) & strWhere
    '自已部門代副理以上要印在點名單的後面
    'Modify by Amy 2018/09/18 增加其他部門顯示,排除虛設編號(第4碼為9)及巨京(st06='5')
    '    strExc(0) = strExc(0) & " union select st01,st02,sb02,sb03,st06,sb07,st15,2 C1 " & _
    '        " from staff,SeminarBookin where st04='1' " & strWhere & _
    '        " and st01>'6' and st01<'F' and substr(st01,1,3)<>'999' and sb02(+)=st01 and sb01(+)=" & pKey & _
    '        " and sb02 is null And st06<>'5' And SubStr(st01,4,1)<>'9' "
    '    strExc(0) = strExc(0) & " order by st06,C1,st15,st01"
    GetCheckInSql = GetCheckInSql & _
                    " Union select st01,st02,sb02,sb03,st06,sb07,st15,'2'||st06||st20||st01 C1 From SeminarBookin A,staff" & stTB & _
                    " Where sb01=" & pKey & " and st01(+)=sb02 And substr(st01,1,3)<>'999' And Nvl(st20,'99')<='44' " & strWhere2(0) & strWhere
    '其他部門以部門、職稱、員編排
    GetCheckInSql = GetCheckInSql & _
                    " Union select st01,st02,sb02,sb03,st06,sb07,st15,'3'||st15||st20||st01 C1 From SeminarBookin A,staff" & stTB & _
                    " Where sb01=" & pKey & " and st01(+)=sb02 And substr(st01,1,3)<>'999' " & strWhere2(1) & strWhere
    GetCheckInSql = GetCheckInSql & " order by C1"
    'end 2019/11/12
    'end 2020/11/27
End Function

Private Sub SetCboChoose()
    CboChoose.Clear
    CboChoose.AddItem "全部"
    CboChoose.AddItem "只印自己部門"
    CboChoose.AddItem "只印其他部門"
    
    CboChoose = "全部"
    If Left(strDeptNo, 2) = "P1" Then
        CboChoose = "只印自己部門"
    End If
End Sub

'單名單-依下拉選單畫面條件印點名單,以議題印,經副理級以上列於最後,不同所別分開印
'原抓MSHFlexGrid2印,改用語法印,因可能不印全部
Private Sub runWord()
    Dim RsQ As New ADODB.Recordset, rsA1 As New ADODB.Recordset, rsA2 As New ADODB.Recordset
    Dim strQ As String, strA As String, strWhere As String, strCondition As String
    Dim intQ As Integer, intA1 As Integer, intA2 As Integer
    Dim ii As Integer, jj As Integer
    Dim iRow As Integer, iCols As Integer '目前列/產生欄
    Dim stSubject As String, stTime As String '議題/時間
    Dim strFontSize As String, stOldOffice As String, stOldSubject As String '字大小/前一筆所別/前一筆議題
    Dim iResumeCnt As Integer
    Dim dblColW As Double
    Dim stTmp As String

On Error GoTo ErrHnd
    
    If CboChoose <> "全部" Then
        If Left(strDeptNo, 1) = "S" Then
            strQ = "S"
        Else
            strQ = Left(strDeptNo, 2)
        End If
        strQ = " (st15 Like '" & strQ & "%' Or st03 Like '" & strQ & "%') "
        If CboChoose = "只印其他部門" Then
            strQ = "Not " & strQ
        End If
        strQ = " And " & strQ
        strWhere = strQ
    End If
    
    strQ = "Select Distinct st06 From SeminarBookIn,Staff Where sb01=" & Text1(1) & " And  sb02=st01(+) And sb02<>'ALL999' " & strQ
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 0 Then Exit Sub
   
    If g_WordAp Is Nothing Then Set g_WordAp = New Word.Application
    g_WordAp.Documents.add
    
    With g_WordAp
        .Visible = True
        .Selection.Font.Name = "標楷體"
          
        With .Options
            .DefaultBorderLineStyle = wdLineStyleSingle
            .DefaultBorderLineWidth = wdLineWidth050pt
        End With
    
        .Selection.PageSetup.PaperSize = wdPaperA4
          
        .Selection.PageSetup.Orientation = wdOrientLandscape
        strFontSize = 14
    
        .Selection.Orientation = wdTextOrientationHorizontal
        .Selection.Font.Size = strFontSize
        '邊界
        .Selection.PageSetup.LeftMargin = .CentimetersToPoints(1.6)
        .Selection.PageSetup.RightMargin = .CentimetersToPoints(1.4)
        .Selection.PageSetup.TopMargin = .CentimetersToPoints(1.5)
        .Selection.PageSetup.BottomMargin = .CentimetersToPoints(1)
        '頁尾
        .Selection.PageSetup.HeaderDistance = .CentimetersToPoints(1)
        
        iTables = 0 'Add by Amy 2021/01/25
        dblColW = 0.8: intMaxTB = 3: intMaxCol = 25 '欄寬/一頁最多表格數/一頁最多欄數
        RsQ.MoveFirst
        Do While RsQ.EOF = False
            '所別不同換頁
            If iTables <> 0 Then
                .Selection.MoveDown Unit:=wdLine, Count:=1
                .Selection.SelectRow
                .Selection.MoveRight Unit:=wdCharacter, Count:=2
                .Selection.Font.Size = 10
                .Selection.InsertBreak Type:=wdPageBreak
                '避免不過高刪除一行
                .Selection.MoveUp Unit:=wdLine, Count:=1
                .Selection.TypeBackspace
                .Selection.MoveDown Unit:=wdLine, Count:=1
                .Selection.Font.Size = 14
            End If
            Call SetWordTitle("" & RsQ.Fields("st06"), strFontSize)
            iTables = 0
            
            '*** 議題 ***
            strA = "Select ss01,ss02,ss03,Decode(Length(ss04),4,ss04,'0'||ss04) ss04,Decode(Length(ss05),4,ss05,'0'||ss05) ss05 " & _
                        "From SeminarSubject Where ss01=" & Text1(1)
            intA1 = 1
            Set rsA1 = ClsLawReadRstMsg(intA1, strA)
            If intA1 = 1 Then
                rsA1.MoveFirst
                Do While rsA1.EOF = False
                    '***  參加人員 ***
                    strCondition = " And st06='" & RsQ.Fields("st06") & "' And InStr(','||sb07,','||" & rsA1.Fields("ss02") & ")>0 " & strWhere
                    strQ = GetCheckInSql(Text1(1), True, strCondition)
                    intA2 = 1
                    Set rsA2 = ClsLawReadRstMsg(intA2, strQ)
                    If intA2 = 1 Then
                        rsA2.MoveFirst
                        Do While rsA2.EOF = False
                            stSubject = "" & rsA1.Fields("ss03")
                            stTime = "(" & Left(rsA1.Fields("ss04"), 2) & ":" & Right(rsA1.Fields("ss04"), 2) & "～" & _
                                                    Left(rsA1.Fields("ss05"), 2) & ":" & Right(rsA1.Fields("ss05"), 2) & ")"
                            '表格超過 換頁
                            'Modify by Amy 2021/01/25 原換頁程式搬至function
                            Call SetOverTB("" & RsQ.Fields("st06"), strFontSize)
                            
                            '議題不同/超過25欄 換表格
                            If (stOldSubject <> rsA1.Fields("ss02") And iNowCol = 0) Or iNowCol Mod intMaxCol = 0 Then
                                If iRow = Val(rsA2.RecordCount) \ 25 Then
                                    iCols = Val(rsA2.RecordCount) Mod 25
                                Else
                                    iCols = 25
                                    iRow = iRow + 1
                                End If
                                'Add by Amy 2021/01/25 表格超過 換頁
                                Call SetOverTB("" & RsQ.Fields("st06"), strFontSize, True)
                                'end 2021/01/25
                                If (iTables >= 1 And iNowCol = 0) Or (iNowCol > 0 And iNowCol Mod intMaxCol = 0) Then
                                   .Selection.MoveDown Unit:=wdLine, Count:=1
                                   .Selection.SelectRow
                                   .Selection.MoveRight Unit:=wdCharacter, Count:=2
                                   .Selection.Font.Size = 10
                                   .Selection.TypeParagraph
                                   .Selection.Font.Size = 14
                                End If
                                If iNowCol <> 0 Then
                                    stSubject = "同上"
                                End If
                                
                                Call SetTable(iTables, iCols, stSubject, stTime)
                            End If
                            '往右跳一欄
                            .Selection.MoveRight Unit:=wdCell, Count:=1
                            '往上跳一行
                            .Selection.MoveUp Unit:=wdLine, Count:=1
                            
                            '參加人姓名
                            .Selection.Cells.SetWidth ColumnWidth:=.CentimetersToPoints(dblColW), RulerStyle:=wdAdjustProportional
                            .Selection.TypeText Text:="" & rsA2.Fields("st02")
                            
                            .Selection.MoveDown Unit:=wdLine, Count:=1
                            .Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
                            If InStr("" & rsA2.Fields("sb03"), "" & rsA1.Fields("ss02")) > 0 Then
                                .Selection.TypeText Text:="V"
                            End If
                            
                            iNowCol = iNowCol + 1
                            rsA2.MoveNext
                        Loop
                    End If
                    '***  End 參加人員 ***
                    iNowCol = 0: iRow = 0
                    stOldSubject = "" & rsA1.Fields("ss02")
                    If iTables = intMaxTB Then iTables = iTables + 1
                    rsA1.MoveNext
                Loop
            End If
            '*** End 議題 ***
            iNowCol = 0
            stOldOffice = "" & RsQ.Fields("st06")
            RsQ.MoveNext
        Loop '所別
        .Selection.WholeStory
        .Selection.Font.Name = "Times New Roman"
        .Selection.MoveRight Unit:=wdCharacter, Count:=1
        '插入頁碼
        If .ActiveWindow.View.SplitSpecial = wdPaneNone Then
           .ActiveWindow.ActivePane.View.Type = wdPageView
        Else
           .ActiveWindow.View.Type = wdPageView
        End If
        .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
        .Selection.HeaderFooter.PageNumbers.add PageNumberAlignment:=wdAlignPageNumberCenter, FirstPage:=True
        .ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
        
        .Activate
    End With
    Exit Sub
   
ErrHnd:
   If Err.Number <> 0 Then
      If iResumeCnt > 3 Then
         MsgBox "錯誤 : " & Err.Description, vbCritical
      Else
         iResumeCnt = iResumeCnt + 1
         Select Case Err.Number
            Case 91:
               g_WordAp.Documents.add
               Resume Next
            Case 462:
               Set g_WordAp = New Word.Application
               Resume
            Case Else:
               MsgBox "錯誤" & " : " & Err.Description, vbCritical '& iLine
         End Select
      End If
   End If
End Sub

Private Sub SetWordTitle(stOffice As String, strFontSize As String)
    Dim stTmp As String
    
    With g_WordAp
        '印表頭
        .Selection.ParagraphFormat.DisableLineHeightGrid = True
        stTmp = lblTitle & " ( " & MaskEdBox1 & " )"
        .Selection.Font.Size = 16
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Selection.TypeText Text:=stTmp

        .Selection.TypeParagraph
        .Selection.Font.Size = strFontSize
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify

        Select Case stOffice
            Case "1": stTmp = "北所"
            Case "2": stTmp = "中所"
            Case "3": stTmp = "南所"
            Case "4": stTmp = "高所"
        End Select
        .Selection.Font.Size = 14
        .Selection.TypeText Text:=stTmp
        .Selection.Font.Size = 10
        .Selection.TypeParagraph
        .Selection.Font.Size = 14
    End With
End Sub

'設定表格及議題
Private Sub SetTable(ByRef iTables As Integer, iCols As Integer, stSubjectTxt As String, stTimeTxt As String)
    Dim iLines As Integer, iRowHeight As Integer
    Dim stTmp As String
    
    With g_WordAp
        '插入表格
        iTables = iTables + 1
        .Selection.Tables.add Range:=.Selection.Range, NumRows:=3, NumColumns:=(iCols + 1)
        '設框線,高寬
        .Selection.Tables(1).Select
        
        With .Selection.Borders(wdBorderTop)
            .LineStyle = g_WordAp.Options.DefaultBorderLineStyle
            .LineWidth = g_WordAp.Options.DefaultBorderLineWidth
        End With
        With .Selection.Borders(wdBorderLeft)
            .LineStyle = g_WordAp.Options.DefaultBorderLineStyle
            .LineWidth = g_WordAp.Options.DefaultBorderLineWidth
        End With
        With .Selection.Borders(wdBorderBottom)
            .LineStyle = g_WordAp.Options.DefaultBorderLineStyle
            .LineWidth = g_WordAp.Options.DefaultBorderLineWidth
        End With
        With .Selection.Borders(wdBorderRight)
            .LineStyle = g_WordAp.Options.DefaultBorderLineStyle
            .LineWidth = g_WordAp.Options.DefaultBorderLineWidth
        End With
        With .Selection.Borders(wdBorderHorizontal)
            .LineStyle = g_WordAp.Options.DefaultBorderLineStyle
        End With
        With .Selection.Borders(wdBorderVertical)
            .LineStyle = g_WordAp.Options.DefaultBorderLineStyle
            .LineWidth = g_WordAp.Options.DefaultBorderLineWidth
        End With
    
        '設定表格高度
        .Selection.MoveLeft Unit:=wdCharacter, Count:=1
        .Selection.SelectRow
        .Selection.Cells.SetHeight RowHeight:=56, HeightRule:=wdRowHeightExactly
    
        .Selection.MoveLeft Unit:=wdCharacter, Count:=1
        .Selection.SelectColumn
        .Selection.Cells.SetWidth ColumnWidth:=.CentimetersToPoints(5.5), RulerStyle:=wdAdjustProportional
        .Selection.SelectRow
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
        .Selection.MoveLeft Unit:=wdCharacter, Count:=1
        
        .Selection.MoveDown Unit:=wdLine, Count:=1
        
        .Selection.SelectRow
        .Selection.Cells.SetHeight RowHeight:=20, HeightRule:=wdRowHeightExactly
        
        '議題內容(取18字)
        stTmp = stSubjectTxt
        If Len(stTmp) > 18 Then
            stTmp = StrToStr(stSubjectTxt, 18) & "..."
        End If
        iLines = 1
        intI = InStr(1, StrToStr(stSubjectTxt, 10), vbCrLf)
        Do While intI > 0
            iLines = iLines + 1
            intI = InStr(intI + 1, stTmp, vbCrLf)
        Loop
        '時間
        If stTmp <> "同上" Then
            stTmp = stTmp & vbCrLf & stTimeTxt
            iLines = iLines + 1
        End If
        
        iRowHeight = iLines * 20
        If iRowHeight < 80 Then iRowHeight = 80
        .Selection.MoveLeft Unit:=wdCharacter, Count:=1
        .Selection.MoveDown Unit:=wdLine, Count:=1
        .Selection.SelectRow
        
        .Selection.Cells.SetHeight RowHeight:=(iRowHeight - 20), HeightRule:=wdRowHeightExactly
        .Selection.MoveLeft Unit:=wdCharacter, Count:=1
        .Selection.MoveUp Unit:=wdLine, Count:=1, Extend:=wdExtend
        .Selection.Cells.Merge
        .Selection.TypeText Text:=stTmp
    End With
End Sub
'end 2020/11/27

'Add by Amy 2020/12/28 多筆主講者顯示
Private Function SetSS06(ByVal stSS06 As String, bolStaffNo As Boolean) As String
    Dim RsQ As ADODB.Recordset
    Dim strQ As String, stTmp As String
    Dim intQ As Integer, ii As Integer
    Dim arrTmp
    
    SetSS06 = ""
    strExc(0) = "Select * From Staff Where 1=1 "
    arrTmp = Split(stSS06, ";")
    For ii = LBound(arrTmp) To UBound(arrTmp)
        strQ = strExc(0) & " And st01='" & arrTmp(ii) & "' "
        intQ = 1
        Set RsQ = ClsLawReadRstMsg(intQ, strQ)
        If intQ = 1 Then
            SetSS06 = SetSS06 & ";" & RsQ.Fields("st02")
            '顯示員編-代表所內員工需發mail,因可能為所外人員同名同姓
            If bolStaffNo = True Then
                SetSS06 = SetSS06 & "(" & arrTmp(ii) & ")"
            End If
        Else
            SetSS06 = SetSS06 & ";" & arrTmp(ii)
        End If
    Next ii
 
    If SetSS06 <> MsgText(601) Then
        SetSS06 = Mid(SetSS06, 2)
    End If
    
    Set RsQ = Nothing
End Function

'取得需寄信之演講者(所內人員)
Private Function GetSendSS06() As String
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String
    Dim ii As Integer, intQ As Integer
    
    GetSendSS06 = ""
    With MSHFlexGrid1
        For ii = 0 To .Rows - 1
            strQ = strQ & ";" & MSHFlexGrid1.TextMatrix(ii, 5)
        Next ii
    End With
    '保留員工,非員工會顯示名稱不寄
    If strQ <> MsgText(601) Then
        strQ = "'" & Replace(Mid(strQ, 2), ";", "','") & "'"
        strQ = "Select st01 From Staff Where st01 In(" & strQ & ") Order by st01"
        intQ = 1
        Set RsQ = ClsLawReadRstMsg(intQ, strQ)
        If intQ = 1 Then
            Do While Not RsQ.EOF
                GetSendSS06 = GetSendSS06 & ";" & RsQ.Fields("st01")
                RsQ.MoveNext
            Loop
        End If
        If GetSendSS06 <> MsgText(601) Then GetSendSS06 = Mid(GetSendSS06, 2)
    End If
End Function
'end 2020/12/28

'Add by Amy 2021/01/25 判斷表格超過換頁
Private Sub SetOverTB(ByVal st06 As String, ByVal stFontSize As String, Optional ByVal IsColMax As Boolean = False)
    Dim intTB As Integer
    
    intTB = iTables
    If IsColMax = True Then intTB = intTB + 1
    
    If intTB > Val(intMaxTB) Then
        With g_WordAp
            .Selection.MoveDown Unit:=wdLine, Count:=1
            .Selection.SelectRow
            .Selection.MoveRight Unit:=wdCharacter, Count:=2
            .Selection.Font.Size = 10
            .Selection.InsertBreak Type:=wdPageBreak
            '避免不過高刪除一行
            .Selection.MoveUp Unit:=wdLine, Count:=1
            .Selection.TypeBackspace
            .Selection.MoveDown Unit:=wdLine, Count:=1
            .Selection.Font.Size = 14
            '表頭
            Call SetWordTitle(st06, stFontSize)
            iTables = 0
            If IsColMax = True Then iNowCol = 0
        End With
    End If
    
End Sub
