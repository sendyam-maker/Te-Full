VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm160016 
   BorderStyle     =   1  '單線固定
   Caption         =   "員工工作評價資料"
   ClientHeight    =   5870
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   8190
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5870
   ScaleWidth      =   8190
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6960
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H00C00000&
      Height          =   940
      Left            =   30
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   54
      Text            =   "frm160016.frx":0000
      Top             =   4890
      Width           =   7960
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4320
      Left            =   30
      TabIndex        =   21
      Top             =   570
      Width           =   8150
      _ExtentX        =   14376
      _ExtentY        =   7620
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "單筆資料"
      TabPicture(0)   =   "frm160016.frx":0114
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "textSTJ01_2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "LblHadData"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "textSTJ01"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "textSTJ02"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "SSTab2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "多筆瀏覽"
      TabPicture(1)   =   "frm160016.frx":0130
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txt1(4)"
      Tab(1).Control(1)=   "txt1(5)"
      Tab(1).Control(2)=   "txt1(2)"
      Tab(1).Control(3)=   "txt1(3)"
      Tab(1).Control(4)=   "Command1"
      Tab(1).Control(5)=   "Combo1"
      Tab(1).Control(6)=   "Combo2"
      Tab(1).Control(7)=   "GRD1"
      Tab(1).Control(8)=   "cmdok"
      Tab(1).Control(9)=   "txt1(1)"
      Tab(1).Control(10)=   "txt1(0)"
      Tab(1).Control(11)=   "Line2"
      Tab(1).Control(12)=   "Label6"
      Tab(1).Control(13)=   "Line1"
      Tab(1).Control(14)=   "Label5"
      Tab(1).Control(15)=   "Label4"
      Tab(1).Control(16)=   "Label16"
      Tab(1).Control(17)=   "Label15"
      Tab(1).Control(18)=   "Line4"
      Tab(1).ControlCount=   19
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   4
         Left            =   -73920
         MaxLength       =   5
         TabIndex        =   43
         Top             =   330
         Width           =   680
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   5
         Left            =   -73110
         MaxLength       =   5
         TabIndex        =   44
         Top             =   330
         Width           =   680
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   2
         Left            =   -71010
         MaxLength       =   3
         TabIndex        =   47
         Top             =   630
         Width           =   530
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   3
         Left            =   -70350
         MaxLength       =   3
         TabIndex        =   48
         Top             =   630
         Width           =   530
      End
      Begin VB.CommandButton Command1 
         Caption         =   "匯出Excel檔"
         Height          =   255
         Left            =   -68610
         TabIndex        =   52
         Top             =   690
         Width           =   1400
      End
      Begin VB.ComboBox Combo1 
         Height          =   260
         ItemData        =   "frm160016.frx":014C
         Left            =   -73920
         List            =   "frm160016.frx":014E
         Style           =   2  '單純下拉式
         TabIndex        =   49
         Top             =   930
         Width           =   2210
      End
      Begin VB.ComboBox Combo2 
         Height          =   260
         ItemData        =   "frm160016.frx":0150
         Left            =   -71010
         List            =   "frm160016.frx":0152
         Style           =   2  '單純下拉式
         TabIndex        =   50
         Top             =   930
         Width           =   2210
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   3280
         Left            =   30
         TabIndex        =   27
         Top             =   990
         Width           =   8080
         _ExtentX        =   14252
         _ExtentY        =   5786
         _Version        =   393216
         TabHeight       =   476
         TabCaption(0)   =   "第一階主管"
         TabPicture(0)   =   "frm160016.frx":0154
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label23(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "textSTJ04_2(0)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label1(1)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label1(7)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Label1(6)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "textSTJ04(0)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "textSTJ05(0)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "cmdOpenAtt(0)"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "cmdAddAtt(0)"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "cmdRemAtt(0)"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "lstAtt(0)"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).ControlCount=   11
         TabCaption(1)   =   "第二階主管"
         TabPicture(1)   =   "frm160016.frx":0170
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "lstAtt(1)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "textSTJ04(1)"
         Tab(1).Control(2)=   "textSTJ05(1)"
         Tab(1).Control(3)=   "cmdOpenAtt(1)"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "cmdAddAtt(1)"
         Tab(1).Control(5)=   "cmdRemAtt(1)"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "Label1(8)"
         Tab(1).Control(7)=   "Label23(1)"
         Tab(1).Control(8)=   "textSTJ04_2(1)"
         Tab(1).Control(9)=   "Label1(3)"
         Tab(1).Control(10)=   "Label1(2)"
         Tab(1).ControlCount=   11
         TabCaption(2)   =   "第三階主管"
         TabPicture(2)   =   "frm160016.frx":018C
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "lstAtt(2)"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "textSTJ04(2)"
         Tab(2).Control(2)=   "textSTJ05(2)"
         Tab(2).Control(3)=   "cmdOpenAtt(2)"
         Tab(2).Control(3).Enabled=   0   'False
         Tab(2).Control(4)=   "cmdAddAtt(2)"
         Tab(2).Control(5)=   "cmdRemAtt(2)"
         Tab(2).Control(5).Enabled=   0   'False
         Tab(2).Control(6)=   "Label1(9)"
         Tab(2).Control(7)=   "Label23(2)"
         Tab(2).Control(8)=   "textSTJ04_2(2)"
         Tab(2).Control(9)=   "Label1(5)"
         Tab(2).Control(10)=   "Label1(4)"
         Tab(2).ControlCount=   11
         Begin VB.ListBox lstAtt 
            Height          =   760
            Index           =   2
            ItemData        =   "frm160016.frx":01A8
            Left            =   -74340
            List            =   "frm160016.frx":01AF
            MultiSelect     =   2  '進階多重選取
            Sorted          =   -1  'True
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   2250
            Width           =   6570
         End
         Begin VB.ListBox lstAtt 
            Height          =   760
            Index           =   1
            ItemData        =   "frm160016.frx":01BB
            Left            =   -74340
            List            =   "frm160016.frx":01C2
            MultiSelect     =   2  '進階多重選取
            Sorted          =   -1  'True
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   2250
            Width           =   6570
         End
         Begin VB.ListBox lstAtt 
            Height          =   760
            Index           =   0
            ItemData        =   "frm160016.frx":01CE
            Left            =   660
            List            =   "frm160016.frx":01D5
            MultiSelect     =   2  '進階多重選取
            Sorted          =   -1  'True
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   2250
            Width           =   6570
         End
         Begin VB.TextBox textSTJ04 
            Height          =   270
            Index           =   2
            Left            =   -74070
            MaxLength       =   6
            TabIndex        =   14
            Top             =   300
            Width           =   735
         End
         Begin VB.TextBox textSTJ05 
            Height          =   1620
            Index           =   2
            Left            =   -74340
            MaxLength       =   500
            MultiLine       =   -1  'True
            ScrollBars      =   2  '垂直捲軸
            TabIndex        =   15
            Top             =   600
            Width           =   7340
         End
         Begin VB.CommandButton cmdOpenAtt 
            Caption         =   "開啟"
            Height          =   255
            Index           =   2
            Left            =   -67740
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   2220
            Width           =   735
         End
         Begin VB.CommandButton cmdAddAtt 
            Caption         =   "<- 新增"
            Height          =   285
            Index           =   2
            Left            =   -67740
            TabIndex        =   18
            Top             =   2490
            Width           =   735
         End
         Begin VB.CommandButton cmdRemAtt 
            Caption         =   "-> 移除"
            Height          =   255
            Index           =   2
            Left            =   -67740
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   2790
            Width           =   735
         End
         Begin VB.TextBox textSTJ04 
            Height          =   270
            Index           =   1
            Left            =   -74070
            MaxLength       =   6
            TabIndex        =   8
            Top             =   300
            Width           =   735
         End
         Begin VB.TextBox textSTJ05 
            Height          =   1620
            Index           =   1
            Left            =   -74340
            MaxLength       =   500
            MultiLine       =   -1  'True
            ScrollBars      =   2  '垂直捲軸
            TabIndex        =   9
            Top             =   600
            Width           =   7340
         End
         Begin VB.CommandButton cmdOpenAtt 
            Caption         =   "開啟"
            Height          =   255
            Index           =   1
            Left            =   -67740
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   2220
            Width           =   735
         End
         Begin VB.CommandButton cmdAddAtt 
            Caption         =   "<- 新增"
            Height          =   285
            Index           =   1
            Left            =   -67740
            TabIndex        =   12
            Top             =   2490
            Width           =   735
         End
         Begin VB.CommandButton cmdRemAtt 
            Caption         =   "-> 移除"
            Height          =   255
            Index           =   1
            Left            =   -67740
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   2790
            Width           =   735
         End
         Begin VB.CommandButton cmdRemAtt 
            Caption         =   "-> 移除"
            Height          =   255
            Index           =   0
            Left            =   7260
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   2790
            Width           =   735
         End
         Begin VB.CommandButton cmdAddAtt 
            Caption         =   "<- 新增"
            Height          =   285
            Index           =   0
            Left            =   7260
            TabIndex        =   6
            Top             =   2490
            Width           =   735
         End
         Begin VB.CommandButton cmdOpenAtt 
            Caption         =   "開啟"
            Height          =   255
            Index           =   0
            Left            =   7260
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   2220
            Width           =   735
         End
         Begin VB.TextBox textSTJ05 
            Height          =   1620
            Index           =   0
            Left            =   660
            MaxLength       =   500
            MultiLine       =   -1  'True
            ScrollBars      =   2  '垂直捲軸
            TabIndex        =   3
            Top             =   600
            Width           =   7340
         End
         Begin VB.TextBox textSTJ04 
            Height          =   270
            Index           =   0
            Left            =   930
            MaxLength       =   6
            TabIndex        =   2
            Top             =   300
            Width           =   735
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "內容："
            Height          =   180
            Index           =   9
            Left            =   -74940
            TabIndex        =   57
            Top             =   600
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "內容："
            Height          =   180
            Index           =   8
            Left            =   -74940
            TabIndex        =   56
            Top             =   600
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "內容："
            Height          =   180
            Index           =   6
            Left            =   60
            TabIndex        =   55
            Top             =   600
            Width           =   540
         End
         Begin MSForms.Label Label23 
            Height          =   200
            Index           =   2
            Left            =   -74910
            TabIndex        =   42
            Top             =   3060
            Width           =   7400
            VariousPropertyBits=   27
            Caption         =   "CREATE :                                                    UPDATE : "
            Size            =   "13053;353"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.Label textSTJ04_2 
            Height          =   230
            Index           =   2
            Left            =   -73290
            TabIndex        =   41
            Top             =   330
            Width           =   1400
            BackColor       =   12632256
            VariousPropertyBits=   27
            Size            =   "2461;397"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "主管員編："
            Height          =   180
            Index           =   5
            Left            =   -74940
            TabIndex        =   40
            Top             =   350
            Width           =   900
         End
         Begin VB.Label Label1 
            Caption         =   "附件："
            Height          =   180
            Index           =   4
            Left            =   -74940
            TabIndex        =   39
            Top             =   2280
            Width           =   600
         End
         Begin MSForms.Label Label23 
            Height          =   200
            Index           =   1
            Left            =   -74910
            TabIndex        =   38
            Top             =   3060
            Width           =   7400
            VariousPropertyBits=   27
            Caption         =   "CREATE :                                                    UPDATE : "
            Size            =   "13053;353"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.Label textSTJ04_2 
            Height          =   230
            Index           =   1
            Left            =   -73290
            TabIndex        =   37
            Top             =   330
            Width           =   1400
            BackColor       =   12632256
            VariousPropertyBits=   27
            Size            =   "2461;397"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "主管員編："
            Height          =   180
            Index           =   3
            Left            =   -74940
            TabIndex        =   36
            Top             =   350
            Width           =   900
         End
         Begin VB.Label Label1 
            Caption         =   "附件："
            Height          =   180
            Index           =   2
            Left            =   -74940
            TabIndex        =   35
            Top             =   2280
            Width           =   600
         End
         Begin VB.Label Label1 
            Caption         =   "附件："
            Height          =   180
            Index           =   7
            Left            =   60
            TabIndex        =   31
            Top             =   2280
            Width           =   600
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "主管員編："
            Height          =   180
            Index           =   1
            Left            =   60
            TabIndex        =   30
            Top             =   350
            Width           =   900
         End
         Begin MSForms.Label textSTJ04_2 
            Height          =   230
            Index           =   0
            Left            =   1710
            TabIndex        =   29
            Top             =   330
            Width           =   1400
            BackColor       =   12632256
            VariousPropertyBits=   27
            Size            =   "2461;397"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.Label Label23 
            Height          =   200
            Index           =   0
            Left            =   90
            TabIndex        =   28
            Top             =   3060
            Width           =   7400
            VariousPropertyBits=   27
            Caption         =   "CREATE :                                                    UPDATE : "
            Size            =   "13053;353"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Bindings        =   "frm160016.frx":01E1
         Height          =   3020
         Left            =   -74970
         TabIndex        =   53
         Top             =   1230
         Width           =   8070
         _ExtentX        =   14235
         _ExtentY        =   5327
         _Version        =   393216
         FixedCols       =   0
         ScrollTrack     =   -1  'True
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
      Begin VB.CommandButton cmdok 
         Caption         =   "查詢"
         Height          =   255
         Left            =   -68610
         TabIndex        =   51
         Top             =   360
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   1
         Left            =   -72870
         MaxLength       =   6
         TabIndex        =   46
         Top             =   630
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   0
         Left            =   -73920
         MaxLength       =   6
         TabIndex        =   45
         Top             =   630
         Width           =   915
      End
      Begin VB.TextBox textSTJ02 
         Height          =   285
         Left            =   1080
         MaxLength       =   5
         TabIndex        =   0
         Top             =   360
         Width           =   740
      End
      Begin VB.TextBox textSTJ01 
         Height          =   270
         Left            =   1080
         MaxLength       =   6
         TabIndex        =   1
         Top             =   660
         Width           =   735
      End
      Begin VB.Label LblHadData 
         Caption         =   "有評價資料"
         ForeColor       =   &H000000C0&
         Height          =   190
         Left            =   2190
         TabIndex        =   58
         Top             =   390
         Width           =   1000
      End
      Begin VB.Line Line2 
         X1              =   -73620
         X2              =   -72930
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "年月："
         Height          =   180
         Left            =   -74490
         TabIndex        =   34
         Top             =   360
         Width           =   540
      End
      Begin VB.Line Line1 
         X1              =   -70710
         X2              =   -70020
         Y1              =   780
         Y2              =   780
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "組別："
         Height          =   180
         Left            =   -71580
         TabIndex        =   33
         Top             =   960
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "職位："
         Height          =   180
         Left            =   -74490
         TabIndex        =   32
         Top             =   960
         Width           =   540
      End
      Begin MSForms.Label textSTJ01_2 
         Height          =   230
         Left            =   1860
         TabIndex        =   26
         Top             =   690
         Width           =   1400
         BackColor       =   12632256
         VariousPropertyBits=   27
         Size            =   "2461;397"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "部門："
         Height          =   180
         Left            =   -71580
         TabIndex        =   25
         Top             =   660
         Width           =   540
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "員工編號："
         Height          =   180
         Left            =   -74850
         TabIndex        =   24
         Top             =   660
         Width           =   900
      End
      Begin VB.Line Line4 
         X1              =   -73230
         X2              =   -72540
         Y1              =   780
         Y2              =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "年月："
         Height          =   180
         Left            =   390
         TabIndex        =   23
         Top             =   420
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "員工編號："
         Height          =   180
         Index           =   0
         Left            =   210
         TabIndex        =   22
         Top             =   710
         Width           =   900
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7500
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160016.frx":01F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160016.frx":0512
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160016.frx":082E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160016.frx":0A0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160016.frx":0D26
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160016.frx":1042
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160016.frx":135E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160016.frx":167A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160016.frx":1996
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160016.frx":1CB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160016.frx":1FCE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   520
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   8190
      _ExtentX        =   14446
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
End
Attribute VB_Name = "frm160016"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create by Sindy 2023/10/16
Option Explicit

Dim RcMain As New ADODB.Recordset, RsAdo As New ADODB.Recordset
' 變數宣告區
Dim m_EditMode As Integer
Dim m_SubMode As Integer
'(執行各項功能的權限)
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
' 第一筆資料的本所案號
Dim m_FirstKEY(2) As String
' 最後一筆資料的本所案號
Dim m_LastKEY(2) As String
' 目前正在顯示的本所案號
Dim m_CurrKEY(2) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim m_STJLimitEmp As String
Dim strSysDtBef1M As String
Dim m_AttachPath As String
Dim m_FilesRemoved() As String
Dim m_FileName As String, m_TempFileName As String


'新增
Private Sub cmdAddAtt_Click(Index As Integer)
   Dim stFileName As String
   Dim sFile
   Dim ii As Integer
   Dim fs, f
   Dim strFile As String
   Dim strFilePath As String
   
On Error GoTo ErrHnd
   
   '取得開啟檔案的路徑
   If lstAtt(Index).ListCount > 0 Then
      ii = 0
      Do While ii < lstAtt(Index).ListCount
         If lstAtt(Index).Selected(ii) = True Then
            If InStr(lstAtt(Index).List(ii), "\") > 0 Then
               strFilePath = Mid(lstAtt(Index).List(ii), 1, InStrRev(lstAtt(Index).List(ii), "\") - 1)
               Exit Do
            End If
         End If
         ii = ii + 1
      Loop
   End If
'   If strFilePath = "" Then
'      If GetSetting("TAIE", "FCP", UCase(Me.Name) & "Dir", "") <> "" Then
'         strFilePath = GetSetting("TAIE", "FCP", UCase(Me.Name) & "Dir", "")
'      Else
'         strFilePath = PUB_Getdesktop
'      End If
'      If PUB_ChkDir(strFilePath) = False Then
'         strFilePath = PUB_Getdesktop
'      End If
'   End If
   
   stFileName = "*.*"
   With CommonDialog1
      .CancelError = True
      .FileName = stFileName
      .Filter = "All Files (*.*)|*.*"
      .InitDir = strFilePath
      .MaxFileSize = 3000
      .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
      .ShowOpen
      If .FileName <> "" Then
         '選取多個檔案時
         If InStr(.FileName, ChrW$(0)) > 0 Then
            sFile = Split(.FileName, ChrW$(0))
            For ii = 1 To UBound(sFile)
               If InStr(CStr(sFile(ii)), "#") > 0 Then
                  MsgBox CStr(sFile(ii)) & vbCrLf & vbCrLf & "【#】符號為系統保留字，不可使用於檔案命名"
                  Exit Sub
               End If
               
               If InStr(sFile(ii), "\") > 0 Then
                  stFileName = sFile(ii)
               Else
                  stFileName = sFile(0) & "\" & sFile(ii)
               End If
               Set fs = CreateObject("Scripting.FileSystemObject")
               Set f = fs.GetFile(stFileName)
               '檔案大小為 0 KB 有誤
               If f.Size = 0 Then
                  ShowMsg sFile(ii) & MsgText(9221)
                  Exit Sub
               ElseIf f.Size > 5242880 Then
                  If MsgBox("檔案過大（容量超過5MB），確認是否要傳送？", vbYesNo, "警告") = vbNo Then
                     Exit Sub
                  End If
               End If
               AddListX lstAtt(Index), stFileName & " (" & Round(f.Size / 1024, 2) & " KB)"
            Next
            
         '選取單檔時
         Else
            strFile = Mid(.FileName, InStrRev(.FileName, "\") + 1)
            If InStr(strFile, "#") > 0 Then
               MsgBox strFile & vbCrLf & vbCrLf & "【#】符號為系統保留字，不可使用於檔案命名"
               Exit Sub
            End If
            stFileName = .FileName
            Set fs = CreateObject("Scripting.FileSystemObject")
            Set f = fs.GetFile(stFileName)
            '檔案大小為 0 KB 有誤
            If f.Size = 0 Then
               ShowMsg strFile & MsgText(9221)
               Exit Sub
            ElseIf f.Size > 5242880 Then
               If MsgBox("檔案過大（容量超過5MB），確認是否要傳送？", vbYesNo, "警告") = vbNo Then
                  Exit Sub
               End If
            End If
            AddListX lstAtt(Index), stFileName & " (" & Round(f.Size / 1024, 2) & " KB)"
         End If
         
         '移除已不存在的電子檔
         If Index = 0 Then
            If lstAtt(Index).ListCount > 0 Then
               For ii = lstAtt(Index).ListCount - 1 To 0 Step -1
                  strFilePath = lstAtt(Index).List(ii)
                  If InStrRev(strFilePath, " (") > 0 Then
                     If UCase(Mid(strFilePath, InStrRev(strFilePath, " (") + 1, Len("(X86)"))) <> "(X86)" Then
                        strFilePath = Left(strFilePath, InStrRev(strFilePath, " (") - 1)
                     End If
                  End If
                  If Dir(strFilePath) = "" Then
                     lstAtt(Index).RemoveItem ii
                  End If
               Next ii
            End If
         End If
      End If
      ChDir App.path '釋放資料夾權限
   End With
   Exit Sub
ErrHnd:
   If Err.Number <> 32755 Then '32755=已選取「取消」。
      MsgBox Err.Description
   End If
End Sub

Private Function AddListX(oList As Object, stNewItem As String) As Boolean
   Dim idx As Integer, stFileName As String
      
   If stNewItem <> "" Then
      For idx = 0 To oList.ListCount - 1
         stFileName = GetFileName(oList.List(idx))
         If UCase(GetFileName(stNewItem)) = UCase(stFileName) Then
            MsgBox "附件 " & stFileName & " 已存在！"
            AddListX = False
            Exit Function
         End If
      Next
      
      oList.AddItem stNewItem, 0
      SetListScroll oList
      AddListX = True
   End If
End Function

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

Private Sub cmdok_Click()
   If txt1(4) = "" Or txt1(5) = "" Then
      MsgBox "起迄年月不可以空白！", vbExclamation, "操作錯誤！"
      Exit Sub
   End If
   GetData
End Sub

'開啟附件
Private Sub cmdOpenAtt_Click(Index As Integer)
   Dim hLocalFile As Long
   Dim stFileName As String
   Dim strAtt As String
   Dim bolIsSelect As Boolean
   Dim ii As Integer
   
   bolIsSelect = False
   Screen.MousePointer = vbHourglass
   
   strAtt = lstAtt(Index).Text
   
   If strAtt = "" Then
      MsgBox "請選擇欲開啟的附件！"
      lstAtt(Index).SetFocus
   Else
      For ii = 0 To lstAtt(Index).ListCount - 1
         If lstAtt(Index).Selected(ii) Then
            bolIsSelect = True
            stFileName = lstAtt(Index).List(ii)
            If InStrRev(stFileName, " (") > 0 Then
               If UCase(Mid(stFileName, InStrRev(stFileName, " (") + 1, Len("(X86)"))) <> "(X86)" Then
                  stFileName = Left(stFileName, InStrRev(stFileName, " (") - 1)
               End If
            End If
            
            If InStr(stFileName, "\") = 0 Then
               If GetAttachFile_STJ(textSTJ01, Val(textSTJ02) + 191100, Index + 1, stFileName, m_AttachPath) = False Then
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If
            End If
            If m_EditMode <> 1 Then '非新增,才能設定為唯讀
               SetAttr stFileName, vbReadOnly '檔案設定為唯讀屬性
            End If
            ShellExecute hLocalFile, "open", stFileName, vbNullString, vbNullString, 1
         End If
      Next ii
      If bolIsSelect = False Then
         MsgBox "請選擇欲開啟的附件！"
         lstAtt(Index).SetFocus
      End If
   End If
   
   Screen.MousePointer = vbDefault
End Sub

Private Function GetAttachFile_STJ(ByVal strKey1 As String, StrKey2 As String, strKey3 As Integer, _
   ByRef pFileName As String, pSavePath As String, Optional bolIsFullPath As Boolean = False) As Boolean
   
   Dim stAttPath As String
   Dim lngSize As Long
   Dim iFileNo As Integer
   Dim bytes() As Byte
   Dim rsQuery As ADODB.Recordset
   Dim bolHadShowMsg As Boolean
   Dim strSql As String
   
On Error GoTo ErrHnd
   
   If bolIsFullPath = True Then
      stAttPath = pSavePath
      
   Else
      If Dir(pSavePath, vbDirectory) = "" Then
         MkDir pSavePath
      End If
      stAttPath = pSavePath & "\" & pFileName
         
      '檔案已存在時
      If Dir(stAttPath) <> "" Then
         '檢查檔案是否正在使用中
         If PUB_ChkFileOpening(stAttPath, bolHadShowMsg) = True Then
            If bolHadShowMsg = False Then
               MsgBox stAttPath & vbCrLf & "檔案正在使用中（請關閉），方可繼續操作。", vbExclamation
            End If
            Screen.MousePointer = vbDefault
            Exit Function
         End If
                     
         SetAttr stAttPath, vbNormal '檔案設定為正常屬性
            
         Kill stAttPath
      End If
   End If
   
   strSql = "select SJF05 from Staff_Evalu_File" & _
            " where SJF01='" & strKey1 & "' and SJF02='" & StrKey2 & "' and SJF03=" & strKey3 & _
            " and SJF04='" & ChgSQL(pFileName) & "' and SJF05 is not null"
   intI = 1
   Set rsQuery = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      If Not IsNull(rsQuery.Fields(0)) Then
         pFileName = stAttPath
         GetAttachFile_STJ = PUB_GetFtpFile(rsQuery.Fields(0), stAttPath, "STAFF_EVALU_FILE", True)
      End If
   End If
   
   Set rsQuery = Nothing
   Exit Function

ErrHnd:
   If Err.Number = 70 Then
      MsgBox ChgSQL(pFileName) & " 檔案已開啟！", vbCritical
   Else
      strExc(10) = Err.Number & ":" & Err.Description
      MsgBox strExc(10), vbCritical
   End If
   Set rsQuery = Nothing
End Function

'刪除
Private Sub cmdRemAtt_Click(Index As Integer)
Dim bolSel As Boolean
Dim ii As Integer
   
   bolSel = False
   If lstAtt(Index).ListCount > 0 Then
      ii = 0
      Do While ii < lstAtt(Index).ListCount
         If lstAtt(Index).Selected(ii) = True Then
            bolSel = True
         End If
         ii = ii + 1
      Loop
   End If
   If bolSel = True Then
      If MsgBox("確定要刪除附件嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
         Exit Sub
      End If
      Call RemoveList(lstAtt(Index), Index)
   End If
End Sub

Private Function RemoveList(oList As ListBox, Index As Integer) As Boolean
   Dim ii As Integer
   If oList.ListCount > 0 Then
      ii = 0
      Do While ii < oList.ListCount
         If oList.Selected(ii) = True Then
            
            If oList.ItemData(ii) > 0 Then  '已存檔的附件要移除
               intI = UBound(m_FilesRemoved) + 1
               ReDim Preserve m_FilesRemoved(intI) As String
               m_FilesRemoved(intI) = GetFileName(oList.List(ii))
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

'產出Excel檔
Private Sub Command1_Click()
Dim xlsReport As New Excel.Application
Dim wksReport As New Worksheet
Dim intRow As Integer, ii As Integer
Dim strText As String
   
On Error GoTo ErrHnd

   If GRD1.Rows - 1 = 0 Then
      MsgBox "無資料! ", vbInformation
      Exit Sub
   ElseIf GRD1.TextMatrix(1, 0) = "" Then
      MsgBox "無資料! ", vbInformation
      Exit Sub
   End If
   
   Screen.MousePointer = vbHourglass
   
'   m_TempFileName = strSrvDate(2) & "-" & ServerTime & "_員工工作評價範本檔.xls"
'   If Dir(PUB_Getdesktop & "\" & m_TempFileName) <> "" Then
'      Kill PUB_Getdesktop & "\" & m_TempFileName
'   End If
   
   xlsReport.Visible = True
   xlsReport.Workbooks.Open App.path & "\" & m_FileName
   Set wksReport = xlsReport.Worksheets(1)
   '條件
   strText = "": intRow = 0
   If txt1(4) <> "" Then
      If strText <> "" Then strText = strText & " + "
      strText = strText & "年月=" & Left(txt1(4), Len(txt1(4)) - 2) & "/" & Right(txt1(4), 2) & _
                                    IIf(txt1(4) <> txt1(5), " ~ " & Left(txt1(5), Len(txt1(5)) - 2) & "/" & Right(txt1(5), 2), "")
   End If
   If txt1(0) <> "" Then
      If strText <> "" Then strText = strText & " + "
      strText = strText & "員編=" & txt1(0) & " ~ " & txt1(1)
   End If
   If txt1(2) <> "" Then
      If strText <> "" Then strText = strText & " + "
      strText = strText & "部門=" & txt1(2) & GetPrjSalesBlack(txt1(2)) & IIf(txt1(2) <> txt1(3), " ~ " & txt1(3) & GetPrjSalesBlack(txt1(3)), "")
   End If
   If Trim(Combo2.Text) <> "" Then
      If strText <> "" Then strText = strText & " + "
      strText = strText & "組別=" & Trim(Combo2.Text)
   End If
   If Trim(Combo1.Text) <> "" Then
      If strText <> "" Then strText = strText & " + "
      strText = strText & "職位=" & Trim(Combo1.Text)
   End If
   wksReport.Range("A1") = "查詢條件：" & strText
   'wksReport.Range("1:1").RowHeight = 75 '(250 * intRow) '調整列高
   
   intRow = 4 '開始新增的列數
   For ii = 1 To GRD1.Rows - 1
      wksReport.Range("A" & intRow) = Left(GRD1.TextMatrix(ii, 0), Len(GRD1.TextMatrix(ii, 0)) - 2) & _
                                      "/" & Right(GRD1.TextMatrix(ii, 0), 2) '月份
      wksReport.Range("B" & intRow) = GRD1.TextMatrix(ii, 2) '員工編號
      wksReport.Range("C" & intRow) = GRD1.TextMatrix(ii, 3) '姓名
      '第一階評價
      wksReport.Range("D" & intRow) = GRD1.TextMatrix(ii, 4) '員工編號
      wksReport.Range("E" & intRow) = GRD1.TextMatrix(ii, 5) '主管
      wksReport.Range("F" & intRow) = GRD1.TextMatrix(ii, 6) '內容
      wksReport.Range("G" & intRow) = GRD1.TextMatrix(ii, 7) '附件
      '第二階評價
      wksReport.Range("H" & intRow) = GRD1.TextMatrix(ii, 8) '員工編號
      wksReport.Range("I" & intRow) = GRD1.TextMatrix(ii, 9) '主管
      wksReport.Range("J" & intRow) = GRD1.TextMatrix(ii, 10) '內容
      wksReport.Range("K" & intRow) = GRD1.TextMatrix(ii, 11) '附件
      '第三階評價
      wksReport.Range("L" & intRow) = GRD1.TextMatrix(ii, 12) '員工編號
      wksReport.Range("M" & intRow) = GRD1.TextMatrix(ii, 13) '主管
      wksReport.Range("N" & intRow) = GRD1.TextMatrix(ii, 14) '內容
      wksReport.Range("O" & intRow) = GRD1.TextMatrix(ii, 15) '附件
      intRow = intRow + 1
   Next ii
   'wksReport.Range(sRow & ":" & sRow).RowHeight = 36
   '框線
   wksReport.Range("A4" & ":" & "O" & intRow - 1).Select
   xlsReport.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
   xlsReport.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
   xlsReport.Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
   xlsReport.Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
   xlsReport.Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
   xlsReport.Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
   wksReport.Range("A4").Select
   'wksReport.Sheets(1).Select '選擇工作表
   
   'xlsReport.Workbooks(1).SaveAs PUB_Getdesktop & "\" & m_TempFileName
   'xlsReport.Workbooks.Close
   'xlsReport.Quit
   'MsgBox "Excel電子檔已存於 " & PUB_Getdesktop & "\" & m_TempFileName, vbExclamation
   
   Screen.MousePointer = vbDefault
   Set wksReport = Nothing
   Set xlsReport = Nothing
   Exit Sub
   
ErrHnd:
   Screen.MousePointer = vbDefault
'   If Err.Number = 462 Then '遠端伺服器不存在或無法使用
'      GoTo RestarWord
'   Else
   If Err.Number <> 0 Then
      MsgBox (Err.Description)
   End If
   
ExitPoint:
'   Set xlsReport = Nothing
End Sub

' 按下按鍵
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   '當focus在下列欄位時,按enter鍵維持換行功能而不是存檔功能
   If KeyCode = vbKeyReturn And _
      UCase(Me.ActiveControl.Name) = UCase("textSTJ05") Then
      Exit Sub
   End If
   
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
      Case vbKeyEscape:
         If m_EditMode = 0 Then
            OnAction KeyCode
         Else
            OnAction vbKeyF10
         End If
   End Select
End Sub

'Enter 事件，等於存檔，做完取消，不然 form 內其他物件有寫 keycode 或是 keyascii 事件的話，也會做到
Private Sub Form_KeyPress(KeyAscii As Integer)
'    Select Case KeyAscii
'      Case vbKeyReturn:
'         If m_EditMode <> 0 Then
'            KeyAscii = 0
'            OnAction vbKeyF9
'         End If
'    End Select
End Sub

Private Sub Form_Load()
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

   m_bInsert = True 'IsUserHasRightOfFunction(Me.Name, strAdd, False)
   m_bUpdate = True 'IsUserHasRightOfFunction(Me.Name, strEdit, False)
   m_bDelete = True 'IsUserHasRightOfFunction(Me.Name, strDel, False)
   m_bQuery = True 'IsUserHasRightOfFunction(Me.Name, strFind, False)
   
   textSTJ01_2.BackColor = &H8000000F
   textSTJ04_2(0).BackColor = &H8000000F
   textSTJ04_2(1).BackColor = &H8000000F
   textSTJ04_2(2).BackColor = &H8000000F
   
   MoveFormToCenter Me
   setCombo1
   Me.Combo2.Clear
   m_STJLimitEmp = PUB_GetSTJLimitEmp '取得有權限查看工作評價的同仁清單
   
   InitialData
   RefreshRange
   ShowFirstRecord
   UpdateToolbarState
   SetCtrlReadOnly True
   
   m_AttachPath = App.path & Pub_GetSpecMan("EmpFlowAttPath")
   If Dir(m_AttachPath, vbDirectory) = "" Then
      MkDir m_AttachPath
   End If
   m_AttachPath = App.path & Pub_GetSpecMan("EmpFlowAttPath") & "\" & strUserNum
   If Dir(m_AttachPath, vbDirectory) = "" Then
      MkDir m_AttachPath
   End If
   ReDim m_FilesRemoved(0)
   
   strSysDtBef1M = DBDATE(DateAdd("m", -1, ChangeWStringToWDateString(Left(strSrvDate(1), 6) & "01")))
   Me.SSTab1.Tab = 0
   m_FileName = "$$員工工作評價範本檔.xls"
   If Dir(App.path & "\" & m_FileName) = "" Then
      Call PUB_GetSampleFile(m_FileName, "M21-000004-0-00", , App.path & "\")
   End If
End Sub

'職位
Private Sub setCombo1()
Me.Combo1.Clear
strSql = "select * from allcode where ac01='02' order by ac02"
intI = 1
Set RsTemp = ClsLawReadRstMsg(intI, strSql)
Me.Combo1.AddItem ""
If intI = 1 Then
   RsTemp.MoveFirst
   While Not RsTemp.EOF
      Me.Combo1.AddItem Trim(RsTemp.Fields("ac02").Value) & " " & Trim(RsTemp.Fields("ac03").Value)
      RsTemp.MoveNext
   Wend
End If
Me.Combo1.ListIndex = 0
End Sub

'組別
Private Sub SetCombo2()
Dim strST16Nm As String

Me.Combo2.Clear
'部門相同才會檢查有無組別
If Not (txt1(2) = txt1(3) And txt1(2) <> "" And txt1(3) <> "") Then
   Exit Sub
End If
If txt1(2) = "F11" Then '外商承辦
   strST16Nm = "decode(st16,'2','英文組','4','日文組','6','CF組',st16)"
ElseIf txt1(2) = "F21" Then '外專工程師
   strST16Nm = "decode(st16,'1','電子電機組','2','化學組','3','日文組','4','機械設計組',st16)"
ElseIf txt1(2) = "F23" Then '外專承辦
   strST16Nm = "decode(st16,'1','英文組','2','日文組',st16)"
Else
   Exit Sub
End If
strSql = "select st93,st16,count(*)," & strST16Nm & " as st16Nm from staff where st04='1' and st16 is not null" & _
         " and st93='" & txt1(2) & "'" & _
         " group by st93,st16 order by st16"
intI = 1
Set RsTemp = ClsLawReadRstMsg(intI, strSql)
Me.Combo2.AddItem ""
If intI = 1 Then
   RsTemp.MoveFirst
   While Not RsTemp.EOF
      Me.Combo2.AddItem Trim(RsTemp.Fields("st16").Value) & " " & Trim(RsTemp.Fields("st16Nm").Value)
      RsTemp.MoveNext
   Wend
End If
Me.Combo2.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm160016 = Nothing
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
      If GRD1.CellBackColor <> &HFFC0C0 Then
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
         textSTJ01.Text = GRD1.TextMatrix(tmpMouseRow, 2)
         textSTJ02.Text = GRD1.TextMatrix(tmpMouseRow, 0)
         QueryRecord
         GRD1.Visible = True
      End If
   End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   If PreviousTab = 0 Then
      cmdok.SetFocus
      cmdok.Default = True
      If txt1(4).Text = "" Then txt1(4).Text = Left(strSysDtBef1M, 4) - 1911 & Mid(strSysDtBef1M, 5, 2) '預設年月
      If txt1(5).Text = "" Then txt1(5).Text = txt1(4).Text
   Else
      cmdok.Default = False
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

' 更新 Create 及 Update 的人
Private Sub UpdateCUID(Index As Integer, ByRef rsSrcTmp As ADODB.Recordset)
Dim strTemp As String
Dim strCName As String
Dim strCDate As String
Dim strCTime As String
Dim strUName As String
Dim strUDate As String
Dim strUTime As String
   
   If IsNull(rsSrcTmp.Fields("STJ06")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("STJ06")) = False Then
         strCName = GetStaffName(rsSrcTmp.Fields("STJ06"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("STJ07")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("STJ07")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("STJ07"))
         strCDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("STJ08")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("STJ08")) = False Then
         strTemp = rsSrcTmp.Fields("STJ08")
         strCTime = Format(strTemp, "##:##:##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("STJ09")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("STJ09")) = False Then
         strUName = GetStaffName(rsSrcTmp.Fields("STJ09"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("STJ10")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("STJ10")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("STJ10"))
         strUDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("STJ11")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("STJ11")) = False Then
         strTemp = rsSrcTmp.Fields("STJ11")
         strUTime = Format(strTemp, "##:##:##")
      End If
   End If
   
   ' 設定CUID中的文字
   Label23(Index).Caption = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime & String(10, " ") & _
              "UPDATE : " & strUName & " " & _
              " " & strUDate & " " & _
              " " & strUTime
End Sub

Private Function txtValidate() As Boolean
Dim Cancel As Boolean
   
   txtValidate = False
   If textSTJ02.Text = "" Then
      MsgBox "年月不可以空白！", vbExclamation
      textSTJ02.SetFocus
      Exit Function
   Else
      Cancel = False
      textSTJ02_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If textSTJ01.Text = "" Then
      MsgBox "員工編號不可以空白！", vbExclamation
      textSTJ01.SetFocus
      Exit Function
   Else
      Cancel = False
      textSTJ01_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If textSTJ04(SSTab2.Tab).Text = "" Then
      MsgBox "主管員編不可以空白！", vbExclamation
      textSTJ04(SSTab2.Tab).SetFocus
      Exit Function
   Else
      Cancel = False
      textSTJ04_Validate SSTab2.Tab, Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If textSTJ05(SSTab2.Tab).Text = "" Then
      MsgBox "內容不可以空白！", vbExclamation
      textSTJ05(SSTab2.Tab).SetFocus
      Exit Function
   End If
   
   txtValidate = True
End Function

Private Function SaveAttFile(strKey1 As String, StrKey2 As String, strKey3 As Integer) As Boolean
Dim stFilePath As String
Dim iFileNo As Integer
Dim lngSize As Long '檔案大小
Dim adoRst As New ADODB.Recordset
Dim strFile As String
Dim stFtpPath As String
Dim ii As Integer
Dim Index As Integer
   
On Error GoTo ErrHand
   
   SaveAttFile = True
   Index = Me.SSTab2.Tab
   For ii = 0 To lstAtt(Index).ListCount - 1
      If lstAtt(Index).ItemData(ii) = 0 Then
         stFilePath = lstAtt(Index).List(ii)
         If InStrRev(stFilePath, " (") > 0 Then
            If UCase(Mid(stFilePath, InStrRev(stFilePath, " (") + 1, Len("(X86)"))) <> "(X86)" Then
               stFilePath = Left(stFilePath, InStrRev(stFilePath, " (") - 1)
            End If
         End If
         If iFileNo > 0 Then Close #iFileNo
         iFileNo = FreeFile
         Open stFilePath For Binary Access Read As #iFileNo
         lngSize = LOF(iFileNo)
         
         If lngSize = 0 Then
            Close #iFileNo
            SaveAttFile = False
            ShowMsg stFilePath & MsgText(9221)
            Exit Function
         End If
         
         With adoRst
            If adoRst.State = adStateClosed Then
               strExc(0) = "select * from STAFF_EVALU_FILE where rownum<1"
               .CursorLocation = adUseClient
               .Open strExc(0), cnnConnection, adOpenStatic, adLockOptimistic
            End If
            strFile = GetFileName(stFilePath)
            .AddNew
            .Fields("SJF01").Value = strKey1
            .Fields("SJF02").Value = StrKey2
            .Fields("SJF03").Value = strKey3
            .Fields("SJF04").Value = strFile
            .Fields("SJF06").Value = lngSize
            Close #iFileNo
            
            PUB_PutFtpFile stFilePath, strKey1 & "-" & StrKey2 & "-" & strKey3, strFile, stFtpPath, "STAFF_EVALU_FILE"
            If stFtpPath <> "" Then
               .Fields("SJF05") = stFtpPath
               .Fields("SJF07") = strSrvDate(1)
            End If
            .UPDATE
         End With
      End If
   Next ii
   
   Exit Function
   
ErrHand:
   Close #iFileNo
   SaveAttFile = False
   MsgBox Err.Description, vbCritical
End Function

' 儲存記錄
Private Function SaveRecord() As Boolean
Dim strSql As String
Dim strYM As String, strST01 As String
Dim bolConn As Boolean
Dim ii As Integer
   
On Error GoTo ErrHand
   
   SaveRecord = False
   cnnConnection.BeginTrans: bolConn = True
   strYM = Val(textSTJ02) + 191100
   strST01 = textSTJ01
   
   '新增
   If m_EditMode = 1 Then
      strSql = "INSERT INTO STAFF_Evaluation (STJ01,STJ02,STJ03,STJ04,STJ05) VALUES (" & _
               CNULL(strST01) & "," & CNULL(strYM, True) & "," & Val(Me.SSTab2.Tab) + 1 & _
               "," & CNULL(textSTJ04(Me.SSTab2.Tab)) & "," & CNULL(textSTJ05(Me.SSTab2.Tab)) & ")"
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql
   '修改
   Else
      '更新
      strSql = "UPDATE STAFF_Evaluation SET STJ05='" & textSTJ05(Me.SSTab2.Tab) & "'" & _
               " WHERE STJ01='" & strST01 & "' and STJ02=" & strYM & " and STJ03=" & Val(Me.SSTab2.Tab) + 1
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql
   End If
   
   '刪除附件
   For ii = 1 To UBound(m_FilesRemoved)
      PUB_DelFtpFile2 strST01 & "-" & strYM & "-" & Val(Me.SSTab2.Tab) + 1, " and SJF04='" & ChgSQL(m_FilesRemoved(ii)) & "'", "Staff_Evalu_File"
      '刪除條件要和前面刪除FTP檔的同步
      strSql = "delete Staff_Evalu_File" & _
               " where SJF01='" & strST01 & "' AND SJF02='" & strYM & "' AND SJF03=" & Val(Me.SSTab2.Tab) + 1 & _
               " and SJF04='" & ChgSQL(m_FilesRemoved(ii)) & "'"
      cnnConnection.Execute strSql
   Next ii
   Call SaveAttFile(strST01, strYM, Val(Me.SSTab2.Tab) + 1)
   
   cnnConnection.CommitTrans: bolConn = False
   Erase m_FilesRemoved
   ReDim m_FilesRemoved(0) As String
   
   ShowCurrRecord strST01, strYM, Val(Me.SSTab2.Tab) + 1
      
   SaveRecord = True
   Exit Function
   
ErrHand:
    If bolConn = True Then cnnConnection.RollbackTrans
    MsgBox (Err.Description)
End Function

' 刪除記錄
Private Function DelRecord() As Boolean
Dim ii As Integer
Dim stFileName As String

   DelRecord = False
   
On Error GoTo ErrHand
   
   cnnConnection.BeginTrans
   
   strSql = "DELETE FROM STAFF_Evaluation" & _
            " WHERE STJ01='" & textSTJ01 & "' and STJ02=" & Val(textSTJ02) + 191100 & " and STJ03=" & Val(Me.SSTab2.Tab) + 1
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql, intI
   
   '刪除附件
   For ii = 0 To lstAtt(Me.SSTab2.Tab).ListCount - 1
      stFileName = GetFileName(lstAtt(Me.SSTab2.Tab).List(ii))
      PUB_DelFtpFile2 textSTJ01 & "-" & Val(textSTJ02) + 191100 & "-" & Val(Me.SSTab2.Tab) + 1, " and SJF04='" & ChgSQL(stFileName) & "'", "Staff_Evalu_File"
      '刪除條件要和前面刪除FTP檔的同步
      strSql = "delete Staff_Evalu_File" & _
               " where SJF01='" & textSTJ01 & "' AND SJF02='" & Val(textSTJ02) + 191100 & "' AND SJF03=" & Val(Me.SSTab2.Tab) + 1 & _
               " and SJF04='" & ChgSQL(stFileName) & "'"
      cnnConnection.Execute strSql, intI
   Next ii
   
   cnnConnection.CommitTrans
   DelRecord = True
   
   RefreshRange
   ShowCurrRecord m_LastKEY(0), m_LastKEY(1), m_LastKEY(2)
   
   Exit Function
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox "刪除失敗！" & vbCrLf & Err.Description
End Function

' 查詢記錄
Private Function QueryRecord() As Boolean
   
   QueryRecord = False
   
   If IsRecordExist(textSTJ01, Val(textSTJ02) + 191100) = True Then
      m_CurrKEY(0) = textSTJ01
      m_CurrKEY(1) = Val(textSTJ02) + 191100
      QueryRecord = True
      UpdateCtrlData
   Else
      QueryRecord = False
   End If

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
            If txtValidate = False Then Exit Function
            If SaveRecord = True Then
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
            If txtValidate = False Then Exit Function
            If SaveRecord = False Then Exit Function
         Else
            GoTo EXITSUB
         End If
      Case 3: '刪除
         If DelRecord = True Then
            RefreshRange
            ClearField
            ShowCurrRecord m_CurrKEY(0), m_CurrKEY(1), m_CurrKEY(2)
         Else
            Exit Function
         End If
      Case 4: '查詢
         If textSTJ01 <> "" And textSTJ02 <> "" Then
            If QueryRecord = False Then
               strMsg = "無此資料"
               strTit = "查詢資料"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               UpdateCtrlData
            End If
         Else
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
      Case 1: If Me.Visible = True Then textSTJ01.SetFocus
      Case 2: If Me.Visible = True Then textSTJ05(Me.SSTab2.Tab).SetFocus
      Case 4: If Me.Visible = True Then textSTJ02.SetFocus
   End Select
End Sub
' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal strKEY01 As String, ByVal strKEY02 As String, _
                               Optional ByVal strKEY03 As String = "") As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   IsRecordExist = False
   strSql = "SELECT * FROM STAFF_Evaluation " & _
            "WHERE STJ01 = '" & strKEY01 & "' and STJ02=" & strKEY02
   If strKEY03 <> "" Then
      strSql = strSql & " and STJ03='" & strKEY03 & "'"
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
Private Sub ShowCurrRecord(ByVal strKEY01 As String, ByVal strKEY02 As String, ByVal strKEY03 As String)
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   If IsRecordExist(strKEY01, strKEY02, strKEY03) = True Then
      m_CurrKEY(0) = strKEY01
      m_CurrKEY(1) = strKEY02
      m_CurrKEY(2) = strKEY03
   Else
      strSql = "SELECT STJ01,STJ02,STJ03 FROM STAFF_Evaluation " & _
               "WHERE STJ01 = '" & m_CurrKEY(0) & "' and STJ02=" & m_CurrKEY(1) & _
               " ORDER BY STJ01 ASC,STJ02 ASC,STJ03 ASC"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("STJ01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("STJ01")
         If IsNull(rsTmp.Fields("STJ02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("STJ02")
         If IsNull(rsTmp.Fields("STJ03")) = False Then: m_CurrKEY(2) = rsTmp.Fields("STJ03")
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      End If
      rsTmp.Close
      
      strSql = "SELECT STJ01,STJ02,STJ03 FROM STAFF_Evaluation " & _
               "WHERE STJ01 = '" & m_CurrKEY(0) & "'" & _
               " ORDER BY STJ01 ASC,STJ02 ASC,STJ03 ASC"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("STJ01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("STJ01")
         If IsNull(rsTmp.Fields("STJ02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("STJ02")
         If IsNull(rsTmp.Fields("STJ03")) = False Then: m_CurrKEY(2) = rsTmp.Fields("STJ03")
      Else
         ShowLastRecord
         GoTo EXITSUB
      End If
      rsTmp.Close
   End If
   UpdateCtrlData
EXITSUB:
End Sub

' 顯示第一筆資料
Private Sub ShowFirstRecord()
   m_CurrKEY(0) = m_FirstKEY(0)
   m_CurrKEY(1) = m_FirstKEY(1)
   m_CurrKEY(2) = m_FirstKEY(2)
   UpdateCtrlData
End Sub

' 顯示上一筆資料
Private Sub ShowPrevRecord()
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   If m_CurrKEY(0) = m_FirstKEY(0) And m_CurrKEY(1) = m_FirstKEY(1) Then
      ShowMsg MsgText(9008)
      GoTo EXITSUB
   End If
   
   strSql = "SELECT STJ01,STJ02,STJ03 FROM STAFF_Evaluation " & _
            "WHERE STJ01 = '" & m_CurrKEY(0) & "' AND " & _
                  "STJ02 = (SELECT MAX(STJ02) FROM STAFF_Evaluation " & _
                          "WHERE STJ01 = '" & m_CurrKEY(0) & "' AND " & _
                                "STJ02 < " & m_CurrKEY(1) & ") " & _
            "ORDER BY STJ03 ASC"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("STJ01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("STJ01")
      If IsNull(rsTmp.Fields("STJ02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("STJ02")
      If IsNull(rsTmp.Fields("STJ03")) = False Then: m_CurrKEY(2) = rsTmp.Fields("STJ03")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT STJ01,STJ02,STJ03 FROM STAFF_Evaluation " & _
            "WHERE STJ01 = (SELECT MAX(STJ01) FROM STAFF_Evaluation " & _
                           "WHERE STJ01 < '" & m_CurrKEY(0) & "' " & IIf(m_STJLimitEmp = "ALL", "", "AND STJ01 in('" & Replace(m_STJLimitEmp, ",", "','") & "') ") & ") AND " & _
                  "STJ02 = (SELECT MAX(STJ02) FROM STAFF_Evaluation " & _
                           "WHERE STJ01 = (SELECT MAX(STJ01) FROM STAFF_Evaluation " & _
                                          "WHERE STJ01 < '" & m_CurrKEY(0) & "' " & IIf(m_STJLimitEmp = "ALL", "", "AND STJ01 in('" & Replace(m_STJLimitEmp, ",", "','") & "') ") & ") ) " & _
            IIf(m_STJLimitEmp = "ALL", "", "AND STJ01 in('" & Replace(m_STJLimitEmp, ",", "','") & "') ") & _
            "ORDER BY STJ03 ASC"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("STJ01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("STJ01")
      If IsNull(rsTmp.Fields("STJ02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("STJ02")
      If IsNull(rsTmp.Fields("STJ03")) = False Then: m_CurrKEY(2) = rsTmp.Fields("STJ03")
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
   
   If m_CurrKEY(0) = m_LastKEY(0) And m_CurrKEY(1) = m_LastKEY(1) Then
      ShowMsg MsgText(9009)
      GoTo EXITSUB
   End If
   
   strSql = "SELECT STJ01,STJ02,STJ03 FROM STAFF_Evaluation " & _
            "WHERE STJ01 = '" & m_CurrKEY(0) & "' AND " & _
                  "STJ02 = (SELECT MIN(STJ02) FROM STAFF_Evaluation " & _
                          "WHERE STJ01 = '" & m_CurrKEY(0) & "' AND " & _
                                "STJ02 > " & m_CurrKEY(1) & ") " & _
            "ORDER BY STJ03 ASC"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("STJ01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("STJ01")
      If IsNull(rsTmp.Fields("STJ02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("STJ02")
      If IsNull(rsTmp.Fields("STJ03")) = False Then: m_CurrKEY(2) = rsTmp.Fields("STJ03")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT STJ01,STJ02,STJ03 FROM STAFF_Evaluation " & _
            "WHERE STJ01 = (SELECT MIN(STJ01) FROM STAFF_Evaluation " & _
                           "WHERE STJ01 > '" & m_CurrKEY(0) & "' " & IIf(m_STJLimitEmp = "ALL", "", "AND STJ01 in('" & Replace(m_STJLimitEmp, ",", "','") & "') ") & ") AND " & _
                  "STJ02 = (SELECT MIN(STJ02) FROM STAFF_Evaluation " & _
                           "WHERE STJ01 = (SELECT MIN(STJ01) FROM STAFF_Evaluation " & _
                                          "WHERE STJ01 > '" & m_CurrKEY(0) & "' " & IIf(m_STJLimitEmp = "ALL", "", "AND STJ01 in('" & Replace(m_STJLimitEmp, ",", "','") & "') ") & ")) " & _
            IIf(m_STJLimitEmp = "ALL", "", "AND STJ01 in('" & Replace(m_STJLimitEmp, ",", "','") & "') ") & _
            "ORDER BY STJ03 ASC"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("STJ01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("STJ01")
      If IsNull(rsTmp.Fields("STJ02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("STJ02")
      If IsNull(rsTmp.Fields("STJ03")) = False Then: m_CurrKEY(2) = rsTmp.Fields("STJ03")
   End If
   rsTmp.Close
   
   UpdateCtrlData
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 顯示最後一筆資料
Private Sub ShowLastRecord()
   m_CurrKEY(0) = m_LastKEY(0)
   m_CurrKEY(1) = m_LastKEY(1)
   m_CurrKEY(2) = m_LastKEY(2)
   UpdateCtrlData
End Sub

' 執行指令
Private Sub OnAction(ByVal KeyCode As Integer)
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim i As Integer
   
   LblHadData.Visible = False
   m_SubMode = 0
   Select Case KeyCode
      ' 新增 => 新的員工編號+年月
      Case vbKeyF2:
         m_EditMode = 1
         ClearField
         Me.textSTJ02 = Left(strSysDtBef1M, 4) - 1911 & Mid(strSysDtBef1M, 5, 2) '預設年月
         Me.SSTab1.TabEnabled(1) = False
         SSTab1.Tab = 0
         SetCtrlReadOnly False
         UpdateToolbarState
         SetInputEntry
         Me.textSTJ01.SetFocus
      ' 修改
      Case vbKeyF3:
         strSql = "select B0130, '1' as Sort from ABS001 where B0101='" & textSTJ01 & "' and B0130='" & strUserNum & "'" & _
                  " union select B0131, '2' as Sort from ABS001 where B0101='" & textSTJ01 & "' and B0131='" & strUserNum & "'" & _
                  " union select B0132, '3' as Sort from ABS001 where B0101='" & textSTJ01 & "' and B0132='" & strUserNum & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 0 Then
            MsgBox "您非" & Me.textSTJ01_2 & "的評價主管，無修改權限！", vbExclamation
            Exit Sub
         ElseIf Me.textSTJ04(Me.SSTab2.Tab) <> "" And Me.textSTJ04(Me.SSTab2.Tab) <> strUserNum Then
            MsgBox "無修改權限！", vbExclamation
            Exit Sub
         End If
         m_EditMode = 2
         Me.SSTab1.TabEnabled(1) = False
         SSTab1.Tab = 0
         SetCtrlReadOnly False
         SetKeyReadOnly True
         UpdateToolbarState
         SetInputEntry
         If Me.textSTJ04(Me.SSTab2.Tab) = "" Then
            m_EditMode = 1 '新增
            textSTJ04(Me.SSTab2.Tab).Locked = False
            textSTJ04(Me.SSTab2.Tab).BackColor = &H80000005
         End If
         For i = 0 To 2
            Me.SSTab2.TabEnabled(i) = False
         Next
         Me.SSTab2.TabEnabled(Me.SSTab2.Tab) = True
         If Me.textSTJ04(Me.SSTab2.Tab).Text = "" Then
            Me.textSTJ04(Me.SSTab2.Tab).Text = strUserNum
            textSTJ04_Validate Me.SSTab2.Tab, False
         End If
      ' 刪除
      Case vbKeyF5:
         If Me.textSTJ04(Me.SSTab2.Tab) <> "" Then
            strTit = "詢問"
            strMsg = "是否要刪除" & Me.textSTJ04_2(Me.SSTab2.Tab) & "評價的此筆資料?"
            nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit)
            If nResponse = vbYes Then
               m_EditMode = 3
               If OnWork = True Then
                   UpdateToolbarState
               Else
                   Exit Sub
               End If
            End If
         End If
      ' 查詢
      Case vbKeyF4:
         m_EditMode = 4
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
                  If m_EditMode = 1 Then ClearField
                  m_EditMode = 0
                  Me.SSTab1.TabEnabled(1) = True
                  UpdateCtrlData
                  SetCtrlReadOnly True
                  UpdateToolbarState
               End If
               If textSTJ01.Locked = True Then
                  For i = 0 To 2
                     If InStr(Me.SSTab2.Tag, "," & i) > 0 Then
                        Me.SSTab2.TabEnabled(i) = True
                     Else
                        Me.SSTab2.TabEnabled(i) = False
                     End If
                  Next
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
   If KeyCode <> vbKeyEscape And KeyCode <> vbKeyF3 Then
'      tabCustomer.Tab = 0
   End If
End Sub

Private Sub RefreshRange()
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   If m_STJLimitEmp = "" Then
      MsgBox "無使用權限!!", vbExclamation
      Unload Me
      Exit Sub
   End If
   
   If m_STJLimitEmp = "ALL" Then
      strSql = "SELECT STJ01,STJ02,STJ03 FROM STAFF_Evaluation" & _
               " ORDER BY STJ01 ASC,STJ02 ASC,STJ03 ASC"
   Else
      strSql = "SELECT STJ01,STJ02,STJ03 FROM STAFF_Evaluation" & _
               " WHERE STJ01 in('" & Replace(m_STJLimitEmp, ",", "','") & "')" & _
               " ORDER BY STJ01 ASC,STJ02 ASC,STJ03 ASC"
   End If
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("STJ01")) = False Then: m_FirstKEY(0) = rsTmp.Fields("STJ01")
      If IsNull(rsTmp.Fields("STJ02")) = False Then: m_FirstKEY(1) = rsTmp.Fields("STJ02")
      If IsNull(rsTmp.Fields("STJ03")) = False Then: m_FirstKEY(2) = rsTmp.Fields("STJ03")
   End If
   rsTmp.Close
   
   If m_STJLimitEmp = "ALL" Then
      strSql = "SELECT STJ01,STJ02,STJ03 FROM STAFF_Evaluation" & _
               " ORDER BY STJ01 DESC,STJ02 DESC,STJ03 DESC"
   Else
      strSql = "SELECT STJ01,STJ02,STJ03 FROM STAFF_Evaluation" & _
               " WHERE STJ01 in('" & Replace(m_STJLimitEmp, ",", "','") & "')" & _
               " ORDER BY STJ01 DESC,STJ02 DESC,STJ03 DESC"
   End If
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("STJ01")) = False Then: m_LastKEY(0) = rsTmp.Fields("STJ01")
      If IsNull(rsTmp.Fields("STJ02")) = False Then: m_LastKEY(1) = rsTmp.Fields("STJ02")
      If IsNull(rsTmp.Fields("STJ03")) = False Then: m_LastKEY(2) = rsTmp.Fields("STJ03")
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

'查詢附件檔
Private Sub ReadAttachFile(strKey1 As String, StrKey2 As String, strKey3 As Integer)
Dim Index As Integer
   
   Index = Val(strKey3) - 1
   lstAtt(Index).Clear
   strExc(0) = "select * from STAFF_EVALU_FILE" & _
               " where SJF01='" & strKey1 & "' and SJF02='" & StrKey2 & "' and SJF03=" & strKey3 & _
               " and SJF05 is not null"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
         Do While Not .EOF
            lstAtt(Index).AddItem .Fields("SJF04") & " (" & Round(.Fields("SJF06") / 1024, 2) & " KB)", 0
            lstAtt(Index).ItemData(0) = 1
            .MoveNext
         Loop
      End With
      Me.cmdOpenAtt(Index).Enabled = True
   End If
   If lstAtt(Index).ListCount > 0 Then SetListScroll lstAtt(Index)
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim i As Integer, j As Integer, int_Idx As Integer
   
   strSql = "SELECT STAFF_Evaluation.*,s1.ST02 as STJ01_NM,s2.ST02 as STJ04_NM FROM STAFF_Evaluation,staff s1,staff s2" & _
            " WHERE STJ01='" & m_CurrKEY(0) & "' and STJ02='" & m_CurrKEY(1) & "'" & _
            " AND STJ01=s1.ST01 AND STJ04=s2.ST01(+)"
   If m_STJLimitEmp <> "ALL" Then
      strSql = strSql & " AND STJ01 in('" & Replace(m_STJLimitEmp, ",", "','") & "')"
   End If
   strSql = strSql & " ORDER BY STJ03 asc"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   LblHadData.Visible = False
   If rsTmp.RecordCount > 0 Then
      LblHadData.Visible = True
      ClearField
      rsTmp.MoveFirst
      Do While Not rsTmp.EOF
         If IsNull(rsTmp.Fields("STJ01")) = False Then: textSTJ01 = rsTmp.Fields("STJ01"): textSTJ01_2 = rsTmp.Fields("STJ01_NM")
         If IsNull(rsTmp.Fields("STJ02")) = False Then: textSTJ02 = Val(Left(rsTmp.Fields("STJ02"), 4)) - 1911 & Mid(rsTmp.Fields("STJ02"), 5)
         If IsNull(rsTmp.Fields("STJ03")) = False Then
            int_Idx = Val(rsTmp.Fields("STJ03")) - 1
            textSTJ04(int_Idx) = rsTmp.Fields("STJ04")
            textSTJ04_2(int_Idx) = rsTmp.Fields("STJ04_NM")
            textSTJ05(int_Idx) = rsTmp.Fields("STJ05")
            '查詢附件檔
            Call ReadAttachFile(rsTmp.Fields("STJ01"), rsTmp.Fields("STJ02"), rsTmp.Fields("STJ03"))
            '更新CUID
            UpdateCUID int_Idx, rsTmp
         End If
   
         rsTmp.MoveNext
      Loop
      '檢查可以查看的Tab資料
      If m_STJLimitEmp = "" Then
         Me.SSTab2.TabEnabled(0) = False
         Me.SSTab2.TabEnabled(1) = False
         Me.SSTab2.TabEnabled(2) = False
      Else
         strSql = "select B0130, '1' as Sort from ABS001 where B0101='" & textSTJ01 & "' and B0130='" & strUserNum & "'" & _
                  " union select B0131, '2' as Sort from ABS001 where B0101='" & textSTJ01 & "' and B0131='" & strUserNum & "'" & _
                  " union select B0132, '3' as Sort from ABS001 where B0101='" & textSTJ01 & "' and B0132='" & strUserNum & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 0 Then
            Me.SSTab2.TabEnabled(0) = False
            Me.SSTab2.TabEnabled(1) = False
            Me.SSTab2.TabEnabled(2) = False
         Else
            For i = 0 To 2
               If i <= Val(RsTemp.Fields("Sort")) - 1 Then
                  Me.SSTab2.TabEnabled(i) = True
                  Me.SSTab2.Tab = i
               Else
                  Me.SSTab2.TabEnabled(i) = False
               End If
            Next i
         End If
         If m_STJLimitEmp = "ALL" Then
            Me.SSTab2.TabEnabled(0) = True
            Me.SSTab2.TabEnabled(1) = True
            Me.SSTab2.TabEnabled(2) = True
            Me.SSTab2.Tab = 0
         End If
      End If
      For i = 0 To 2
         If Me.SSTab2.TabEnabled(i) = True Then
            Me.SSTab2.Tag = Me.SSTab2.Tag & "," & i
         End If
      Next i
   End If
   rsTmp.Close

EXITSUB:
   Set rsTmp = Nothing
End Sub

Sub GetData()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String, strSTJsql As String
Dim strQueryEmp As String

   strSql = "": strSTJsql = ""
   If txt1(0) <> "" Then
      strSTJsql = strSTJsql & " and STJ01>='" & txt1(0) & "'"
   End If
   If txt1(1) <> "" Then
      strSTJsql = strSTJsql & " and STJ01<='" & txt1(1) & "'"
   End If
   If txt1(2) <> "" Then
      strSql = strSql & " and s1.ST93>='" & txt1(2) & "'"
   End If
   If txt1(3) <> "" Then
      strSql = strSql & " and s1.ST93<='" & txt1(3) & "'"
   End If
   If txt1(4) <> "" Then
      strSTJsql = strSTJsql & " and STJ02>=" & Val(txt1(4)) + 191100
   End If
   If txt1(5) <> "" Then
      strSTJsql = strSTJsql & " and STJ02<=" & Val(txt1(5)) + 191100
   End If
   If Trim(Combo1.Text) <> "" Then
      strSql = strSql & " and s1.ST21='" & Left(Trim(Combo1.Text), 2) & "'"
   End If
   If Trim(Combo2.Text) <> "" Then
      strSql = strSql & " and s1.ST16='" & Left(Trim(Combo2.Text), 1) & "'"
   End If
   If m_STJLimitEmp = "ALL" Then
      strQueryEmp = "ALL"
   Else
      strQueryEmp = strUserNum
      strSTJsql = strSTJsql & " and STJ01 in('" & Replace(m_STJLimitEmp, ",", "','") & "')"
   End If
   
   '抓取資料
   strSql = "SELECT STJ02-191100,A0922,STJ01,s1.st02," & _
            "GetSTJ05_Text(STJ01,STJ02,'" & strQueryEmp & "',1,'員編'),GetSTJ05_Text(STJ01,STJ02,'" & strQueryEmp & "',1,'姓名'),GetSTJ05_Text(STJ01,STJ02,'" & strQueryEmp & "',1,''),GetSTJ05_Text(STJ01,STJ02,'" & strQueryEmp & "',1,'附件')," & _
            "GetSTJ05_Text(STJ01,STJ02,'" & strQueryEmp & "',2,'員編'),GetSTJ05_Text(STJ01,STJ02,'" & strQueryEmp & "',2,'姓名'),GetSTJ05_Text(STJ01,STJ02,'" & strQueryEmp & "',2,''),GetSTJ05_Text(STJ01,STJ02,'" & strQueryEmp & "',2,'附件')," & _
            "GetSTJ05_Text(STJ01,STJ02,'" & strQueryEmp & "',3,'員編'),GetSTJ05_Text(STJ01,STJ02,'" & strQueryEmp & "',3,'姓名'),GetSTJ05_Text(STJ01,STJ02,'" & strQueryEmp & "',3,''),GetSTJ05_Text(STJ01,STJ02,'" & strQueryEmp & "',3,'附件')" & _
            ",ST93" & _
            " From STAFF_Evaluation,staff s1,acc090NEW" & _
            " where STJ01=s1.st01(+) and A0921(+)=s1.st93" & strSTJsql & strSql & _
            " group by STJ02,A0922,STJ01,s1.st02,ST93" & _
            " order by STJ02,ST93,STJ01"
   If rsTmp.State = 1 Then rsTmp.Close
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount = 0 Then MsgBox "無資料！", vbInformation
   Set GRD1.Recordset = rsTmp
   SetGrd
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
   
'   nResponse = False
'   textSTJ01_Validate nResponse
'   If nResponse = True Then GoTo EXITSUB
'   nResponse = False
'   textSTJ02_Validate nResponse
'   If nResponse = True Then GoTo EXITSUB
   
   CheckDataValid = True
EXITSUB:
End Function

' 更新Key的狀態
Private Sub SetKeyReadOnly(ByVal bEnable As Boolean)
Dim i As Integer
   
   textSTJ01.Locked = bEnable
   textSTJ02.Locked = bEnable
   If bEnable Then textSTJ01.BackColor = &H8000000F Else textSTJ01.BackColor = &H80000005
   If bEnable Then textSTJ02.BackColor = &H8000000F Else textSTJ02.BackColor = &H80000005
   If m_EditMode = 2 Then
      textSTJ04(Me.SSTab2.Tab).Locked = bEnable
      If bEnable Then textSTJ04(Me.SSTab2.Tab).BackColor = &H8000000F Else textSTJ04(Me.SSTab2.Tab).BackColor = &H80000005
   End If
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
Dim i As Integer
   
   textSTJ01.Locked = bEnable
   textSTJ02.Locked = bEnable
   If bEnable Then textSTJ01.BackColor = &H8000000F Else textSTJ01.BackColor = &H80000005
   If bEnable Then textSTJ02.BackColor = &H8000000F Else textSTJ02.BackColor = &H80000005
   If m_EditMode = 2 Then
      textSTJ04(Me.SSTab2.Tab).Locked = bEnable
      textSTJ05(Me.SSTab2.Tab).Locked = bEnable
      If bEnable Then textSTJ04(Me.SSTab2.Tab).BackColor = &H8000000F Else textSTJ04(Me.SSTab2.Tab).BackColor = &H80000005
      If bEnable Then textSTJ05(Me.SSTab2.Tab).BackColor = &H8000000F Else textSTJ05(Me.SSTab2.Tab).BackColor = &H80000005
      If bEnable Then lstAtt(Me.SSTab2.Tab).BackColor = &H8000000F Else lstAtt(Me.SSTab2.Tab).BackColor = &H80000005
      Me.cmdAddAtt(Me.SSTab2.Tab).Enabled = Not bEnable
      Me.cmdRemAtt(Me.SSTab2.Tab).Enabled = Not bEnable
   Else
      For i = 0 To 2
         textSTJ04(i).Locked = bEnable
         textSTJ05(i).Locked = bEnable
         If bEnable Then textSTJ04(i).BackColor = &H8000000F Else textSTJ04(i).BackColor = &H80000005
         If bEnable Then textSTJ05(i).BackColor = &H8000000F Else textSTJ05(i).BackColor = &H80000005
         If bEnable Then lstAtt(i).BackColor = &H8000000F Else lstAtt(i).BackColor = &H80000005
         Me.cmdAddAtt(i).Enabled = Not bEnable
         Me.cmdRemAtt(i).Enabled = Not bEnable
      Next i
   End If
End Sub

Private Sub ClearField()
Dim nIndex As Integer
Dim i As Integer
   
   textSTJ01 = Empty
   textSTJ01_2 = Empty
   textSTJ02 = Empty
   For i = 0 To 2
      textSTJ04(i) = Empty
      textSTJ04_2(i) = Empty
      textSTJ05(i) = Empty
      lstAtt(i).Clear
      Label23(i) = Empty
      Me.SSTab2.Tag = Empty
   Next i
   SetGrd
   Me.SSTab2.TabEnabled(0) = True
   Me.SSTab2.TabEnabled(1) = True
   Me.SSTab2.TabEnabled(2) = True
End Sub

'帶預設資料
Private Sub InitialData()
   SetGrd
End Sub

Private Sub textSTJ01_GotFocus()
   If m_EditMode <> 0 Then
      InverseTextBox textSTJ01
      CloseIme
   End If
End Sub

Private Sub textSTJ01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textSTJ01_Validate(Cancel As Boolean)
Dim i As Integer

   If textSTJ01 = "" Then textSTJ01_2 = ""
   textSTJ01_2 = GetStaffName(textSTJ01, True)
   If m_EditMode = 1 And textSTJ01 <> "" Then
      If IsRecordExist(textSTJ01, Val(textSTJ02) + 191100) = True And _
         textSTJ01.Enabled = True And textSTJ01.Locked = False Then
         
         MsgBox "該員工當月已有資料！請先查詢再修改。", vbInformation
         Cancel = True
         textSTJ01 = ""
         m_EditMode = 0 '執行取消時,才不會彈詢問訊息
         OnAction vbKeyF10 '取消
         Exit Sub
      End If
      If textSTJ01_2 = "" Then
         MsgBox "員工編號錯誤！查無此員工！", vbInformation
         textSTJ01.SetFocus
         Cancel = True
         Exit Sub
      Else
         If InStr(m_STJLimitEmp, textSTJ01) = 0 And m_STJLimitEmp <> "ALL" Then
            MsgBox "您非" & Me.textSTJ01_2 & "的評價主管！", vbInformation
            textSTJ01.SetFocus
            Cancel = True
            Exit Sub
         ElseIf textSTJ01 = strUserNum Then
            MsgBox "欲評價的員工不可以是自己！", vbInformation
            textSTJ01.SetFocus
            Cancel = True
            Exit Sub
         Else
            strSql = "select B0130, '1' as Sort from ABS001 where B0101='" & textSTJ01 & "' and B0130='" & strUserNum & "'" & _
                     " union select B0131, '2' as Sort from ABS001 where B0101='" & textSTJ01 & "' and B0131='" & strUserNum & "'" & _
                     " union select B0132, '3' as Sort from ABS001 where B0101='" & textSTJ01 & "' and B0132='" & strUserNum & "'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If RsTemp.RecordCount = 0 Then
               MsgBox "您非" & Me.textSTJ01_2 & "的評價主管！", vbInformation
               textSTJ01.SetFocus
               Cancel = True
               Exit Sub
            Else
               Me.SSTab2.Tag = ""
               For i = 0 To 2
                  If i <= Val(RsTemp.Fields("Sort")) - 1 Then
                     Me.SSTab2.Tag = Me.SSTab2.Tag & "," & i
                  End If
                  If i = Val(RsTemp.Fields("Sort")) - 1 Then
                     Me.SSTab2.TabEnabled(i) = True
                     Me.SSTab2.Tab = i
                  Else
                     Me.SSTab2.TabEnabled(i) = False
                  End If
               Next i
               '預帶目前階層的評價主管
               strSql = "select B0130,B0131,B0132 from ABS001 where B0101='" & textSTJ01 & "'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strSql)
               If RsTemp.RecordCount = 1 Then
                  If "" & RsTemp.Fields(Me.SSTab2.Tab) <> "" Then
                     Me.textSTJ04(Me.SSTab2.Tab).Text = RsTemp.Fields(Me.SSTab2.Tab)
                  End If
               End If
               If Me.textSTJ04(Me.SSTab2.Tab) <> "" Then
                  textSTJ04_Validate Me.SSTab2.Tab, False
                  Me.textSTJ05(Me.SSTab2.Tab).SetFocus
               End If
            End If
         End If
      End If
   Else
      If m_EditMode = 4 Then
         OnAction vbKeyF9
         Exit Sub
      End If
   End If
End Sub

Private Sub textSTJ02_GotFocus()
   If m_EditMode <> 0 Then
      InverseTextBox textSTJ02
   End If
End Sub

Private Sub textSTJ02_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textSTJ02_Validate(Cancel As Boolean)
   If m_EditMode = 1 And textSTJ02 <> "" Then
      If IsRecordExist(textSTJ01, Val(textSTJ02) + 191100) = True And _
         textSTJ02.Enabled = True And textSTJ02.Locked = False Then
         
         MsgBox "該員工當月已有資料！請先查詢再修改。", vbInformation
         Cancel = True
         textSTJ02 = ""
         m_EditMode = 0 '執行取消時,才不會彈詢問訊息
         OnAction vbKeyF10 '取消
         Exit Sub
      End If
      If CheckIsTaiwanDate(textSTJ02 & "01", False) = False Then
         Cancel = True
         MsgBox "請輸入正確的年月！", vbInformation, "輸入年月錯誤"
         Exit Sub
      End If
   End If
End Sub

Private Sub textSTJ04_GotFocus(Index As Integer)
   If m_EditMode <> 0 Then
      InverseTextBox textSTJ04(Index)
      CloseIme
   End If
End Sub

Private Sub textSTJ04_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textSTJ04_Validate(Index As Integer, Cancel As Boolean)
   If textSTJ04(Index) = "" Then textSTJ04_2(Index) = ""
   If m_EditMode = 1 And textSTJ04(Index) <> "" Then
      textSTJ04_2(Index) = GetStaffName(textSTJ04(Index), True)
'      If IsRecordExist(textSTJ01, Val(textSTJ02) + 191100, Index + 1) = True And _
'         textSTJ04(Index).Enabled = True And textSTJ04(Index).Locked = False Then
'
'         MsgBox "該員工當月已有資料！請先查詢再修改。", vbInformation
'         Cancel = True
'         textSTJ04(Index) = ""
'         Exit Sub
'      End If
      If textSTJ04_2(Index) = "" Then
         MsgBox "主管員編錯誤！查無此員工！", vbInformation
         textSTJ04(Index).SetFocus
         Cancel = True
         Exit Sub
      Else
         strSql = "select B0130, '1' as Sort from ABS001 where B0101='" & textSTJ01 & "' and B0130='" & textSTJ04(Index) & "'" & _
                  " union select B0131, '2' as Sort from ABS001 where B0101='" & textSTJ01 & "' and B0131='" & textSTJ04(Index) & "'" & _
                  " union select B0132, '3' as Sort from ABS001 where B0101='" & textSTJ01 & "' and B0132='" & textSTJ04(Index) & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 0 Then
            MsgBox "您非" & Me.textSTJ01_2 & "的評價主管！", vbInformation
            textSTJ04(Index).Text = ""
            Cancel = True
            Exit Sub
         Else
            If Me.SSTab2.Tab <> Val(RsTemp.Fields("Sort")) - 1 Then
               MsgBox "您是主管" & RsTemp.Fields("Sort") & "，請輸入在正確的頁面上，謝謝！", vbInformation, "輸入錯誤！"
               textSTJ04(Index).Text = ""
               Cancel = True
               Exit Sub
            End If
         End If
      End If
   End If
End Sub

Private Sub textSTJ05_GotFocus(Index As Integer)
   InverseTextBox textSTJ05(Index)
   CloseIme
End Sub

Private Sub SetGrd()
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer
   
   arrGridHeadText = Array("年月", "部門", "員工代號", "姓名", "主管1_id", "主管1", "第一階評價", "附件1", _
      "主管2_id", "主管2", "第二階評價", "附件2", "主管3_id", "主管3", "第三階評價", "附件3")
   arrGridHeadWidth = Array(700, 1000, 0, 700, 0, 700, 1200, 500, _
      0, 700, 1200, 500, 0, 700, 1200, 500)
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

Private Sub txt1_GotFocus(Index As Integer)
   InverseTextBox txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
   Case 0, 1, 2, 3
           KeyAscii = UpperCase(KeyAscii)
   Case 4, 5
           KeyAscii = Pub_NumAscii(KeyAscii)
   Case Else
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   If txt1(Index) = "" Then Exit Sub
   Select Case Index
      Case 0, 1 '員工編號
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
            If RunNick(txt1(Index - 1), txt1(Index)) Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
      Case 2, 3 '部門
         If Index = 2 Then
            If txt1(Index) <> "" And txt1(Index + 1) = "" Then
               txt1(Index + 1) = txt1(Index)
            End If
         ElseIf Index = 3 Then
            If RunNick(txt1(Index - 1), txt1(Index)) Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
         SetCombo2
      Case 4, 5 '年月
         If Index = 4 Then
            If txt1(Index) <> "" And txt1(Index + 1) = "" Then
               txt1(Index + 1) = txt1(Index)
            End If
         ElseIf Index = 5 Then
            If RunNick2(txt1(Index - 1), txt1(Index)) Then
               Cancel = True
               Exit Sub
            End If
         End If
      Case Else
   End Select
End Sub
