VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm160010 
   BorderStyle     =   1  '單線固定
   Caption         =   "新年度特別假維護"
   ClientHeight    =   5430
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   8710
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   8710
   Begin TabDlg.SSTab SSTab2 
      Height          =   5415
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   8685
      _ExtentX        =   15311
      _ExtentY        =   9543
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "試算作業"
      TabPicture(0)   =   "frm160010.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmd(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "textYV01_1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmd(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "修改作業"
      TabPicture(1)   =   "frm160010.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "TBar1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "ImageList1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "SSTab1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "更新特別假"
      TabPicture(2)   =   "frm160010.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label3"
      Tab(2).Control(1)=   "cmd(3)"
      Tab(2).Control(2)=   "cmd(2)"
      Tab(2).Control(3)=   "Frame2"
      Tab(2).Control(4)=   "textYV01_2"
      Tab(2).ControlCount=   5
      Begin VB.TextBox textYV01_2 
         Height          =   285
         Left            =   -73890
         MaxLength       =   3
         TabIndex        =   37
         Top             =   510
         Width           =   825
      End
      Begin VB.Frame Frame2 
         Caption         =   "目前進度資料"
         Height          =   4425
         Left            =   -74880
         TabIndex        =   33
         Top             =   900
         Width           =   8445
         Begin VB.ListBox ListInfo2 
            Height          =   3640
            Left            =   60
            TabIndex        =   34
            Top             =   270
            Width           =   8325
         End
         Begin MSComctlLib.ProgressBar PB2 
            Height          =   255
            Left            =   30
            TabIndex        =   35
            Top             =   4110
            Width           =   8385
            _ExtentX        =   14781
            _ExtentY        =   459
            _Version        =   393216
            Appearance      =   0
            Scrolling       =   1
         End
      End
      Begin VB.CommandButton cmd 
         Caption         =   "更新(&S)"
         Height          =   405
         Index           =   2
         Left            =   -68400
         TabIndex        =   38
         Top             =   420
         Width           =   900
      End
      Begin VB.CommandButton cmd 
         Caption         =   "結束(&X)"
         Height          =   405
         Index           =   3
         Left            =   -67400
         TabIndex        =   39
         Top             =   420
         Width           =   900
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   4365
         Left            =   90
         TabIndex        =   20
         Top             =   990
         Width           =   8505
         _ExtentX        =   15011
         _ExtentY        =   7691
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "單筆資料"
         TabPicture(0)   =   "frm160010.frx":0054
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label2"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label9"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label11"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label12"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Label13"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Label14"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Label15"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "lblST13"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "Label4"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "LblBackDate_T"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "LblBackDate"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "Label23"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "textYV02_2"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "Grd2"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "textYV04"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "textYV01"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "textYV02"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "textYV03"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "text_m_01"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "textYV03_2"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).Control(20)=   "textYV11"
         Tab(0).Control(20).Enabled=   0   'False
         Tab(0).ControlCount=   21
         TabCaption(1)   =   "多筆瀏覽"
         TabPicture(1)   =   "frm160010.frx":0070
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "cmdok"
         Tab(1).Control(1)=   "txt1(3)"
         Tab(1).Control(2)=   "txt1(2)"
         Tab(1).Control(3)=   "txt1(1)"
         Tab(1).Control(4)=   "txt1(0)"
         Tab(1).Control(5)=   "GRD1"
         Tab(1).Control(6)=   "Label16"
         Tab(1).Control(7)=   "Label17"
         Tab(1).Control(8)=   "Line4"
         Tab(1).Control(9)=   "Line5"
         Tab(1).ControlCount=   10
         Begin VB.TextBox textYV11 
            Height          =   285
            Left            =   1350
            Locked          =   -1  'True
            MaxLength       =   7
            TabIndex        =   8
            Top             =   2640
            Width           =   1155
         End
         Begin VB.TextBox textYV03_2 
            Appearance      =   0  '平面
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  '沒有框線
            Height          =   195
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   1050
            Width           =   2235
         End
         Begin VB.CommandButton cmdok 
            Caption         =   "查詢"
            Height          =   255
            Left            =   -68580
            TabIndex        =   13
            Top             =   405
            Width           =   915
         End
         Begin VB.TextBox txt1 
            Height          =   270
            Index           =   3
            Left            =   -69960
            MaxLength       =   3
            TabIndex        =   12
            Top             =   390
            Width           =   915
         End
         Begin VB.TextBox txt1 
            Height          =   270
            Index           =   2
            Left            =   -70920
            MaxLength       =   3
            TabIndex        =   11
            Top             =   390
            Width           =   915
         End
         Begin VB.TextBox txt1 
            Height          =   270
            Index           =   1
            Left            =   -72840
            MaxLength       =   6
            TabIndex        =   10
            Top             =   390
            Width           =   915
         End
         Begin VB.TextBox txt1 
            Height          =   270
            Index           =   0
            Left            =   -73890
            MaxLength       =   6
            TabIndex        =   9
            Top             =   390
            Width           =   915
         End
         Begin VB.TextBox text_m_01 
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            MaxLength       =   4
            TabIndex        =   6
            Top             =   1320
            Width           =   1155
         End
         Begin VB.TextBox textYV03 
            Height          =   270
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   1020
            Width           =   495
         End
         Begin VB.TextBox textYV02 
            Height          =   270
            Left            =   1320
            MaxLength       =   6
            TabIndex        =   4
            Top             =   720
            Width           =   1005
         End
         Begin VB.TextBox textYV01 
            Height          =   285
            Left            =   1320
            MaxLength       =   3
            TabIndex        =   3
            Top             =   390
            Width           =   765
         End
         Begin VB.TextBox textYV04 
            Height          =   285
            Left            =   1320
            MaxLength       =   4
            TabIndex        =   7
            Top             =   1650
            Width           =   1155
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grd2 
            Height          =   3405
            Left            =   4590
            TabIndex        =   21
            Top             =   600
            Width           =   3825
            _ExtentX        =   6756
            _ExtentY        =   5997
            _Version        =   393216
            Cols            =   1
            FixedCols       =   0
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
            _Band(0).Cols   =   1
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
            Bindings        =   "frm160010.frx":008C
            Height          =   3585
            Left            =   -74940
            TabIndex        =   29
            Top             =   735
            Width           =   8340
            _ExtentX        =   14711
            _ExtentY        =   6332
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
         Begin MSForms.Label textYV02_2 
            Height          =   225
            Left            =   2400
            TabIndex        =   49
            Top             =   750
            Width           =   1395
            BackColor       =   12632256
            VariousPropertyBits=   27
            Size            =   "2461;397"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.Label Label23 
            Height          =   195
            Left            =   360
            TabIndex        =   48
            Top             =   4050
            Width           =   7785
            VariousPropertyBits=   27
            Caption         =   "CREATE :                                                    UPDATE : "
            Size            =   "13732;344"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label LblBackDate 
            Height          =   225
            Left            =   1860
            TabIndex        =   47
            Top             =   2280
            Width           =   2175
         End
         Begin VB.Label LblBackDate_T 
            AutoSize        =   -1  'True
            Caption         =   "留職停薪特休起算日："
            Height          =   180
            Left            =   60
            TabIndex        =   46
            Top             =   2310
            Width           =   1800
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "更新員工檔日期："
            Height          =   180
            Left            =   60
            TabIndex        =   45
            Top             =   2700
            Width           =   1440
         End
         Begin VB.Label lblST13 
            Height          =   225
            Left            =   1320
            TabIndex        =   32
            Top             =   2010
            Width           =   2175
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "年假年度："
            Height          =   180
            Left            =   -71790
            TabIndex        =   31
            Top             =   435
            Width           =   900
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "員工編號："
            Height          =   180
            Left            =   -74820
            TabIndex        =   30
            Top             =   435
            Width           =   900
         End
         Begin VB.Line Line4 
            X1              =   -73110
            X2              =   -72420
            Y1              =   540
            Y2              =   540
         End
         Begin VB.Line Line5 
            X1              =   -70230
            X2              =   -69630
            Y1              =   510
            Y2              =   510
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "任職時間："
            Height          =   180
            Left            =   4590
            TabIndex        =   28
            Top             =   390
            Width           =   900
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "入所日期："
            Height          =   180
            Left            =   420
            TabIndex        =   27
            Top             =   2010
            Width           =   900
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "今年特別假："
            Height          =   180
            Left            =   240
            TabIndex        =   26
            Top             =   1380
            Width           =   1080
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "年　　資："
            Height          =   180
            Left            =   420
            TabIndex        =   25
            Top             =   1050
            Width           =   900
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "員工編號："
            Height          =   180
            Left            =   420
            TabIndex        =   24
            Top             =   750
            Width           =   900
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "年度：                      (ex:97)"
            Height          =   180
            Left            =   780
            TabIndex        =   23
            Top             =   450
            Width           =   2040
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "新年度特別假："
            Height          =   180
            Left            =   60
            TabIndex        =   22
            Top             =   1680
            Width           =   1260
         End
      End
      Begin VB.CommandButton cmd 
         Caption         =   "結束(&X)"
         Height          =   405
         Index           =   1
         Left            =   -67770
         TabIndex        =   2
         Top             =   420
         Width           =   1245
      End
      Begin VB.TextBox textYV01_1 
         Height          =   285
         Left            =   -73890
         MaxLength       =   3
         TabIndex        =   0
         Top             =   420
         Width           =   825
      End
      Begin VB.CommandButton cmd 
         Caption         =   "開始試算(&S)"
         Height          =   405
         Index           =   0
         Left            =   -69090
         TabIndex        =   1
         Top             =   420
         Width           =   1245
      End
      Begin VB.Frame Frame1 
         Caption         =   "目前進度資料"
         Height          =   4425
         Left            =   -74880
         TabIndex        =   15
         Top             =   900
         Width           =   8445
         Begin VB.ListBox ListInfo 
            Height          =   2920
            Left            =   60
            TabIndex        =   19
            Top             =   270
            Width           =   8325
         End
         Begin MSComctlLib.ProgressBar PB1 
            Height          =   255
            Left            =   60
            TabIndex        =   16
            Top             =   3390
            Width           =   8325
            _ExtentX        =   14676
            _ExtentY        =   459
            _Version        =   393216
            Appearance      =   0
            Scrolling       =   1
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "注意："
            ForeColor       =   &H000000C0&
            Height          =   180
            Left            =   240
            TabIndex        =   43
            Top             =   3690
            Width           =   540
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "1. 試算資料為當時在職人員"
            ForeColor       =   &H000000C0&
            Height          =   180
            Left            =   240
            TabIndex        =   42
            Top             =   3930
            Width           =   2160
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "2. 若資料已做修改，不可再重新執行試算，否則修改資料將被還原"
            ForeColor       =   &H000000C0&
            Height          =   180
            Left            =   240
            TabIndex        =   41
            Top             =   4170
            Width           =   5220
         End
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   7950
         Top             =   360
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
               Picture         =   "frm160010.frx":00A1
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm160010.frx":03BD
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm160010.frx":06D9
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm160010.frx":08B5
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm160010.frx":0BD1
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm160010.frx":0EED
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm160010.frx":1209
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm160010.frx":1525
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm160010.frx":1841
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm160010.frx":1B5D
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm160010.frx":1E79
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar TBar1 
         Height          =   520
         Left            =   30
         TabIndex        =   17
         Top             =   330
         Width           =   8610
         _ExtentX        =   15187
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
               Enabled         =   0   'False
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
               Enabled         =   0   'False
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
               Enabled         =   0   'False
               Caption         =   "結束"
               Key             =   "keyExit"
               ImageIndex      =   11
            EndProperty
         EndProperty
         BorderStyle     =   1
      End
      Begin VB.Label Label5 
         Caption         =   "請於11月薪資作業完畢後再試算次年否則11月到職者無工作天可計算！"
         ForeColor       =   &H000000C0&
         Height          =   360
         Left            =   -72480
         TabIndex        =   44
         Top             =   480
         Width           =   2880
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "新年度：                   (ex:97)"
         Height          =   180
         Left            =   -74610
         TabIndex        =   36
         Top             =   550
         Width           =   2265
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "試算年度：                   (ex:97)"
         Height          =   180
         Left            =   -74790
         TabIndex        =   18
         Top             =   480
         Width           =   2265
      End
   End
End
Attribute VB_Name = "frm160010"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/6/16 Form2.0已修改
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Modify By Sindy 2010/7/20 日期欄已修改
'Create by nickc 2008/01/29 copy from frm140401
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
' 宣告欄位內容結構
Private Type FIELDITEM
   fiName As String
   fiOldData As String
   fiNewData As String
   fiType As Integer
End Type
Dim m_FieldList() As FIELDITEM
' 第一筆資料的本所案號
Dim m_FirstKEY(2) As String
' 最後一筆資料的本所案號
Dim m_LastKEY(2) As String
' 目前正在顯示的本所案號
Dim m_CurrKEY(2) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim tf_YV As Integer
Dim MyKind As String


Private Sub Form_Activate()
   SSTab2.Tab = 0
   textYV01_1.SetFocus
End Sub

Private Sub cmd_Click(Index As Integer)
Dim m_rs As New ADODB.Recordset
Dim m_StrSQL As String
   
   Select Case Index
      Case 0
        '檢查條件
        If textYV01_1 = "" Then
            MsgBox "請輸入年假年度！", vbExclamation, "操作錯誤！"
            textyv01_1_GotFocus
            Exit Sub
        End If
        '檢查是否已有資料
        Set m_rs = New ADODB.Recordset
        m_StrSQL = "select * from YearVacation where yv01='" & Mid(Trim(DBDATE(textYV01_1 & "0101")), 1, 4) & "'  "
        If m_rs.State = 1 Then m_rs.Close
        m_rs.CursorLocation = adUseClient
        m_rs.Open m_StrSQL, cnnConnection, adOpenStatic, adLockReadOnly
        If m_rs.RecordCount <> 0 Then
            If MsgBox("已經曾經產生過，是否全部重新產生??", vbExclamation + vbYesNo, "嚴重警告！") = vbNo Then
               textyv01_1_GotFocus
               Exit Sub
            End If
        End If
                
        CalBonus '計算年假
        InitialField
        InitialData
        RefreshRange
        ShowFirstRecord
        UpdateToolbarState
        SetCtrlReadOnly True
        Me.SSTab1.Tab = 0
      Case 1, 3
        Unload Me
      Case 2
        '檢查條件
        If textYV01_2 = "" Then
            MsgBox "請輸入年假年度！", vbExclamation, "操作錯誤！"
            textyv01_2_GotFocus
            Exit Sub
        End If
        '檢查是否已有新年度特別假資料
        Set m_rs = New ADODB.Recordset
        m_StrSQL = "select * from YearVacation where yv01='" & Mid(Trim(DBDATE(textYV01_2 & "0101")), 1, 4) & "'  "
        If m_rs.State = 1 Then m_rs.Close
        m_rs.CursorLocation = adUseClient
        m_rs.Open m_StrSQL, cnnConnection, adOpenStatic, adLockReadOnly
        If m_rs.RecordCount <= 0 Then
            MsgBox "無新年度特別假資料!!!", vbExclamation, "操作錯誤！"
            textyv01_2_GotFocus
            Exit Sub
        End If
        
        CalUpdVacation '更新特別假
        InitialField
        InitialData
        RefreshRange
        ShowFirstRecord
        UpdateToolbarState
        SetCtrlReadOnly True
        Me.SSTab1.Tab = 0
      Case Else
   End Select
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
       MsgBox "查詢條件不可以空白！", vbExclamation, "操作錯誤！"
   End If
End Sub

Private Sub Form_Initialize()
   Set rsA = New ADODB.Recordset
   If rsA.State = 1 Then rsA.Close
   rsA.CursorLocation = adUseClient
   rsA.Open "select * from YearVacation where rownum <2 ", cnnConnection, adOpenStatic, adLockReadOnly
   tf_YV = rsA.Fields.Count
   SetGrd
End Sub

' 按下按鍵
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
'edit by nickc 2006/11/13
'      Case vbKeyReturn:
'         If m_EditMode <> 0 Then
'            OnAction vbKeyF9
'         End If
      Case vbKeyEscape:
         If m_EditMode = 0 Then
            OnAction KeyCode
         Else
            OnAction vbKeyF10
         End If
   End Select
End Sub
'add by nickc 2006/11/13 Enter 事件，等於存檔，做完取消，不然 form 內其他物件有寫 keycode 或是 keyascii 事件的話，也會做到
Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
      Case vbKeyReturn:
         If m_EditMode <> 0 Then
            KeyAscii = 0
            OnAction vbKeyF9
         End If
    End Select
End Sub

Private Sub Form_Load()
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

   ReDim m_FieldList(tf_YV) As FIELDITEM
   
   m_bInsert = IsUserHasRightOfFunction(Me.Name, strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
   m_bDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   m_bQuery = IsUserHasRightOfFunction(Me.Name, strFind, False)
   
   textYV01.BackColor = &H8000000F
   textYV02.BackColor = &H8000000F
   
   MoveFormToCenter Me
   
   InitialField
   InitialData
   RefreshRange
   ShowFirstRecord
   UpdateToolbarState
   SetCtrlReadOnly True
   Me.SSTab1.Tab = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm160010 = Nothing
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
            '2008/12/12 ADD BY SONIA
            textYV01.Text = GRD1.TextMatrix(tmpMouseRow, 0)
            textYV02.Text = GRD1.TextMatrix(tmpMouseRow, 1)
            QueryRecord
            '2008/12/12 END
            GRD1.Visible = True
       End If
   End If
End Sub

'Add By Sindy 2019/8/27
Private Sub SSTab1_Click(PreviousTab As Integer)
   If PreviousTab = 0 Then
      cmdok.SetFocus
      cmdok.Default = True
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
Private Sub UpdateCUID(ByRef rsSrcTmp As ADODB.Recordset)
Dim strTemp As String
Dim strCName As String
Dim strCDate As String
Dim strCTime As String
Dim strUName As String
Dim strUDate As String
Dim strUTime As String
   
   If IsNull(rsSrcTmp.Fields("yv05")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("yv05")) = False Then
         strCName = GetStaffName(rsSrcTmp.Fields("yv05"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("yv06")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("yv06")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("yv06"))
         strCDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("yv07")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("yv07")) = False Then
         strTemp = rsSrcTmp.Fields("yv07")
         strCTime = Format(strTemp, "##:##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("yv08")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("yv08")) = False Then
         strUName = GetStaffName(rsSrcTmp.Fields("yv08"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("yv09")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("yv09")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("yv09"))
         strUDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("yv10")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("yv10")) = False Then
         strTemp = rsSrcTmp.Fields("yv10")
         strUTime = Format(strTemp, "##:##")
      End If
   End If
   
   ' 設定CUID中的文字
   Label23.Caption = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime & String(10, " ") & _
              "UPDATE : " & strUName & " " & _
              " " & strUDate & " " & _
              " " & strUTime
End Sub

Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

   TxtValidate = False
   If Me.textYV03.Enabled = True Then
      Cancel = False
      textyv03_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textYV01.Enabled = True Then
      Cancel = False
      textyv01_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If textYV01.Text = "" Then
       MsgBox "年假年度不可以空白！", vbExclamation
       textYV01.SetFocus
       Exit Function
   End If
   If Me.textYV02.Enabled = True Then
      Cancel = False
      textyv02_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   TxtValidate = True
End Function

' 設定欄位的內容
Private Sub SetFieldNewData(ByVal strName As String, Optional ByVal strData As String = "#==#")
Dim nIndex As Integer
   
   For nIndex = 0 To tf_YV - 1
      If strName = m_FieldList(nIndex).fiName Then
         If strData = "#==#" Then
            m_FieldList(nIndex).fiNewData = m_FieldList(nIndex).fiOldData
         Else
            m_FieldList(nIndex).fiNewData = strData
         End If
         Exit For
      End If
   Next nIndex
End Sub

' 從記錄中更新欄位內容
Private Sub UpdateFieldOldData(ByRef rsTmp As ADODB.Recordset)
Dim nIndex As Integer
Dim strTmp As String
   
   For nIndex = 0 To tf_YV - 1
      If m_FieldList(nIndex).fiName <> Empty Then
         If IsNull(rsTmp.Fields(m_FieldList(nIndex).fiName)) = False Then
            m_FieldList(nIndex).fiOldData = rsTmp.Fields(m_FieldList(nIndex).fiName)
            m_FieldList(nIndex).fiNewData = rsTmp.Fields(m_FieldList(nIndex).fiName)
         Else
            m_FieldList(nIndex).fiOldData = Empty
            m_FieldList(nIndex).fiNewData = Empty
         End If
      End If
   Next nIndex
EXITSUB:
End Sub

' 新增記錄
Private Function AddRecord() As Boolean
Dim strSql As String
Dim strTmp As String
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim nIndex As Integer
Dim bDifference As Boolean
Dim bFirst As Boolean
Dim strYV01 As String
Dim strYV02 As String
   
   AddRecord = False
   
   strYV02 = textYV02
   strYV01 = Mid(DBDATE(textYV01.Text & "0101"), 1, 4)

   ' 檢查記錄是否已存在
   If IsRecordExist(strYV01, strYV02) = True Then
      strTit = "新增資料"
      strMsg = "該筆記錄已存在"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      UpdateCtrlData
      Exit Function
   End If
   
   bFirst = True
   bDifference = False
   strSql = "INSERT INTO YearVacation ("
   For nIndex = 0 To tf_YV - 1
      strTmp = Empty
      If m_FieldList(nIndex).fiOldData <> m_FieldList(nIndex).fiNewData Then
         strTmp = m_FieldList(nIndex).fiName
      End If
      If strTmp <> Empty Then
         bDifference = True
         If bFirst = True Then
            strSql = strSql & strTmp
            bFirst = False
         Else
            strSql = strSql & "," & strTmp
         End If
      End If
   Next nIndex
   strSql = strSql & ") "
   strSql = strSql & "VALUES ("
   
   bFirst = True
   For nIndex = 0 To tf_YV - 1
      strTmp = Empty
      If m_FieldList(nIndex).fiOldData <> m_FieldList(nIndex).fiNewData Then
         If m_FieldList(nIndex).fiType = 0 Then
            strTmp = "'" & ChgSQL(m_FieldList(nIndex).fiNewData) & "'"
         Else
            strTmp = m_FieldList(nIndex).fiNewData
         End If
      End If
      If strTmp <> Empty Then
         If bFirst = True Then
            strSql = strSql & strTmp
            bFirst = False
         Else
            strSql = strSql & "," & strTmp
         End If
      End If
   Next nIndex
   strSql = strSql & ")"
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   
   cnnConnection.Execute strSql
   
   If ((strYV01 & strYV02) < (m_FirstKEY(0) & m_FirstKEY(1))) Or ((strYV01 & strYV02) > (m_LastKEY(0) & m_LastKEY(1))) Then
      RefreshRange
   End If
   cnnConnection.CommitTrans
   
   ShowCurrRecord strYV01, strYV02
   AddRecord = True
   Exit Function
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox " 新增失敗！" & vbCrLf & Err.Description
    
End Function

' 修改記錄
Private Function ModRecord() As Boolean
Dim strSql As String
Dim strTmp As String
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim nIndex As Integer
Dim bDifference As Boolean
Dim bFirst As Boolean
Dim strYV01 As String
Dim strYV02 As String
Dim strOB03 As String
       
   ModRecord = False
   
   strYV01 = m_CurrKEY(0)
   strYV02 = m_CurrKEY(1)
   
   strSql = "begin user_data.user_enabled:=1; UPDATE YearVacation SET "

   bFirst = True
   bDifference = False
   For nIndex = 0 To tf_YV - 1
      strTmp = Empty
      If nIndex < 4 Or nIndex > 9 Then
            If m_FieldList(nIndex).fiOldData <> m_FieldList(nIndex).fiNewData Then
               If m_FieldList(nIndex).fiType = 0 Then
                  If m_FieldList(nIndex).fiNewData = Empty Then
                     strTmp = m_FieldList(nIndex).fiName & " = NULL "
                  Else
                     strTmp = m_FieldList(nIndex).fiName & " = '" & ChgSQL(m_FieldList(nIndex).fiNewData) & "'"
                  End If
               Else
                  If m_FieldList(nIndex).fiNewData = Empty Then
                     strTmp = m_FieldList(nIndex).fiName & " = NULL "
                  Else
                     strTmp = m_FieldList(nIndex).fiName & " = " & m_FieldList(nIndex).fiNewData
                  End If
               End If
            End If
            If strTmp <> Empty Then
               bDifference = True
               If bFirst = True Then
                  strSql = strSql & strTmp
                  bFirst = False
               Else
                  strSql = strSql & "," & strTmp
               End If
            End If
        End If
   Next nIndex
   
   'Add By Sindy 2019/6/25
   If Val(textYV04.Tag) <> Val(textYV04.Text) Then
      '計算方案:C.人事處調整
      strSql = strSql & ",yv12='C'"
   End If
   '2019/6/25 END
   
   strSql = strSql & " " & _
                  "WHERE yv01 = '" & strYV01 & "' and yv02='" & strYV02 & "' ; end; "
On Error GoTo ErrHand
      cnnConnection.BeginTrans
        If bDifference = True Then
           Pub_SeekTbLog strSql
           cnnConnection.Execute strSql
        End If
        cnnConnection.CommitTrans

   ShowCurrRecord strYV01, strYV02
      
    ModRecord = True
   Exit Function
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox (Err.Description)

End Function

' 刪除記錄
Private Function DelRecord() As Boolean
Dim strSql As String
Dim strYV01 As String
Dim strYV02 As String

   DelRecord = False
   
On Error GoTo ErrHand
   
   cnnConnection.BeginTrans
   
   strYV01 = m_CurrKEY(0)
   strYV02 = m_CurrKEY(1)

   strSql = "DELETE FROM YearVacation " & _
            "WHERE yv01 = '" & strYV01 & "'  and yv02='" & strYV02 & "'  "

   cnnConnection.Execute strSql

   If (strYV01 = m_LastKEY(0) And strYV02 = m_LastKEY(1)) Or (strYV01 = m_FirstKEY(0) And strYV02 = m_FirstKEY(1)) Then
      RefreshRange
   End If
   ShowCurrRecord strYV01, strYV02
   DelRecord = True
   cnnConnection.CommitTrans
   
   Exit Function
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox "刪除失敗！" & vbCrLf & Err.Description
End Function

' 查詢記錄
Private Function QueryRecord() As Boolean
Dim strYV01 As String
Dim strYV02 As String
   
   QueryRecord = False
   strYV01 = Mid(DBDATE(textYV01.Text & "0101"), 1, 4)
   strYV02 = textYV02
   If IsRecordExist(strYV01, strYV02) = True Then
      m_CurrKEY(0) = strYV01
      m_CurrKEY(1) = strYV02
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
            If TxtValidate = False Then Exit Function
            UpdateFieldNewData
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
            'add by sonia 2019/1/28
            If Pub_StrUserSt03 = "M51" And Val(textYV01) + 1911 = Val(Left(strSrvDate(1), 4)) Then
               MsgBox "若為新人可休假日數移至次年，請先將前一年可休假日數扣除(取消此程式限制當年控管)，否則年終未休假代金未多算！次年日數請至員工檔修改！"
            End If
            'end 2019/1/28
            'Add By Cheng 2002/05/22
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Function
            UpdateFieldNewData
            If ModRecord = False Then Exit Function
         Else
            GoTo EXITSUB
         End If
      Case 3: '刪除
         If DelRecord = True Then
            RefreshRange
            ClearField
            ShowCurrRecord m_CurrKEY(0), m_CurrKEY(1)
         Else
            Exit Function
         End If
      Case 4: '查詢
         If textYV01 <> "" And textYV02 <> "" Then
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
      'Case 1: If Me.Visible = True Then textyv01.SetFocus
      Case 2: If Me.Visible = True Then textYV03.SetFocus
      Case 4: If Me.Visible = True Then textYV01.SetFocus
   End Select
End Sub

' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal strKEY01 As String, ByVal strKEY02 As String) As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   IsRecordExist = False
   strSql = "SELECT * FROM YearVacation " & _
            "WHERE yv01 = '" & strKEY01 & "'  and yv02='" & strKEY02 & "'  "
                  
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
Private Sub ShowCurrRecord(ByVal strKEY01 As String, ByVal strKEY02 As String)
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   If IsRecordExist(strKEY01, strKEY02) = True Then
      m_CurrKEY(0) = strKEY01
      m_CurrKEY(1) = strKEY02
   Else
      strSql = "SELECT yv01,yv02 FROM YearVacation " & _
               "WHERE yv01 = '" & m_CurrKEY(0) & "' and yv02='" & m_CurrKEY(1) & "'  "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("YV01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("YV01")
         If IsNull(rsTmp.Fields("YV02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("YV02")
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      End If
      rsTmp.Close
      
      strSql = "SELECT YV01,YV02 FROM YearVacation " & _
               "WHERE yv02 = (SELECT MIN(yv02) FROM YearVacation where yv01=(select min(yv01) from YearVacation) ) and yv01=(select min(yv01) from YearVacation) "
   
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("YV01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("YV01")
         If IsNull(rsTmp.Fields("YV02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("YV02")
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
   
   strSql = "SELECT yv01,yv02 FROM YearVacation " & _
            "WHERE yv01 = '" & m_CurrKEY(0) & "' AND " & _
                  "yv02 = (select max(yv02) from YearVacation where  yv01 = '" & m_CurrKEY(0) & "' and " & _
                                "yv02 < '" & m_CurrKEY(1) & "' ) "
                                
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("YV01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("YV01")
      If IsNull(rsTmp.Fields("YV02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("YV02")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
  
   strSql = "SELECT yv01,yv02 FROM YearVacation " & _
            "WHERE yv01 = (select max(yv01) from YearVacation where yv01< '" & m_CurrKEY(0) & "') AND " & _
                  "yv02 = (SELECT MAX(yv02) FROM YearVacation " & _
                          "WHERE yv01 = (select max(yv01) from YearVacation where yv01< '" & m_CurrKEY(0) & "') ) "
                                
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("YV01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("YV01")
      If IsNull(rsTmp.Fields("YV02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("YV02")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close

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
   
   strSql = "SELECT yv01,yv02 FROM YearVacation " & _
            "WHERE yv01 = '" & m_CurrKEY(0) & "' AND " & _
                  "yv02 = (select min(yv02) from YearVacation where  yv01  = '" & m_CurrKEY(0) & "' AND " & _
                                "yv02 > '" & m_CurrKEY(1) & "' )            "
                                
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("YV01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("YV01")
      If IsNull(rsTmp.Fields("YV02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("YV02")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
  
   strSql = "SELECT yv01,yv02 FROM YearVacation " & _
            "WHERE yv01 =(select min(yv01) from YearVacation where yv01>'" & m_CurrKEY(0) & "') AND " & _
                  "yv02 = (select min(yv02) from YearVacation where yv01 =(select min(yv01) from YearVacation where yv01>'" & m_CurrKEY(0) & "'))  "
                                
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("YV01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("YV01")
      If IsNull(rsTmp.Fields("YV02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("YV02")
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
   m_CurrKEY(1) = m_LastKEY(1)
   UpdateCtrlData
End Sub

' 執行指令
Private Sub OnAction(ByVal KeyCode As Integer)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   m_SubMode = 0
   Select Case KeyCode
      ' 新增
      Case vbKeyF2:
'         m_EditMode = 1
'         ClearField
'         Me.SSTab1.TabEnabled(1) = False
'         SSTab1.Tab = 0
'         SetCtrlReadOnly False
'         UpdateToolbarState
'         SetInputEntry
      ' 修改
      Case vbKeyF3:
         '當月資料才能修改
         If Val(textYV01) + 1911 < Val(Left(strSrvDate(1), 4)) Then
            strTit = "修改"
            strMsg = "非當年資料不可修改!!"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            Exit Sub
         End If
         m_EditMode = 2
         Me.SSTab1.TabEnabled(1) = False
         SSTab1.Tab = 0
         SSTab2.TabEnabled(0) = False
         SSTab2.Tab = 1
         SetCtrlReadOnly False
         SetKeyReadOnly True
         UpdateToolbarState
         SetInputEntry
      ' 刪除
      Case vbKeyF5:
'         strTit = "詢問"
'         strMsg = "是否要刪除此筆資料?"
'         nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit)
'         If nResponse = vbYes Then
'            m_EditMode = 3
'            If OnWork = True Then
'                UpdateToolbarState
'            Else
'                Exit Sub
'            End If
'         End If
      ' 查詢
      Case vbKeyF4:
         m_EditMode = 4
         Me.SSTab1.TabEnabled(1) = False
         SSTab1.Tab = 0
         SSTab2.TabEnabled(0) = False
         SSTab2.Tab = 1
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
         ' 將所有欄位的內容更新到欄位串列中的欄位內容項目
         UpdateFieldNewData
         If OnWork = True Then
            Me.SSTab1.TabEnabled(1) = True
            SSTab2.TabEnabled(0) = True
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
                  SSTab2.TabEnabled(0) = True
                  UpdateCtrlData
                  SetCtrlReadOnly True
                  UpdateToolbarState
               End If
            Case Else
               m_EditMode = 0
               Me.SSTab1.TabEnabled(1) = True
               SSTab2.TabEnabled(0) = True
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
   
   strSql = "SELECT yv01,yv02 FROM YearVacation " & _
            "WHERE yv01 = (SELECT MIN(yv01) FROM YearVacation) AND " & _
                  "yv02 = (SELECT MIN(yv02) FROM YearVacation " & _
                           "WHERE yv01 = (SELECT MIN(yv01) FROM YearVacation)) "
                           
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("YV01")) = False Then: m_FirstKEY(0) = rsTmp.Fields("YV01")
      If IsNull(rsTmp.Fields("YV02")) = False Then: m_FirstKEY(1) = rsTmp.Fields("YV02")
   End If
   rsTmp.Close

   strSql = "SELECT yv01,yv02 FROM YearVacation " & _
            "WHERE yv01 = (SELECT MAX(yv01) FROM YearVacation) AND " & _
                  "yv02 = (SELECT MAX(yv02) FROM YearVacation " & _
                           "WHERE yv01 = (SELECT MAX(yv01) FROM YearVacation)) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("YV01")) = False Then: m_LastKEY(0) = rsTmp.Fields("YV01")
      If IsNull(rsTmp.Fields("YV02")) = False Then: m_LastKEY(1) = rsTmp.Fields("YV02")
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim i As Integer, j As Integer
Dim strBackTaieDate As String
   
   strSql = "SELECT * FROM YearVacation,staff " & _
            "WHERE yv01='" & m_CurrKEY(0) & "' and YV02 = '" & m_CurrKEY(1) & "' and yv02(+)=st01  "
                  
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      ClearField
      If IsNull(rsTmp.Fields("YV01")) = False Then
         textYV01.Text = CheckStr(rsTmp.Fields("YV01")) - 1911
      End If
      If IsNull(rsTmp.Fields("YV02")) = False Then: textYV02 = rsTmp.Fields("YV02")
      If IsNull(rsTmp.Fields("YV03")) = False Then: textYV03 = rsTmp.Fields("YV03")
      If IsNull(rsTmp.Fields("YV04")) = False Then: textYV04 = rsTmp.Fields("YV04")
      If IsNull(rsTmp.Fields("ST40")) = False Then text_m_01 = rsTmp.Fields("ST40")
      If IsNull(rsTmp.Fields("YV11")) = False Then: textYV11 = ChangeWStringToTString(rsTmp.Fields("YV11")) 'Add By Sindy 2018/5/16
      
      textYV02_2 = GetStaffName(textYV02, True)
      If Val(textYV03) > 0 Then
         textYV03_2 = "為  " & PUB_ChangeNianZi(Val(textYV03))
      Else
         textYV03_2 = ""
      End If
      
      ' 更新CUID
      UpdateCUID rsTmp
      ' 更新暫存區的資料
      UpdateFieldOldData rsTmp

      '讀取入所日
      strSql = "select * from staff where st01='" & textYV02 & "' "
      If rsTmp.State = 1 Then rsTmp.Close
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
          lblST13.Caption = ChangeWStringToTDateString(CheckStr(rsTmp.Fields("st13")))
      Else
          lblST13.Caption = ""
      End If
      '讀取任職時間
      ' 2008/12/26 Modify BY SINDY
      'strSQL = "select sqldatet(sc02) as 日期,ac03 as 原因 from staff_change,allcode where sc01='" & textYV02 & "' and sc03 in ('0001','0003','0006','0008') and ac01='05' and sc03=ac02(+) order by sc02 "
      strSql = "select sqldatet(sc02) as 日期,ac03 as 原因 from staff_change,allcode where sc01='" & textYV02 & "' and ac01='05' and sc03=ac02(+) order by sc02 ASC "
      ' 2008/12/26 M
      If rsTmp.State = 1 Then rsTmp.Close
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      Set Grd2.Recordset = rsTmp
      'Add By Sindy 2021/6/16
      If rsTmp.RecordCount > 0 Then
         Grd2.row = 1
         Grd2.col = 0
      End If
      '2021/6/16 END
      
      'Add By Sindy 2019/6/25
      strBackTaieDate = Pub_BackTaieToDate(m_CurrKEY(1), textYV01)
      If Val(strBackTaieDate) > 0 Then
         LblBackDate.Caption = ChangeWStringToTDateString(strBackTaieDate)
      End If
      textYV04.Tag = textYV04.Text '記錄原本系統計算出來的天數
      '2019/6/25 END
   End If
   
   rsTmp.Close
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

Sub GetData()
Dim rsTmp As New ADODB.Recordset

   strSql = ""
   If txt1(0) <> "" Then
       strSql = strSql & " and yv02>='" & txt1(0) & "' "
   End If
   If txt1(1) <> "" Then
       strSql = strSql & " and yv02<='" & txt1(1) & "' "
   End If
   If txt1(2) <> "" Then
       strSql = strSql & " and yv01>='" & Val(txt1(2)) + 1911 & "' "
   End If
   If txt1(3) <> "" Then
       strSql = strSql & " and yv01<='" & Val(txt1(3)) + 1911 & "' "
   End If
   '抓取資料
   strSql = "SELECT yv01-1911,yv02,st02,yv03,yv04 FROM YearVacation,staff where yv02(+)=st01 " & strSql & _
           " order by yv01,yv02 "
   If rsTmp.State = 1 Then rsTmp.Close
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   Set GRD1.Recordset = rsTmp
   SetGrd
End Sub

' 更新toolbar上按紐的狀態
Private Sub UpdateToolbarState()
   Select Case m_EditMode
      ' 無任何動作
      Case 0:
'         If m_bInsert Then
'            TBar1.Buttons(1).Enabled = True
'         Else
'            TBar1.Buttons(1).Enabled = False
'         End If
         If m_bUpdate Then
            TBar1.Buttons(2).Enabled = True
         Else
            TBar1.Buttons(2).Enabled = False
         End If
'         If m_bDelete Then
'            TBar1.Buttons(3).Enabled = True
'         Else
'            TBar1.Buttons(3).Enabled = False
'         End If
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
   textyv01_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textyv02_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textyv03_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   CheckDataValid = True
EXITSUB:
End Function

' 更新Key的狀態
Private Sub SetKeyReadOnly(ByVal bEnable As Boolean)
   textYV01.Locked = bEnable
   textYV02.Locked = bEnable
   If bEnable Then textYV01.BackColor = &H8000000F Else textYV01.BackColor = &H80000005
   If bEnable Then textYV02.BackColor = &H8000000F Else textYV02.BackColor = &H80000005
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
Dim i As Integer
   
   textYV01.Locked = bEnable
   textYV02.Locked = bEnable
   textYV03.Locked = bEnable
   textYV04.Locked = bEnable
   If bEnable Then textYV01.BackColor = &H8000000F Else textYV01.BackColor = &H80000005
   If bEnable Then textYV02.BackColor = &H8000000F Else textYV02.BackColor = &H80000005
End Sub

Private Sub ClearField()
Dim nIndex As Integer
   
   textYV01 = Empty
   textYV02 = Empty
   textYV02_2 = Empty
   textYV03 = Empty
   textYV04 = Empty
   textYV11 = Empty 'Add By Sindy 2018/5/16
   text_m_01 = Empty
   Label23 = Empty
   SetGrd
   For nIndex = 0 To tf_YV - 1
      m_FieldList(nIndex).fiOldData = Empty
      m_FieldList(nIndex).fiNewData = Empty
   Next nIndex
   LblBackDate.Caption = Empty 'Add By Sindy 2019/6/25
End Sub

Private Sub UpdateFieldNewData()
Dim MyArr As Variant
   
   '若新增資料
   If m_EditMode = 1 Then
      SetFieldNewData "YV01", Mid(DBDATE(textYV01.Text & "0101"), 1, 4)
      SetFieldNewData "YV02", textYV02
   End If
   SetFieldNewData "YV03", textYV03
   SetFieldNewData "YV04", textYV04 '2009/12/15 add by sonia
End Sub

' 初始化欄位陣列
Private Sub InitialField()
Dim nIndex As Integer
Dim strTmp As String
   
   ' 初始化欄位陣列
   For nIndex = 1 To tf_YV
      strTmp = Format(nIndex, "00")
      m_FieldList(nIndex - 1).fiName = "YV" & strTmp
      m_FieldList(nIndex - 1).fiOldData = Empty
      m_FieldList(nIndex - 1).fiNewData = Empty
      m_FieldList(nIndex - 1).fiType = 0 '文字型態
      Select Case nIndex
         Case 1, 3, 4, 11
            m_FieldList(nIndex - 1).fiType = 1 '數值型態
      End Select
   Next nIndex
End Sub

'帶預設資料
Private Sub InitialData()
   SetGrd
   SetGrd2
End Sub

Private Sub textyv01_2_GotFocus()
   InverseTextBox textYV01_2
End Sub

Private Sub textyv01_2_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textyv01_2_Validate(Cancel As Boolean)
   If textYV01_2.Text = "" Then Exit Sub
   If CheckIsTaiwanDate(textYV01_2.Text & "0101", False) = False Then
       Cancel = True
       MsgBox "請輸入民國年度！", vbInformation, "輸入新年度錯誤"
       Exit Sub
   End If
End Sub

Private Sub textyv01_GotFocus()
   InverseTextBox textYV01
End Sub

Private Sub textyv01_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textyv01_Validate(Cancel As Boolean)
   If m_EditMode = 1 And textYV01.Text <> "" Then
       If IsRecordExist(Mid(DBDATE(textYV01.Text & "0101"), 1, 4), textYV02) = True And textYV01.Enabled = True And textYV01.Locked = False Then
           MsgBox "該員工當年度已有資料，請修改！", vbInformation
           Cancel = True
           Exit Sub
       End If
       If CheckIsTaiwanDate(textYV01.Text & "0101", False) = False Then
           Cancel = True
           MsgBox "請輸入民國年度！", vbInformation, "輸入年假年度錯誤"
           Exit Sub
       End If
   End If
End Sub

Private Sub textyv01_1_GotFocus()
   InverseTextBox textYV01_1
End Sub

Private Sub textyv01_1_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textyv01_1_Validate(Cancel As Boolean)
   If textYV01_1.Text = "" Then Exit Sub
   If CheckIsTaiwanDate(textYV01_1.Text & "0101", False) = False Then
       Cancel = True
       MsgBox "請輸入民國年度！", vbInformation, "輸入年假年度錯誤"
       Exit Sub
   End If
End Sub

Private Sub textyv02_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textYV02
   End If
End Sub

Private Sub textyv02_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textyv02_Validate(Cancel As Boolean)
   If textYV02.Text = "" Then textYV02_2.Caption = "" ' 2008/12/18 ADD BY SINDY
   
   If m_EditMode <> 0 And textYV02 <> "" Then
       textYV02_2 = GetStaffName(textYV02, True)
       ' 2008/12/18 ADD BY SINDY
       ' 檢查員工編號規則
       If ChkStaffID(textYV02) Then
          Call textyv02_GotFocus
          Cancel = True
          Exit Sub
       End If
       ' 2008/12/18 END
       If textYV02_2 = "" Then
           MsgBox "員工編號錯誤！查無此員工！", vbInformation
           Call textyv02_GotFocus ' 2008/12/18 ADD BY SINDY
           Cancel = True
           Exit Sub
       End If
   End If
   
   If m_EditMode = 1 And textYV01 <> "" And textYV02 <> "" Then
       If IsRecordExist(textYV01, textYV02) = True And textYV02.Enabled = True And textYV02.Locked = False Then
           MsgBox "該員工當天已有資料，請修改！", vbInformation
           Call textyv02_GotFocus ' 2008/12/18 ADD BY SINDY
           Cancel = True
           Exit Sub
       End If
   End If
End Sub

Private Sub SetGrd()
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer
   
   arrGridHeadText = Array("年假年度", "員工編號", "姓名", "年資", "年假天數")
   arrGridHeadWidth = Array(800, 800, 1200, 800, 800)
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

Private Sub textyv03_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textYV03
   End If
End Sub

Private Sub textyv03_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii, True)
End Sub

Private Sub textyv03_Validate(Cancel As Boolean)
   If m_EditMode <> 0 And textYV03 <> "" Then
       If IsNumeric(textYV03) = False Then
           MsgBox "請輸入數字！", vbExclamation, "操作錯誤！"
           Cancel = True
           Exit Sub
       End If
   End If
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   InverseTextBox txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
   Case 0, 1
           KeyAscii = UpperCase(KeyAscii)
   Case 2, 3
           KeyAscii = Pub_NumAscii(KeyAscii)
   Case Else
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   If txt1(Index) = "" Then Exit Sub
   Select Case Index
      Case 0, 1
            ' 2008/12/26 ADD BY SINDY
            If Index = 0 Then
               If txt1(Index) <> "" And txt1(Index + 1) = "" Then
                  txt1(Index + 1) = txt1(Index)
               End If
            ' 2008/12/26 END
            ElseIf Index = 1 Then
               If RunNick(txt1(Index - 1), txt1(Index)) Then
                   Call txt1_GotFocus(Index)
                   Cancel = True
                   Exit Sub
               End If
            End If
      Case 2, 3
           If CheckIsTaiwanDate(txt1(Index) & "0101", False) = False Then
               Cancel = True
               MsgBox "請輸入民國年度不含/！", vbInformation, "輸入年假年度錯誤"
               Call txt1_GotFocus(Index)
               Exit Sub
           End If
           ' 2008/12/26 ADD BY SINDY
            If Index = 2 Then
               If txt1(Index) <> "" And txt1(Index + 1) = "" Then
                  txt1(Index + 1) = txt1(Index)
               End If
            ' 2008/12/26 END
            ElseIf Index = 3 Then
               If RunNick2(txt1(Index - 1), txt1(Index)) Then
                   Call txt1_GotFocus(Index)
                   Cancel = True
                   Exit Sub
               End If
            End If
      Case Else
   End Select
End Sub

Private Sub SetGrd2()
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer
   
   arrGridHeadText = Array("日期", "狀態")
   arrGridHeadWidth = Array(1000, 1000)
   Grd2.Visible = False
   Grd2.Cols = UBound(arrGridHeadText) + 1
   For iRow = 0 To Grd2.Cols - 1
      Grd2.row = 0
      Grd2.col = iRow
      Grd2.Text = arrGridHeadText(iRow)
      Grd2.ColWidth(iRow) = arrGridHeadWidth(iRow)
      Grd2.CellAlignment = flexAlignCenterCenter
   Next
   Grd2.Visible = True
End Sub

'106年新特別假計算方式
'計算年假
Sub CalBonus()
Dim m_rs As New ADODB.Recordset
Dim m_rs2 As New ADODB.Recordset
Dim m_rs3 As New ADODB.Recordset 'Add By Sindy 2017/1/3
Dim m_StrSQL As String
Dim m_StrSQL2 As String
Dim m_Year As String
Dim m_item As Long
Dim m_Days As Double
Dim strStarDate As String, strEndDate As String, strScSql As String
Dim m_YearDay As Long       '年度總天數
Dim LongWorkDay As Long, m_Year1_Std As String 'Add By Sindy 2017/1/3
Dim m_Type As String, m_Formulation As String 'Add By Sindy 2019/6/18
   
   Screen.MousePointer = vbHourglass
   
   Me.Enabled = False
   ListInfo.Clear
   
   PB1.Value = 0
   m_item = 0
   ListInfo.AddItem "開始統計   " & textYV01_1.Text & " 年度年假資料", m_item
   ListInfo.Selected(m_item) = True
   m_item = m_item + 1
'On Error GoTo err1
   cnnConnection.BeginTrans
   
   '若先已經存在-->刪除
   Set m_rs = New ADODB.Recordset
   m_StrSQL = "select * from YearVacation where yv01='" & Mid(Trim(DBDATE(textYV01_1 & "0101")), 1, 4) & "'  "
   If m_rs.State = 1 Then m_rs.Close
   m_rs.CursorLocation = adUseClient
   m_rs.Open m_StrSQL, cnnConnection, adOpenStatic, adLockReadOnly
   If m_rs.RecordCount <> 0 Then
       'Pub_SeekTbLog "delete from YearVacation where yv01='" & Mid(Trim(DBDATE(textYV01_1 & "0101")), 1, 4) & "' "
       cnnConnection.Execute "delete from YearVacation where yv01='" & Mid(Trim(DBDATE(textYV01_1 & "0101")), 1, 4) & "' "
   End If
   
   '取得計算年度之前年總天數
   If PUB_GetMonthDays((Val(textYV01_1) + 1911 - 1), 2) = 28 Then
      m_YearDay = 365
   Else
      m_YearDay = 366
   End If
   
   strStarDate = CStr(Val(CStr(Val(textYV01_1) - 1) & "0101") + 19110000)
   strEndDate = CStr(Val(CStr(Val(textYV01_1) - 1) & "1231") + 19110000)
   strScSql = ""
   '抓留職停薪員工
   Set m_rs2 = New ADODB.Recordset
   m_StrSQL2 = "select sc01 " & _
                          "from staff_change " & _
                          "where sc03='04' and sc02 between '" & strStarDate & "' and '" & strEndDate & "' " & _
                          "and sc02 = (select max(a.sc02) from staff_change a where a.sc01=staff_change.sc01) " & _
                          "order by sc01,sc02 asc "
   If m_rs2.State = 1 Then m_rs2.Close
   m_rs2.CursorLocation = adUseClient
   m_rs2.Open m_StrSQL2, cnnConnection, adOpenStatic, adLockReadOnly
   If m_rs2.RecordCount <> 0 Then
       m_rs2.MoveFirst
       Do While Not m_rs2.EOF
       If strScSql = "" Then
            strScSql = "'" & CheckStr(m_rs2.Fields("sc01")) & "'"
       Else
            strScSql = strScSql & ",'" & CheckStr(m_rs2.Fields("sc01")) & "'"
       End If
       m_rs2.MoveNext
       Loop
   End If
   
   '開始計算
   Set m_rs = New ADODB.Recordset
   ' 2008/12/26 Modify BY SINDY
   'm_StrSQL = "select * from staff where st04='1' and ascii(substr(st01,1,1))>=48 and ascii(substr(st01,1,1))<=57  " & _
   '                    " and ((length(st01)=5 and substr(st01,1,1) not in ('0','1','2','3','4','5')) and st01 not in ('99998','99999')) " & _
   '                    " order by st01  "
   'Modify By Sindy 2024/4/30 + and not(substr(st01,5,1)>='A') 排除 B309A=宗家澔
   If strScSql = "" Then
      m_StrSQL = "select * from staff,SalaryData " & _
                           "where ST01=SD01 " & _
                           "and ST04='1' and sd02 not in('P','F') and not(substr(st01,5,1)>='A') " & _
                           "order by st01 ASC "
   Else
      m_StrSQL = "select * from staff,SalaryData " & _
                           "where ST01=SD01 " & _
                           "and ((ST04='1' and sd02 not in('P','F')) or st01 in (" & strScSql & ")) " & _
                           "and not(substr(st01,5,1)>='A') " & _
                           "order by st01 ASC "
   End If
   ' 2008/12/26 END
   If m_rs.State = 1 Then m_rs.Close
   m_rs.CursorLocation = adUseClient
   m_rs.Open m_StrSQL, cnnConnection, adOpenStatic, adLockReadOnly
   If m_rs.RecordCount <> 0 Then
       PB1.Min = 0
       PB1.max = m_rs.RecordCount
       m_rs.MoveFirst
       Do While Not m_rs.EOF
           PB1.Value = m_rs.AbsolutePosition
           If CheckStr(m_rs.Fields("st13")) = "" Then
               ListInfo.AddItem "員工 " & CheckStr(m_rs.Fields("st01")) & " " & CheckStr(m_rs.Fields("st02")) & "-->計算失敗，原因  沒有到職日 ", m_item
               ListInfo.Selected(m_item) = True
               m_item = m_item + 1
           Else
               'Modify By Sindy 2019/6/18 計算年資和特別假天數
               m_Days = PUB_GetSeniorityYearVacation(CheckStr(m_rs.Fields("st01")), textYV01_1, m_Year, m_Type, m_Formulation)
'               '計算年資
'               m_Year = Trim(CalYear(CheckStr(m_rs.Fields("st01")), strEndDate))
'               If Val(m_Year) < 0 Then m_Year = "0"
               If m_Year = "" Then
                   ListInfo.AddItem "員工 " & CheckStr(m_rs.Fields("st01")) & " " & CheckStr(m_rs.Fields("st02")) & "-->計算失敗，原因  統計錯誤 ，請聯絡電腦中心檢查  ", m_item
                   ListInfo.Selected(m_item) = True
                   m_item = m_item + 1
               Else
'                   m_Days = 0
'                   If Val(m_Year) >= 10 Then '滿 10 年
'                       m_Days = 16 + (Int(m_Year) - 10) '滿10年者以16天起算
'                       If m_Days > 30 Then
'                           m_Days = 30
'                       End If
'                   ElseIf Val(m_Year) >= 5 Then    '滿 5 年
'                       m_Days = 15
'                   ElseIf Val(m_Year) >= 3 Then ' 滿 3 年
'                       m_Days = 14
'                   ElseIf Val(m_Year) >= 2 Then ' 滿 2 年
'                       m_Days = 10
'                   ElseIf Val(m_Year) >= 1 Then  '滿1年
'                       m_Days = 7
'                  'Modify by Sindy 2017/1/3 未滿一年都是每天跑批次計算特別假
''                   ElseIf Val(m_Year) >= 0.5 Then   '滿6月
''                       m_Days = 3
'                   End If
'                   If m_Days < 0 Then m_Days = 0
'
''                   '到職日為前一年者
''                   If CheckStr(m_rs.Fields("st13")) > Val((Val(textYV01_1) + 1911 - 1) & "0101") Then
''                        '2009/12/3 add by sonia 每年12/1到職者特別假給0.5天
''                        If CheckStr(m_rs.Fields("st13")) = Val((Val(textYV01_1) + 1911 - 1) & "1201") Then
''                           m_Days = 0.5
''                        '每年12/1以後到職者無特別假
''                        ElseIf CheckStr(m_rs.Fields("st13")) > Val((Val(textYV01_1) + 1911 - 1) & "1201") Then
''                           m_Days = 0
'''                        Else
'''                        '2009/12/3 end
'''2009/12/15 CANCEL BY SONIA 上面已處理不必再做
'''                           '取得計算年度之前年工作總時數
'''                           m_StrSQL2 = "select sum(nvl(sm27,0))+31 " & _
'''                                       "from salarymonth " & _
'''                                       "where sm01='" & CheckStr(m_rs.Fields("st01")) & "' " & _
'''                                       "and sm02>='" & CStr(Val(textYV01_1) + 1911 - 1) & "01" & "' " & _
'''                                       "and sm02<='" & CStr(Val(textYV01_1) + 1911 - 1) & "12" & "' "
'''                           If m_rs2.State = 1 Then m_rs2.Close
'''                           m_rs2.CursorLocation = adUseClient
'''                           m_rs2.Open m_StrSQL2, cnnConnection, adOpenStatic, adLockReadOnly
'''                           If m_rs2.RecordCount <> 0 Then
'''                               m_rs2.MoveFirst
'''                               If m_YearDay <> CLng(CheckStr(m_rs2.Fields(0))) Then
'''                                    '因前年未工作滿一年, 所以特別假要按照工作總時數計算
'''                                    'm_Days = Round(m_Days * CLng(CheckStr(m_rs2.Fields(0))) / m_YearDay, 1)
'''                                    m_Days = Round(7 * CLng(CheckStr(m_rs2.Fields(0))) / m_YearDay, 0)
'''                               End If
'''                           End If
'''2009/12/15 end
''                        End If
''                   End If
'
'                   'Add By Sindy 2017/1/3 滿一年以上,當年復職者,依工作天數比例給假
'                   If Val(m_Year) >= 1 Then
'                     '為當年復職者
'                     strSql = "select sc02 " & _
'                              "from staff_change " & _
'                              "where sc03='02' and substr(sc02,1,4)='" & CStr(Val(textYV01_1) + 1911 - 1) & "' " & _
'                              "and sc01='" & m_rs.Fields("st01") & "' "
'                     If m_rs3.State = 1 Then m_rs3.Close
'                     m_rs3.CursorLocation = adUseClient
'                     m_rs3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'                     LongWorkDay = 0
'                     If m_rs3.RecordCount > 0 Then
'                        '抓該年的第一個工作天
'                        m_Year1_Std = GetYearStdDay(CStr(Val(textYV01_1) + 1911 - 1))
'                        '假如 該年的第一個工作天與復職日同一天, 算做滿整年不用算比例給假
'                        If m_rs3.Fields("sc02") <> m_Year1_Std Then
'                           If Mid(m_rs3.Fields("sc02"), 5, 2) = "12" Then '復職日為12月份時
'                              Call PUB_NianZiDaysYear(m_rs3.Fields("sc02"), CStr(Val(textYV01_1) + 1911 - 1) & "1231", LongWorkDay, 0) '工作天數
'                           Else
'                              '檢查12月份薪水是否已產生
'                              strSql = "select sum(nvl(sm27,0)) from SalaryMonth where sm01='" & m_rs.Fields("st01") & "' " & _
'                                       "and sm02='" & CStr(Val(textYV01_1) + 1911 - 1) & "12' "
'                              If m_rs3.State = 1 Then m_rs3.Close
'                              m_rs3.CursorLocation = adUseClient
'                              m_rs3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'                              If m_rs3.RecordCount > 0 Then
'                                 If Val("" & m_rs3.Fields(0)) > 0 Then
'                                    LongWorkDay = m_rs3.Fields(0)
'                                 Else
'                                    LongWorkDay = 31
'                                 End If
'                              Else
'                                 LongWorkDay = 31
'                              End If
'                              strSql = "select sum(nvl(sm27,0)) from SalaryMonth where sm01='" & m_rs.Fields("st01") & "' " & _
'                                       "and sm02>='" & CStr(Val(textYV01_1) + 1911 - 1) & "01' " & _
'                                       "and sm02<='" & CStr(Val(textYV01_1) + 1911 - 1) & "11' "
'                              If m_rs3.State = 1 Then m_rs3.Close
'                              m_rs3.CursorLocation = adUseClient
'                              m_rs3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'                              If m_rs3.RecordCount > 0 Then
'                                 If m_rs3.Fields(0) > 0 Then
'                                    LongWorkDay = LongWorkDay + m_rs3.Fields(0)
'                                 End If
'                              End If
'                           End If
'                           m_Days = Round(m_Days * (LongWorkDay / m_YearDay), 1)
'                        End If
'                     End If
'                   End If
'                   '2017/1/3 END
                   '2019/6/18 END
                   
                   'Modify By Sindy 2019/6/18 + ,yv12,yv13
                   cnnConnection.Execute " insert into YearVacation (yv01,yv02,yv03,yv04,yv12,yv13) values ('" & Mid(Trim(DBDATE(textYV01_1 & "0101")), 1, 4) & "','" & CheckStr(m_rs.Fields("st01")) & "','" & m_Year & "','" & m_Days & "','" & m_Type & "','" & m_Formulation & "') "
                   ListInfo.AddItem "員工 " & CheckStr(m_rs.Fields("st01")) & " " & CheckStr(m_rs.Fields("st02")) & "-->計算成功          入所日：" & ChangeWStringToTDateString(CheckStr(m_rs.Fields("st13"))) & "    年資：" & m_Year & "；年假：" & m_Days, m_item
                   ListInfo.Selected(m_item) = True
                   m_item = m_item + 1
               End If
           End If
           DoEvents
           m_rs.MoveNext
       Loop
   End If
   ListInfo.AddItem "統計結束;　共 " & m_rs.RecordCount & " 筆", m_item
   ListInfo.Selected(m_item) = True
   m_item = m_item + 1
   cnnConnection.CommitTrans
   Me.Enabled = True
   Screen.MousePointer = vbDefault
   Exit Sub
err1:
    cnnConnection.RollbackTrans
    ListInfo.AddItem "發生錯誤！所有剛剛統計的資料將會還原！", m_item
    m_item = m_item + 1
    ListInfo.AddItem "錯誤碼：" & Trim(Err.Number) & "   " & Err.Description, m_item
    m_item = m_item + 1
    ListInfo.AddItem "請通知電腦中心處理！謝謝！", m_item
    m_item = m_item + 1
    Me.Enabled = True
    Screen.MousePointer = vbDefault
End Sub

'105年(含)以前使用之計算方式
''計算年假
'Sub CalBonus()
'Dim m_rs As New ADODB.Recordset
'Dim m_rs2 As New ADODB.Recordset
'Dim m_StrSQL As String
'Dim m_StrSQL2 As String
'Dim m_Year As String
'Dim m_item As Long
'Dim m_Days As Double
'Dim strStarDate As String, strEndDate As String, strScSql As String
'Dim m_YearDay As Long       '年度總天數
'
'   Screen.MousePointer = vbHourglass
'
'   Me.Enabled = False
'   ListInfo.Clear
'
'   PB1.Value = 0
'   m_item = 0
'   ListInfo.AddItem "開始統計   " & textYV01_1.Text & " 年度年假資料", m_item
'   ListInfo.Selected(m_item) = True
'   m_item = m_item + 1
'   On Error GoTo err1
'   cnnConnection.BeginTrans
'
'   '若先已經存在-->刪除
'   Set m_rs = New ADODB.Recordset
'   m_StrSQL = "select * from YearVacation where yv01='" & Mid(Trim(DBDATE(textYV01_1 & "0101")), 1, 4) & "'  "
'   If m_rs.State = 1 Then m_rs.Close
'   m_rs.CursorLocation = adUseClient
'   m_rs.Open m_StrSQL, cnnConnection, adOpenStatic, adLockReadOnly
'   If m_rs.RecordCount <> 0 Then
'       'Pub_SeekTbLog "delete from YearVacation where yv01='" & Mid(Trim(DBDATE(textYV01_1 & "0101")), 1, 4) & "' "
'       cnnConnection.Execute "delete from YearVacation where yv01='" & Mid(Trim(DBDATE(textYV01_1 & "0101")), 1, 4) & "' "
'   End If
'
'   '取得計算年度之前年總天數
'   If PUB_GetMonthDays((Val(textYV01_1) + 1911 - 1), 2) = 28 Then
'      m_YearDay = 365
'   Else
'      m_YearDay = 366
'   End If
'
'   strStarDate = CStr(Val(CStr(Val(textYV01_1) - 1) & "0101") + 19110000)
'   strEndDate = CStr(Val(CStr(Val(textYV01_1) - 1) & "1231") + 19110000)
'   strScSql = ""
'   '抓留職停薪員工
'   Set m_rs2 = New ADODB.Recordset
'   m_StrSQL2 = "select sc01 " & _
'                          "from staff_change " & _
'                          "where sc03='04' and sc02 between '" & strStarDate & "' and '" & strEndDate & "' " & _
'                          "and sc02 = (select max(a.sc02) from staff_change a where a.sc01=staff_change.sc01) " & _
'                          "order by sc01,sc02 asc "
'   If m_rs2.State = 1 Then m_rs2.Close
'   m_rs2.CursorLocation = adUseClient
'   m_rs2.Open m_StrSQL2, cnnConnection, adOpenStatic, adLockReadOnly
'   If m_rs2.RecordCount <> 0 Then
'       m_rs2.MoveFirst
'       Do While Not m_rs2.EOF
'       If strScSql = "" Then
'            strScSql = "'" & CheckStr(m_rs2.Fields("sc01")) & "'"
'       Else
'            strScSql = strScSql & ",'" & CheckStr(m_rs2.Fields("sc01")) & "'"
'       End If
'       m_rs2.MoveNext
'       Loop
'   End If
'
'   '開始計算
'   Set m_rs = New ADODB.Recordset
'   ' 2008/12/26 Modify BY SINDY
'   'm_StrSQL = "select * from staff where st04='1' and ascii(substr(st01,1,1))>=48 and ascii(substr(st01,1,1))<=57  " & _
'   '                    " and ((length(st01)=5 and substr(st01,1,1) not in ('0','1','2','3','4','5')) and st01 not in ('99998','99999')) " & _
'   '                    " order by st01  "
'   If strScSql = "" Then
'      m_StrSQL = "select * from staff,SalaryData " & _
'                           "where ST01=SD01 " & _
'                           "and ST04='1' and sd02 not in('P','F') " & _
'                           "order by st01 ASC "
'   Else
'      m_StrSQL = "select * from staff,SalaryData " & _
'                           "where ST01=SD01 " & _
'                           "and ((ST04='1' and sd02 not in('P','F')) or st01 in (" & strScSql & ")) " & _
'                           "order by st01 ASC "
'   End If
'   ' 2008/12/26 END
'   If m_rs.State = 1 Then m_rs.Close
'   m_rs.CursorLocation = adUseClient
'   m_rs.Open m_StrSQL, cnnConnection, adOpenStatic, adLockReadOnly
'   If m_rs.RecordCount <> 0 Then
'       PB1.Min = 0
'       PB1.max = m_rs.RecordCount
'       m_rs.MoveFirst
'       Do While Not m_rs.EOF
'           PB1.Value = m_rs.AbsolutePosition
'           If CheckStr(m_rs.Fields("st13")) = "" Then
'               ListInfo.AddItem "員工 " & CheckStr(m_rs.Fields("st01")) & " " & CheckStr(m_rs.Fields("st02")) & "-->計算失敗，原因  沒有到職日 ", m_item
'               ListInfo.Selected(m_item) = True
'               m_item = m_item + 1
'           Else
'               '計算年資
'               m_Year = Trim(CalYear(CheckStr(m_rs.Fields("st01")), strEndDate))
'               If Val(m_Year) < 0 Then m_Year = "0"
'               If m_Year = "" Then
'                   ListInfo.AddItem "員工 " & CheckStr(m_rs.Fields("st01")) & " " & CheckStr(m_rs.Fields("st02")) & "-->計算失敗，原因  統計錯誤 ，請聯絡電腦中心檢查  ", m_item
'                   ListInfo.Selected(m_item) = True
'                   m_item = m_item + 1
'               Else
'                   m_Days = 0
'                   If Val(m_Year) >= 10 Then '滿 10 年
'                       m_Days = 14 + (Int(m_Year) - 10)
'                       If m_Days > 30 Then
'                           m_Days = 30
'                       End If
'                   ElseIf Val(m_Year) >= 5 Then    '滿 5 年
'                       m_Days = 14
'                   ElseIf Val(m_Year) >= 3 Then ' 滿 3 年
'                       m_Days = 10
'                   ElseIf Val(m_Year) >= 1 Then  '滿1年
'                       m_Days = 7
'                   ElseIf Val(m_Year) >= 0.92 Then   '滿11月
'                       m_Days = 5.5
'                   ElseIf Val(m_Year) >= 0.83 Then   '滿10月
'                       m_Days = 5
'                   ElseIf Val(m_Year) >= 0.75 Then   '滿9月
'                       m_Days = 4.5
'                   ElseIf Val(m_Year) >= 0.67 Then   '滿8月
'                       m_Days = 4
'                   ElseIf Val(m_Year) >= 0.58 Then   '滿7月
'                       m_Days = 3.5
'                   ElseIf Val(m_Year) >= 0.5 Then   '滿6月
'                       m_Days = 3
'                   ElseIf Val(m_Year) >= 0.42 Then   '滿5月
'                       m_Days = 2.5
'                   ElseIf Val(m_Year) >= 0.33 Then   '滿4月
'                       m_Days = 2
'                   ElseIf Val(m_Year) >= 0.25 Then   '滿3月
'                       m_Days = 1.5
'                   ElseIf Val(m_Year) >= 0.17 Then   '滿2月
'                       m_Days = 1
'                   ElseIf Val(m_Year) >= 0.08 Then   '滿1月
'                       m_Days = 0.5
'                   End If
'                   If m_Days < 0 Then m_Days = 0
'
'                   '到職日為前一年者
'                   If CheckStr(m_rs.Fields("st13")) > Val((Val(textYV01_1) + 1911 - 1) & "0101") Then
'                        '2009/12/3 add by sonia 每年12/1到職者特別假給0.5天
'                        If CheckStr(m_rs.Fields("st13")) = Val((Val(textYV01_1) + 1911 - 1) & "1201") Then
'                           m_Days = 0.5
'                        '每年12/1以後到職者無特別假
'                        ElseIf CheckStr(m_rs.Fields("st13")) > Val((Val(textYV01_1) + 1911 - 1) & "1201") Then
'                           m_Days = 0
''                        Else
''                        '2009/12/3 end
''2009/12/15 CANCEL BY SONIA 上面已處理不必再做
''                           '取得計算年度之前年工作總時數
''                           m_StrSQL2 = "select sum(nvl(sm27,0))+31 " & _
''                                       "from salarymonth " & _
''                                       "where sm01='" & CheckStr(m_rs.Fields("st01")) & "' " & _
''                                       "and sm02>='" & CStr(Val(textYV01_1) + 1911 - 1) & "01" & "' " & _
''                                       "and sm02<='" & CStr(Val(textYV01_1) + 1911 - 1) & "12" & "' "
''                           If m_rs2.State = 1 Then m_rs2.Close
''                           m_rs2.CursorLocation = adUseClient
''                           m_rs2.Open m_StrSQL2, cnnConnection, adOpenStatic, adLockReadOnly
''                           If m_rs2.RecordCount <> 0 Then
''                               m_rs2.MoveFirst
''                               If m_YearDay <> CLng(CheckStr(m_rs2.Fields(0))) Then
''                                    '因前年未工作滿一年, 所以特別假要按照工作總時數計算
''                                    'm_Days = Round(m_Days * CLng(CheckStr(m_rs2.Fields(0))) / m_YearDay, 1)
''                                    m_Days = Round(7 * CLng(CheckStr(m_rs2.Fields(0))) / m_YearDay, 0)
''                               End If
''                           End If
''2009/12/15 end
'                        End If
'                   End If
'
'                   cnnConnection.Execute " insert into YearVacation (yv01,yv02,yv03,yv04) values ('" & Mid(Trim(DBDATE(textYV01_1 & "0101")), 1, 4) & "','" & CheckStr(m_rs.Fields("st01")) & "','" & m_Year & "','" & m_Days & "' ) "
'                   ListInfo.AddItem "員工 " & CheckStr(m_rs.Fields("st01")) & " " & CheckStr(m_rs.Fields("st02")) & "-->計算成功          入所日：" & ChangeWStringToTDateString(CheckStr(m_rs.Fields("st13"))) & "    年資：" & m_Year & "；年假：" & m_Days, m_item
'                   ListInfo.Selected(m_item) = True
'                   m_item = m_item + 1
'               End If
'           End If
'           DoEvents
'           m_rs.MoveNext
'       Loop
'   End If
'   ListInfo.AddItem "統計結束;　共 " & m_rs.RecordCount & " 筆", m_item
'   ListInfo.Selected(m_item) = True
'   m_item = m_item + 1
'   cnnConnection.CommitTrans
'   Me.Enabled = True
'   Screen.MousePointer = vbDefault
'   Exit Sub
'err1:
'    cnnConnection.RollbackTrans
'    ListInfo.AddItem "發生錯誤！所有剛剛統計的資料將會還原！", m_item
'    m_item = m_item + 1
'    ListInfo.AddItem "錯誤碼：" & Trim(Err.Number) & "   " & Err.Description, m_item
'    m_item = m_item + 1
'    ListInfo.AddItem "請通知電腦中心處理！謝謝！", m_item
'    m_item = m_item + 1
'    Me.Enabled = True
'    Screen.MousePointer = vbDefault
'End Sub

'更新特別假
Sub CalUpdVacation()
Dim m_rs As New ADODB.Recordset
Dim m_StrSQL As String
Dim m_item As Long
Dim strStarDate As String, strEndDate As String
   
   Screen.MousePointer = vbHourglass
   
   Me.Enabled = False
   ListInfo2.Clear
   
   PB2.Value = 0
   m_item = 0
   ListInfo2.AddItem "開始更新   " & textYV01_1.Text & " 年度特別假", m_item
   ListInfo2.Selected(m_item) = True
   m_item = m_item + 1
   
   strStarDate = CStr(Val(CStr(Val(textYV01_2) - 1) & "0101") + 19110000)
   strEndDate = CStr(Val(CStr(Val(textYV01_2) - 1) & "1231") + 19110000)
   
   On Error GoTo err1
   cnnConnection.BeginTrans
   
   '開始更新
   Set m_rs = New ADODB.Recordset
   'Modify By Sindy 2022/1/3 在職的才更新,但含留職停薪
   m_StrSQL = "select * from staff,YearVacation where st01=yv02(+) and yv01='" & Mid(Trim(DBDATE(textYV01_2 & "0101")), 1, 4) & "'" & _
              " and (st04='1'" & _
              " or st01 in(select sc01 From staff_change where sc03='04' and sc02 between '" & strStarDate & "' and '" & strEndDate & "'" & _
                          " and sc02 = (select max(a.sc02) from staff_change a where a.sc01=staff_change.sc01))" & _
              ")" & _
              " Order BY st01 ASC"
   If m_rs.State = 1 Then m_rs.Close
   m_rs.CursorLocation = adUseClient
   m_rs.Open m_StrSQL, cnnConnection, adOpenStatic, adLockReadOnly
   If m_rs.RecordCount <> 0 Then
       PB2.Min = 0
       PB2.max = m_rs.RecordCount
       m_rs.MoveFirst
       
       Do While Not m_rs.EOF
           PB2.Value = m_rs.AbsolutePosition
           
           strSql = " st40=" & CheckStr(m_rs.Fields("yv04"))
'           If Left(Trim(strSql), 1) = "," Then
'               strSql = Mid(Trim(strSql), 2)
'           End If
           cnnConnection.Execute "update staff set " & strSql & " where st01='" & CheckStr(m_rs.Fields("st01")) & "' "
           'Add By Sindy 2018/5/16 增加yv11.更新員工檔日期
           cnnConnection.Execute "update YearVacation set yv11=" & strSrvDate(1) & " where yv01=" & m_rs.Fields("yv01") & " and yv02='" & CheckStr(m_rs.Fields("st01")) & "'"
           '2018/5/16 END
           ListInfo2.AddItem "員工 " & CheckStr(m_rs.Fields("st01")) & " " & CheckStr(m_rs.Fields("st02")) & "-->更新成功　　原年假：" & CheckStr(m_rs.Fields("st40")) & "　新年假：" & CheckStr(m_rs.Fields("yv04")), m_item
           ListInfo2.Selected(m_item) = True
           m_item = m_item + 1
           m_rs.MoveNext
       Loop
   End If
   ListInfo2.AddItem "更新結束;　共 " & m_rs.RecordCount & " 筆", m_item
   ListInfo2.Selected(m_item) = True
   m_item = m_item + 1
   cnnConnection.CommitTrans
   Me.Enabled = True
   Screen.MousePointer = vbDefault
   Exit Sub
err1:
    cnnConnection.RollbackTrans
    ListInfo2.AddItem "發生錯誤！所有剛剛更新的資料將會還原！", m_item
    m_item = m_item + 1
    ListInfo2.AddItem "錯誤碼：" & Trim(Err.Number) & "   " & Err.Description, m_item
    m_item = m_item + 1
    ListInfo2.AddItem "請通知電腦中心處理！謝謝！", m_item
    m_item = m_item + 1
    Me.Enabled = True
    Screen.MousePointer = vbDefault
End Sub
