VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm160009 
   BorderStyle     =   1  '單線固定
   Caption         =   "端午、中秋獎金維護"
   ClientHeight    =   5450
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
   ScaleHeight     =   5450
   ScaleWidth      =   8710
   Begin TabDlg.SSTab SSTab2 
      Height          =   5415
      Left            =   30
      TabIndex        =   24
      Top             =   0
      Width           =   8685
      _ExtentX        =   15311
      _ExtentY        =   9543
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "試算作業"
      TabPicture(0)   =   "frm160009.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label7"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label8"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label18"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label19"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label20"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label21"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "textBD02"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "textBD03"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "textBD04"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "textBD05"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "textBD06"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "textBD07"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "textBD08"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "textBD09"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "textBD10"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "textBD11"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "textBD12"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "cmd(0)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Frame1"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "cmd(1)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "textBD01"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).ControlCount=   27
      TabCaption(1)   =   "修改作業"
      TabPicture(1)   =   "frm160009.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "TBar1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "ImageList1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "SSTab1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.TextBox textBD01 
         Height          =   285
         Left            =   1110
         MaxLength       =   5
         TabIndex        =   0
         Top             =   450
         Width           =   705
      End
      Begin VB.CommandButton cmd 
         Caption         =   "結束(&X)"
         Height          =   405
         Index           =   1
         Left            =   5940
         TabIndex        =   13
         Top             =   450
         Width           =   1245
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   4365
         Left            =   -74940
         TabIndex        =   36
         Top             =   960
         Width           =   8535
         _ExtentX        =   15064
         _ExtentY        =   7691
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         Tab             =   1
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "單筆資料"
         TabPicture(0)   =   "frm160009.frx":0038
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "textOB04_2"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "textOB01"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "textOB02"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "textOB03"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "textOB04"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "textOB05"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Grd2"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "Label23"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "textOB03_2"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "lblST13"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "Label15"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "Label9"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "Label10"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "Label11"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "Label12"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "Label13"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "Label14"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).ControlCount=   17
         TabCaption(1)   =   "多筆瀏覽"
         TabPicture(1)   =   "frm160009.frx":0054
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Label16"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Label17"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Line4"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "Line5"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "GRD1"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "cmdok"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "txt1(3)"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "txt1(2)"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).Control(8)=   "txt1(1)"
         Tab(1).Control(8).Enabled=   0   'False
         Tab(1).Control(9)=   "txt1(0)"
         Tab(1).Control(9).Enabled=   0   'False
         Tab(1).ControlCount=   10
         Begin VB.TextBox textOB04_2 
            Appearance      =   0  '平面
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  '沒有框線
            Height          =   195
            Left            =   -72990
            Locked          =   -1  'True
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   1560
            Width           =   2145
         End
         Begin VB.TextBox textOB01 
            Height          =   270
            Left            =   -73950
            MaxLength       =   5
            TabIndex        =   14
            Top             =   630
            Width           =   675
         End
         Begin VB.TextBox txt1 
            Height          =   270
            Index           =   0
            Left            =   1140
            MaxLength       =   6
            TabIndex        =   19
            Top             =   375
            Width           =   915
         End
         Begin VB.TextBox txt1 
            Height          =   270
            Index           =   1
            Left            =   2190
            MaxLength       =   6
            TabIndex        =   20
            Top             =   375
            Width           =   915
         End
         Begin VB.TextBox txt1 
            Height          =   270
            Index           =   2
            Left            =   4110
            MaxLength       =   5
            TabIndex        =   21
            Top             =   375
            Width           =   915
         End
         Begin VB.TextBox txt1 
            Height          =   270
            Index           =   3
            Left            =   5070
            MaxLength       =   5
            TabIndex        =   22
            Top             =   375
            Width           =   915
         End
         Begin VB.CommandButton cmdok 
            Caption         =   "查詢"
            Height          =   345
            Left            =   6450
            TabIndex        =   23
            Top             =   360
            Width           =   915
         End
         Begin VB.TextBox textOB02 
            Height          =   270
            Left            =   -73950
            TabIndex        =   15
            Top             =   900
            Width           =   405
         End
         Begin VB.TextBox textOB03 
            Height          =   270
            Left            =   -73950
            MaxLength       =   6
            TabIndex        =   16
            Top             =   1200
            Width           =   1005
         End
         Begin VB.TextBox textOB04 
            Height          =   270
            Left            =   -73950
            MaxLength       =   5
            TabIndex        =   17
            Top             =   1500
            Width           =   915
         End
         Begin VB.TextBox textOB05 
            Height          =   285
            Left            =   -73950
            TabIndex        =   18
            Top             =   1800
            Width           =   1155
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grd2 
            Height          =   3345
            Left            =   -70170
            TabIndex        =   37
            Top             =   600
            Width           =   3555
            _ExtentX        =   6279
            _ExtentY        =   5891
            _Version        =   393216
            Cols            =   1
            FixedCols       =   0
            AllowBigSelection=   0   'False
            HighLight       =   0
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
            Bindings        =   "frm160009.frx":0070
            Height          =   3585
            Left            =   90
            TabIndex        =   45
            Top             =   720
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
         Begin MSForms.Label Label23 
            Height          =   195
            Left            =   -74730
            TabIndex        =   56
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
         Begin MSForms.Label textOB03_2 
            Height          =   225
            Left            =   -72900
            TabIndex        =   55
            Top             =   1230
            Width           =   1395
            BackColor       =   12632256
            VariousPropertyBits=   27
            Size            =   "2461;397"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label lblST13 
            Height          =   285
            Left            =   -73950
            TabIndex        =   49
            Top             =   2130
            Width           =   2175
         End
         Begin VB.Line Line5 
            X1              =   4800
            X2              =   5400
            Y1              =   503
            Y2              =   503
         End
         Begin VB.Line Line4 
            X1              =   1950
            X2              =   2640
            Y1              =   503
            Y2              =   503
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "員工編號："
            Height          =   180
            Left            =   210
            TabIndex        =   47
            Top             =   420
            Width           =   900
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "獎金年月："
            Height          =   180
            Left            =   3240
            TabIndex        =   46
            Top             =   420
            Width           =   900
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "任職時間："
            Height          =   180
            Left            =   -70140
            TabIndex        =   44
            Top             =   420
            Width           =   900
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "獎金年月：                (9701)"
            Height          =   180
            Left            =   -74880
            TabIndex        =   43
            Top             =   630
            Width           =   2100
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "獎金類別：          (1:端午   2:中秋)"
            Height          =   180
            Left            =   -74880
            TabIndex        =   42
            Top             =   930
            Width           =   2595
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "員工編號："
            Height          =   180
            Left            =   -74880
            TabIndex        =   41
            Top             =   1230
            Width           =   900
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "年　　資："
            Height          =   180
            Left            =   -74880
            TabIndex        =   40
            Top             =   1530
            Width           =   900
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "獎金金額："
            Height          =   180
            Left            =   -74880
            TabIndex        =   39
            Top             =   1860
            Width           =   900
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "入所日期："
            Height          =   180
            Left            =   -74880
            TabIndex        =   38
            Top             =   2160
            Width           =   900
         End
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   -67050
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
               Picture         =   "frm160009.frx":0085
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm160009.frx":03A1
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm160009.frx":06BD
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm160009.frx":0899
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm160009.frx":0BB5
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm160009.frx":0ED1
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm160009.frx":11ED
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm160009.frx":1509
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm160009.frx":1825
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm160009.frx":1B41
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm160009.frx":1E5D
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame1 
         Caption         =   "目前進度資料"
         Height          =   2805
         Left            =   120
         TabIndex        =   33
         Top             =   2520
         Width           =   8445
         Begin MSComctlLib.ProgressBar PB1 
            Height          =   225
            Left            =   30
            TabIndex        =   34
            Top             =   2550
            Width           =   8385
            _ExtentX        =   14781
            _ExtentY        =   406
            _Version        =   393216
            Appearance      =   0
            Scrolling       =   1
         End
         Begin VB.ListBox ListInfo 
            Height          =   2380
            Left            =   60
            TabIndex        =   48
            Top             =   180
            Width           =   8325
         End
      End
      Begin VB.CommandButton cmd 
         Caption         =   "開始試算(&S)"
         Height          =   405
         Index           =   0
         Left            =   4650
         TabIndex        =   12
         Top             =   450
         Width           =   1245
      End
      Begin VB.TextBox textBD12 
         Height          =   270
         Left            =   2580
         TabIndex        =   11
         Top             =   2190
         Width           =   1005
      End
      Begin VB.TextBox textBD11 
         Height          =   270
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   10
         Top             =   2190
         Width           =   375
      End
      Begin VB.TextBox textBD10 
         Height          =   270
         Left            =   2580
         TabIndex        =   9
         Top             =   1920
         Width           =   1005
      End
      Begin VB.TextBox textBD09 
         Height          =   270
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   8
         Top             =   1920
         Width           =   375
      End
      Begin VB.TextBox textBD08 
         Height          =   270
         Left            =   2580
         TabIndex        =   7
         Top             =   1650
         Width           =   1005
      End
      Begin VB.TextBox textBD07 
         Height          =   270
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   6
         Top             =   1650
         Width           =   375
      End
      Begin VB.TextBox textBD06 
         Height          =   270
         Left            =   2580
         TabIndex        =   5
         Top             =   1380
         Width           =   1005
      End
      Begin VB.TextBox textBD05 
         Height          =   270
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   4
         Top             =   1380
         Width           =   375
      End
      Begin VB.TextBox textBD04 
         Height          =   270
         Left            =   2580
         TabIndex        =   3
         Top             =   1110
         Width           =   1005
      End
      Begin VB.TextBox textBD03 
         Height          =   270
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   2
         Top             =   1110
         Width           =   375
      End
      Begin VB.TextBox textBD02 
         Height          =   285
         Left            =   1110
         MaxLength       =   1
         TabIndex        =   1
         Top             =   750
         Width           =   405
      End
      Begin MSComctlLib.Toolbar TBar1 
         Height          =   520
         Left            =   -74970
         TabIndex        =   35
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
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "修改資料將被還原"
         ForeColor       =   &H000000C0&
         Height          =   180
         Left            =   4590
         TabIndex        =   54
         Top             =   2220
         Width           =   1440
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "2. 若資料已做修改，不可再重新執行試算，否則"
         ForeColor       =   &H000000C0&
         Height          =   180
         Left            =   4410
         TabIndex        =   53
         Top             =   2010
         Width           =   3780
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "1. 試算資料為當時在職人員"
         ForeColor       =   &H000000C0&
         Height          =   180
         Left            =   4410
         TabIndex        =   52
         Top             =   1680
         Width           =   2160
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "注意："
         ForeColor       =   &H000000C0&
         Height          =   180
         Left            =   4410
         TabIndex        =   51
         Top             =   1320
         Width           =   540
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "年資            年以上                        元"
         Height          =   180
         Left            =   1110
         TabIndex        =   32
         Top             =   2235
         Width           =   2700
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "年資            年以上                        元"
         Height          =   180
         Left            =   1110
         TabIndex        =   31
         Top             =   1965
         Width           =   2700
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "年資            年以上                        元"
         Height          =   180
         Left            =   1110
         TabIndex        =   30
         Top             =   1695
         Width           =   2700
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "年資            年以上                        元"
         Height          =   180
         Left            =   1110
         TabIndex        =   29
         Top             =   1425
         Width           =   2700
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "年資            年以上                        元"
         Height          =   180
         Left            =   1110
         TabIndex        =   28
         Top             =   1155
         Width           =   2700
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "發放金額："
         Height          =   180
         Left            =   210
         TabIndex        =   27
         Top             =   1140
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "獎金類別：          (1:端午   2:中秋)"
         Height          =   180
         Left            =   210
         TabIndex        =   26
         Top             =   810
         Width           =   2595
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "獎金年月：                (9701)"
         Height          =   180
         Left            =   210
         TabIndex        =   25
         Top             =   480
         Width           =   2100
      End
   End
End
Attribute VB_Name = "frm160009"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/6/16 Form2.0已修改
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Modify By Sindy 2010/7/20 日期欄已修改
'Create by nickc 2006/12/25 copy from frm140401
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
Dim m_FirstKEY(3) As String
' 最後一筆資料的本所案號
Dim m_LastKEY(3) As String
' 目前正在顯示的本所案號
Dim m_CurrKEY(3) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim tf_OB As Integer
Dim MyKind As String


Private Sub Form_Activate()
   SSTab2.Tab = 0
   textBD01.SetFocus
End Sub

Private Sub cmd_Click(Index As Integer)
Dim m_rs As New ADODB.Recordset
Dim m_StrSQL As String, nResponse As Variant

Select Case Index
Case 0
        '檢查條件
        If textBD01 = "" Then
            MsgBox "請輸入獎金年月！", vbExclamation, "操作錯誤！"
            textBD01_GotFocus
            Exit Sub
        End If
        If Trim(textBD02) = "" Then
            MsgBox "請輸入獎金類別！", vbExclamation, "操作錯誤！"
            textBD02_GotFocus
            Exit Sub
        End If
        If Trim(textBD03) & Trim(textBD05) & Trim(textBD07) & Trim(textBD09) & Trim(textBD11) = "" Then
            MsgBox "請輸入年資分級！", vbExclamation, "操作錯誤！"
            textBD03_GotFocus
            Exit Sub
        End If
        If Trim(textBD04) & Trim(textBD06) & Trim(textBD08) & Trim(textBD10) & Trim(textBD12) = "" Then
            MsgBox "請輸入獎金分級！", vbExclamation, "操作錯誤！"
            textBD04_GotFocus
            Exit Sub
        End If
        If Trim(textBD03) <> "" And Trim(textBD04) = "" Then
            MsgBox "請輸入對應的獎金！", vbExclamation, "操作錯誤！"
            textBD04_GotFocus
            Exit Sub
        End If
        If Trim(textBD05) <> "" And Trim(textBD06) = "" Then
            MsgBox "請輸入對應的獎金！", vbExclamation, "操作錯誤！"
            textBD06_GotFocus
            Exit Sub
        End If
        If Trim(textBD07) <> "" And Trim(textBD08) = "" Then
            MsgBox "請輸入對應的獎金！", vbExclamation, "操作錯誤！"
            textBD08_GotFocus
            Exit Sub
        End If
        If Trim(textBD09) <> "" And Trim(textBD10) = "" Then
            MsgBox "請輸入對應的獎金！", vbExclamation, "操作錯誤！"
            textBD10_GotFocus
            Exit Sub
        End If
        If Trim(textBD11) <> "" And Trim(textBD12) = "" Then
            MsgBox "請輸入對應的獎金！", vbExclamation, "操作錯誤！"
            textBD12_GotFocus
            Exit Sub
        End If
        If Trim(textBD03) = "" And Trim(textBD04) <> "" Then
            MsgBox "請輸入對應的年資！", vbExclamation, "操作錯誤！"
            textBD03_GotFocus
            Exit Sub
        End If
        If Trim(textBD05) = "" And Trim(textBD06) <> "" Then
            MsgBox "請輸入對應的年資！", vbExclamation, "操作錯誤！"
            textBD05_GotFocus
            Exit Sub
        End If
        If Trim(textBD07) = "" And Trim(textBD08) <> "" Then
            MsgBox "請輸入對應的年資！", vbExclamation, "操作錯誤！"
            textBD07_GotFocus
            Exit Sub
        End If
        If Trim(textBD09) = "" And Trim(textBD10) <> "" Then
            MsgBox "請輸入對應的年資！", vbExclamation, "操作錯誤！"
            textBD09_GotFocus
            Exit Sub
        End If
        If Trim(textBD11) = "" And Trim(textBD12) <> "" Then
            MsgBox "請輸入對應的年資！", vbExclamation, "操作錯誤！"
            textBD11_GotFocus
            Exit Sub
        End If
        If Trim(textBD11) <> "" Then
            If Trim(textBD09) <> "" Then
                If Val(textBD11) <= Val(textBD09) Then
                    MsgBox "等級 5 的年資應該比等級 4 的年資大！", vbExclamation, "操作錯誤！"
                    textBD11_GotFocus
                    Exit Sub
                End If
            End If
            If Trim(textBD07) <> "" Then
                If Val(textBD11) <= Val(textBD07) Then
                    MsgBox "等級 5 的年資應該比等級 3 的年資大！", vbExclamation, "操作錯誤！"
                    textBD11_GotFocus
                    Exit Sub
                End If
            End If
            If Trim(textBD05) <> "" Then
                If Val(textBD11) <= Val(textBD05) Then
                    MsgBox "等級 5 的年資應該比等級 2 的年資大！", vbExclamation, "操作錯誤！"
                    textBD11_GotFocus
                    Exit Sub
                End If
            End If
            If Trim(textBD03) <> "" Then
                If Val(textBD11) <= Val(textBD03) Then
                    MsgBox "等級 5 的年資應該比等級 1 的年資大！", vbExclamation, "操作錯誤！"
                    textBD11_GotFocus
                    Exit Sub
                End If
            End If
        End If
        If Trim(textBD09) <> "" Then
            If Trim(textBD07) <> "" Then
                If Val(textBD09) <= Val(textBD07) Then
                    MsgBox "等級 4 的年資應該比等級 3 的年資大！", vbExclamation, "操作錯誤！"
                    textBD09_GotFocus
                    Exit Sub
                End If
            End If
            If Trim(textBD05) <> "" Then
                If Val(textBD09) <= Val(textBD05) Then
                    MsgBox "等級 4 的年資應該比等級 2 的年資大！", vbExclamation, "操作錯誤！"
                    textBD09_GotFocus
                    Exit Sub
                End If
            End If
            If Trim(textBD03) <> "" Then
                If Val(textBD09) <= Val(textBD03) Then
                    MsgBox "等級 4 的年資應該比等級 1 的年資大！", vbExclamation, "操作錯誤！"
                    textBD09_GotFocus
                    Exit Sub
                End If
            End If
        End If
        If Trim(textBD07) <> "" Then
            If Trim(textBD05) <> "" Then
                If Val(textBD07) <= Val(textBD05) Then
                    MsgBox "等級 3 的年資應該比等級 2 的年資大！", vbExclamation, "操作錯誤！"
                    textBD07_GotFocus
                    Exit Sub
                End If
            End If
            If Trim(textBD03) <> "" Then
                If Val(textBD07) <= Val(textBD03) Then
                    MsgBox "等級 3 的年資應該比等級 1 的年資大！", vbExclamation, "操作錯誤！"
                    textBD07_GotFocus
                    Exit Sub
                End If
            End If
        End If
        If Trim(textBD05) <> "" Then
            If Trim(textBD03) <> "" Then
                If Val(textBD05) <= Val(textBD03) Then
                    MsgBox "等級 2 的年資應該比等級 1 的年資大！", vbExclamation, "操作錯誤！"
                    textBD05_GotFocus
                    Exit Sub
                End If
            End If
        End If
        If Trim(textBD12) <> "" Then
            If Trim(textBD10) <> "" Then
                If Val(textBD12) <= Val(textBD10) Then
                    MsgBox "等級 5 的獎金應該比等級 4 的年資大！", vbExclamation, "操作錯誤！"
                    textBD12_GotFocus
                    Exit Sub
                End If
            End If
            If Trim(textBD08) <> "" Then
                If Val(textBD12) <= Val(textBD08) Then
                    MsgBox "等級 5 的獎金應該比等級 3 的年資大！", vbExclamation, "操作錯誤！"
                    textBD12_GotFocus
                    Exit Sub
                End If
            End If
            If Trim(textBD06) <> "" Then
                If Val(textBD12) <= Val(textBD06) Then
                    MsgBox "等級 5 的獎金應該比等級 2 的年資大！", vbExclamation, "操作錯誤！"
                    textBD12_GotFocus
                    Exit Sub
                End If
            End If
            If Trim(textBD04) <> "" Then
                If Val(textBD12) <= Val(textBD04) Then
                    MsgBox "等級 5 的獎金應該比等級 1 的年資大！", vbExclamation, "操作錯誤！"
                    textBD12_GotFocus
                    Exit Sub
                End If
            End If
        End If
        If Trim(textBD10) <> "" Then
            If Trim(textBD08) <> "" Then
                If Val(textBD10) <= Val(textBD08) Then
                    MsgBox "等級 4 的獎金應該比等級 3 的年資大！", vbExclamation, "操作錯誤！"
                    textBD10_GotFocus
                    Exit Sub
                End If
            End If
            If Trim(textBD06) <> "" Then
                If Val(textBD10) <= Val(textBD06) Then
                    MsgBox "等級 4 的獎金應該比等級 2 的年資大！", vbExclamation, "操作錯誤！"
                    textBD10_GotFocus
                    Exit Sub
                End If
            End If
            If Trim(textBD04) <> "" Then
                If Val(textBD10) <= Val(textBD04) Then
                    MsgBox "等級 4 的獎金應該比等級 1 的年資大！", vbExclamation, "操作錯誤！"
                    textBD10_GotFocus
                    Exit Sub
                End If
            End If
        End If
        If Trim(textBD08) <> "" Then
            If Trim(textBD06) <> "" Then
                If Val(textBD08) <= Val(textBD06) Then
                    MsgBox "等級 3 的獎金應該比等級 2 的年資大！", vbExclamation, "操作錯誤！"
                    textBD08_GotFocus
                    Exit Sub
                End If
            End If
            If Trim(textBD04) <> "" Then
                If Val(textBD08) <= Val(textBD04) Then
                    MsgBox "等級 3 的獎金應該比等級 1 的年資大！", vbExclamation, "操作錯誤！"
                    textBD08_GotFocus
                    Exit Sub
                End If
            End If
        End If
        If Trim(textBD06) <> "" Then
            If Trim(textBD04) <> "" Then
                If Val(textBD06) <= Val(textBD04) Then
                    MsgBox "等級 2 的獎金應該比等級 1 的年資大！", vbExclamation, "操作錯誤！"
                    textBD06_GotFocus
                    Exit Sub
                End If
            End If
        End If
        '檢查是否已有資料
        Set m_rs = New ADODB.Recordset
        'm_StrSQL = "select * from OhBonus where ob01='" & Mid(Trim(DBDATE(textBD01 & "01")), 1, 6) & "' and ob02='" & textBD02 & "' "
        m_StrSQL = "select * from OhBonus where substr(ob01,1,4)='" & Mid(Trim(DBDATE(textBD01 & "01")), 1, 4) & "' and ob02='" & textBD02 & "' "
        If m_rs.State = 1 Then m_rs.Close
        m_rs.CursorLocation = adUseClient
        m_rs.Open m_StrSQL, cnnConnection, adOpenStatic, adLockReadOnly
        If m_rs.RecordCount <> 0 Then
            If Val(m_rs.Fields(0)) < Val(Left(strSrvDate(1), 6)) Then
               nResponse = MsgBox("已過當月，不可重新產生資料!!", vbOKOnly, "試算")
               Exit Sub
            End If
            If MsgBox(Val(m_rs.Fields(0)) - 191100 & "已經產生過資料，是否要全部重新產生??", vbExclamation + vbYesNo, "嚴重警告！") = vbNo Then
                Exit Sub
            End If
        End If
        '開始計算
        CalBonus
        InitialField
        InitialData
        RefreshRange
        ShowFirstRecord
        UpdateToolbarState
        SetCtrlReadOnly True
        Me.SSTab1.Tab = 0
Case 1
        Unload Me
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
rsA.Open "select * from OhBonus where rownum <2 ", cnnConnection, adOpenStatic, adLockReadOnly
tf_OB = rsA.Fields.Count
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

   ReDim m_FieldList(tf_OB) As FIELDITEM
   
   m_bInsert = IsUserHasRightOfFunction(Me.Name, strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
   m_bDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   m_bQuery = IsUserHasRightOfFunction(Me.Name, strFind, False)
   
   textOB01.BackColor = &H8000000F
   textOB02.BackColor = &H8000000F
   textOB03.BackColor = &H8000000F
   
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
   Set frm160009 = Nothing
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
Dim strDate As String

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
         strDate = ChangeTDateStringToTString(GRD1.TextMatrix(tmpMouseRow, 0) & "/01")
         textOB01.Text = Left(strDate, Len(strDate) - 2)
         textOB02.Text = Left(GRD1.TextMatrix(tmpMouseRow, 1), 1)
         textOB03.Text = GRD1.TextMatrix(tmpMouseRow, 2)
         QueryRecord
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
   If IsNull(rsSrcTmp.Fields("ob06")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("ob06")) = False Then
         strCName = GetStaffName(rsSrcTmp.Fields("ob06"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("ob07")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("ob07")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("ob07"))
         strCDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("ob08")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("ob08")) = False Then
         strTemp = rsSrcTmp.Fields("ob08")
         strCTime = Format(strTemp, "##:##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("ob09")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("ob09")) = False Then
         strUName = GetStaffName(rsSrcTmp.Fields("ob09"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("ob10")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("ob10")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("ob10"))
         strUDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("ob11")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("ob11")) = False Then
         strTemp = rsSrcTmp.Fields("ob11")
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
If Me.textOB03.Enabled = True Then
   Cancel = False
   textOB03_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If textOB03.Text = "" Then
    MsgBox "員工編號不可以空白！", vbExclamation
    textOB03.SetFocus
    Exit Function
End If
If Me.textOB01.Enabled = True Then
   Cancel = False
   textOB01_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If textOB01.Text = "" Then
    MsgBox "獎金年月不可以空白！", vbExclamation
    textOB01.SetFocus
    Exit Function
End If
If Me.textOB02.Enabled = True Then
   Cancel = False
   textOB02_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

TxtValidate = True
End Function

' 設定欄位的內容
Private Sub SetFieldNewData(ByVal strName As String, Optional ByVal strData As String = "#==#")
   Dim nIndex As Integer
   For nIndex = 0 To tf_OB - 1
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
   
   For nIndex = 0 To tf_OB - 1
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
   Dim strOB01 As String
   Dim strOB02 As String
   Dim strOB03 As String
   'add by sonia 2016/1/18
   Dim rsTmp As New ADODB.Recordset
   Dim strSD19 As String
   'end 2016/1/18
   
   AddRecord = False
   
   strOB02 = textOB02
   strOB01 = Mid(DBDATE(textOB01.Text & "01"), 1, 6)
   strOB03 = textOB03

   ' 檢查記錄是否已存在
   If IsRecordExist(strOB01, strOB02, strOB03) = True Then
      strTit = "新增資料"
      strMsg = "該筆記錄已存在"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      UpdateCtrlData
      Exit Function
   End If
   
   'add by sonia 2016/1/18
   m_FieldList(12).fiNewData = ""
   strSql = "select sd19 from SalaryData where SD01='" & strOB03 & " ') "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      m_FieldList(12).fiNewData = "" & rsTmp.Fields("SD19")
   End If
   rsTmp.Close
   'end 2016/1/18
   
   bFirst = True
   bDifference = False
   strSql = "INSERT INTO OhBonus ("
   For nIndex = 0 To tf_OB - 1
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
   For nIndex = 0 To tf_OB - 1
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
   
   If ((strOB01 & strOB02 & strOB03) < (m_FirstKEY(0) & m_FirstKEY(1) & m_FirstKEY(2))) Or ((strOB01 & strOB02 & strOB03) > (m_LastKEY(0) & m_LastKEY(1) & m_LastKEY(2))) Then
      RefreshRange
   End If
   cnnConnection.CommitTrans
   
   ShowCurrRecord strOB01, strOB02, strOB03
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
   Dim strOB01 As String
   Dim strOB02 As String
   Dim strOB03 As String
       
   ModRecord = False
   
   strOB01 = m_CurrKEY(0)
   strOB02 = m_CurrKEY(1)
   strOB03 = m_CurrKEY(2)
   
   strSql = "begin user_data.user_enabled:=1; UPDATE OhBonus SET "

   bFirst = True
   bDifference = False
   For nIndex = 0 To tf_OB - 1
      strTmp = Empty
      If nIndex < 5 Or nIndex > 10 Then
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

   strSql = strSql & " " & _
                  "WHERE ob01 = '" & strOB01 & "' and ob02='" & strOB02 & "' and ob03='" & strOB03 & "' ; end; "
On Error GoTo ErrHand
      cnnConnection.BeginTrans
        If bDifference = True Then
           Pub_SeekTbLog strSql
           cnnConnection.Execute strSql
        End If
        cnnConnection.CommitTrans

      ShowCurrRecord strOB01, strOB02, strOB03
      
    ModRecord = True
   Exit Function
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox (Err.Description)

End Function

' 刪除記錄
Private Function DelRecord() As Boolean
   Dim strSql As String
   Dim strOB01 As String
   Dim strOB02 As String
   Dim strOB03 As String
   
   DelRecord = False
   
On Error GoTo ErrHand
   
   cnnConnection.BeginTrans
   
   strOB01 = m_CurrKEY(0)
   strOB02 = m_CurrKEY(1)
   strOB03 = m_CurrKEY(2)

   strSql = "DELETE FROM OhBonus " & _
            "WHERE ob01 = '" & strOB01 & "'  and ob02='" & strOB02 & "' and ob03='" & strOB03 & "' "

   cnnConnection.Execute strSql

   If (strOB01 = m_LastKEY(0) And strOB02 = m_LastKEY(1) And strOB03 = m_LastKEY(2)) Or (strOB01 = m_FirstKEY(0) And strOB02 = m_FirstKEY(1) And strOB03 = m_FirstKEY(2)) Then
      RefreshRange
   End If
   ShowCurrRecord strOB01, strOB02, strOB03
   DelRecord = True
   cnnConnection.CommitTrans
   
   Exit Function
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox "刪除失敗！" & vbCrLf & Err.Description
End Function

' 查詢記錄
Private Function QueryRecord() As Boolean
   Dim strOB01 As String
   Dim strOB02 As String
   Dim strOB03 As String
   
   QueryRecord = False
   strOB01 = Mid(DBDATE(textOB01.Text & "01"), 1, 6)
   strOB02 = textOB02
   strOB03 = textOB03
   If IsRecordExist(strOB01, strOB02, strOB03) = True Then
      m_CurrKEY(0) = strOB01
      m_CurrKEY(1) = strOB02
      m_CurrKEY(2) = strOB03
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
            ShowCurrRecord m_CurrKEY(0), m_CurrKEY(1), m_CurrKEY(2)
         Else
            Exit Function
         End If
      Case 4: '查詢
         If textOB01 <> "" And textOB02 <> "" And textOB03 <> "" Then
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
      'Case 1: If Me.Visible = True Then textOB01.SetFocus
      Case 2: If Me.Visible = True Then textOB04.SetFocus
      Case 4: If Me.Visible = True Then textOB01.SetFocus
   End Select
End Sub

' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal strKEY01 As String, ByVal strKEY02 As String, ByVal strKEY03 As String) As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   IsRecordExist = False
   strSql = "SELECT * FROM OhBonus " & _
            "WHERE ob01 = '" & strKEY01 & "'  and ob02='" & strKEY02 & "' and ob03='" & strKEY03 & "'  "
                  
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
      strSql = "SELECT ob01,ob02,ob03 FROM OhBonus " & _
               "WHERE ob01 = '" & m_CurrKEY(0) & "' and ob02='" & m_CurrKEY(1) & "' and ob03='" & m_CurrKEY(2) & "' "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("OB01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("OB01")
         If IsNull(rsTmp.Fields("OB02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("OB02")
         If IsNull(rsTmp.Fields("OB03")) = False Then: m_CurrKEY(2) = rsTmp.Fields("OB03")
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      End If
      rsTmp.Close
      
      strSql = "SELECT OB01,OB02,OB03 FROM OhBonus " & _
               "WHERE ob02 = (SELECT MIN(ob02) FROM OhBonus where ob01=(select min(ob01) from OhBonus) ) and ob01=(select min(ob01) from OhBonus) " & _
               " and ob03=(select min(ob03) from OhBonus where ob02 = (SELECT MIN(ob02) FROM OhBonus where ob01=(select min(ob01) from OhBonus) )) "
   
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("OB01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("OB01")
         If IsNull(rsTmp.Fields("OB02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("OB02")
         If IsNull(rsTmp.Fields("OB03")) = False Then: m_CurrKEY(2) = rsTmp.Fields("OB03")
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
   
   If m_CurrKEY(0) = m_FirstKEY(0) And m_CurrKEY(1) = m_FirstKEY(1) And m_CurrKEY(2) = m_FirstKEY(2) Then
      ShowMsg MsgText(9008)
      GoTo EXITSUB
   End If
   
   
   strSql = "SELECT ob01,ob02,ob03 FROM OhBonus " & _
            "WHERE ob01 = '" & m_CurrKEY(0) & "' AND " & _
                  "ob02 = '" & m_CurrKEY(1) & "'  " & _
                   " and ob03=(select max(ob03) from OhBonus where  ob01 = '" & m_CurrKEY(0) & "' and " & _
                                "ob02 = '" & m_CurrKEY(1) & "' and  ob03< '" & m_CurrKEY(2) & "')            "
                                
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("OB01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("OB01")
      If IsNull(rsTmp.Fields("OB02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("OB02")
      If IsNull(rsTmp.Fields("OB03")) = False Then: m_CurrKEY(2) = rsTmp.Fields("OB03")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT ob01,ob02,ob03 FROM OhBonus " & _
            "WHERE ob01 = '" & m_CurrKEY(0) & "' AND " & _
                  "ob02 = (SELECT MAX(ob02) FROM OhBonus " & _
                          "WHERE ob01 = " & m_CurrKEY(0) & " AND " & _
                                "ob02 < '" & m_CurrKEY(1) & "' ) " & _
                   " and ob03=(select max(ob03) from OhBonus where  ob01 = '" & m_CurrKEY(0) & "' and ob02 = (SELECT MAX(ob02) FROM OhBonus " & _
                          "WHERE ob01 = '" & m_CurrKEY(0) & "' AND " & _
                                "ob02 < '" & m_CurrKEY(1) & "' )            "
                                
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("OB01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("OB01")
      If IsNull(rsTmp.Fields("OB02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("OB02")
      If IsNull(rsTmp.Fields("OB03")) = False Then: m_CurrKEY(2) = rsTmp.Fields("OB03")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT ob01,ob02,ob03 FROM OhBonus " & _
            "WHERE ob01 = (select max(ob01) from ohbonus where ob01< '" & m_CurrKEY(0) & "') AND " & _
                  "ob02 = (SELECT MAX(ob02) FROM OhBonus " & _
                          "WHERE ob01 = (select max(ob01) from ohbonus where ob01< '" & m_CurrKEY(0) & "') ) " & _
                   " and ob03=(select max(ob03) from OhBonus where  ob02 = (SELECT MAX(ob02) FROM OhBonus " & _
                          "WHERE ob01 = (select max(ob01) from ohbonus where ob01< '" & m_CurrKEY(0) & "') ))            "
                                
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("OB01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("OB01")
      If IsNull(rsTmp.Fields("OB02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("OB02")
      If IsNull(rsTmp.Fields("OB03")) = False Then: m_CurrKEY(2) = rsTmp.Fields("OB03")
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
   
   If m_CurrKEY(0) = m_LastKEY(0) And m_CurrKEY(1) = m_LastKEY(1) And m_CurrKEY(2) = m_LastKEY(2) Then
      ShowMsg MsgText(9009)
      GoTo EXITSUB
   End If
   
   strSql = "SELECT ob01,ob02,ob03 FROM OhBonus " & _
            "WHERE ob01 = '" & m_CurrKEY(0) & "' AND " & _
                  "ob02 =  '" & m_CurrKEY(1) & "'  " & _
                   " and ob03=(select min(ob03) from OhBonus where  ob01  = '" & m_CurrKEY(0) & "' AND " & _
                                "ob02 = '" & m_CurrKEY(1) & "' and  ob03> '" & m_CurrKEY(2) & "')            "
                                
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("OB01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("OB01")
      If IsNull(rsTmp.Fields("OB02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("OB02")
      If IsNull(rsTmp.Fields("OB03")) = False Then: m_CurrKEY(2) = rsTmp.Fields("OB03")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT ob01,ob02,ob03 FROM OhBonus " & _
            "WHERE ob01 = '" & m_CurrKEY(0) & "' AND " & _
                  "ob02 = (select min(ob02) from OhBonus where ob02>'" & m_CurrKEY(1) & "' and ob01 = '" & m_CurrKEY(0) & "')  " & _
                   " and ob03=(select min(ob03) from OhBonus where  ob01  = '" & m_CurrKEY(0) & "' AND " & _
                                "ob02= (select min(ob02) from OhBonus where ob02>'" & m_CurrKEY(1) & "' and ob01 = '" & m_CurrKEY(0) & "') ) "
                                
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("OB01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("OB01")
      If IsNull(rsTmp.Fields("OB02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("OB02")
      If IsNull(rsTmp.Fields("OB03")) = False Then: m_CurrKEY(2) = rsTmp.Fields("OB03")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT ob01,ob02,ob03 FROM OhBonus " & _
            "WHERE ob01 =(select min(ob01) from OhBonus where ob01>'" & m_CurrKEY(0) & "') AND " & _
                  "ob02 = (select min(ob02) from OhBonus where ob01 =(select min(ob01) from OhBonus where ob01>'" & m_CurrKEY(0) & "'))  " & _
                   " and ob03=(select min(ob03) from OhBonus where  ob01 =(select min(ob01) from OhBonus where ob01>'" & m_CurrKEY(0) & "') AND " & _
                                "ob02 = (select min(ob02) from OhBonus where ob01 =(select min(ob01) from OhBonus where ob01>'" & m_CurrKEY(0) & "'))) "
                                
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("OB01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("OB01")
      If IsNull(rsTmp.Fields("OB02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("OB02")
      If IsNull(rsTmp.Fields("OB03")) = False Then: m_CurrKEY(2) = rsTmp.Fields("OB03")
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
   m_CurrKEY(2) = m_LastKEY(2)
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
         If Val(textOB01) + 191100 < Val(Left(strSrvDate(1), 6)) Then
            strTit = "修改"
            strMsg = "非當月資料不可修改!!"
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
   
   strSql = "SELECT ob01,ob02,ob03 FROM OhBonus " & _
            "WHERE ob01 = (SELECT MIN(ob01) FROM OhBonus) AND " & _
                  "ob02 = (SELECT MIN(ob02) FROM OhBonus " & _
                           "WHERE ob01 = (SELECT MIN(ob01) FROM OhBonus)) and " & _
                   "ob03=(select min(ob03) from OhBonus where  ob01 = (SELECT MIN(ob01) FROM OhBonus) AND " & _
                  "ob02 = (SELECT MIN(ob02) FROM OhBonus " & _
                           "WHERE ob01 = (SELECT MIN(ob01) FROM OhBonus))) "
                           
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("OB01")) = False Then: m_FirstKEY(0) = rsTmp.Fields("OB01")
      If IsNull(rsTmp.Fields("OB02")) = False Then: m_FirstKEY(1) = rsTmp.Fields("OB02")
      If IsNull(rsTmp.Fields("OB03")) = False Then: m_FirstKEY(2) = rsTmp.Fields("OB03")
   End If
   rsTmp.Close

   strSql = "SELECT ob01,ob02,ob03 FROM OhBonus " & _
            "WHERE ob01 = (SELECT MAX(ob01) FROM OhBonus) AND " & _
                  "ob02 = (SELECT MAX(ob02) FROM OhBonus " & _
                           "WHERE ob01 = (SELECT MAX(ob01) FROM OhBonus)) and " & _
                   "ob03=(select max(ob03) from OhBonus where  ob01 = (SELECT MAX(ob01) FROM OhBonus) AND " & _
                  "ob02 = (SELECT MAX(ob02) FROM OhBonus " & _
                           "WHERE ob01 = (SELECT MAX(ob01) FROM OhBonus))) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("OB01")) = False Then: m_LastKEY(0) = rsTmp.Fields("OB01")
      If IsNull(rsTmp.Fields("OB02")) = False Then: m_LastKEY(1) = rsTmp.Fields("OB02")
      If IsNull(rsTmp.Fields("OB03")) = False Then: m_LastKEY(2) = rsTmp.Fields("OB03")
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim i As Integer, j As Integer
   
   strSql = "SELECT * FROM OhBonus " & _
            "WHERE OB01='" & m_CurrKEY(0) & "' and OB02 = '" & m_CurrKEY(1) & "' and OB03='" & m_CurrKEY(2) & "'  "
                  
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      ClearField
      If IsNull(rsTmp.Fields("OB01")) = False Then
            If Len(ChangeWStringToTDateString(CheckStr(rsTmp.Fields("OB01")) & "01")) = 8 Then
                 textOB01.Text = Mid(Trim(ChangeWStringToTString(CheckStr(rsTmp.Fields("OB01")) & "01")), 1, 4)
            Else
                textOB01.Text = Mid(ChangeWStringToTString(CheckStr(rsTmp.Fields("OB01")) & "01"), 1, 5)
            End If
      End If
      If IsNull(rsTmp.Fields("OB02")) = False Then: textOB02 = rsTmp.Fields("OB02")
      If IsNull(rsTmp.Fields("OB03")) = False Then: textOB03 = rsTmp.Fields("OB03")
      If IsNull(rsTmp.Fields("OB04")) = False Then: textOB04 = rsTmp.Fields("OB04")
      If IsNull(rsTmp.Fields("OB05")) = False Then: textOB05 = rsTmp.Fields("OB05")
      
      textOB03_2 = GetStaffName(textOB03, True)
      If Val(textOB04) > 0 Then
         textOB04_2 = "為  " & PUB_ChangeNianZi(Val(textOB04))
      Else
         textOB04_2 = ""
      End If
      
      ' 更新CUID
      UpdateCUID rsTmp
      ' 更新暫存區的資料
      UpdateFieldOldData rsTmp

      '更新入所日
      strSql = "select * from staff where st01='" & textOB03 & "' "
      If rsTmp.State = 1 Then rsTmp.Close
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
          lblST13.Caption = ChangeWStringToTDateString(CheckStr(rsTmp.Fields("st13")))
      Else
          lblST13.Caption = ""
      End If
      '更新任職時間
      strSql = "select sqldatet(sc02) as 日期,ac03 as 原因 from staff_change,allcode where sc01='" & textOB03 & "' and ac01='05' and sc03=ac02(+) order by sc02 ASC "
      If rsTmp.State = 1 Then rsTmp.Close
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      Set Grd2.Recordset = rsTmp
   End If

   rsTmp.Close
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

Sub GetData()
Dim rsTmp As New ADODB.Recordset
strSql = ""
If txt1(0) <> "" Then
    strSql = strSql & " and ob03>='" & txt1(0) & "' "
End If
If txt1(1) <> "" Then
    strSql = strSql & " and ob03<='" & txt1(1) & "' "
End If
If txt1(2) <> "" Then
    strSql = strSql & " and ob01>='" & txt1(2) + 191100 & "' "
End If
If txt1(3) <> "" Then
    strSql = strSql & " and ob01<='" & txt1(3) + 191100 & "' "
End If
'抓取資料
strSql = "SELECT decode(length(ob01-191100),4,substr(ob01-191100,1,2)||'/'||substr(ob01-191100,3,2),substr(ob01-191100,1,3)||'/'||substr(ob01-191100,4,2))," & _
                "ob02||' '||decode(ob02,'1','端午','2','中秋',''),ob03,st02,ob04,ob05 FROM OhBonus,staff where ob03=st01(+) " & strSql & _
        " order by ob01,ob02,ob03 "
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
   textOB01_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textOB02_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textOB03_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textOB04_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textOB05_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   CheckDataValid = True
EXITSUB:
End Function

' 更新Key的狀態
Private Sub SetKeyReadOnly(ByVal bEnable As Boolean)
   textOB01.Locked = bEnable
   textOB02.Locked = bEnable
   textOB03.Locked = bEnable
   If bEnable Then textOB01.BackColor = &H8000000F Else textOB01.BackColor = &H80000005
   If bEnable Then textOB02.BackColor = &H8000000F Else textOB02.BackColor = &H80000005
   If bEnable Then textOB03.BackColor = &H8000000F Else textOB03.BackColor = &H80000005
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
Dim i As Integer
   textOB01.Locked = bEnable
   textOB02.Locked = bEnable
   textOB03.Locked = bEnable
   If bEnable Then textOB01.BackColor = &H8000000F Else textOB01.BackColor = &H80000005
   If bEnable Then textOB02.BackColor = &H8000000F Else textOB02.BackColor = &H80000005
   If bEnable Then textOB03.BackColor = &H8000000F Else textOB03.BackColor = &H80000005
   textOB04.Locked = bEnable
   textOB05.Locked = bEnable
End Sub

Private Sub ClearField()
   Dim nIndex As Integer
   textOB01 = Empty
   textOB02 = Empty
   textOB03 = Empty
   textOB03_2 = Empty
   textOB04 = Empty
   textOB05 = Empty
   Label23 = Empty
   SetGrd
   For nIndex = 0 To tf_OB - 1
      m_FieldList(nIndex).fiOldData = Empty
      m_FieldList(nIndex).fiNewData = Empty
   Next nIndex
End Sub

Private Sub UpdateFieldNewData()
    Dim MyArr As Variant
   '若新增資料
   If m_EditMode = 1 Then
      SetFieldNewData "OB01", Mid(DBDATE(textOB01.Text & "01"), 1, 6)
      SetFieldNewData "OB02", textOB02
      SetFieldNewData "OB03", textOB03
   End If
   SetFieldNewData "OB04", textOB04
   SetFieldNewData "OB05", textOB05
End Sub

' 初始化欄位陣列
Private Sub InitialField()
   Dim nIndex As Integer
   Dim strTmp As String
   ' 初始化欄位陣列
   For nIndex = 1 To tf_OB
      strTmp = Format(nIndex, "00")
      m_FieldList(nIndex - 1).fiName = "OB" & strTmp
      m_FieldList(nIndex - 1).fiOldData = Empty
      m_FieldList(nIndex - 1).fiNewData = Empty
      m_FieldList(nIndex - 1).fiType = 0 '文字型態
      Select Case nIndex
         Case 1, 4, 5:
            m_FieldList(nIndex - 1).fiType = 1 '數值型態
      End Select
   Next nIndex
End Sub

'帶預設資料
Private Sub InitialData()
SetGrd
SetGrd2
End Sub

Private Sub textBD01_GotFocus()
InverseTextBox textBD01
End Sub

Private Sub textBD01_KeyPress(KeyAscii As Integer)
KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textBD01_Validate(Cancel As Boolean)
If textBD01.Text = "" Then Exit Sub
If CheckIsTaiwanDate(textBD01.Text & "01", False) = False Then
    Cancel = True
    MsgBox "請輸入民國年月！", vbInformation, "輸入獎金年月錯誤"
    Exit Sub
End If
End Sub

Private Sub textBD02_GotFocus()
InverseTextBox textBD02
End Sub

Private Sub textBD02_KeyPress(KeyAscii As Integer)
KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textBD02_Validate(Cancel As Boolean)
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   Select Case textBD02
   Case "1", "2", ""
   Case Else
           MsgBox "獎金類別請輸入 1 或 2 或是空白！", vbExclamation, "操作錯誤！"
           textBD02_GotFocus
           Cancel = True
           Exit Sub
   End Select
   
   strSql = "SELECT * FROM BonusDefinition " & _
            "WHERE BD01='" & CStr(Val(textBD01) + 191100) & "' and BD02 = '" & textBD02 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("BD03")) = False Then: textBD03 = rsTmp.Fields("BD03")
      If IsNull(rsTmp.Fields("BD04")) = False Then: textBD04 = rsTmp.Fields("BD04")
      If IsNull(rsTmp.Fields("BD05")) = False Then: textBD05 = rsTmp.Fields("BD05")
      If IsNull(rsTmp.Fields("BD06")) = False Then: textBD06 = rsTmp.Fields("BD06")
      If IsNull(rsTmp.Fields("BD07")) = False Then: textBD07 = rsTmp.Fields("BD07")
      If IsNull(rsTmp.Fields("BD08")) = False Then: textBD08 = rsTmp.Fields("BD08")
      If IsNull(rsTmp.Fields("BD09")) = False Then: textBD09 = rsTmp.Fields("BD09")
      If IsNull(rsTmp.Fields("BD10")) = False Then: textBD10 = rsTmp.Fields("BD10")
      If IsNull(rsTmp.Fields("BD11")) = False Then: textBD11 = rsTmp.Fields("BD11")
      If IsNull(rsTmp.Fields("BD12")) = False Then: textBD12 = rsTmp.Fields("BD12")
   End If
End Sub

Private Sub textBD03_GotFocus()
InverseTextBox textBD03
End Sub

Private Sub textBD03_KeyPress(KeyAscii As Integer)
KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textBD04_GotFocus()
InverseTextBox textBD04
End Sub

Private Sub textBD04_KeyPress(KeyAscii As Integer)
KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textBD05_GotFocus()
InverseTextBox textBD05
End Sub

Private Sub textBD05_KeyPress(KeyAscii As Integer)
KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textBD06_GotFocus()
InverseTextBox textBD06
End Sub

Private Sub textBD06_KeyPress(KeyAscii As Integer)
KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textBD07_GotFocus()
InverseTextBox textBD07
End Sub

Private Sub textBD07_KeyPress(KeyAscii As Integer)
KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textBD08_GotFocus()
InverseTextBox textBD08
End Sub

Private Sub textBD08_KeyPress(KeyAscii As Integer)
KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textBD09_GotFocus()
InverseTextBox textBD09
End Sub

Private Sub textBD09_KeyPress(KeyAscii As Integer)
KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textBD10_GotFocus()
InverseTextBox textBD10
End Sub

Private Sub textBD10_KeyPress(KeyAscii As Integer)
KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textBD11_GotFocus()
InverseTextBox textBD11
End Sub

Private Sub textBD11_KeyPress(KeyAscii As Integer)
KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textBD12_GotFocus()
InverseTextBox textBD12
End Sub

Private Sub textBD12_KeyPress(KeyAscii As Integer)
KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textOB01_GotFocus()
If m_EditMode <> 0 Then
    InverseTextBox textOB01
End If
End Sub

Private Sub textOB01_KeyPress(KeyAscii As Integer)
KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textOB01_Validate(Cancel As Boolean)
If m_EditMode = 1 And textOB01.Text <> "" Then
    If IsRecordExist(Mid(DBDATE(textOB01.Text & "01"), 1, 6), textOB02, textOB03) = True And textOB01.Enabled = True And textOB01.Locked = False Then
        MsgBox "該員工當年度已有資料，請修改！", vbInformation
        Cancel = True
        Exit Sub
    End If
    If CheckIsTaiwanDate(textOB01.Text & "01", False) = False Then
        Cancel = True
        MsgBox "請輸入民國年月不含/！", vbInformation, "輸入獎金年月錯誤"
        Exit Sub
    End If
End If
End Sub

Private Sub textOB03_GotFocus()
If m_EditMode <> 0 Then
    InverseTextBox textOB03
    CloseIme
End If
End Sub

Private Sub textOB03_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textOB03_Validate(Cancel As Boolean)
If m_EditMode = 1 And textOB03 <> "" Then
     textOB03_2 = GetStaffName(textOB03, True)
    If IsRecordExist(Mid(DBDATE(textOB01.Text & "01"), 1, 6), textOB02, textOB03) = True And textOB03.Enabled = True And textOB03.Locked = False Then
        MsgBox "該員工當年度已有資料，請修改！", vbInformation
        Cancel = True
        Exit Sub
    End If
    If textOB03_2 = "" Then
        MsgBox "員工編號錯誤！查無此員工！", vbInformation
        Cancel = True
        Exit Sub
    End If
End If
End Sub

Private Sub SetGrd()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   arrGridHeadText = Array("獎金年月", "獎金類別", "員工編號", "姓名", "年資", "金額")
   arrGridHeadWidth = Array(800, 800, 800, 1200, 800, 1000)
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

Private Sub textOB02_GotFocus()
If m_EditMode <> 0 Then
    InverseTextBox textOB02
End If
End Sub

Private Sub textOB02_KeyPress(KeyAscii As Integer)
KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textOB02_Validate(Cancel As Boolean)
If m_EditMode <> 1 And textOB02 <> "" Then
    If IsRecordExist(Mid(DBDATE(textOB01.Text & "01"), 1, 6), textOB02, textOB03) = True And textOB02.Enabled = True And textOB03.Locked = False Then
        MsgBox "該員工當年度已有資料，請修改！", vbInformation
        Cancel = True
        Exit Sub
    End If
    Select Case textOB02
    Case "1", "2", ""
    Case Else
        MsgBox "獎金類別請輸入 1 或 2 ！", vbExclamation, "輸入錯誤！"
        Cancel = True
        Exit Sub
    End Select
End If
CloseIme
End Sub

Private Sub textOB04_GotFocus()
If m_EditMode <> 0 Then
    InverseTextBox textOB04
End If
End Sub

Private Sub textOB04_KeyPress(KeyAscii As Integer)
KeyAscii = Pub_NumAscii(KeyAscii, True)
End Sub

Private Sub textOB04_Validate(Cancel As Boolean)
If m_EditMode <> 0 And textOB04 <> "" Then
    If IsNumeric(textOB04) = False Then
        MsgBox "請輸入數字！", vbExclamation, "操作錯誤！"
        Cancel = True
        Exit Sub
    End If
End If
End Sub

Private Sub textOB05_GotFocus()
If m_EditMode <> 0 Then
    InverseTextBox textOB05
End If
End Sub

Private Sub textOB05_KeyPress(KeyAscii As Integer)
KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textOB05_Validate(Cancel As Boolean)
If m_EditMode <> 0 And textOB05 <> "" Then
    If IsNumeric(textOB05) = False Then
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
Case 2, 3
        If CheckIsTaiwanDate(txt1(Index) & "01", False) = False Then
            Cancel = True
            MsgBox "請輸入民國年度不含/！", vbInformation, "輸入年假年度錯誤"
            Call txt1_GotFocus(Index)
            Exit Sub
        End If
         If Index = 2 Then
            If txt1(Index) <> "" And txt1(Index + 1) = "" Then
               txt1(Index + 1) = txt1(Index)
            End If
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

'計算獎金
Sub CalBonus()
Dim m_rs As New ADODB.Recordset
Dim m_StrSQL As String
Dim m_Year As String
Dim m_item As Long
Dim m_money As String
Dim strEndDate As String

Screen.MousePointer = vbHourglass
Me.Enabled = False
ListInfo.Clear
PB1.Value = 0
m_item = 0
ListInfo.AddItem "開始統計   " & textBD01.Text & " " & IIf(textBD02 = "1", "端午", "中秋") & " 獎金資料", m_item
ListInfo.Selected(m_item) = True
m_item = m_item + 1
On Error GoTo err1
cnnConnection.BeginTrans

'若先已經存在-->刪除
Set m_rs = New ADODB.Recordset
'm_StrSQL = "select * from OhBonus where ob01='" & Mid(Trim(DBDATE(textBD01 & "01")), 1, 6) & "' and ob02='" & textBD02 & "' "
m_StrSQL = "select * from OhBonus where substr(ob01,1,4)='" & Mid(Trim(DBDATE(textBD01 & "01")), 1, 4) & "' and ob02='" & textBD02 & "' "
If m_rs.State = 1 Then m_rs.Close
m_rs.CursorLocation = adUseClient
m_rs.Open m_StrSQL, cnnConnection, adOpenStatic, adLockReadOnly
If m_rs.RecordCount <> 0 Then
'    Pub_SeekTbLog "delete from OhBonus where ob01='" & Mid(Trim(DBDATE(textBD01 & "01")), 1, 6) & "' and ob02='" & textBD02 & "' "
    'cnnConnection.Execute "delete from OhBonus where ob01='" & Mid(Trim(DBDATE(textBD01 & "01")), 1, 6) & "' and ob02='" & textBD02 & "' "
    cnnConnection.Execute "delete from OhBonus where substr(ob01,1,4)='" & Mid(Trim(DBDATE(textBD01 & "01")), 1, 4) & "' and ob02='" & textBD02 & "' "
'    Pub_SeekTbLog "delete from BonusDefinition where bd01='" & Mid(Trim(DBDATE(textBD01 & "01")), 1, 6) & "' and bd02='" & textBD02 & "' "
    'cnnConnection.Execute "delete from BonusDefinition where bd01='" & Mid(Trim(DBDATE(textBD01 & "01")), 1, 6) & "' and bd02='" & textBD02 & "' "
    cnnConnection.Execute "delete from BonusDefinition where substr(bd01,1,4)='" & Mid(Trim(DBDATE(textBD01 & "01")), 1, 4) & "' and bd02='" & textBD02 & "' "
End If
'Pub_SeekTbLog "insert into BonusDefinition (bd01,bd02,bd03,bd04,bd05,bd06,bd07,bd08,bd09,bd10,bd11,bd12) values ('" & Mid(Trim(DBDATE(textBD01 & "01")), 1, 6) & "','" & textBD02 & "','" & textBD03 & "','" & textBD04 & "','" & textBD05 & "','" & textBD06 & "','" & textBD07 & "','" & textBD08 & "','" & textBD09 & "','" & textBD10 & "','" & textBD11 & "','" & textBD12 & "') "
cnnConnection.Execute "insert into BonusDefinition (bd01,bd02,bd03,bd04,bd05,bd06,bd07,bd08,bd09,bd10,bd11,bd12) values ('" & Mid(Trim(DBDATE(textBD01 & "01")), 1, 6) & "','" & textBD02 & "','" & textBD03 & "','" & textBD04 & "','" & textBD05 & "','" & textBD06 & "','" & textBD07 & "','" & textBD08 & "','" & textBD09 & "','" & textBD10 & "','" & textBD11 & "','" & textBD12 & "') "

'開始計算
Set m_rs = New ADODB.Recordset
'2011/4/19 MODIFY BY SONIA 改同frm160010 但加入68007
'm_StrSQL = "select * from staff,SalaryData where st04='1' and ST01=SD01 and ((sd02 not in('P','F') or sd02 is null) or ST01='68007') and ascii(substr(st01,1,1))>=48 and ascii(substr(st01,1,1))<=57  " & _
                    " and ((length(st01)=5 and substr(st01,1,1) not in ('0','1','2','3','4','5')) and st01 not in ('99998','99999')) " & _
                    " order by st01  "
'Modify By Sindy 2021/6/4 + 排除 98099.江郁仁
'Modify By Sindy 2024/4/30 + and not(substr(st01,5,1)>='A') 排除 B309A=宗家澔
m_StrSQL = "select * from staff,SalaryData where st04='1' and ST01=SD01 and (sd02 not in('P','F') or ST01='68007') " & _
           " and ST01 not in('98099')" & _
           " and not(substr(st01,5,1)>='A')" & _
           " order by st01 asc"
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
            strEndDate = CStr(Val(Left(CStr(Val(textBD01 & "31") + 19110000), 4)) - 1) & "1231"
            '計算年資
            m_Year = Trim(CalYear(CheckStr(m_rs.Fields("st01")), strEndDate))
            If Val(m_Year) < 0 Then m_Year = "0"
            If m_Year = "" Then
                ListInfo.AddItem "員工 " & CheckStr(m_rs.Fields("st01")) & " " & CheckStr(m_rs.Fields("st02")) & "-->計算失敗，原因  統計錯誤 ，請聯絡電腦中心檢查  ", m_item
                ListInfo.Selected(m_item) = True
                m_item = m_item + 1
            Else
                m_money = ""
                If Val(m_Year) >= Val(textBD11) And Val(textBD11) > 0 Then
                    m_money = textBD12
                ElseIf Val(m_Year) >= Val(textBD09) And Val(textBD09) > 0 Then
                    m_money = textBD10
                ElseIf Val(m_Year) >= Val(textBD07) And Val(textBD07) > 0 Then
                    m_money = textBD08
                ElseIf Val(m_Year) >= Val(textBD05) And Val(textBD05) > 0 Then
                    m_money = textBD06
                ElseIf Val(m_Year) >= Val(textBD03) And Val(textBD03) > 0 Then
                    m_money = textBD04
                ElseIf Val(textBD03) = 0 Then
                    m_money = textBD04
                Else
                    m_money = 0
                End If
                'Pub_SeekTbLog " insert into OhBonus (ob01,ob02,ob03,ob04,ob05) values ('" & Mid(Trim(DBDATE(textBD01 & "01")), 1, 6) & "','" & textBD02 & "','" & CheckStr(m_rs.Fields("st01")) & "','" & m_Year & "','" & m_money & "' ) "
                'modify by sonia 2016/1/18 +ob12
                'cnnConnection.Execute " insert into OhBonus (ob01,ob02,ob03,ob04,ob05) values ('" & Mid(Trim(DBDATE(textBD01 & "01")), 1, 6) & "','" & textBD02 & "','" & CheckStr(m_rs.Fields("st01")) & "','" & m_Year & "','" & m_money & "' ) "
                cnnConnection.Execute " insert into OhBonus (ob01,ob02,ob03,ob04,ob05,ob12) values ('" & Mid(Trim(DBDATE(textBD01 & "01")), 1, 6) & "','" & textBD02 & "','" & CheckStr(m_rs.Fields("st01")) & "','" & m_Year & "','" & m_money & "','" & CheckStr(m_rs.Fields("sd19")) & "' ) "
                ListInfo.AddItem "員工 " & CheckStr(m_rs.Fields("st01")) & " " & CheckStr(m_rs.Fields("st02")) & "-->計算成功          入所日：" & ChangeWStringToTDateString(CheckStr(m_rs.Fields("st13"))) & "    年資：" & m_Year & "；獎金：" & m_money, m_item
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
