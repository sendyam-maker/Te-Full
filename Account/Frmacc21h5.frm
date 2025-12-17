VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Frmacc21h5 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "國內收據點數分配輸入"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8790
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4845
   ScaleWidth      =   8790
   Begin VB.CommandButton Command10 
      Height          =   300
      Left            =   2300
      Picture         =   "Frmacc21h5.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   120
      Width           =   350
   End
   Begin VB.CommandButton Command9 
      Caption         =   "點數重新分配"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6930
      TabIndex        =   39
      Top             =   540
      Width           =   1770
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4155
      Left            =   90
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   630
      Width           =   8610
      _ExtentX        =   15187
      _ExtentY        =   7329
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "專業點數"
      TabPicture(0)   =   "Frmacc21h5.frx":0102
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label9"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label6"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label13"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label14"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Shape1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "DataGrid1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Adodc1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtSum"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Command5"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtA1N03"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtA1N04"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtST02"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtA1N05"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Command1"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Command2"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Command7"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      TabCaption(1)   =   "業務點數"
      TabPicture(1)   =   "Frmacc21h5.frx":011E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command6"
      Tab(1).Control(1)=   "Command4"
      Tab(1).Control(2)=   "txtA1N05_1"
      Tab(1).Control(3)=   "txtST02_1"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "txtA1N04_1"
      Tab(1).Control(5)=   "txtA1N03_1"
      Tab(1).Control(6)=   "Command3"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "txtSum_1"
      Tab(1).Control(8)=   "Adodc2"
      Tab(1).Control(9)=   "DataGrid2"
      Tab(1).Control(10)=   "Shape2"
      Tab(1).Control(11)=   "Label11"
      Tab(1).Control(12)=   "Label10"
      Tab(1).Control(13)=   "Label4"
      Tab(1).Control(14)=   "Label2"
      Tab(1).ControlCount=   15
      Begin VB.CommandButton Command7 
         Caption         =   "複製到業務點數"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6075
         TabIndex        =   37
         Top             =   2670
         Width           =   2400
      End
      Begin VB.CommandButton Command6 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.5
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   -67875
         Picture         =   "Frmacc21h5.frx":013A
         Style           =   1  '圖片外觀
         TabIndex        =   14
         ToolTipText     =   "清除畫面"
         Top             =   3315
         Width           =   550
      End
      Begin VB.CommandButton Command4 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.5
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   -67275
         Picture         =   "Frmacc21h5.frx":0A04
         Style           =   1  '圖片外觀
         TabIndex        =   15
         ToolTipText     =   "取消"
         Top             =   3315
         Width           =   550
      End
      Begin VB.TextBox txtA1N05_1 
         Alignment       =   1  '靠右對齊
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.5
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -72330
         MaxLength       =   8
         TabIndex        =   10
         Top             =   3495
         Width           =   945
      End
      Begin VB.TextBox txtST02_1 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.5
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -73770
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   3495
         Width           =   1332
      End
      Begin VB.TextBox txtA1N04_1 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.5
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74730
         MaxLength       =   6
         TabIndex        =   9
         Top             =   3495
         Width           =   972
      End
      Begin VB.TextBox txtA1N03_1 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.5
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -71280
         TabIndex        =   11
         Top             =   3495
         Width           =   1125
      End
      Begin VB.CommandButton Command3 
         Height          =   300
         Left            =   -70155
         Picture         =   "Frmacc21h5.frx":106E
         Style           =   1  '圖片外觀
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   3510
         Width           =   350
      End
      Begin VB.TextBox txtSum_1 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.5
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -72480
         TabIndex        =   32
         Top             =   2730
         Width           =   1005
      End
      Begin VB.CommandButton Command2 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.5
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   7125
         Picture         =   "Frmacc21h5.frx":1170
         Style           =   1  '圖片外觀
         TabIndex        =   7
         ToolTipText     =   "清除畫面"
         Top             =   3315
         Width           =   550
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.5
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   7725
         Picture         =   "Frmacc21h5.frx":1A3A
         Style           =   1  '圖片外觀
         TabIndex        =   8
         ToolTipText     =   "取消"
         Top             =   3315
         Width           =   550
      End
      Begin VB.TextBox txtA1N05 
         Alignment       =   1  '靠右對齊
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.5
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2670
         MaxLength       =   8
         TabIndex        =   4
         Top             =   3495
         Width           =   945
      End
      Begin VB.TextBox txtST02 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.5
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1230
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   3495
         Width           =   1332
      End
      Begin VB.TextBox txtA1N04 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.5
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   270
         MaxLength       =   6
         TabIndex        =   2
         Top             =   3495
         Width           =   972
      End
      Begin VB.TextBox txtA1N03 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.5
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3720
         TabIndex        =   3
         Top             =   3495
         Width           =   1125
      End
      Begin VB.CommandButton Command5 
         Height          =   300
         Left            =   4845
         Picture         =   "Frmacc21h5.frx":20A4
         Style           =   1  '圖片外觀
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   3510
         Width           =   350
      End
      Begin VB.TextBox txtSum 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.5
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2520
         TabIndex        =   26
         Top             =   2730
         Width           =   1005
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   315
         Left            =   360
         Top             =   1590
         Visible         =   0   'False
         Width           =   960
         _ExtentX        =   2117
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2145
         Left            =   180
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   450
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   3784
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   0   'False
         BackColor       =   -2147483624
         HeadLines       =   1
         RowHeight       =   16
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "a1n04"
            Caption         =   "承辦人代號"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1028
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "st02"
            Caption         =   "姓名"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1028
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "a1n05"
            Caption         =   "點數"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1028
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "a1n03"
            Caption         =   "收文號"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1028
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "cpm03"
            Caption         =   "案件性質"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1028
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            AllowRowSizing  =   0   'False
            BeginProperty Column00 
               ColumnWidth     =   1170.142
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1035.213
            EndProperty
            BeginProperty Column02 
               Alignment       =   1
               ColumnWidth     =   720
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1275.024
            EndProperty
            BeginProperty Column04 
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   315
         Left            =   -74640
         Top             =   1590
         Visible         =   0   'False
         Width           =   960
         _ExtentX        =   2117
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Height          =   2145
         Left            =   -74820
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   450
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   3784
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   0   'False
         BackColor       =   -2147483624
         HeadLines       =   1
         RowHeight       =   16
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "a1n04"
            Caption         =   "智權人員"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1028
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "st02"
            Caption         =   "姓名"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1028
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "a1n05"
            Caption         =   "點數"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1028
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "a1n03"
            Caption         =   "收文號"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1028
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "cpm03"
            Caption         =   "案件性質"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1028
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            AllowRowSizing  =   0   'False
            BeginProperty Column00 
               ColumnWidth     =   1170.142
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1035.213
            EndProperty
            BeginProperty Column02 
               Alignment       =   1
               ColumnWidth     =   720
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1275.024
            EndProperty
            BeginProperty Column04 
            EndProperty
         EndProperty
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FF0000&
         Height          =   915
         Left            =   -74865
         Top             =   3120
         Width           =   8340
      End
      Begin VB.Label Label11 
         Alignment       =   2  '置中對齊
         BackStyle       =   0  '透明
         Caption         =   "點數"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.5
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -72330
         TabIndex        =   36
         Top             =   3240
         Width           =   945
      End
      Begin VB.Label Label10 
         Alignment       =   2  '置中對齊
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "智權人員"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.5
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -74040
         TabIndex        =   35
         Top             =   3240
         Width           =   915
      End
      Begin VB.Label Label4 
         Alignment       =   2  '置中對齊
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "收文號"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.5
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -71025
         TabIndex        =   34
         Top             =   3240
         Width           =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "點數合計"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.5
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -73515
         TabIndex        =   33
         Top             =   2745
         Width           =   900
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FF0000&
         Height          =   915
         Left            =   135
         Top             =   3120
         Width           =   8340
      End
      Begin VB.Label Label14 
         Alignment       =   2  '置中對齊
         BackStyle       =   0  '透明
         Caption         =   "點數"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.5
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2670
         TabIndex        =   30
         Top             =   3240
         Width           =   945
      End
      Begin VB.Label Label13 
         Alignment       =   2  '置中對齊
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "承辦人"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.5
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1065
         TabIndex        =   29
         Top             =   3240
         Width           =   705
      End
      Begin VB.Label Label6 
         Alignment       =   2  '置中對齊
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "收文號"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.5
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3975
         TabIndex        =   28
         Top             =   3240
         Width           =   705
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "點數合計"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.5
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1485
         TabIndex        =   27
         Top             =   2745
         Width           =   900
      End
   End
   Begin VB.CommandButton Command8 
      Caption         =   "離開"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7605
      TabIndex        =   38
      Top             =   90
      Width           =   1095
   End
   Begin VB.TextBox txtPts 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6300
      TabIndex        =   22
      Top             =   120
      Width           =   1212
   End
   Begin VB.TextBox txtA0K01 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1125
      MaxLength       =   9
      TabIndex        =   0
      Top             =   120
      Width           =   1515
   End
   Begin VB.TextBox txtCP01 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3690
      TabIndex        =   19
      Top             =   120
      Width           =   492
   End
   Begin VB.TextBox txtCP02 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4170
      TabIndex        =   18
      Top             =   120
      Width           =   852
   End
   Begin VB.TextBox txtCP03 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5010
      TabIndex        =   17
      Top             =   120
      Width           =   252
   End
   Begin VB.TextBox txtCP04 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5250
      TabIndex        =   16
      Top             =   120
      Width           =   372
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "點數"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5820
      TabIndex        =   23
      Top             =   150
      Width           =   465
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "收據號碼"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   165
      TabIndex        =   21
      Top             =   150
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "本所案號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2730
      TabIndex        =   20
      Top             =   150
      Width           =   975
   End
End
Attribute VB_Name = "Frmacc21h5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create by Lydia 2015/05/08 國內收據點數分配輸入
'Memo by Lydia 2020/04/06
'1.原本請作單為"新增CFL案件分配點數功能"
'　1.1新增程式:法務\外法\檔案維護\國內收據點數分配輸入，並配合財務室的銷帳退費作業，由本處成員填寫銷帳後之剩餘點收分配情形。
'　1.2修改相關查詢報表：外法\報表列印\收發文明細表和收發文統計表。
'2.後來請作單修改為”法務工作點數分配Frm071021”，國內收據點數分配輸入(Frmacc21h5)程式移到Computer做保留。
'end 2020/04/06
Option Explicit

Dim adoacc1n0 As ADODB.Recordset
Dim m_bolAddNew As Boolean
Dim m_A1N03_CPM03 As String '收文號的案件性質

Dim adoacc1n0_1 As ADODB.Recordset
Dim m_bolAddNew_1 As Boolean
Dim m_A1N03_CPM03_1 As String '收文號的案件性質
Dim m_bolQuery As Boolean '是否為查詢
Public m_PrevForm As Form  '前一畫面
Public m_bolPrev As Boolean '是否為外部呼叫
Private Sub Command1_Click()
   AdodcDelete adoacc1n0
   AdodcClear
   DataGrid1.Refresh
   SumShow
End Sub

Private Sub Command2_Click()
   AdodcClear
   txtA1N04.SetFocus
End Sub

Private Sub Command3_Click()
   Dim bCancel As Boolean
   strExc(0) = "select '',cp09,sqldatet(cp05) cp05,decode(cpm03,'（無）',cpm04,cpm03) cpm03,st02,cp13" & _
     " from caseprogress,casepropertymap,staff where cp60='" & txtA0K01 & "'" & _
      " and cpm01(+)=cp01 and cpm02(+)=cp10 and st01(+)=cp13"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      Set Frmacc21h4.grdDataList.Recordset = RsTemp
      Frmacc21h4.grdDataList.FormatString = Replace(Frmacc21h4.grdDataList.FormatString, "承辦人", "智權人員")
      Set Frmacc21h4.fmParent = Me
      Frmacc21h4.Show vbModal
      strFormName = Me.Name
      If Me.Tag <> "" Then
         txtA1N03_1 = Me.Tag
      End If
      txtA1N03_1.SetFocus
   End If
End Sub

Private Sub Command4_Click()
   AdodcDelete adoacc1n0_1
   AdodcClear_1
   DataGrid2.Refresh
   SumShow_1
End Sub

Private Sub Command5_Click()
   Dim bCancel As Boolean
   strExc(0) = "select '',cp09,sqldatet(cp05) cp05,decode(cpm03,'（無）',cpm04,cpm03) cpm03,st02,cp13" & _
      " from caseprogress,casepropertymap,staff where cp60='" & txtA0K01 & "'" & _
      " and cpm01(+)=cp01 and cpm02(+)=cp10 and st01(+)=cp14"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      Set Frmacc21h4.grdDataList.Recordset = RsTemp
      Set Frmacc21h4.fmParent = Me
      Frmacc21h4.Show vbModal
      strFormName = Me.Name
      If Me.Tag <> "" Then
         txtA1N03 = Me.Tag
      End If
      txtA1N03.SetFocus
   End If
End Sub

Private Sub Command6_Click()
   AdodcClear_1
   txtA1N04_1.SetFocus
End Sub

Private Sub Command7_Click()
   Set RsTemp = adoacc1n0.Clone
   '+FormName 改暫存TB
   Set adoacc1n0_1 = PUB_CreateRecordset(RsTemp, , , , Me.Name)
   adoacc1n0_1.Sort = "a1n04,a1n03"
   Set RsTemp = Nothing
   Set Adodc2.Recordset = adoacc1n0_1
   Set DataGrid2.DataSource = Adodc2
   DataGrid2.Refresh
   AdodcClear_1
   SumShow_1
   If FormSave Then
      MsgBox "複製完成！"
      SSTab1.Tab = 1
   End If
   
End Sub

Private Sub Command8_Click()
   If m_bolQuery Or (Val(txtSum) = 0 And Val(txtSum_1) = 0) Then
      Unload Me
      If m_bolPrev = True Then
         m_PrevForm.Visible = True
      End If
   Else
      If Val(txtSum) <> Val(txtSum_1) Then
           MsgBox "專業點數合計與業務點數合計不符！"
           SSTab1.Tab = 0
      ElseIf Val(txtPts) > 0 And (Val(txtPts) <> Val(txtSum) Or Val(txtPts) <> Val(txtSum_1)) Then
           If MsgBox("輸入點數與預計點數不符，是否要繼續輸入？", vbYesNo + vbDefaultButton2) = vbYes Then
              SSTab1.Tab = 0
           ElseIf FormSave Then
                Unload Me
           End If
      ElseIf FormSave Then
             Unload Me
      End If
   End If
End Sub
Private Sub Command9_Click()
   If txtA0K01.Tag <> txtA0K01.Text Then
      Call Command10_Click
   End If
   
   If MsgBox("系統將清除目前分配並依照規則重新分配，是否確定要繼續？", vbYesNo + vbDefaultButton2) = vbYes Then
      Get_PointAutoassign (txtA0K01)
      OpenTable
      MsgBox "重新分配完畢，若有特殊分配請再人工調整！"
   End If
End Sub


Private Sub DataGrid1_SelChange(Cancel As Integer)
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   AdodcShow
End Sub

Private Sub DataGrid2_SelChange(Cancel As Integer)
   If Adodc2.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   AdodcShow_1
End Sub

Private Sub Form_Activate()
   If txtA0K01 <> "" Then
      Call Command10_Click
   End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   If Not m_bolQuery Then
      KeyDefine KeyCode
   End If
End Sub

Private Sub Form_Load()
   PUB_InitForm Me, Me.Width, Me.Height
   
   Screen.MousePointer = vbDefault
   tool4_enabled
   SSTab1.Tab = 0

End Sub

Private Sub SetFormEnable(bolEnabled As Boolean)
   Dim oControl As Control
   For Each oControl In Me.Controls
      If TypeName(oControl) = "CommandButton" Then
         oControl.Enabled = bolEnabled
      
      ElseIf TypeName(oControl) = "TextBox" Then
         oControl.Locked = Not bolEnabled
         
      End If
   Next
   '一律開放單據查詢和離開鈕
   Command8.Enabled = True
   Command10.Enabled = True
   txtA0K01.Enabled = True
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set Frmacc21h5 = Nothing
   If m_bolPrev = True Then
      tool1_enabled
      m_PrevForm.Visible = True
   End If
End Sub

Private Function FormSave() As Boolean
   cnnConnection.BeginTrans
On Error GoTo ErrHnd
   strSql = "delete acc1n0 where a1n01='" & txtA0K01.Tag & "'"
   cnnConnection.Execute strSql, intI
   With adoacc1n0
   If .RecordCount > 0 Then
      .MoveFirst
      Do While Not .EOF
         strSql = "insert into acc1n0(a1n01,a1n02,a1n03,a1n04,a1n05)" & _
            " values('" & txtA0K01.Tag & "','2','" & .Fields("a1n03") & "'" & _
            ",'" & .Fields("a1n04") & "'," & .Fields("a1n05") & ")"
         cnnConnection.Execute strSql, intI
         .MoveNext
      Loop
   End If
   End With
   
   With adoacc1n0_1
   If .RecordCount > 0 Then
      .MoveFirst
      Do While Not .EOF
         strSql = "insert into acc1n0(a1n01,a1n02,a1n03,a1n04,a1n05)" & _
            " values('" & txtA0K01.Tag & "','1','" & .Fields("a1n03") & "'" & _
            ",'" & .Fields("a1n04") & "'," & .Fields("a1n05") & ")"
         cnnConnection.Execute strSql, intI
         .MoveNext
      Loop
   End If
   End With
   
   cnnConnection.CommitTrans
   FormSave = True
   Exit Function
   
ErrHnd:
   cnnConnection.RollbackTrans
   MsgBox Err.Description
   
End Function
Public Sub Command10_Click()
If m_bolQuery = False Then
   If Val(txtSum) <> Val(txtSum_1) Then
      MsgBox "專業點數合計與業務點數合計不符！"
      If txtA0K01.Tag <> txtA0K01.Text Then txtA0K01.Text = txtA0K01.Tag
      SSTab1.Tab = 0
      Exit Sub
   ElseIf Val(txtPts) > 0 And Val(txtPts) <> Val(txtSum) Then
      If MsgBox("輸入點數與預計點數不符，是否要繼續輸入？", vbYesNo + vbDefaultButton2) = vbYes Then
         If txtA0K01.Tag <> txtA0K01.Text Then txtA0K01.Text = txtA0K01.Tag
         SSTab1.Tab = 0
         Exit Sub
      End If
   Else
      If Val(txtPts) > 0 And Val(txtSum) > 0 Then
         FormSave
      End If
   End If
End If
txtA0K01.Tag = txtA0K01.Text
OpenTable '收據號碼(重新查詢,載入基本資料)

End Sub
'*************************************************
'  開啟資料表
'*************************************************
Private Function OpenTable() As Boolean
Dim amt1 As Double
On Error GoTo Checking
   
   strExc(0) = "select a1.*,c1.cp01,cp02,cp03,cp04,cp13,nvl(a1u07,0) a1u07 from acc0k0 a1,caseprogress c1,acc1u0 " & _
               "where a0k01='" & txtA0K01 & "' and a0k01=cp60(+) and cp60=a1u02(+) and cp09=a1u03(+)"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      RsTemp.MoveFirst
      txtCP01.Text = "" & RsTemp.Fields("cp01")
      txtCP02.Text = "" & RsTemp.Fields("cp02")
      txtCP03.Text = "" & RsTemp.Fields("cp03")
      txtCP04.Text = "" & RsTemp.Fields("cp04")
      amt1 = Val("" & RsTemp.Fields("a0k06")) - RsTemp.Fields("a1u07")
      '減收據有財務處銷帳或銷退後的點數
      Do While Not RsTemp.EOF
         If RsTemp.AbsolutePosition > 1 Then
            amt1 = amt1 - RsTemp.Fields("a1u07")
         End If
         RsTemp.MoveNext
      Loop
      txtPts = Format(amt1 / 1000, "###0.000")
   Else
      MsgBox "資料庫查無資料!!"
      Exit Function
   End If
   '2:專業點數
   strExc(0) = "select st02,a1n02,a1n03,decode(cpm03,'（無）',cpm04,cpm03) cpm03,a1n04,a1n05,a1n06" & _
      " from acc1n0,staff,caseprogress,casepropertymap where a1n01='" & txtA0K01 & "' and a1n02='2'" & _
      " and cp09(+)=a1n03 and cpm01(+)=cp01 and cpm02(+)=cp10" & _
      " and st01(+)=a1n04 order by a1n04,a1n03"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   '改暫存TB
   Set adoacc1n0 = PUB_CreateRecordset(RsTemp, , , , Me.Name)
   Set RsTemp = Nothing
   Set Adodc1.Recordset = adoacc1n0
   
   Set DataGrid1.DataSource = Adodc1
   DataGrid1.Refresh
   DataGrid1.col = 0
   DataGrid1.CurrentCellVisible = True
   SumShow
   AdodcClear
   '1:業務點數
   strExc(0) = "select st02,a1n02,a1n03,decode(cpm03,'（無）',cpm04,cpm03) cpm03,a1n04,a1n05,a1n06" & _
      " from acc1n0,staff,caseprogress,casepropertymap where a1n01='" & txtA0K01 & "' and a1n02='1'" & _
      " and cp09(+)=a1n03 and cpm01(+)=cp01 and cpm02(+)=cp10 and st01(+)=a1n04 order by a1n04,a1n03"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   '改暫存TB
   Set adoacc1n0_1 = PUB_CreateRecordset(RsTemp, , , , Me.Name)
   Set RsTemp = Nothing
   Set Adodc2.Recordset = adoacc1n0_1
   
   Set DataGrid2.DataSource = Adodc2
   DataGrid2.Refresh
   DataGrid2.col = 0
   DataGrid2.CurrentCellVisible = True
   SumShow_1
   AdodcClear_1
   
   OpenTable = True
   Exit Function
   
Checking:
   If Err.Number <> 0 Then
      MsgBox Err.Description, , MsgText(5)
   End If
   
End Function
'*************************************************
'  顯示 Adodc 之資料
'
'*************************************************
Private Sub AdodcShow()
   With adoacc1n0
   txtA1N04 = .Fields("a1n04").Value
   txtA1N05 = Round(.Fields("a1n05").Value, 3)
   txtA1N03 = "" & .Fields("a1n03").Value
   txtST02 = "" & .Fields("st02").Value
   End With
   m_bolAddNew = False
End Sub



Private Sub SSTab1_Click(PreviousTab As Integer)
   If PreviousTab = 0 Then
      txtA1N04_1.TabStop = True
      txtA1N05_1.TabStop = True
      txtA1N03_1.TabStop = True
      Command6.TabStop = True
      Command4.TabStop = True
   End If
End Sub

Private Sub txtA0K01_GotFocus()
   TextInverse txtA0K01
End Sub

Private Sub txtA0K01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtA1N03_GotFocus()
   TextInverse txtA1N03
End Sub

Private Sub txtA1N03_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtA1N03_Validate(Cancel As Boolean)
   m_A1N03_CPM03 = ""
   If Trim(txtA1N03) <> "" Then
      strExc(0) = "select decode(cpm03,'（無）',cpm04,cpm03) cpm03 from caseprogress,casepropertymap" & _
         " where cp60='" & txtA0K01 & "' and cp09='" & txtA1N03 & "' and cpm01(+)=cp01 and cpm02(+)=cp10"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         m_A1N03_CPM03 = "" & RsTemp.Fields(0)
      Else
         MsgBox "收文號輸入錯誤！"
         Cancel = True
         Exit Sub
      End If
   End If
End Sub

Private Sub txtA1N04_GotFocus()
   TextInverse txtA1N04
End Sub

Private Sub txtA1N04_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtA1N04_Validate(Cancel As Boolean)
   txtST02 = ""
   If txtA1N04 <> "" Then
      txtST02 = StaffQuery(txtA1N04)
      If txtST02 = "" Then
         MsgBox "承辦人輸入錯誤！"
         Cancel = True
         txtA1N04_GotFocus
         Exit Sub
      End If
      strExc(1) = PUB_GetStaffST15(txtA1N04, 1)
      If strExc(1) = "F51" Or strExc(1) = "F52" Then
         MsgBox "不可輸入外翻編號！"
         Cancel = True
         txtA1N04_GotFocus
         Exit Sub
      End If
      
      If chkA0910(strExc(1)) = False Then
         MsgBox "承辦人作帳部門未設定或無法讀取！"
         Cancel = True
         txtA1N04_GotFocus
         Exit Sub
      End If
      
      '收文號只有一個時預設
      If Cancel = False And Trim(txtA1N03) = "" Then
         If txtA1N04 < "F" Then
            strExc(0) = "SELECT DISTINCT CP09 FROM CASEPROGRESS WHERE CP60='" & txtA0K01 & "'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               If RsTemp.RecordCount = 1 Then
                  txtA1N03 = RsTemp(0)
               End If
            End If
         End If
      End If
   End If
End Sub

Private Sub txtA1N05_GotFocus()
   TextInverse txtA1N05
End Sub

Private Sub SumShow()
   Dim ii As Integer
   txtSum = 0
   Set RsTemp = Adodc1.Recordset.Clone
   With RsTemp
   If .RecordCount > 0 Then
      .MoveFirst
      Do While Not .EOF
         txtSum = Val(txtSum) + Val("" & .Fields("a1n05"))
         .MoveNext
      Loop
   End If
   End With
End Sub

Private Sub SumShow_1()
   Dim ii As Integer
   txtSum_1 = 0
   Set RsTemp = Adodc2.Recordset.Clone
   With RsTemp
   If .RecordCount > 0 Then
      .MoveFirst
      Do While Not .EOF
         txtSum_1 = Val(txtSum_1) + Val("" & .Fields("a1n05"))
         .MoveNext
      Loop
   End If
   End With
End Sub
'*************************************************
'  刪除 Adodc 之資料
'
'*************************************************
Private Sub AdodcDelete(adoRst As ADODB.Recordset)
   With adoRst
   If Not (.EOF Or .BOF) Then
      .Delete
      '.UpdateBatch
      .UPDATE
   End If
   End With
End Sub

'*************************************************
'  清除查詢顯示
'
'*************************************************
Public Sub AdodcClear()
   txtA1N04 = ""
   txtST02 = ""
   txtA1N05 = ""
   txtA1N03 = " " '預設空白
   m_bolAddNew = True
   If txtA1N04.Enabled And txtA1N04.Visible Then txtA1N04.SetFocus
End Sub

Private Sub AdodcAdd()
   Dim bolAdd As Boolean
   bolAdd = True
   With adoacc1n0
   
   If .RecordCount > 0 Then
      .Sort = "a1n04,a1n03"
      .MoveFirst
      .Find "a1n04='" & txtA1N04 & "'"
      If Not .EOF Then
         .Find "a1n03='" & txtA1N03 & "'"
         If Not .EOF Then
            If txtA1N04 = .Fields("a1n04") Then
               bolAdd = False
               If MsgBox("資料已存在，是否要更新！", vbYesNo + vbDefaultButton2) = vbNo Then
                  GoTo eXitPort
               End If
            End If
         End If
      End If
   End If
   If bolAdd Then .AddNew
   .Fields("a1n04").Value = txtA1N04
   .Fields("a1n05").Value = Val(txtA1N05)
   .Fields("a1n03").Value = txtA1N03
   .Fields("st02").Value = txtST02
   .Fields("cpm03").Value = m_A1N03_CPM03
   '.UpdateBatch
   .UPDATE
   .Sort = "a1n04,a1n03"
   AdodcClear
   SumShow
   
eXitPort:
   End With
   
End Sub

Private Sub AdodcUpdate()
   Dim iPos As Integer
   
   With adoacc1n0
   iPos = .AbsolutePosition
   .MoveFirst
   .Find "a1n04='" & txtA1N04 & "'"
   If Not .EOF Then
      .Find "a1n03='" & txtA1N03 & "'"
      If Not .EOF Then
         If txtA1N04 = .Fields("a1n04") And iPos <> .AbsolutePosition Then
            MsgBox "承辦人+收文號資料重複，請重新輸入！"
            If iPos = 1 Then
               .MoveFirst
            Else
               .Move iPos - 1, 1
            End If
            GoTo eXitPort
         End If
      End If
   End If
   If iPos = 1 Then
      .MoveFirst
   Else
      .Move iPos - 1, 1
   End If
   .Fields("a1n04").Value = txtA1N04
   .Fields("a1n05").Value = Val(txtA1N05)
   .Fields("a1n03").Value = txtA1N03
   .Fields("st02").Value = txtST02
   .Fields("cpm03").Value = m_A1N03_CPM03
   '.UpdateBatch
   .UPDATE
   AdodcClear
   SumShow
   
eXitPort:
   End With
   
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)
   Select Case KeyCode
      Case vbKeyInsert '新增記錄按Insert鍵
         If SSTab1.Tab = 0 Then
            If TxtValidate Then
               If m_bolAddNew Then
                  AdodcAdd
               Else
                  AdodcUpdate
               End If
            End If
         Else
            If TxtValidate_1 Then
               If m_bolAddNew_1 Then
                  AdodcAdd_1
               Else
                  AdodcUpdate_1
               End If
            End If
         End If
   End Select
   KeyEnter KeyCode
End Sub

Private Function TxtValidate() As Boolean
   Dim bCancel As Boolean
   If txtA1N04 = "" Then
      MsgBox "承辦人不可空白！"
      txtA1N04.SetFocus
      Exit Function
   Else
      txtA1N04_Validate bCancel
      If bCancel Then
         txtA1N04_GotFocus
         txtA1N04.SetFocus
         Exit Function
      End If
   End If
   
   If txtA1N04 < "F" And Trim(txtA1N03) = "" Then
      MsgBox "承辦人非部門時收文號不可空白！"
      txtA1N03_GotFocus
      txtA1N03.SetFocus
      Exit Function
      
   ElseIf txtA1N04 > "F" And Trim(txtA1N03) <> "" Then
      MsgBox "承辦人為部門時不可輸入收文號！"
      txtA1N03_GotFocus
      txtA1N03.SetFocus
      Exit Function
      
   End If
   
   If (txtCP01 = "FCL" Or txtCP01 = "LIN" Or txtCP01 = "CFL") And txtA1N04 = "97009" Then
      MsgBox "FCL,LIN,CFL案件承辦人不可用 97009 編號！"
      txtA1N04_GotFocus
      txtA1N04.SetFocus
      Exit Function
   End If
   
   txtA1N03_Validate bCancel
   If bCancel Then
      txtA1N03_GotFocus
      txtA1N03.SetFocus
      Exit Function
   End If

   If Val(txtA1N05) = 0 Then
      MsgBox "點數必須大於 0 ！", vbExclamation
      txtA1N05.SetFocus
      Exit Function
   End If
   
   TxtValidate = True
End Function

Private Sub txtA1N04_1_GotFocus()
   TextInverse txtA1N04_1
End Sub

Private Sub txtA1N04_1_Validate(Cancel As Boolean)
   txtST02_1 = ""
   If txtA1N04_1 <> "" Then
      txtST02_1 = StaffQuery(txtA1N04_1)
      If txtST02_1 = "" Then
         MsgBox "智權人員輸入錯誤！"
         Cancel = True
      End If
      '收文號只有一個時預設
      If Cancel = False And Trim(txtA1N03) = "" Then
         If txtA1N04_1 < "F" Then
            strExc(0) = "SELECT DISTINCT CP09 FROM CASEPROGRESS WHERE CP60='" & txtA0K01 & "'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               If RsTemp.RecordCount = 1 Then
                  txtA1N03_1 = RsTemp(0)
               End If
            End If
         End If
      End If
   End If
End Sub

Private Sub txtA1N05_1_GotFocus()
   TextInverse txtA1N05_1
End Sub

Private Sub txtA1N03_1_GotFocus()
   TextInverse txtA1N03_1
End Sub

Private Sub txtA1N03_1_Validate(Cancel As Boolean)
   m_A1N03_CPM03_1 = ""
   If Trim(txtA1N03_1) <> "" Then
      strExc(0) = "select decode(cpm03,'（無）',cpm04,cpm03) cpm03 from caseprogress,casepropertymap" & _
         " where cp60='" & txtA0K01 & "' and cp09='" & txtA1N03_1 & "' and cpm01(+)=cp01 and cpm02(+)=cp10"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         m_A1N03_CPM03_1 = "" & RsTemp.Fields(0)
      Else
         MsgBox "收文號輸入錯誤！"
         Cancel = True
         Exit Sub
      End If
   End If
End Sub

'*************************************************
'  清除查詢顯示
'
'*************************************************
Public Sub AdodcClear_1()
   txtA1N04_1 = ""
   txtST02_1 = ""
   txtA1N05_1 = ""
   txtA1N03_1 = " " '預設空白
   m_bolAddNew_1 = True
   If txtA1N04_1.Enabled And txtA1N04_1.Visible Then txtA1N04_1.SetFocus
End Sub

Private Function TxtValidate_1() As Boolean
   Dim bCancel As Boolean
   If txtA1N04_1 = "" Then
      MsgBox "智權人員不可空白！"
      txtA1N04_1.SetFocus
      Exit Function
   Else
      txtA1N04_1_Validate bCancel
      If bCancel Then
         txtA1N04_1.SetFocus
         Exit Function
      End If
   End If
   
   If Val(txtA1N05_1) = 0 Then
      MsgBox "點數必須大於 0 ！", vbExclamation
      txtA1N05_1.SetFocus
      Exit Function
   End If
   
   If Trim(txtA1N03_1) = "" Then
      txtA1N03_1 = " "
   End If
   txtA1N03_1_Validate bCancel
   If bCancel Then
      txtA1N03_1.SetFocus
      Exit Function
   End If
   
   If txtA1N04_1 < "F" And Trim(txtA1N03_1) = "" Then
      MsgBox "智權人員非部門時收文號不可空白！"
      txtA1N03_1_GotFocus
      txtA1N03_1.SetFocus
      Exit Function
   End If
   TxtValidate_1 = True
End Function

Private Sub AdodcAdd_1()
   Dim bolAdd As Boolean
   bolAdd = True
   With adoacc1n0_1
   If Not (.EOF And .BOF) And adoacc1n0_1.RecordCount > 0 Then
      .Sort = "a1n04,a1n03"
      .MoveFirst
      .Find "a1n04='" & txtA1N04_1 & "'"
      If Not .EOF Then
         .Find "a1n03='" & txtA1N03_1 & "'"
         If Not .EOF Then
            bolAdd = False
            If MsgBox("資料已存在，是否要更新！", vbYesNo + vbDefaultButton2) = vbNo Then
               GoTo eXitPort
            End If
         End If
      End If
   End If
   If bolAdd Then .AddNew
   .Fields("a1n04").Value = txtA1N04_1
   .Fields("a1n05").Value = Val(txtA1N05_1)
   .Fields("a1n03").Value = txtA1N03_1
   .Fields("st02").Value = txtST02_1
   .Fields("cpm03").Value = m_A1N03_CPM03_1
   '.UpdateBatch
   .UPDATE
   .Sort = "a1n04,a1n03"
   AdodcClear_1
   SumShow_1
   
eXitPort:
   End With
   
End Sub

Private Sub AdodcUpdate_1()
   Dim iPos As Integer
   
   With adoacc1n0_1
   iPos = .AbsolutePosition
   .MoveFirst
   .Find "a1n04='" & txtA1N04_1 & "'"
   If Not .EOF Then
      .Find "a1n03='" & txtA1N03_1 & "'"
      If Not .EOF Then
         If txtA1N04_1 = .Fields("a1n04") And iPos <> .AbsolutePosition Then
            MsgBox "承辦人+收文號資料重複，請重新輸入！"
            If iPos = 1 Then
               .MoveFirst
            Else
               .Move iPos - 1, 1
            End If
            GoTo eXitPort
         End If
      End If
   End If
   If iPos = 1 Then
      .MoveFirst
   Else
      .Move iPos - 1, 1
   End If
   .Fields("a1n04").Value = txtA1N04_1
   .Fields("a1n05").Value = Val(txtA1N05_1)
   .Fields("a1n03").Value = txtA1N03_1
   .Fields("st02").Value = txtST02_1
   .Fields("cpm03").Value = m_A1N03_CPM03_1
   '.UpdateBatch
   .UPDATE
   AdodcClear_1
   SumShow_1
   
eXitPort:
   End With
   
End Sub

'*************************************************
'  顯示 Adodc 之資料
'
'*************************************************
Private Sub AdodcShow_1()
   With adoacc1n0_1
   txtA1N04_1 = .Fields("a1n04").Value
   txtA1N05_1 = Round(.Fields("a1n05").Value, 3)
   txtA1N03_1 = "" & .Fields("a1n03").Value
   txtST02_1 = "" & .Fields("st02").Value
   End With
   m_bolAddNew_1 = False
End Sub

Private Function chkA0910(p_A0901 As String) As Boolean
   Dim stSQL As String, intR As Integer
   Dim adoRst As ADODB.Recordset
   stSQL = "select a0910 from acc090 where a0901='" & p_A0901 & "'"
   intR = 1
   Set adoRst = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      If Not IsNull(adoRst(0)) Then
         chkA0910 = True
      End If
   End If
   Set adoRst = Nothing
End Function

'依照預設規則分配點數
Private Function Get_PointAutoassign(strA0K01 As String) As Boolean
   Dim stSQL As String, intR As Integer, ii As Integer
   Dim adoRst As ADODB.Recordset, adoRst2 As ADODB.Recordset
   Dim douPtSum As Double '總點數
   Dim douPt As Double
   Dim bolFMP As Boolean '是否FMP案
   Dim bolFMPnewcase As Boolean '是否FMP新案請款(要扣安全基金)
   Dim bolOurFMP As Boolean '新案是否寰華案件
   Dim bolTcase As Boolean '是否內商人員承辦案件
On Error GoTo ErrHnd
   

      stSQL = "select a0k01,sum(a0k06) damt1,sum(nvl(a1u07,0)) damt2 from acc0k0,acc1u0 " & _
              "where A0K01='" & strA0K01 & "' and a0k01=a1u02(+) and substr(a1u01,1,1)='I' group by a0k01"
      intR = 1
      Set adoRst = ClsLawReadRstMsg(intR, stSQL)
      If intR = 1 Then 'strA0k01
        cnnConnection.BeginTrans
          cnnConnection.Execute "delete acc1n0 where a1n01='" & strA0K01 & "'"
         '小數一律捨去
         douPtSum = Fix(adoRst.Fields("damt1") - adoRst.Fields("damt2")) / 1000 '減收據有財務處銷帳或銷退後的點數
         douPt = douPtSum
         If douPtSum > 0 Then
            '智權人員點數放在最後收文的智權人員
            stSQL = "select cp01,cp02,cp03,cp04,cp12,cp13,cp09,st15,pa09 from caseprogress,staff,patent " & _
                    "where cp60='" & strA0K01 & "' and st01(+)=cp14 and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 order by cp05,cp09"
            intR = 1
            Set adoRst = ClsLawReadRstMsg(intR, stSQL)
            If intR = 1 Then
               With adoRst
                 .MoveLast
                 If .Fields("cp01") = "P" And Left(.Fields("cp12"), 1) = "F" And "" & .Fields("pa09") <> "000" Then
                   bolFMP = True
                   bolOurFMP = PUB_FMPtoCheck(1, 2, Pub_strUserST05, .Fields("cp01"), .Fields("cp02"), .Fields("cp03"), .Fields("cp04"))  'Added by Morgan 2014/11/6
                 End If
                 '業務點數
                 stSQL = "insert into ACC1N0(a1n01,a1n02,a1n03,a1n04,a1n05)" & _
                       " values ('" & strA0K01 & "','1','" & .Fields("cp09") & "','" & .Fields("cp13") & "'," & douPtSum & ")"
                 cnnConnection.Execute stSQL, intR
                 
                 If .Fields("cp01") = "FCT" Or .Fields("cp01") = "T" Then
                   Do While Not .EOF
                      If Left("" & .Fields("st15"), 2) = "P2" Then
                         bolTcase = True
                         Exit Do
                      End If
                      .MoveNext
                   Loop
                 End If
               End With
            End If
               
            '專業點數
            'FCT,FMT的請款單只要有一個承辦人是內商人員則整張請款單專業點數都歸內商--陳鳳英
            If bolTcase Then
               stSQL = "insert into ACC1N0(a1n01,a1n02,a1n03,a1n04,a1n05)" & _
                     " values ('" & strA0K01 & "','2',' ','P2001'," & douPtSum & ")"
               cnnConnection.Execute stSQL, intR
            Else
                stSQL = "select cp09,cp01,cp02,cp03,cp04,cp14,nvl(a0j09,0) a0j09,st15,nvl(a1u07,0) a1u07 " & _
                        "from caseprogress,acc0j0,staff,(select * from acc1u0 where substr(a1u01,1,1)='I' and a1u02='" & strA0K01 & "') " & _
                        "where cp60='" & strA0K01 & "' and cp09=a0j01(+) and cp60=a0j13(+) and cp14=st01(+) and a1u02(+)=cp60 and a1u03(+)=cp09"
                intR = 1
                Set adoRst = ClsLawReadRstMsg(intR, stSQL)
                If intR = 1 Then
                   adoRst.MoveFirst
                   Do While Not adoRst.EOF
                        douPt = (adoRst.Fields("a0j09") - adoRst.Fields("a1u07")) / 1000

                        stSQL = "insert into ACC1N0(a1n01,a1n02,a1n03,a1n04,a1n05)" & _
                              " values ('" & strA0K01 & "','2','" & adoRst.Fields("cp09") & "','" & adoRst.Fields("cp14") & "'," & douPt & ")"
                        cnnConnection.Execute stSQL, intR
                        adoRst.MoveNext
                   Loop
                End If
            End If
         End If
        cnnConnection.CommitTrans
      End If 'strA0k01
   
   GoTo eXitPort

ErrHnd:
      cnnConnection.RollbackTrans
      MsgBox Err.Description
      
eXitPort:
   Set adoRst = Nothing
   Set adoRst2 = Nothing
   
End Function


Private Sub txtcp01_Change()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse

   m_bolQuery = True
   If IsEmptyText(txtCP01) = False Then
      ' 檢查系統類別
      If IsCorrectSysKind(txtCP01) = False Then
         strTit = "資料檢核"
         strMsg = "本所案號中的系統別不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         txtA0K01_GotFocus
         GoTo FrmChange
      End If
      ' 檢查使用者權限
      If IsUserHasRightOfSystem(strUserNum, txtCP01) = False Then
         strTit = "資料檢核"
         strMsg = "您沒有使用該系統類別的權限"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         txtA0K01_GotFocus
         GoTo FrmChange
      End If
   End If
   
   m_bolQuery = False
   
FrmChange:
   If m_bolQuery Then
      SetFormEnable False
   Else
      SetFormEnable True
   End If
   
End Sub

'Private Sub txtcp01_Validate(Cancel As Boolean)
'   Dim strTit As String
'   Dim strMsg As String
'   Dim nResponse
'
'   Cancel = False
'   If IsEmptyText(txtCP01) = False Then
'      ' 檢查系統類別
'      If IsCorrectSysKind(txtCP01) = False Then
'         Cancel = True
'         strTit = "資料檢核"
'         strMsg = "本所案號中的系統別不正確"
'         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'         txtA0K01_GotFocus
'         GoTo EXITSUB
'      End If
'      ' 檢查使用者權限
'      If IsUserHasRightOfSystem(strUserNum, txtCP01) = False Then
'         Cancel = True
'         strTit = "資料檢核"
'         strMsg = "您沒有使用該系統類別的權限"
'         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'         txtA0K01_GotFocus
'         GoTo EXITSUB
'      End If
'   End If
'
'End Sub
