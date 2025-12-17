VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc21h3 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "點數分配輸入"
   ClientHeight    =   4836
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8784
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4836
   ScaleWidth      =   8784
   Begin VB.CommandButton Command9 
      Caption         =   "點數重新分配"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6930
      TabIndex        =   38
      Top             =   540
      Width           =   1770
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4155
      Left            =   90
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   630
      Width           =   8610
      _ExtentX        =   15177
      _ExtentY        =   7324
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "專業點數"
      TabPicture(0)   =   "Frmacc21h3.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label9"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label8"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label6"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label13"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label14"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Shape1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtST02"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "DataGrid1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Adodc1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtSum"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Command5"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtA1N06"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtA1N03"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtA1N04"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtA1N05"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Command1"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Command2"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Command7"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).ControlCount=   18
      TabCaption(1)   =   "業務點數"
      TabPicture(1)   =   "Frmacc21h3.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command6"
      Tab(1).Control(1)=   "Command4"
      Tab(1).Control(2)=   "txtA1N05_1"
      Tab(1).Control(3)=   "txtA1N04_1"
      Tab(1).Control(4)=   "txtA1N03_1"
      Tab(1).Control(5)=   "Command3"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "txtSum_1"
      Tab(1).Control(7)=   "Adodc2"
      Tab(1).Control(8)=   "DataGrid2"
      Tab(1).Control(9)=   "txtST02_1"
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
            Size            =   11.4
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6075
         TabIndex        =   36
         Top             =   2670
         Width           =   2400
      End
      Begin VB.CommandButton Command6 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.8
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   -67875
         Picture         =   "Frmacc21h3.frx":0038
         Style           =   1  '圖片外觀
         TabIndex        =   11
         ToolTipText     =   "清除畫面"
         Top             =   3315
         Width           =   550
      End
      Begin VB.CommandButton Command4 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.8
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   -67275
         Picture         =   "Frmacc21h3.frx":0902
         Style           =   1  '圖片外觀
         TabIndex        =   12
         ToolTipText     =   "取消"
         Top             =   3315
         Width           =   550
      End
      Begin VB.TextBox txtA1N05_1 
         Alignment       =   1  '靠右對齊
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.8
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -72330
         MaxLength       =   8
         TabIndex        =   8
         Top             =   3495
         Width           =   945
      End
      Begin VB.TextBox txtA1N04_1 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.8
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -74730
         MaxLength       =   6
         TabIndex        =   7
         Top             =   3495
         Width           =   972
      End
      Begin VB.TextBox txtA1N03_1 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.8
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -71280
         TabIndex        =   9
         Top             =   3495
         Width           =   1125
      End
      Begin VB.CommandButton Command3 
         Height          =   300
         Left            =   -70155
         Picture         =   "Frmacc21h3.frx":0F6C
         Style           =   1  '圖片外觀
         TabIndex        =   10
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
            Size            =   10.8
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -72480
         TabIndex        =   31
         Top             =   2730
         Width           =   1005
      End
      Begin VB.CommandButton Command2 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.8
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   7125
         Picture         =   "Frmacc21h3.frx":106E
         Style           =   1  '圖片外觀
         TabIndex        =   5
         ToolTipText     =   "清除畫面"
         Top             =   3315
         Width           =   550
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.8
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   7725
         Picture         =   "Frmacc21h3.frx":1938
         Style           =   1  '圖片外觀
         TabIndex        =   6
         ToolTipText     =   "取消"
         Top             =   3315
         Width           =   550
      End
      Begin VB.TextBox txtA1N05 
         Alignment       =   1  '靠右對齊
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.8
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2670
         MaxLength       =   8
         TabIndex        =   1
         Top             =   3495
         Width           =   945
      End
      Begin VB.TextBox txtA1N04 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.8
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   270
         MaxLength       =   6
         TabIndex        =   0
         Top             =   3495
         Width           =   972
      End
      Begin VB.TextBox txtA1N03 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.8
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3720
         TabIndex        =   2
         Top             =   3495
         Width           =   1125
      End
      Begin VB.TextBox txtA1N06 
         Alignment       =   1  '靠右對齊
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.8
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5250
         MaxLength       =   1
         TabIndex        =   4
         Top             =   3495
         Width           =   1575
      End
      Begin VB.CommandButton Command5 
         Height          =   300
         Left            =   4845
         Picture         =   "Frmacc21h3.frx":1FA2
         Style           =   1  '圖片外觀
         TabIndex        =   3
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
            Size            =   10.8
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2520
         TabIndex        =   24
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
         _ExtentY        =   572
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
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   450
         Width           =   8295
         _ExtentX        =   14626
         _ExtentY        =   3789
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   0   'False
         BackColor       =   -2147483624
         HeadLines       =   1
         RowHeight       =   16
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體-ExtB"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   6
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
         BeginProperty Column05 
            DataField       =   "a1n06"
            Caption         =   "是否為核稿點數"
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
               ColumnWidth     =   1175.811
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1031.811
            EndProperty
            BeginProperty Column02 
               Alignment       =   1
               ColumnWidth     =   720
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1272.189
            EndProperty
            BeginProperty Column04 
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               ColumnWidth     =   1811.906
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
         _ExtentY        =   572
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
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   450
         Width           =   8295
         _ExtentX        =   14626
         _ExtentY        =   3789
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   0   'False
         BackColor       =   -2147483624
         HeadLines       =   1
         RowHeight       =   16
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體-ExtB"
            Size            =   9.6
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
               ColumnWidth     =   1175.811
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1031.811
            EndProperty
            BeginProperty Column02 
               Alignment       =   1
               ColumnWidth     =   720
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1272.189
            EndProperty
            BeginProperty Column04 
            EndProperty
         EndProperty
      End
      Begin MSForms.TextBox txtST02_1 
         Height          =   330
         Left            =   -73740
         TabIndex        =   40
         Top             =   3495
         Width           =   1335
         VariousPropertyBits=   671105055
         BackColor       =   14737632
         Size            =   "2355;582"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   225
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtST02 
         Height          =   330
         Left            =   1260
         TabIndex        =   39
         Top             =   3495
         Width           =   1335
         VariousPropertyBits=   671105055
         BackColor       =   14737632
         Size            =   "2355;582"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   225
         FontCharSet     =   136
         FontPitchAndFamily=   34
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
            Size            =   10.8
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -72330
         TabIndex        =   35
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
            Size            =   10.8
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -74040
         TabIndex        =   34
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
            Size            =   10.8
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -71025
         TabIndex        =   33
         Top             =   3240
         Width           =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "點數合計"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.8
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -73515
         TabIndex        =   32
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
            Size            =   10.8
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2670
         TabIndex        =   29
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
            Size            =   10.8
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1065
         TabIndex        =   28
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
            Size            =   10.8
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3975
         TabIndex        =   27
         Top             =   3240
         Width           =   705
      End
      Begin VB.Label Label8 
         Alignment       =   2  '置中對齊
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "是否為核稿點數"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.8
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   5235
         TabIndex        =   26
         Top             =   3240
         Width           =   1605
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "點數合計"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.8
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1485
         TabIndex        =   25
         Top             =   2745
         Width           =   900
      End
   End
   Begin VB.CommandButton Command8 
      Caption         =   "確定"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7605
      TabIndex        =   37
      Top             =   90
      Width           =   1095
   End
   Begin VB.TextBox txtPts 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6300
      TabIndex        =   20
      Top             =   150
      Width           =   1212
   End
   Begin VB.TextBox txtA1K01 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1125
      MaxLength       =   15
      TabIndex        =   17
      Top             =   120
      Width           =   1395
   End
   Begin VB.TextBox txtA1K13 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3690
      TabIndex        =   16
      Top             =   150
      Width           =   492
   End
   Begin VB.TextBox txtA1K14 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4170
      TabIndex        =   15
      Top             =   150
      Width           =   852
   End
   Begin VB.TextBox txtA1K15 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5010
      TabIndex        =   14
      Top             =   150
      Width           =   252
   End
   Begin VB.TextBox txtA1K16 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5250
      TabIndex        =   13
      Top             =   150
      Width           =   372
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "點數"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5820
      TabIndex        =   21
      Top             =   150
      Width           =   465
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "請款編號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   165
      TabIndex        =   19
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "本所案號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2730
      TabIndex        =   18
      Top             =   150
      Width           =   975
   End
End
Attribute VB_Name = "Frmacc21h3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/08 改成Form2.0 ; DataGrid1改字型=新細明體-ExtB、txtST02、txtST02_1
'Memo By Sonia 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit

Dim adoacc1n0 As ADODB.Recordset
Dim m_bolAddNew As Boolean
Dim m_A1N03_CPM03 As String '收文號的案件性質

Dim adoacc1n0_1 As ADODB.Recordset
Dim m_bolAddNew_1 As Boolean
Dim m_A1N03_CPM03_1 As String '收文號的案件性質
Public m_bolQuery As Boolean '是否為查詢

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
   'Modified by Morgan 2022/5/2
   strExc(0) = "select '',cp09,sqldatet(cp05) cp05,decode(cpm03,'（無）',cpm04,cpm03) cpm03,st02 from caseprogress,casepropertymap,staff where cp60='" & txtA1K01 & "'" & _
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
   strExc(0) = "select '',cp09,sqldatet(cp05) cp05,decode(cpm03,'（無）',cpm04,cpm03) cpm03,st02,cp14 from caseprogress,casepropertymap,staff where cp60='" & txtA1K01 & "'" & _
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
         'Modified by Morgan 2022/5/2
         'txtA1N03.SetFocus
         RsTemp.MoveFirst
         RsTemp.Find " cp09='" & txtA1N03 & "'"
         If Not RsTemp.EOF Then
            txtA1N04 = RsTemp.Fields("cp14")
            txtA1N04_Validate False
         End If
         txtA1N05.SetFocus
         'end 2022/5/2
      End If
   End If
End Sub

Private Sub Command6_Click()
   AdodcClear_1
   txtA1N04_1.SetFocus
End Sub

Private Sub Command7_Click()
   Set RsTemp = adoacc1n0.Clone
   'Modify by Amy 2014/06/26 +FormName 改暫存TB
   Set adoacc1n0_1 = PUB_CreateRecordset(RsTemp, , , , Me.Name)
   adoacc1n0_1.Sort = "a1n04,a1n03"
   Set RsTemp = Nothing
   Set Adodc2.Recordset = adoacc1n0_1
   Set DataGrid2.DataSource = Adodc2
   DataGrid2.Refresh
   AdodcClear_1
   SumShow_1
   MsgBox "複製完成！"
   SSTab1.Tab = 1
End Sub

Private Sub Command8_Click()
   If m_bolQuery Then
      Unload Me
   Else
      If Val(txtSum) <> Val(txtPts) Then
         MsgBox "專業點數合計與請款點數不符！"
         SSTab1.Tab = 0
         
      ElseIf Val(txtSum_1) <> Val(txtPts) Then
         MsgBox "業務點數合計與請款點數不符！"
         SSTab1.Tab = 1
      
      Else
         'Added by Morgan 2016/7/22
         If adoacc1n0.RecordCount > 0 Then
            adoacc1n0.MoveFirst
            adoacc1n0.Find "a1n05=0"
            If Not adoacc1n0.EOF Then
               SSTab1.Tab = 0
               MsgBox adoacc1n0("st02") & "有分配點數為 0，若確定無需分配請刪除！", vbExclamation
               Exit Sub
            End If
         End If
         
         If adoacc1n0_1.RecordCount > 0 Then
            adoacc1n0_1.MoveFirst
            adoacc1n0_1.Find "a1n05=0"
            If Not adoacc1n0_1.EOF Then
               SSTab1.Tab = 1
               MsgBox adoacc1n0_1("st02") & "有分配點數為 0，若確定無需分配請刪除！", vbExclamation
               Exit Sub
            End If
         End If
         'end 2016/7/22
         
         If FormSave Then
            Unload Me
         End If
      
      End If
   End If
End Sub
'Add by Morgan 2010/6/17
Private Sub Command9_Click()
   If MsgBox("系統將清除目前分配並依照規則重新分配，是否確定要繼續？", vbYesNo + vbDefaultButton2) = vbYes Then
      'Modify by Morgan 2010/6/24
      PUB_PointAutoassign txtA1K01
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

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   If Not m_bolQuery Then
      KeyDefine KeyCode
   End If
End Sub

Private Sub Form_Load()
   PUB_InitForm Me, Me.Width, Me.Height
   txtA1K01 = strItemNo
   OpenTable
   If m_bolQuery Then
      SetFormEnable False
   Else
      SetFormEnable True
      'tool3_enabled
   End If
   Screen.MousePointer = vbDefault
   tool4_enabled
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
   Command8.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set Frmacc21h3 = Nothing
End Sub

Private Function FormSave() As Boolean
   cnnConnection.BeginTrans
On Error GoTo ErrHnd
   strSql = "delete acc1n0 where a1n01='" & txtA1K01 & "'"
   cnnConnection.Execute strSql, intI
   With adoacc1n0
   If .RecordCount > 0 Then
      .MoveFirst
      Do While Not .EOF
         strSql = "insert into acc1n0(a1n01,a1n02,a1n03,a1n04,a1n05,a1n06)" & _
            " values('" & txtA1K01 & "','2','" & .Fields("a1n03") & "'" & _
            ",'" & .Fields("a1n04") & "'," & .Fields("a1n05") & ",'" & .Fields("a1n06") & "')"
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
            " values('" & txtA1K01 & "','1','" & .Fields("a1n03") & "'" & _
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

'*************************************************
'  開啟資料表
'*************************************************
Private Function OpenTable() As Boolean

On Error GoTo Checking
   
   'Added by Morgan 2017/1/6
   '維護時先刪除收文號不是該張請款單者(請款單有改過)
   If Not m_bolQuery Then
      cnnConnection.Execute "delete acc1n0 where a1n01='" & txtA1K01 & "' and rtrim(a1n03) is not null and not exists(select * from caseprogress where cp60=a1n01 and cp09=a1n03)", intI
   End If
   'end 2017/1/6
   
   strExc(0) = "select * from acc1k0 where a1k01='" & txtA1K01 & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      txtA1K13 = "" & .Fields("a1k13")
      txtA1K14 = "" & .Fields("a1k14")
      txtA1K15 = "" & .Fields("a1k15")
      txtA1K16 = "" & .Fields("a1k16")
      '台幣金額捨去小數(和點數分配相同原則)
      'Modify By Sindy 2013/1/10
      'txtPts = Fix(Val("" & .Fields("a1k11")) - Val("" & .Fields("a1k09")) - Val("" & .Fields("a1k10")) * Val("" & .Fields("a1k06"))) / 1000
      'Modified by Morgan 2018/3/30
      'txtPts = Fix(Val("" & .Fields("a1k11")) - Val("" & .Fields("a1k09")) - Val("" & .Fields("a1k06"))) / 1000
      txtPts = Fix(Val("" & .Fields("a1k11")) - Val("" & .Fields("a1k09")) - Val("" & .Fields("a1k06")) + Val("" & .Fields("a1k36"))) / 1000
      'end 2018/3/30
      '2013/1/10 End
      End With
   End If
   
   strExc(0) = "select st02,a1n02,a1n03,decode(cpm03,'（無）',cpm04,cpm03) cpm03,a1n04,a1n05,a1n06" & _
      " from acc1n0,staff,caseprogress,casepropertymap where a1n01='" & txtA1K01 & "' and a1n02='2'" & _
      " and cp09(+)=a1n03 and cpm01(+)=cp01 and cpm02(+)=cp10" & _
      " and st01(+)=a1n04 order by a1n04,a1n03"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   'Modify by Amy 2014/06/26 +FormName 改暫存TB
   Set adoacc1n0 = PUB_CreateRecordset(RsTemp, , , , Me.Name)
   Set RsTemp = Nothing
   Set Adodc1.Recordset = adoacc1n0
   
   Set DataGrid1.DataSource = Adodc1
   DataGrid1.Refresh
   DataGrid1.col = 0
   DataGrid1.CurrentCellVisible = True
   SumShow
   AdodcClear
   
   strExc(0) = "select st02,a1n02,a1n03,decode(cpm03,'（無）',cpm04,cpm03) cpm03,a1n04,a1n05,a1n06" & _
      " from acc1n0,staff,caseprogress,casepropertymap where a1n01='" & txtA1K01 & "' and a1n02='1'" & _
      " and cp09(+)=a1n03 and cpm01(+)=cp01 and cpm02(+)=cp10" & _
      " and st01(+)=a1n04 order by a1n04,a1n03"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   'Modify by Amy 2014/06/26 +FormName 改暫存TB
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
   txtA1N06 = "" & .Fields("a1n06").Value
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
         " where cp60='" & txtA1K01 & "' and cp09='" & txtA1N03 & "' and cpm01(+)=cp01 and cpm02(+)=cp10"
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
      strExc(1) = PUB_GetST03(txtA1N04)
      If strExc(1) = "F51" Or strExc(1) = "F52" Then
         MsgBox "不可輸入外翻編號！"
         Cancel = True
         txtA1N04_GotFocus
         Exit Sub
'Removed by Morgan 2015/9/22 不必再限制
'      'Added by Morgan 2015/5/26
'      ElseIf strExc(1) = "F31" Or strExc(1) = "F41" Or strExc(1) = "L02" Then
'         MsgBox "承辦人部門錯誤！"
'         Cancel = True
'         txtA1N04_GotFocus
'         Exit Sub
'      'end 2015/5/26
'end 2015/9/22
      End If
      
      'Modified by Lydia 2015/06/01 改共用模組
      'If chkA0910(strExc(1)) = False Then
      strExc(2) = GetDeptA09(strExc(1), "10")
      If Len(strExc(2)) = 0 Then
         MsgBox "承辦人作帳部門未設定或無法讀取！"
         Cancel = True
         txtA1N04_GotFocus
         Exit Sub
      End If
      
      '收文號只有一個時預設
      If Cancel = False And Trim(txtA1N03) = "" Then
         If txtA1N04 < "F" Then
            strExc(0) = "SELECT DISTINCT CP09 FROM CASEPROGRESS WHERE CP60='" & txtA1K01 & "'"
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

Private Sub txtA1N06_GotFocus()
   TextInverse txtA1N06
End Sub

Private Sub txtA1N06_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
      Beep
      KeyAscii = 0
   End If
End Sub

Private Sub SumShow()
   Dim ii As Integer
   txtSum = 0
   Set RsTemp = Adodc1.Recordset.Clone
   With RsTemp
   If .RecordCount > 0 Then
      .MoveFirst
      Do While Not .EOF
         'Modified by Morgan 2025/8/26 修正有時會回傳科學記號問題
         'txtSum = Val(txtSum) + Val("" & .Fields("a1n05"))
         txtSum = Round(Val(txtSum) + Val("" & .Fields("a1n05")), 3)
         'end 2025/8/26
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
      'Modify by Amy 2014/06/26
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
   txtA1N06 = ""
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
                  GoTo ExitPort
               End If
            End If
         End If
      End If
   End If
   If bolAdd Then .AddNew
   .Fields("a1n04").Value = txtA1N04
   .Fields("a1n05").Value = Val(txtA1N05)
   .Fields("a1n03").Value = txtA1N03
   .Fields("a1n06").Value = txtA1N06
   .Fields("st02").Value = txtST02
   .Fields("cpm03").Value = m_A1N03_CPM03
   'Modify by Amy 2014/06/26
   '.UpdateBatch
   .UPDATE
   .Sort = "a1n04,a1n03"
   AdodcClear
   SumShow
   
ExitPort:
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
            GoTo ExitPort
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
   .Fields("a1n06").Value = txtA1N06
   .Fields("st02").Value = txtST02
   .Fields("cpm03").Value = m_A1N03_CPM03
   'Modify by Amy 2014/06/26
   '.UpdateBatch
   .UPDATE
   AdodcClear
   SumShow
   
ExitPort:
   End With
   
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)
   Select Case KeyCode
      Case vbKeyInsert
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
   
   If (txtA1K13 = "FCL" Or txtA1K13 = "LIN" Or txtA1K13 = "CFL") And txtA1N04 = "97009" Then
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
   
   
   'Added by Morgan 2013/7/29
   If Val(txtA1N05) = 0 Then
      MsgBox "點數必須大於 0 ！", vbExclamation
      txtA1N05.SetFocus
      Exit Function
   End If
   'end 2013/7/29
   
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
      
      'Added by Morgan 2015/5/26
      strExc(1) = GetST15(txtA1N04_1)
      If strExc(1) = "F51" Or strExc(1) = "F52" Then
         MsgBox "不可輸入外翻編號！"
         Cancel = True
         txtA1N04_GotFocus
         Exit Sub
'Removed by Morgan 2015/9/22 不必再限制
'      ElseIf strExc(1) = "F31" Or strExc(1) = "F41" Or strExc(1) = "L02" Then
'         MsgBox "智權人員收文部門錯誤！"
'         Cancel = True
'         txtA1N04_1_GotFocus
'         Exit Sub
'end 2015/9/22
      End If
      'end 2015/5/26
         
      '收文號只有一個時預設
      If Cancel = False And Trim(txtA1N03) = "" Then
         If txtA1N04_1 < "F" Then
            strExc(0) = "SELECT DISTINCT CP09 FROM CASEPROGRESS WHERE CP60='" & txtA1K01 & "'"
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
         " where cp60='" & txtA1K01 & "' and cp09='" & txtA1N03_1 & "' and cpm01(+)=cp01 and cpm02(+)=cp10"
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
   
   'Added by Morgan 2013/7/29
   If Val(txtA1N05_1) = 0 Then
      MsgBox "點數必須大於 0 ！", vbExclamation
      txtA1N05_1.SetFocus
      Exit Function
   End If
   'end 2013/7/29
   
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
   If Not (.EOF And .BOF) Then
      .Sort = "a1n04,a1n03"
      .MoveFirst
      .Find "a1n04='" & txtA1N04_1 & "'"
      If Not .EOF Then
         .Find "a1n03='" & txtA1N03_1 & "'"
         If Not .EOF Then
            If txtA1N04_1 = .Fields("a1n04") Then 'Added by Morgan 2016/5/16
               bolAdd = False
               If MsgBox("資料已存在，是否要更新！", vbYesNo + vbDefaultButton2) = vbNo Then
                  GoTo ExitPort
               End If
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
   'Modify by Amy 2014/06/26
   '.UpdateBatch
   .UPDATE
   .Sort = "a1n04,a1n03"
   AdodcClear_1
   SumShow_1
   
ExitPort:
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
            GoTo ExitPort
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
   'Modify by Amy 2014/06/26
   '.UpdateBatch
   .UPDATE
   AdodcClear_1
   SumShow_1
   
ExitPort:
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
'Modified by Lydia 2015/06/01 改共用模組
'Private Function chkA0910(p_A0901 As String) As Boolean
'   Dim stSQL As String, intR As Integer
'   Dim adoRst As ADODB.Recordset
'   stSQL = "select a0910 from acc090 where a0901='" & p_A0901 & "'"
'   intR = 1
'   Set adoRst = ClsLawReadRstMsg(intR, stSQL)
'   If intR = 1 Then
'      If Not IsNull(adoRst(0)) Then
'         chkA0910 = True
'      End If
'   End If
'   Set adoRst = Nothing
'End Function

