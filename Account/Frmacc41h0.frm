VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Frmacc41h0 
   AutoRedraw      =   -1  'True
   Caption         =   "智權期末結餘保留傳票產生"
   ClientHeight    =   5184
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8928
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5184
   ScaleWidth      =   8928
   Begin VB.CommandButton CmdRead 
      BackColor       =   &H00C0FFC0&
      Caption         =   "讀取資料"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3240
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   0
      Width           =   1200
   End
   Begin VB.CommandButton CmdExcel 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Excel"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7680
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   10
      Width           =   1150
   End
   Begin VB.CommandButton CmdSaveTmp 
      BackColor       =   &H00C0FFC0&
      Caption         =   "暫存檔案"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6480
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   10
      Width           =   1150
   End
   Begin VB.CommandButton CmdSaveAcc 
      BackColor       =   &H00C0FFC0&
      Caption         =   "產生傳票"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   6480
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   360
      Width           =   1150
   End
   Begin VB.CommandButton CmdSaveAcc 
      BackColor       =   &H00C0FFC0&
      Caption         =   "更正傳票"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   7680
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   360
      Width           =   1150
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc41h0.frx":0000
      Height          =   2645
      Left            =   100
      TabIndex        =   16
      Top             =   1440
      Width           =   8700
      _ExtentX        =   15346
      _ExtentY        =   4657
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   16
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.8
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
      ColumnCount     =   12
      BeginProperty Column00 
         DataField       =   "Dept"
         Caption         =   "部門"
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
         DataField       =   "StName"
         Caption         =   "姓名"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Type"
         Caption         =   "資料來源"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "T"
         Caption         =   "T"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "P"
         Caption         =   "P"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "CFT"
         Caption         =   "CFT"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "CFP"
         Caption         =   "CFP"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "Total"
         Caption         =   "合計"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "ID"
         Caption         =   "ID"
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
      BeginProperty Column09 
         DataField       =   "R001"
         Caption         =   "R001"
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
      BeginProperty Column10 
         DataField       =   "R002"
         Caption         =   "R002"
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
      BeginProperty Column11 
         DataField       =   "R003"
         Caption         =   "R003"
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
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   756.284
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   780.095
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   972.284
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1091.906
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1091.906
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   1091.906
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            ColumnWidth     =   1091.906
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column08 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column09 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column10 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column11 
            Object.Visible         =   0   'False
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
      Left            =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   550
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
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   2160
      TabIndex        =   22
      Top             =   10
      Width           =   1000
      _ExtentX        =   1757
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "標楷體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   315
      Left            =   0
      Top             =   1080
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   550
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
      Bindings        =   "Frmacc41h0.frx":0015
      Height          =   1060
      Left            =   105
      TabIndex        =   23
      Top             =   4130
      Width           =   8700
      _ExtentX        =   15346
      _ExtentY        =   1884
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      BackColor       =   -2147483624
      ColumnHeaders   =   0   'False
      HeadLines       =   1
      RowHeight       =   16
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   12
      BeginProperty Column00 
         DataField       =   "Dept"
         Caption         =   "部門"
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
         DataField       =   "StName"
         Caption         =   "姓名"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Type"
         Caption         =   "資料來源"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "T"
         Caption         =   "T"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "P"
         Caption         =   "P"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "CFT"
         Caption         =   "CFT"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "CFP"
         Caption         =   "CFP"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "Total"
         Caption         =   "合計"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "ID"
         Caption         =   "ID"
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
      BeginProperty Column09 
         DataField       =   "R001"
         Caption         =   "R001"
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
      BeginProperty Column10 
         DataField       =   "R002"
         Caption         =   "R002"
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
      BeginProperty Column11 
         DataField       =   "R003"
         Caption         =   "R003"
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
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   756.284
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   780.095
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   972.284
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1091.906
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1091.906
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   1091.906
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            ColumnWidth     =   1091.906
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column08 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column09 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column10 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column11 
            Object.Visible         =   0   'False
         EndProperty
      EndProperty
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "隔月初轉回傳票日期："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   21
      Top             =   360
      Width           =   2205
   End
   Begin VB.Label Lbl1 
      BackStyle       =   0  '透明
      Caption         =   "Lbl(6)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   5400
      TabIndex        =   20
      Top             =   360
      Width           =   1155
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "轉期末傳票日期："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   19
      Top             =   360
      Width           =   1680
   End
   Begin VB.Label Lbl1 
      BackStyle       =   0  '透明
      Caption         =   "Lbl(5)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   2160
      TabIndex        =   18
      Top             =   360
      Width           =   1155
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "PS:若有轉撥資料,產生傳票後會自動開啟該傳票,請               自行修改轉撥傳票之部門及摘要！"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   390
      Left            =   4560
      TabIndex        =   17
      Top             =   960
      Width           =   4995
      WordWrap        =   -1  'True
   End
   Begin VB.Label Lbl1 
      BackStyle       =   0  '透明
      Caption         =   "Lbl(4)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   3370
      TabIndex        =   15
      Top             =   1080
      Width           =   1155
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   14
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Lbl1 
      BackStyle       =   0  '透明
      Caption         =   "Lbl(3)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   2160
      TabIndex        =   13
      Top             =   1080
      Width           =   1155
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "隔月初轉回傳票號碼："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   12
      Top             =   1080
      Width           =   2100
   End
   Begin VB.Label Lbl1 
      BackStyle       =   0  '透明
      Caption         =   "XXXXXXXXXX"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   6360
      TabIndex        =   11
      Top             =   720
      Width           =   1995
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "　轉撥傳票號碼："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   10
      Top             =   720
      Width           =   1680
   End
   Begin VB.Label Lbl1 
      BackStyle       =   0  '透明
      Caption         =   "Lbl(1)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3370
      TabIndex        =   9
      Top             =   720
      Width           =   1155
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   8
      Top             =   765
      Width           =   255
   End
   Begin VB.Label Lbl1 
      BackStyle       =   0  '透明
      Caption         =   "Lbl(0)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2160
      TabIndex        =   7
      Top             =   720
      Width           =   1155
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "轉期末傳票號碼："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   720
      Width           =   1680
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   960
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "轉期末傳票年月："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   10
      Width           =   1680
   End
End
Attribute VB_Name = "Frmacc41h0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/08 Form2.0已修改 DataGrid1
'Create by Amy 2016/04/18
Option Explicit
Dim ado41H0 As New ADODB.Recordset
Dim adoQ As New ADODB.Recordset
Dim i As Integer, j As Integer
Dim strAcDate As String, strQ As String, strCmd As String
Dim strAxb(9 To 13) As String
Dim strAxb03 As String
Dim strA0b01 As String, strA0b05 As String '會計過帳日/業績輸入關閉年月
Dim bol0b1HasDt As Boolean, bolHasAx210 As Boolean 'Acc0b1是否有資料/是否已過帳(非轉撥)
Dim bolTrans As Boolean '是否有產生轉撥資料 for 是否run 傳票畫面
Dim bolNotUpd As Boolean '是否改過資料 for 結束show Msg用
Dim dgRowB As Integer '記錄前一筆Datagrid 的row
Dim dgRow As Integer '記錄目前更新那筆的Datagrid 的row
Dim dgFirstRow As Integer '記錄目前更新那筆Datagrid的第一筆
Dim strF(), intWidth()
Dim intField As Integer, intCounter As Integer, intTitle As Integer
'Add by Amy 2017/09/13
Const MskFormat As String = "###/##"
Dim adoQ2 As New ADODB.Recordset
Dim strQ2 As String
Dim strDate_S As String, strDate_E As String 'for 期初/本月放出
Dim strAxb02 As String '當月期末保留傳票日
Dim bolHasAx210_T As Boolean '轉撥傳票是否已過帳(可能產生傳票時沒後補的)
Dim bolUpdAxb11 As Boolean '轉撥傳票是否需更新
Dim bolShowMsg As Boolean, bolBT As Boolean '是否show 訊息/是否按按鈕 for run SetBTAndLab
Dim RsSum As New ADODB.Recordset 'Add by Amy 2019/01/15 加智權部合計

Private Sub CmdExcel_Click()
    Dim xlsAgentPoint As New Excel.Application
    Dim wksrpt As New Worksheet
    Dim xlsFileName As String, strTmp As String, strOldDept As String
    Dim strSum(1 To 4) As String, strTotal(1 To 4)  As String '記錄區合計/記錄所合計
    Dim intStart As Integer, intColor As Integer
    Dim bolFormula As Boolean, bolFontColor As Boolean
    
On Error GoTo ErrHand
    
    '檢查合計是否有誤
    If ChkGridSum = False Then
        Screen.MousePointer = vbDefault: Exit Sub
    End If
    
    If bolNotUpd = True Then
        If SaveTmp = False Then Screen.MousePointer = vbDefault: Exit Sub
    End If
    
    ReDim strF(12)
    ReDim intWidth(12)
    strF = Array("部門", "編號", "姓名", "ty1", "資料來源", "T", "P", "CFT", "CFP", _
                      "期初保留合計", "本月放出合計", "本月報出合計", "期末保留合計")
    intWidth = Array(6.5, 4.88, 6.25, 2.38, 8.25, 11.38, 11.38, 11.38, 11.38, _
                                13.75, 13.75, 13.75, 13.75)
    
    intField = 65: intCounter = 1
    xlsFileName = strAcDate & "結餘期初保留及本月放出" & ServerDate & MsgText(43)
     If Dir(strExcelPath & xlsFileName) = MsgText(601) Then
       If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
            MkDir strExcelPath
       End If
    Else
         Kill strExcelPath & xlsFileName
    End If
    
    xlsAgentPoint.SheetsInNewWorkbook = 1 'Added by Lydia 2019/03/13 預設工作表數量
    xlsAgentPoint.Workbooks.add
    xlsAgentPoint.Application.WindowState = xlMinimized
    Set wksrpt = xlsAgentPoint.Worksheets(1)
    
    ado41H0.MoveFirst
    Call SetField(wksrpt)
    intTitle = intCounter: intCounter = intCounter + 1: intStart = intCounter
    
    Do While ado41H0.EOF = False
        For i = LBound(strF) To UBound(strF)
            bolFormula = False: strTmp = "": intColor = Empty: bolFontColor = False
            Select Case strF(i)
                Case "部門"
                    strTmp = "" & ado41H0.Fields("Dept")
                    If Mid("" & ado41H0.Fields("R001"), 4, 1) = "X" Or Mid("" & ado41H0.Fields("R001"), 3, 1) = "X" Or _
                        "" & ado41H0.Fields("R001") = "SZZZZ" Or "" & ado41H0.Fields("R001") = "ZZZZZ" Then
                        strTmp = ""
                    End If
                Case "編號"
                    If "" & ado41H0.Fields("R003") = "1" Then strTmp = "" & ado41H0.Fields("R002")
                    If Mid("" & ado41H0.Fields("R001"), 4, 1) = "X" Or Mid("" & ado41H0.Fields("R001"), 3, 1) = "X" Or _
                        "" & ado41H0.Fields("R001") = "SZZZZ" Or "" & ado41H0.Fields("R001") = "ZZZZZ" Then
                        strTmp = ""
                    End If
                Case "姓名"
                    strTmp = "" & ado41H0.Fields("StName")
                    If Mid("" & ado41H0.Fields("R001"), 4, 1) = "X" Or Mid("" & ado41H0.Fields("R001"), 3, 1) = "X" Or "" & ado41H0.Fields("R001") = "SZZZZ" Or "" & ado41H0.Fields("R001") = "ZZZZZ" Then
                        If "" & ado41H0.Fields("R001") = "SZZZZ" Or "" & ado41H0.Fields("R001") = "ZZZZZ" Then
                            strTmp = "" & ado41H0.Fields("Dept") & strTmp
                        Else
                            strTmp = Right("" & ado41H0.Fields("Dept"), 1) & strTmp
                        End If
                        bolFontColor = True
                    End If
                Case "ty1"
                    strTmp = "" & ado41H0.Fields("R003")
                Case "資料來源"
                    strTmp = "" & ado41H0.Fields("Type")
                    If "" & ado41H0.Fields("R003") = "3" Then intColor = 37
                    If "" & ado41H0.Fields("R003") = "4" Then intColor = 40
                Case "T"
                    strTmp = Val(ado41H0.Fields("T"))
                Case "P"
                    strTmp = Val(ado41H0.Fields("P"))
                Case "CFT"
                    strTmp = Val(ado41H0.Fields("CFT"))
                Case "CFP"
                    strTmp = Val(ado41H0.Fields("CFP"))
                Case "期初保留合計"
                    bolFormula = True: strTmp = "1"
                Case "本月放出合計"
                    bolFormula = True: strTmp = "2"
                Case "本月報出合計"
                    bolFormula = True: strTmp = "3"
                Case "期末保留合計"
                    bolFormula = True: strTmp = "4"
            End Select
            '北中區小計
            If Mid("" & ado41H0.Fields("R001"), 4, 1) = "X" And i >= GetValue("T") And i <= GetValue("CFP") Then
                bolFormula = True: strTmp = Chr(i + intField) & intStart & ":" & Chr(i + intField) & intCounter - 1
            '所合計
            ElseIf Mid("" & ado41H0.Fields("R001"), 3, 1) = "X" And i >= GetValue("T") And i <= GetValue("CFP") Then
                bolFormula = True
                If i <> GetValue("T") Then
                    strTmp = Replace(strSum("" & ado41H0.Fields("R003")), Chr(GetValue("T") + intField), Chr(i + intField))
                ElseIf Mid("" & ado41H0.Fields("R001"), 2, 1) >= "3" Then
                    strTmp = "," & Chr(i + intField) & intStart & ":" & Chr(i + intField) & intCounter - 1
                Else
                    strTmp = strSum("" & ado41H0.Fields("R003"))
                End If
            '智權小計
            ElseIf ("" & ado41H0.Fields("R001") = "SZZZZ" Or "" & ado41H0.Fields("R001") = "ZZZZZ") And i >= GetValue("T") And i <= GetValue("CFP") Then
                bolFormula = True
                If i <> GetValue("T") Then
                    strTmp = Replace(strTotal("" & ado41H0.Fields("R003")), Chr(GetValue("T") + intField), Chr(i + intField))
                Else
                    strTmp = strTotal("" & ado41H0.Fields("R003"))
                End If
            End If
                
            If bolFormula = True Then
                '區小計
                If Mid("" & ado41H0.Fields("R001"), 4, 1) = "X" And i >= GetValue("T") And i <= GetValue("CFP") Then
                    strTmp = "=SumIF($" & Chr(GetValue("ty1") + intField) & intStart & ":$" & Chr(GetValue("ty1") + intField) & intCounter - 1 & "," & "$" & Chr(GetValue("ty1") + intField) & intCounter & "," & strTmp & ")"
                '所/智權小計
                ElseIf (Mid("" & ado41H0.Fields("R001"), 3, 1) = "X" Or "" & ado41H0.Fields("R001") = "SZZZZ" Or "" & ado41H0.Fields("R001") = "ZZZZZ") And i >= GetValue("T") And i <= GetValue("CFP") Then
                    If Mid("" & ado41H0.Fields("R001"), 3, 1) = "X" And Mid("" & ado41H0.Fields("R001"), 2, 1) >= "3" Then
                        strTmp = "=SumIF($" & Chr(GetValue("ty1") + intField) & intStart & ":$" & Chr(GetValue("ty1") + intField) & intCounter - 1 & "," & "$" & Chr(GetValue("ty1") + intField) & intCounter & "," & Mid(strTmp, 2) & ")"
                    Else
                        strTmp = "=Sum(" & Mid(strTmp, 2) & ")"
                    End If
                Else
                    strTmp = "=IF($" & Chr(GetValue("ty1") + intField) & intCounter & "=" & strTmp & ",Sum($" & Chr(GetValue("T") + intField) & intCounter & ":$" & Chr(GetValue("CFP") + intField) & intCounter & "),0)"
                End If
                wksrpt.Range(Chr(i + intField) & intCounter).Formula = strTmp
            Else
                If ("" & ado41H0.Fields("R001") = "SZZZZ" Or "" & ado41H0.Fields("R001") = "ZZZZZ") And i = GetValue("姓名") Then
                    wksrpt.Range(Chr(i + intField - 1) & intCounter).Value = strTmp
                Else
                    wksrpt.Range(Chr(i + intField) & intCounter).Value = strTmp
                End If
                If bolFontColor = True Then
                    If ("" & ado41H0.Fields("R001") = "SZZZZ" Or "" & ado41H0.Fields("R001") = "ZZZZZ") Then
                        wksrpt.Range(Chr(i + intField - 1) & intCounter).Font.ColorIndex = 3
                    Else
                        wksrpt.Range(Chr(i + intField) & intCounter).Font.ColorIndex = 3
                    End If
                End If
            End If
            If intColor <> Empty Then
                '設置儲存格填充色(藍)
                If intColor = 37 Then
                    wksrpt.Range(Chr(i + intField) & intCounter).Interior.ColorIndex = intColor
                '設置儲存格填充色(橘)
                Else
                    wksrpt.Range(Chr(i + intField) & intCounter & ":" & Chr(GetValue("CFP") + intField) & intCounter).Interior.ColorIndex = intColor
                End If
            End If
            '記錄所合計=區合計列
            If (Mid("" & ado41H0.Fields("R001"), 4, 1) = "X" Or (Mid("" & ado41H0.Fields("R001"), 3, 1) = "X" And Mid("" & ado41H0.Fields("R001"), 2, 1) >= "3")) And i = GetValue("T") Then
                '南所及高所沒區不需 區合計
                If Mid("" & ado41H0.Fields("R001"), 2, 1) >= "3" Then
                    strSum("" & ado41H0.Fields("R003")) = "," & Chr(i + intField) & intStart & ":" & Chr(i + intField) & intCounter - 1
                Else
                    strSum("" & ado41H0.Fields("R003")) = strSum("" & ado41H0.Fields("R003")) & "," & Chr(i + intField) & intCounter
                End If
            End If
            '記錄智權合計=所合計列
            If Left(strOldDept, 1) = "S" And Left("" & ado41H0.Fields("R001"), 1) <> "S" And i = GetValue("T") Then
                strTotal("" & ado41H0.Fields("R003")) = "," & Chr(i + intField) & intCounter
                strTotal(2) = "": strTotal(3) = "": strTotal(4) = ""
            ElseIf (Mid("" & ado41H0.Fields("R001"), 3, 1) = "X" Or Mid("" & ado41H0.Fields("R001"), 1, 1) <> "S") And i = GetValue("T") And "" & ado41H0.Fields("R001") <> "ZZZZZ" Then
                strTotal("" & ado41H0.Fields("R003")) = strTotal("" & ado41H0.Fields("R003")) & "," & Chr(i + intField) & intCounter
            End If
        Next i
      
        If "" & ado41H0.Fields("R003") = "4" Then
            intCounter = intCounter + 2
        Else
            intCounter = intCounter + 1
        End If
        If Mid("" & ado41H0.Fields("R001"), 4, 1) = "X" And "" & ado41H0.Fields("R003") = "4" Then
            intStart = intCounter
        ElseIf Mid("" & ado41H0.Fields("R001"), 3, 1) = "X" And "" & ado41H0.Fields("R003") = "4" Then
            strSum(1) = "": strSum(2) = "": strSum(3) = "": strSum(4) = ""
            intStart = intCounter
        End If
        strOldDept = "" & ado41H0.Fields("R001")
        ado41H0.MoveNext
        If ado41H0.EOF = False Then
            '小計欄不多空一列
            If Mid("" & ado41H0.Fields("R001"), 4, 1) = "X" And "" & ado41H0.Fields("R003") = "1" Then
                intCounter = intCounter - 1
            '南所及高所沒區不需 區合計
            ElseIf Left(strOldDept, 2) <> "" & Left(ado41H0.Fields("R001"), 2) And Mid("" & ado41H0.Fields("R001"), 2, 1) >= "3" And "" & ado41H0.Fields("R003") = "1" Then
                intStart = intCounter
            End If
        End If
    Loop
    wksrpt.Range(Chr(GetValue("T") + intField) & intCounter & ":" & Chr(GetValue("期末保留合計") + intField) & intCounter).NumberFormatLocal = "#,##0.0_ "
    ado41H0.MoveFirst
    
    '更改部分欄位名稱為全型
    Call SetField(wksrpt, True)
    
    If Val(xlsAgentPoint.Version) < 12 Then
       xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & xlsFileName, FileFormat:=-4143
    Else
       xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & xlsFileName, FileFormat:=56
    End If
    xlsAgentPoint.Workbooks.Close
    xlsAgentPoint.Quit
    MsgBox "Excel已產生！"
    StatusClear
    Exit Sub
    
ErrHand:
    MsgBox Err.Description, , MsgText(5)
    If Val(xlsAgentPoint.Version) < 12 Then
        xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & xlsFileName, FileFormat:=-4143
    Else
        xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & xlsFileName, FileFormat:=56
    End If
    xlsAgentPoint.Workbooks.Close
    xlsAgentPoint.Quit
    Set xlsAgentPoint = Nothing
    Set wksrpt = Nothing
End Sub

Private Sub CmdSaveAcc_Click(Index As Integer)
    Dim strMsg As String
    
    Screen.MousePointer = vbHourglass
    
    '存檔前檢查
    bolBT = True
    '2.產生傳票/3.更正傳票
    If FormCheck(Index) = False Then
        bolBT = False
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    bolBT = False
   
    If SaveTmp = False Then
        DataGrid1.Visible = True
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    If SaveAcc(Index) = True Then
        If bolTrans = True Then
            strMsg = "已產生轉撥傳票，請自行修改轉撥傳票之部門及摘要！"
        Else
            strMsg = "結餘傳票已產生！"
        End If
        MsgBox strMsg
        If bolTrans = True Then
            '開啟傳票輸入畫面
            If Lbl1(2) <> MsgText(601) Then
                With Frmacc4120
                    .Tag = Me.Name
                    Me.Hide
                    .Text1 = "1"
                    .Text2 = Lbl1(2)
                    'Add by Amy 2024/08/05 整合檢查,避免彈訊息後又可以操作
                    .MaskEdBox1 = CFDate(Pub_GetField("Acc020", "a0201='1' And a0202='" & Lbl1(2) & "'", "a0205"))
                    'Modify by Amy 2022/05/16 +if '按修改->Insert->取消->修改->Insert->存檔會出現error
                    .bolF3 = True
                    .Command3_Click
                    .bolF3 = False
                    'end 2022/05/16
                End With
            End If
        End If
    End If
    Call ShowBt
    Screen.MousePointer = vbDefault
End Sub

Private Sub CmdSaveTmp_Click()
    
    bolBT = True
    If FormCheck(1) = False Then bolBT = False: Exit Sub
    bolBT = False
    
    Screen.MousePointer = vbHourglass
    If SaveTmp = True Then
        MsgBox "檔案已暫存！"
    End If
    Call ShowBt
    Screen.MousePointer = vbDefault
End Sub

Private Sub CmdRead_Click()
    
    If FormCheck(0) = False Then Exit Sub
   
    Screen.MousePointer = vbHourglass
    strAcDate = Replace(MaskEdBox1, "/", "") 'Modify by Amy 2017/09/13 原:輸年月日
    
    Call ReadData
    strQ = "Select Decode(R003,1,R009,'') Dept,Decode(R003,1,Decode(length(R001)||SubStr(R001,length(R001),1),'4X','小計',Decode(length(R001)||SubStr(R001,length(R001),1),'3X','合計',Decode(R001,'SZZZZ','合計',Decode(R001,'ZZZZZ','合計',St02)))),'') StName," & _
                "Decode(R003,1,'期初保留',Decode(R003,2,'本月放出',Decode(R003,3,'本月報出','期末保留'))) Type,R004 T,R005 P,R006 CFT,R007 CFP,R008 Total,ID,R001,R002,R003 From Accrpt41h0,Acc090,Staff " & _
                "Where ID='" & strUserNum & "' And R001=A0901(+) And R002=ST01(+) " & _
                "Order by Decode(SubStr(R001,1,1),'S',1,2),R001,R002,R003"
    If ado41H0.State = adStateOpen Then ado41H0.Close
    ado41H0.Open strQ, adoTaie, adOpenDynamic, adLockBatchOptimistic
    
    DataGrid1.Enabled = True
    Adodc1.Recordset.Requery
    DataGrid1.ScrollBars = dbgAutomatic
    
    Call ReadSum 'Add by Amy 2019/01/15
    CmdExcel.Enabled = True
    Call ShowBt
    Screen.MousePointer = vbDefault
End Sub

'將資料寫入TempTable
Private Sub ReadData()
    Dim strVal(4 To 7) As String, strTmp(1) As String
    Dim strDiff As String, strMaxP As String, strTot As String 'Add by Amy 2018/06/01
    Dim RsQ As New ADODB.Recordset, intQ As Integer, bolNoVoucher As Boolean 'Add by Amy 2022/06/10 無傳票

    strDate_S = Replace(strAcDate, "/", "") 'Modify by Amy 2017/09/13 原:Left(strAcDate, 5)
    If Val(Right(strDate_S, 2)) = 12 Then
        strDate_E = Val(Left(strDate_S, 3)) + 1 & "01"
    Else
        strDate_E = Left(strDate_S, 3) & IIf(Val(Right(strDate_S, 2)) + 1 <= 9, "0" & Val(Right(strDate_S, 2)) + 1, Val(Right(strDate_S, 2)) + 1)
    End If
    
    'SalesBance 沒 資料時先寫入人員資料
    strExc(0) = ""
    If ExistCheck("SalesBalance", "SB01", strDate_S, strExc(0), False) = False Then
        'Memo 文雄掛於北四區(201705月改st15=S14)
        strCmd = "Insert Into SalesBalance (SB01,SB02,SB03) " & _
                        "Select Distinct '" & strDate_S & "',Nvl(SP48,ST15) as SP48,AX209 From Acc021,Staff,SalesPoint " & _
                        "Where ax202>='D'||" & strDate_S & " And ax202<'D'||" & Val(strDate_E) & " And SP01(+)=" & Val(strDate_S) + 191100 & " And ax209=SP02(+) " & _
                        "And ax207>0 And ax209 Is not null And ax209=ST01(+) And SubStr(AX205,1,1)='4' And (ax205='4194' Or (InStr(ax213,'結餘')>0 And ax205<>'4194'))"
        adoTaie.Execute strCmd
        
        'Add by Amy 2022/06/10 +上述傳票沒有且SalePoint 結餘資料<>0 的人員 ex:11105月簡協理轉撥給10052,10052做結餘保留(新增人員當下無10052傳票)
        strCmd = "Insert Into SalesBalance (SB01,SB02,SB03) " & _
                        "Select SP01-191100,SP48,SP02 From SalesPoint " & _
                        "Where SP01=" & Val(strDate_S) + 191100 & " And Decode(SP36,null,Decode(SP32,null,Decode(SP28,null,SP24,SP28),SP32),SP36)<>0 " & _
                        "And SP02 Not In(Select SB03 From SalesBalance Where SB01=" & strDate_S & ") "
        adoTaie.Execute strCmd
    End If
    
    strCmd = "Delete From Accrpt41H0 Where ID='" & strUserNum & "' "
    adoTaie.Execute strCmd
        
    '期初保留
    strCmd = "Insert Into Accrpt41H0 (ID,R001,R002,R003,R004,R005,R006,R007,R008) " & _
                "Select '" & strUserNum & "',SB02,SB03,1,Nvl(T,0) T,Nvl(P,0) P,Nvl(CFT,0) CFT,Nvl(CFP,0) CFP, Nvl(Tol,0) Tol " & _
                "From SalesBalance,(" & GetBalanceSQL(1, strDate_S, strDate_E) & ") a " & _
                "Where SB01=" & strDate_S & " And SB02=Dept(+) And SB03=ST01(+) "
    adoTaie.Execute strCmd
    
    '本月放出
    strCmd = "Insert Into Accrpt41H0 (ID,R001,R002,R003,R004,R005,R006,R007,R008) " & _
                    "Select '" & strUserNum & "',SB02,SB03,2,Nvl(T,0) T,Nvl(P,0) P,Nvl(CFT,0) CFT,Nvl(CFP,0) CFP, Nvl(Tol,0) Tol " & _
                    "From SalesBalance,(" & GetBalanceSQL(2, strDate_S, strDate_E) & ") a " & _
                    "Where SB01=" & strDate_S & " And SB02=Dept(+) And SB03=ST01(+) "
    adoTaie.Execute strCmd
    
    '本月報出
    strCmd = ""
    If ChkSBVal(strDate_S, strCmd) = True Then
        'SalesBance 有 資料將資料更新至畫面暫存檔「本月報出」欄位中
        strCmd = "Insert Into Accrpt41H0 (ID,R001,R002,R003,R004,R005,R006,R007,R008) " & _
                        "Select '" & strUserNum & "',SB02,SB03,3,Nvl(SB04,0),Nvl(SB05,0),Nvl(SB06,0),Nvl(SB07,0),Nvl(SB04,0)+Nvl(SB05,0)+Nvl(SB06,0)+Nvl(SB07,0) " & _
                        "From SalesBalance Where SB01=" & strDate_S
        adoTaie.Execute strCmd
        
        strCmd = "Insert Into Accrpt41H0 (ID,R001,R002,R003,R004,R005,R006,R007,R008) " & _
                    "Select '" & strUserNum & "',SB02,SB03,4,NS4-LS4,NS5-LS5,NS6-LS6,NS7-LS7,NS8-LS8 From " & _
                     "(Select R001,R002,Sum(Nvl(R004,0)) NS4,Sum(Nvl(R005,0)) NS5,Sum(Nvl(R006,0)) NS6,Sum(Nvl(R007,0)) NS7,Sum(Nvl(R008,0)) NS8 " & _
                     "From Accrpt41H0 Where ID='" & strUserNum & "' And R003<=2 Group by R001,R002 ), " & _
                     "(Select SB02,SB03,Nvl(SB04,0) LS4,Nvl(SB05,0) LS5,Nvl(SB06,0) LS6,Nvl(SB07,0) LS7,Nvl(SB04,0)+Nvl(SB05,0)+Nvl(SB06,0)+Nvl(SB07,0) LS8 " & _
                     "From SalesBalance Where  SB01=" & strDate_S & " ) " & _
                    "Where R001=SB02(+) And R002=SB03(+) "
        adoTaie.Execute strCmd
    Else
        '增加本月報出及期末保留合計(抓取SalesPoint)
        strCmd = "Insert Into Accrpt41H0 (ID,R001,R002,R003,R004,R005,R006,R007,R008) " & _
                        "Select '" & strUserNum & "',SB02,SB03,3,0,0,0,0,0 From SalesBalance Where  SB01=" & strDate_S & _
           " Union Select Distinct '" & strUserNum & "',R001,R002,4,0,0,0,0,Decode(SP36,null,Decode(SP32,null,Decode(SP28,null,Nvl(SP24,0),SP28),SP32),SP36)*1000 " & _
                        "From Accrpt41H0,SalesPoint " & _
                        "Where ID='" & strUserNum & "' And R001=SP48(+) And R002=SP02(+) And SP01= " & Val(strDate_S) + 191100
        adoTaie.Execute strCmd
        
        '更新本月報出=B8(期初＋放出)-F8(期末)
        strQ = "Select B.*,Nvl(F8,0) F8 From " & _
                "(Select R001 B1,R002 B2,Sum(Nvl(R008,0)) B8 " & _
                "From Accrpt41H0 Where ID='" & strUserNum & "' And R003<=2 Group by R001,R002 " & _
                ") B,(Select R001 F1,R002 F2,R008 F8 From Accrpt41H0 Where ID='" & strUserNum & "' And R003=4) " & _
                "Where B1=F1(+) And B2=F2(+) "
        If adoQ.State = adStateOpen Then adoQ.Close
        adoQ.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
        If adoQ.RecordCount <> 0 Then
            With adoQ
                Do While .EOF = False
                    'Modify by Amy 2018/06/01 原先扣合計較小者,直至扣完為止,改為比例分攤,差額尾數調整於比例最大者
'                    If Val(.Fields("B8")) = Val(.Fields("F8")) Then
'                        strCmd = "Update Accrpt41H0 Set R004=0,R005=0,R006=0,R007=0,R008=0" & _
'                                      " Where ID='" & strUserNum & "' And R002='" & .Fields("B2") & "' And R003=3 "
'                        adoTaie.Execute strCmd
'                    Else
'                        strTmp(0) = Val(.Fields("B8")) - Val(.Fields("F8"))
'                        strTmp(1) = strTmp(0)
'                        '先扣合計較小者,直至扣完為止
'                        strQ2 = "Select * From (" & _
'                                    "Select '4' Ord,Sum(Nvl(R004,0)) StVal From Accrpt41H0 Where ID='" & strUserNum & "' And R002='" & .Fields("B2") & "' And R003<=2 " & _
'                         "Union Select '5' Ord,Sum(Nvl(R005,0)) StVal From Accrpt41H0 Where ID='" & strUserNum & "' And R002='" & .Fields("B2") & "' And R003<=2 " & _
'                         "Union Select '6' Ord,Sum(Nvl(R006,0)) StVal From Accrpt41H0 Where ID='" & strUserNum & "' And R002='" & .Fields("B2") & "' And R003<=2 " & _
'                         "Union Select '7' Ord,Sum(Nvl(R007,0)) StVal From Accrpt41H0 Where ID='" & strUserNum & "' And R002='" & .Fields("B2") & "' And R003<=2 ) " & _
'                         "Order by StVal Asc"
'                        If adoQ2.State = adStateOpen Then adoQ2.Close
'                        adoQ2.Open strQ2, adoTaie, adOpenStatic, adLockReadOnly
'                        If adoQ2.RecordCount <> 0 Then
'                            adoQ2.MoveFirst
'                            Do While adoQ2.EOF = False
'                                If Val(strTmp(1)) > 0 Then
'                                    If Val(strTmp(1)) - Val(adoQ2.Fields("StVal")) > 0 Then
'                                        strVal(adoQ2.Fields("Ord")) = adoQ2.Fields("StVal")
'                                        strTmp(1) = Val(strTmp(1)) - Val(adoQ2.Fields("StVal"))
'                                    Else
'                                        strVal(adoQ2.Fields("Ord")) = strTmp(1)
'                                        strTmp(1) = 0
'                                    End If
'                                Else
'                                    strVal(adoQ2.Fields("Ord")) = 0
'                                End If
'                                adoQ2.MoveNext
'                            Loop
'                            strCmd = "Update Accrpt41H0 Set R004=" & Val(strVal(4)) & ",R005=" & Val(strVal(5)) & ",R006=" & Val(strVal(6)) & _
'                                           ",R007=" & Val(strVal(7)) & ",R008=" & Val(strTmp(0)) & _
'                                      " Where ID='" & strUserNum & "' And R002='" & .Fields("B2") & "' And R003=3 "
'                            adoTaie.Execute strCmd
'                        End If
'                    End If
                strTmp(0) = Val(.Fields("B8")) '期初+本月
                strTmp(1) = Val(.Fields("B8")) - Val(.Fields("F8")) '本月報出
                strQ2 = "Select Sum(Nvl(R004,0)) T,Sum(Nvl(R005,0)) P,Sum(Nvl(R006,0)) CFT,Sum(Nvl(R007,0)) CFP " & _
                            "From Accrpt41H0 Where ID='" & strUserNum & "' And R002='" & .Fields("B2") & "' And R003<=2 "
                If adoQ2.State = adStateOpen Then adoQ2.Close
                adoQ2.Open strQ2, adoTaie, adOpenStatic, adLockReadOnly
                If adoQ2.RecordCount <> 0 Then
                    adoQ2.MoveFirst
                    Do While adoQ2.EOF = False
                        'Modify by Amy 2022/06/10
                        bolNoVoucher = False
                        '期初+本月=0 確認是否無傳票資料
                        If Val(strTmp(0)) = 0 Then
                            strExc(1) = GetBalanceSQL(0, strDate_S, strDate_E, "" & .Fields("B2"))
                            intQ = 1
                            Set RsQ = ClsLawReadRstMsg(intQ, strExc(1))
                            If intQ = 0 Then
                                bolNoVoucher = True
                            End If
                        End If
                        '無傳票資料,全部列至CFP
                        If bolNoVoucher = True Then
                            strVal(4) = "0" 'T
                            strVal(5) = "0" 'P
                            strVal(6) = "0" 'CFT
                            strVal(7) = strTmp(1) 'CFP
                        'end 2022/06/10
                        '期末為0依各部門期初+本月值報
                        ElseIf Val(.Fields("F8")) = 0 Then
                            strVal(4) = Val("" & adoQ2.Fields("T"))
                            strVal(5) = Val("" & adoQ2.Fields("P"))
                            strVal(6) = Val("" & adoQ2.Fields("CFT"))
                            strVal(7) = Val("" & adoQ2.Fields("CFP"))
                        '期末不是0
                        Else
                            '本月總報出*各部門/(期初保留+本月放出)
                            strVal(4) = Round(Val(strTmp(1)) * Round(Val("" & adoQ2.Fields("T")) / Val(strTmp(0)), 5), 0)
                            strVal(5) = Round(Val(strTmp(1)) * Round(Val("" & adoQ2.Fields("P")) / Val(strTmp(0)), 5), 0)
                            strVal(6) = Round(Val(strTmp(1)) * Round(Val("" & adoQ2.Fields("CFT")) / Val(strTmp(0)), 5), 0)
                            strVal(7) = Round(Val(strTmp(1)) * Round(Val("" & adoQ2.Fields("CFP")) / Val(strTmp(0)), 5), 0)
                            strDiff = Val(strTmp(1)) - Val(strVal(4)) - Val(strVal(5)) - Val(strVal(6)) - Val(strVal(7))
                            '尾數放於比例最大且未全報
                            If Val(strDiff) <> 0 Then
                                strMaxP = ""
                                For i = 0 To 3
                                    If Round(Val("" & adoQ2.Fields(i)) / Val(strTmp(0)), 5) > Val(strMaxP) And Val("" & adoQ2.Fields(i)) - Val(strVal(i + 4)) > 0 Then
                                        strMaxP = Round(Val("" & adoQ2.Fields(i)) / Val(strTmp(0)), 5)
                                        strTot = i
                                    End If
                                Next i
                                strVal(strTot + 4) = Val(strVal(strTot + 4)) + Val(strDiff)
                            End If
                        End If '期末是否為0
                        adoQ2.MoveNext
                    Loop
                    strCmd = "Update Accrpt41H0 Set R004=" & Val(strVal(4)) & ",R005=" & Val(strVal(5)) & ",R006=" & Val(strVal(6)) & _
                                    ",R007=" & Val(strVal(7)) & ",R008=" & Val(strTmp(1)) & _
                                    " Where ID='" & strUserNum & "' And R002='" & .Fields("B2") & "' And R003=3 "
                    adoTaie.Execute strCmd
                End If
                'end 2018/06/01
                .MoveNext
                Loop
            End With
        End If
        
        '更新期末保留資料
        strCmd = "Delete Accrpt41H0 Where ID='" & strUserNum & "' And R003=4 "
        adoTaie.Execute strCmd
        
        strCmd = "Insert Into Accrpt41H0 (ID,R001,R002,R003,R004,R005,R006,R007,R008) " & _
                        "Select '" & strUserNum & "',R001,R002,4,NS4-LS4,NS5-LS5,NS6-LS6,NS7-LS7,NS8-LS8 From " & _
                          "(Select R001,R002,Sum(Nvl(R004,0)) NS4,Sum(Nvl(R005,0)) NS5,Sum(Nvl(R006,0)) NS6,Sum(Nvl(R007,0)) NS7,Sum(Nvl(R008,0)) NS8 " & _
                          "From Accrpt41H0 Where ID='" & strUserNum & "' And R003<=2 Group by R001,R002 )," & _
                          "(Select R001 LS1,R002 LS2,Sum(Nvl(R004,0)) LS4,Sum(Nvl(R005,0)) LS5,Sum(Nvl(R006,0)) LS6,Sum(Nvl(R007,0)) LS7,Sum(Nvl(R008,0)) LS8 " & _
                          "From Accrpt41H0 Where ID='" & strUserNum & "' And R003=3 Group by R001,R002 ) " & _
                        "Where R001=LS1(+) And R002=LS2(+) "
        adoTaie.Execute strCmd
    End If
    
    '更新部門名稱
    strCmd = "Update Accrpt41H0 Set R009=(Select A0902 From Acc090 Where ID='" & strUserNum & "' And R001=A0901(+)) Where ID='" & strUserNum & "' "
    adoTaie.Execute strCmd
    
    '區小計(只有北、中所有)
    strCmd = "Insert Into Accrpt41H0 (ID,R001,R002,R003,R004,R005,R006,R007,R008,R009) " & _
                    "Select '" & strUserNum & "',R001||'X',R001||'X',R003,Sum(Nvl(R004,0)),Sum(Nvl(R005,0)),Sum(Nvl(R006,0)),Sum(Nvl(R007,0)),Sum(Nvl(R008,0)),A0902 " & _
                    "From Accrpt41H0,Acc090 Where ID='" & strUserNum & "' And Substr(R001,1,1)='S' And Substr(R001,2,1)<='2' And R001=A0901(+) " & _
                    "Group by R001,A0902,R003"
    adoTaie.Execute strCmd
    
    '所合計
    strCmd = "Insert Into Accrpt41H0 (ID,R001,R002,R003,R004,R005,R006,R007,R008,R009) " & _
                    "Select '" & strUserNum & "',Substr(R001,1,2)||'X',Substr(R001,1,2)||'X',R003,Sum(Nvl(R004,0)),Sum(Nvl(R005,0)),Sum(Nvl(R006,0)),Sum(Nvl(R007,0)),Sum(Nvl(R008,0))," & _
                    "Decode(SubStr(R001,2,1),'1','台北所','2','台中所','3','台南所','高雄所') " & _
                    "From Accrpt41H0 Where ID='" & strUserNum & "' And Substr(R001,1,1)='S' And InStr(R001,'X')=0 " & _
                    "Group by Substr(R001,1,2),Decode(SubStr(R001,2,1),'1','台北所','2','台中所','3','台南所','高雄所'),R003"
    adoTaie.Execute strCmd
    
    '智權部合計
    strCmd = "Insert Into Accrpt41H0 (ID,R001,R002,R003,R004,R005,R006,R007,R008,R009) " & _
                    "Select '" & strUserNum & "','SZZZZ','SZZZZ',R003,Sum(Nvl(R004,0)),Sum(Nvl(R005,0)),Sum(Nvl(R006,0)),Sum(Nvl(R007,0)),Sum(Nvl(R008,0)),'智權部' " & _
                    "From Accrpt41H0 Where ID='" & strUserNum & "' And Substr(R001,1,1)='S' And InStr(R001,'X')=0 " & _
                    "Group by R003"
    adoTaie.Execute strCmd
    '非智權部合計
    strCmd = "Insert Into Accrpt41H0 (ID,R001,R002,R003,R004,R005,R006,R007,R008,R009) " & _
                    "Select '" & strUserNum & "','ZZZZZ','ZZZZZ',R003,Sum(Nvl(R004,0)),Sum(Nvl(R005,0)),Sum(Nvl(R006,0)),Sum(Nvl(R007,0)),Sum(Nvl(R008,0)),'其　他' " & _
                    "From Accrpt41H0 Where ID='" & strUserNum & "' And Substr(R001,1,1)<>'S' And InStr(R001,'X')=0 " & _
                    "Group by R003"
    adoTaie.Execute strCmd
End Sub

Private Sub DataGrid1_AfterColUpdate(ByVal ColIndex As Integer)
    Dim FieldName As String, strUpd As String, strVal As String
    Dim intRow As Integer '目前列數
    
On Error GoTo Checking

    If Adodc1.Recordset.RecordCount = 0 Then Exit Sub
  
    If Right(DataGrid1.Columns(9), 1) = "X" Or DataGrid1.Columns(9) = "SZZZZ" Or _
        DataGrid1.Columns(11) < "3" Or DataGrid1.Columns(11) = "4" Or (DataGrid1.Columns(11) = "3" And (ColIndex <= 2 Or ColIndex = "7")) Then
        Adodc1.Recordset.Requery
        DataGrid1.Scroll 0, dgFirstRow - 1
        DataGrid1.row = Val(dgRow)
        Exit Sub
    End If
    
    With DataGrid1
        intRow = .row
        
        strVal = Val(Replace(.Columns(ColIndex).Text, ",", ""))
        .row = intRow - 2
        strVal = -Val(strVal) + Val(Replace(.Columns(ColIndex).Text, ",", ""))
        .row = intRow - 1
        strVal = Val(strVal) + Val(Replace(.Columns(ColIndex).Text, ",", ""))
        '本月報出不可大於 期初+本月
        If Val(strVal) < 0 Then
            Adodc1.Recordset.Requery
            .Scroll 0, dgFirstRow - 1
            .row = Val(dgRow)
            Exit Sub
        End If
        
        .row = intRow
        FieldName = ""
        If Trim(.Columns(ColIndex).Text) = MsgText(601) Then .Columns(ColIndex).Text = 0
        '避免最後本月報出合計無法確認,故不更新合計
        Adodc1.Recordset.UpdateBatch
        dgRowB = intRow
        
        Select Case ColIndex
            Case 3
                FieldName = "R004"
            Case 4
                FieldName = "R005"
            Case 5
                FieldName = "R006"
            Case 6
                FieldName = "R007"
        End Select

        'Memo 避免最後本月報出合計無法確認,故不更所有 合計 欄
        '更新「個人」期末保留
        strUpd = "Update Accrpt41H0 Set " & FieldName & "=" & _
                        "(Select NS-LS From " & _
                          "(Select R001,R002,Sum(Nvl(" & FieldName & ",0)) NS From Accrpt41H0 Where ID='" & strUserNum & "' And R002='" & .Columns(10).Text & "' And R003<='2' Group by R001,R002), " & _
                          "(Select R001 LS1,R002 LS2, Sum(Nvl(" & FieldName & ",0)) LS From Accrpt41H0 Where ID='" & strUserNum & "' And R002='" & .Columns(10).Text & "' And R003='3' Group by R001,R002) " & _
                        "Where R001=LS1(+) And R002=LS2(+) ) " & _
                      "Where ID='" & strUserNum & "' And R002='" & .Columns(10).Text & "' And R003='4' "
        adoTaie.Execute strUpd
        
        '更新「部門」本月報出
        strUpd = "Update Accrpt41H0 Set " & FieldName & "=" & _
                         "(Select Sum(Nvl(" & FieldName & ",0)) NS From Accrpt41H0 Where ID='" & strUserNum & "' And R001='" & .Columns(9).Text & "' And R003='3') " & _
                       "Where ID='" & strUserNum & "' And R001='" & .Columns(9).Text & "X' And R003='3' "
        adoTaie.Execute strUpd
        '更新「部門」期末保留
        strUpd = "Update Accrpt41H0 Set " & FieldName & "=" & _
                        "(Select NS-LS From " & _
                          "(Select  R001,Sum(Nvl(" & FieldName & ",0)) NS From Accrpt41H0 Where ID='" & strUserNum & "' And R001='" & .Columns(9).Text & "' And R003<='2' Group by R001), " & _
                          "(Select  R001 LS1,Sum(Nvl(" & FieldName & ",0)) LS From Accrpt41H0 Where ID='" & strUserNum & "' And R001='" & .Columns(9).Text & "' And R003='3' Group by R001) " & _
                        "Where R001=LS1(+) ) " & _
                      "Where ID='" & strUserNum & "' And R001='" & .Columns(9).Text & "X' And R003='4' "
        adoTaie.Execute strUpd
        
        '更新「所」本月報出
         strUpd = "Update Accrpt41H0 Set " & FieldName & "=" & _
                         "(Select Sum(Nvl(" & FieldName & ",0)) From Accrpt41H0 Where ID='" & strUserNum & "' And SubStr(R001,1,2)='" & Left(.Columns(9).Text, 2) & "' And InStr(R001,'X')=0 And R003='3') " & _
                       "Where ID='" & strUserNum & "' And SubStr(R001,1,3)='" & Left(.Columns(9).Text, 2) & "X' And R003='3' "
        adoTaie.Execute strUpd
        '更新「所」期末保留
         strUpd = "Update Accrpt41H0 Set " & FieldName & "=" & _
                        "(Select Sum(Nvl(NS,0))-Sum(Nvl(LS,0)) From " & _
                          "(Select R001,Sum(Nvl(" & FieldName & ",0)) NS From Accrpt41H0 Where ID='" & strUserNum & "' And SubStr(R001,1,2)='" & Left(.Columns(9).Text, 2) & "' And InStr(R001,'X')=0 And R003<='2' Group by R001), " & _
                          "(Select R001 LS1,Sum(Nvl(" & FieldName & ",0)) LS From Accrpt41H0 Where ID='" & strUserNum & "' And SubStr(R001,1,2)='" & Left(.Columns(9).Text, 2) & "' And InStr(R001,'X')=0 And R003='3' Group by R001) " & _
                        "Where R001=LS1(+) ) " & _
                      "Where ID='" & strUserNum & "' And SubStr(R001,1,3)='" & Left(.Columns(9).Text, 2) & "X' And R003='4' "
        adoTaie.Execute strUpd
        
        '更新「智權部」本月報出
        strUpd = "Update Accrpt41H0 Set " & FieldName & "=" & _
                        "(Select Sum(Nvl(" & FieldName & ",0)) From Accrpt41H0 Where ID='" & strUserNum & "'  And R001<>'SZZZZ' And InStr(R001,'X')=0 And R003='3') " & _
                      " Where ID='" & strUserNum & "' And R001='SZZZZ' And R003='3' "
        adoTaie.Execute strUpd
        '更新「智權部」期末保留
        strUpd = "Update Accrpt41H0 Set " & FieldName & "=" & _
                        "(Select Sum(Nvl(NS,0))-Sum(Nvl(LS,0)) From " & _
                          "(Select 'SZZZZ' R001,Sum(Nvl(" & FieldName & ",0)) NS From Accrpt41H0 Where ID='" & strUserNum & "'  And R001<>'SZZZZ' And InStr(R001,'X')=0 And R003<='2'), " & _
                          "(Select 'SZZZZ' LS1,Sum(Nvl(" & FieldName & ",0)) LS From Accrpt41H0 Where ID='" & strUserNum & "' And R001<>'SZZZZ' And InStr(R001,'X')=0 And R003='3' ) " & _
                        "Where R001=LS1(+) ) " & _
                      " Where ID='" & strUserNum & "' And R001='SZZZZ' And R003='4' "
        adoTaie.Execute strUpd
       
        Adodc1.Recordset.Requery
        Call ReadSum 'Add by Amy 2019/01/15
        .Scroll 0, dgFirstRow - 1 '設定更新那筆的第一筆
        .row = Val(dgRow) '跳至更新的那一筆
        bolNotUpd = True
    End With
    Exit Sub
       
Checking:
    MsgBox Err.Description, , MsgText(5)
    Screen.MousePointer = vbDefault
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    '輸入年月<=最大筆Axb02則不可修改
    If Val(Left(strAxb02, 5)) > Val(Replace(MaskEdBox1, "/", "")) Then KeyAscii = 0: Exit Sub
   
    If Not (DataGrid1.Columns(11) = "3") Then KeyAscii = 0: Exit Sub
    If Right(DataGrid1.Columns(9), 1) = "X" Or DataGrid1.Columns(9) = "SZZZZ" Then KeyAscii = 0: Exit Sub
    If KeyAscii = 9 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
    If Adodc1.Recordset.RecordCount = 0 Then Exit Sub
    
    dgRow = DataGrid1.row
    dgFirstRow = DataGrid1.FirstRow
   
    Select Case KeyCode
        Case vbKeyReturn
            Select Case DataGrid1.col
                Case 3
                    SendKeys "{RIGHT}"
                Case 4
                    SendKeys "{RIGHT}"
                Case 5
                    SendKeys "{RIGHT}"
                Case 6
                    SendKeys "{DOWN}"
                    For i = 1 To 4
                        SendKeys "{LEFT}"
                    Next i
         End Select
   End Select
End Sub

Private Sub Form_Activate()
    tool3_enabled
End Sub

Private Sub Form_Load()
    Dim strDate As String
    Dim stAxb(0) As String 'Add by Amy 2021/09/23
    
    strFormName = Name
    Me.Width = 9048
    Me.Height = 5700
    Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
    
    Call ClearLabel
    dgRowB = 0
    'Modify by Amy 2017/09/13
    '原:預設上個月最後一個工作日,改預設上個月
    strDate = Left(GetPreMonLastDate(strSrvDate(1)), 5)
    MaskEdBox1.Mask = ""
    MaskEdBox1.Text = Mid(strDate, 1, 3) & "/" & Mid(strDate, 4, 2)
    MaskEdBox1.Mask = MskFormat
    
    bolShowMsg = False
    strA0b01 = GetA0b01(strA0b05)
    OpenTable
    Call ShowBt
    If Val(Replace(MaskEdBox1, "/", "")) <= Val(strA0b05) Then Call CmdRead_Click
    bolShowMsg = True
    'Add by Amy 2021/09/23 判斷每月結餘資料有修改需彈訊息
    Call bolAcc0b1(8, Replace(MaskEdBox1, "/", ""), stAxb())
    If stAxb(0) = "Y" Then
        MsgBox "結餘資料有修改，請確認是否刪除「智權期末結餘保留資料」" & vbCrLf & _
                      "若需刪除請至「智權期末結餘保留資料刪除」作業刪除！"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Add by Amy 2017/09/13 未存檔先存檔
    If bolNotUpd = True Then
        '檢查合計是否有誤
        If ChkGridSum = False Then
            Cancel = True
            Exit Sub
        End If
        If SaveTmp = False Then
            DataGrid1.Visible = True
            Screen.MousePointer = vbDefault
            Cancel = True
            Exit Sub
        End If
    End If
    'end 2017/09/13
    If bolNotSave = True Then
        MsgBox "有修改資料但尚未更正傳票,請按「更正傳票」鈕！", , MsgText(5)
        Cancel = True
        Exit Sub
    'Mark by Amy 2024/06/25  拿掉不彈-婉莘
'    ElseIf bolHasAx210 = False Then
'        MsgBox "請記得過帳！" & vbCrLf & _
'                "在確認專業點數及業務點數相同後, 再通知智權主管寫報告 ！"
    End If
                    
    strFormName = MsgText(601)
    KeyEnter vbKeyEscape
    MenuEnabled
    Call PUB_GetLock("", "Frmacc41h0") 'Add by Amy 2017/09/13
    Set Frmacc41h0 = Nothing
End Sub

Private Sub ClearLabel()
    Dim objLbl As Label
    
    For Each objLbl In Lbl1
        objLbl = ""
    Next
End Sub

'intCmd:0-讀取資料/1-暫存檔案/2-產生傳票/3-更正傳票/4-新增人員
Private Function FormCheck(ByVal intCmd As Integer) As Boolean
    Dim strSql As String
    Dim strLabel As String, bolCancel As Boolean
    Dim strMsg As String
    
    FormCheck = False: bolCancel = False

    'Modify by Amy 2017/09/13 轉期末傳票年月 原輸日期
    strLabel = "轉期末傳票年月"
    If MaskEdBox1.Text = MsgText(601) Or Replace(MaskEdBox1.Text, "/", "") = MsgText(601) Then
        MsgBox strLabel & "不可為空值！", , MsgText(5)
        MaskEdBox1.SetFocus
        Exit Function
    End If
    '智權實績與結餘輸入財務需關閉才可讀取資料
    If Val(Replace(MaskEdBox1.Text, "/", "")) > strA0b05 Then
        MsgBox MaskEdBox1.Text & "月 智權實績與結餘輸入尚未關閉,不可讀取資料！", , MsgText(5)
        SetNoData
        MaskEdBox1.SetFocus
        Exit Function
    End If
    Call MaskEdBox1_Validate(bolCancel)
    If bolCancel = True Then Exit Function
    
    '按「產生傳票」鈕
    If intCmd = 2 Then
        If Val(Replace(MaskEdBox1, "/", "")) < Val(業績自動轉傳票啟用年月) Then
            MsgBox "此程式於 " & 業績自動轉傳票啟用年月 & " 月開始使用！", , MsgText(5)
            MaskEdBox1.SetFocus
            Exit Function
        End If
        If Val(Left(strAxb(9), 5)) >= Val(Replace(MaskEdBox1, "/", "")) Then
            MsgBox "此年月已產生期末傳票！", , MsgText(5)
            MaskEdBox1.SetFocus
            Exit Function
        End If
    End If

    If intCmd >= 1 Then
        '檢查合計是否有誤
        If ChkGridSum = False Then
            Exit Function
        End If
    End If
    'end 2017/09/13

    '判斷A0b10 是否有值
    If intCmd >= 2 Then
        If Pub_GetAcc0b0("a0b10", "1") <> MsgText(601) Then
            MsgBox MsgText(197), , MsgText(5)
            Exit Function
        End If
    End If

    FormCheck = True
End Function

'結餘資料表報出或期未是否有值
Private Function ChkSBVal(ByVal stBS01 As String, ByRef stMsg As String) As Boolean
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String, intQ As Integer
   
    intQ = 1: ChkSBVal = False
   
    'Modify by Amy 2017/11/06 報出或期未是否有值
    strQ = "Select * From SalesBalance Where SB01=" & stBS01 & _
           " And SB04||SB05||SB06||SB07||SB10||SB11||SB12||SB13 is not null"
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        ChkSBVal = True
    End If
    RsQ.Close
End Function

Private Sub OpenTable()

On Error GoTo Checking
    
    If ado41H0.State = adStateOpen Then ado41H0.Close
    ado41H0.CursorLocation = adUseClient
    strQ = "Select * From Accrpt41H0 Where ID='" & strUserNum & "' And Rownum=0 "
    ado41H0.Open strQ, adoTaie, adOpenStatic, adLockReadOnly

   Set Adodc1.Recordset = ado41H0
   
   'Add by Amy 2019/01/15 加智權部合計
   If RsSum.State = adStateOpen Then RsSum.Close
    RsSum.CursorLocation = adUseClient
    strQ = "Select * From Accrpt41H0 Where ID='" & strUserNum & "' And Rownum=0 "
    RsSum.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
   Set Adodc2.Recordset = RsSum
   DataGrid2.Enabled = False
   'end 2019/01/15

Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

Private Function ChkGridSum() As Boolean
    DataGrid1.Visible = False
    
    ChkGridSum = False
    
    With Adodc1.Recordset
        .MoveFirst
        Do While .EOF = False
            If Right(.Fields("R001"), 1) <> "X" And .Fields("R001") <> "SZZZZ" And .Fields("R003") = 3 Then
                If Val(.Fields("Total")) <> Val(.Fields("T")) + Val(.Fields("P")) + Val(.Fields("CFT")) + Val(.Fields("CFP")) Then
                    DataGrid1.SelBookmarks.add .Bookmark
                    DataGrid1.Visible = True
                    MsgBox GetStaffName(.Fields(10), True) & " 本月報出合計不正確,請確認各項的值"
                    Exit Function
                End If
            End If
            .MoveNext
        Loop
    End With
    
    DataGrid1.Visible = True
    ChkGridSum = True
End Function

'儲存傳票資料
Private Function SaveAcc(ByVal intCmd As Integer) As Boolean
    Dim strTmp(3) As String, strField As String
    Dim stA0202 As String, OldData As String, stAx203 As String, stAx204 As String
    
On Error GoTo Checking
    SaveAcc = False
    
    strAcDate = Replace(MaskEdBox1, "/", "")
    
    adoTaie.BeginTrans
    adoTaie.Execute "Update Acc0b0 Set a0b10= '01' Where a0b04 = '1'"
    
    '*** 期末傳票 ***
    '按「更正傳票」鈕(刪除傳票檔資料再新增加)
    If intCmd = 3 Then
        strCmd = "Update Acc020 set A0209=" & Val(strSrvDate(2)) & ",A0210=" & ServerTime & ",A0211='" & strUserNum & "' " & _
                       "Where A0201='1' And (A0202>='" & strAxb(9) & "' And A0202<='" & IIf(strAxb(11) = "", strAxb(10), strAxb(11)) & "' " & _
                       "Or A0202>='" & strAxb(12) & "' And A0202<='" & strAxb(13) & "') "
        adoTaie.Execute strCmd
        
        '期末傳票
        strCmd = "Delete From Acc021 Where Ax201='1' And Ax202>='" & strAxb(9) & "' And Ax202<='" & strAxb(10) & "'"
        adoTaie.Execute strCmd
        '隔月回轉傳票
        strCmd = "Delete From Acc021 Where Ax201='1' And Ax202>='" & strAxb(12) & "' And Ax202<='" & strAxb(13) & "'"
        adoTaie.Execute strCmd
        '轉撥傳票
        'Add by Amy 2017/09/13 需判斷SalesPoint修改日>傳票修改日才改
        If strAxb(11) <> MsgText(601) Then
            bolUpdAxb11 = False
            If bolNotUpdAxb11 = True Then
                bolUpdAxb11 = True
                strCmd = "Delete From Acc021 Where Ax201='1' And Ax202='" & strAxb(11) & "'"
                adoTaie.Execute strCmd
            End If
        End If
        'end 2017/09/13
    End If
    
    '抓取「智權」報出結餘資料產生傳票
    strQ = "Select SB.*,St02 From SalesBalance SB,Staff " & _
                "Where SB01=" & strAcDate & " And SubStr(SB02,1,1)='S' And SB10+SB11+SB12+SB13>0 " & _
                "And SB03=ST01(+) Order by SB02,SB03"
    If adoQ.State = adStateOpen Then adoQ.Close
    adoQ.CursorLocation = adUseClient
    adoQ.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
    If adoQ.RecordCount <> 0 Then
        If intCmd = 3 Then stA0202 = strAxb(9)
        adoQ.MoveFirst
        Do While adoQ.EOF = False
            '依所別寫入傳票
            If Left(OldData, 2) <> Left(adoQ.Fields("SB02"), 2) Then
                '寫入貸方資料
                If OldData <> MsgText(601) Then
                    '取得流水號
                    stAx203 = GetSeqNo("1", stA0202)
                    strCmd = GetCreditSQL(stA0202, stAx203, Left(OldData, 2))
                    If strCmd <> MsgText(601) Then adoTaie.Execute strCmd
                    If intCmd = 3 Then stA0202 = Left(stA0202, 1) & Val(Mid(stA0202, 2)) + 1
                End If
                '取得傳票號
                If intCmd = 2 Then
                    stA0202 = AccAutoNo(MsgText(801), 4, Val(Left(strAcDate, 3)), Val(Mid(strAcDate, 4, 2)))
                    'Modify by Amy 2017/09/13 傳票日原:strAcDate
                    strCmd = "Insert Into Acc020 (a0201,a0202,a0205,a0208,a0206,a0207) " & _
                                    "Values('1','" & stA0202 & "', " & Val(strAxb02) & ",'" & strUserNum & "'," & Val(strSrvDate(2)) & "," & ServerTime & ")"
                    adoTaie.Execute strCmd
                    strCmd = AccSaveAutoNo(MsgText(801), Right(stA0202, 4), Val(Left(strAcDate, 3)), Val(Mid(strAcDate, 4, 2)))
                End If
                If OldData = MsgText(601) Then strAxb(9) = stA0202: Lbl1(0) = strAxb(9)
            End If
                
            For i = 0 To 3
                stAx204 = ""
                '取得流水號
                stAx203 = GetSeqNo("1", stA0202)
                Select Case i
                    Case 0
                        stAx204 = "T"
                    Case 1
                        stAx204 = "P"
                    Case 2
                        stAx204 = "CFT"
                    Case 3
                        stAx204 = "CFP"
                End Select
                If Val(adoQ.Fields(i + 9)) > 0 Then
                    strCmd = "Insert Into Acc021 (ax201,ax202,ax203,ax204,ax205,ax206,ax207,ax209,ax212) " & _
                                "Values('1','" & stA0202 & "', '" & stAx203 & "','" & stAx204 & "','4194'," & _
                                 Val(adoQ.Fields(i + 9)) & ",0,'" & adoQ.Fields("SB03") & "','" & adoQ.Fields("St02") & "/結餘保留')"
                    adoTaie.Execute strCmd
                End If
            Next i
            OldData = adoQ.Fields("SB02")
            adoQ.MoveNext
        Loop
        '寫入最後一筆貸方資料
        stAx203 = GetSeqNo("1", stA0202)
        strCmd = GetCreditSQL(stA0202, stAx203, Left(OldData, 2))
        If strCmd <> MsgText(601) Then adoTaie.Execute strCmd
        
        strAxb(10) = stA0202:  Lbl1(1) = strAxb(10)
    End If
    adoQ.Close
    
    '抓取「非智權」報出結餘資料產生傳票-全部寫成一張傳票,貸方ax213顯示北所ex:D104093821
    strQ = "Select SB.*,St02 From SalesBalance SB,Staff " & _
                "Where SB01=" & strAcDate & " And SubStr(SB02,1,1)<>'S' And SB10+SB11+SB12+SB13>0 " & _
                "And SB03=ST01(+) Order by SB02,SB03"
    If adoQ.State = adStateOpen Then adoQ.Close
    adoQ.CursorLocation = adUseClient
    adoQ.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
    If adoQ.RecordCount <> 0 Then
        '取得傳票號
        If intCmd = 2 Then
            stA0202 = AccAutoNo(MsgText(801), 4, Val(Left(strAcDate, 3)), Val(Mid(strAcDate, 4, 2)))
            'Modify by Amy 2017/09/13 傳票日原:strAcDate
            strCmd = "Insert Into Acc020 (a0201,a0202,a0205,a0208,a0206,a0207) " & _
                            "Values('1','" & stA0202 & "', " & Val(strAxb02) & ",'" & strUserNum & "'," & Val(strSrvDate(2)) & "," & ServerTime & ")"
            adoTaie.Execute strCmd
            strCmd = AccSaveAutoNo(MsgText(801), Right(stA0202, 4), Val(Left(strAcDate, 3)), Val(Mid(strAcDate, 4, 2)))
        Else
            stA0202 = Left(stA0202, 1) & Val(Mid(stA0202, 2)) + 1
        End If
        
        adoQ.MoveFirst
         Do While adoQ.EOF = False
            For i = 0 To 3
                stAx204 = ""
                '取得流水號
                stAx203 = GetSeqNo("1", stA0202)
                Select Case i
                    Case 0
                        stAx204 = "T"
                    Case 1
                        stAx204 = "P"
                    Case 2
                        stAx204 = "CFT"
                    Case 3
                        stAx204 = "CFP"
                End Select
                If Val(adoQ.Fields(i + 9)) > 0 Then
                    strCmd = "Insert Into Acc021 (ax201,ax202,ax203,ax204,ax205,ax206,ax207,ax209,ax212) " & _
                                "Values('1','" & stA0202 & "', '" & stAx203 & "','" & stAx204 & "','4194'," & _
                                 Val(adoQ.Fields(i + 9)) & ",0,'" & adoQ.Fields("SB03") & "','" & adoQ.Fields("St02") & "/結餘保留')"
                    adoTaie.Execute strCmd
                End If
            Next i
            adoQ.MoveNext
        Loop
        '寫入最後一筆貸方資料
        stAx203 = GetSeqNo("1", stA0202)
        strCmd = GetCreditSQL(stA0202, stAx203, "台北所")
        If strCmd <> MsgText(601) Then adoTaie.Execute strCmd
        
        strAxb(10) = stA0202:  Lbl1(1) = strAxb(10)
    End If
    adoQ.Close
    
    '*** 回轉傳票 ***
    OldData = "": stAx203 = "": stA0202 = ""
    strQ = "Select * From Acc021,Staff " & _
                "Where ax201='1' And ax209=st01(+) And ax202>='" & strAxb(9) & "' And ax202<='" & strAxb(10) & "' " & _
                "Order by ax202,Decode(ax206,0,'000',ax203)"
    If adoQ.State = adStateOpen Then adoQ.Close
    adoQ.CursorLocation = adUseClient
    adoQ.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
    If adoQ.RecordCount <> 0 Then
        
        adoQ.MoveFirst
        Do While adoQ.EOF = False
            If OldData <> adoQ.Fields("Ax202") Then
                '取得傳票號
                If intCmd = 3 Then
                    If OldData = MsgText(601) Then
                        stA0202 = strAxb(12)
                    Else
                        stA0202 = Left(stA0202, 1) & Val(Mid(stA0202, 2)) + 1
                    End If
                Else
                    'Modify by Amy 2017/09/13 原傳票日做於strSrvDate(2)
                    stA0202 = AccAutoNo(MsgText(801), 4, Val(Mid(strAxb03, 1, 3)), Val(Mid(strAxb03, 4, 2)))
                    strCmd = "Insert Into Acc020 (a0201,a0202,a0205,a0208,a0206,a0207) " & _
                                    "Values('1','" & stA0202 & "', " & Val(strAxb03) & ",'" & strUserNum & "'," & Val(strSrvDate(2)) & "," & ServerTime & ")"
                    adoTaie.Execute strCmd
                    strCmd = AccSaveAutoNo(MsgText(801), Right(stA0202, 4), Val(Mid(strAxb03, 1, 3)), Val(Mid(strAxb03, 4, 2)))
                    'end 2017/09/13
                End If
                If OldData = MsgText(601) Then strAxb(12) = stA0202: Lbl1(3) = strAxb(12)
            End If
            '取得流水號
            stAx203 = GetSeqNo("1", stA0202)
            '借貸相反
            If Val(adoQ.Fields("Ax206")) > 0 Then
                strCmd = "Insert Into Acc021 (ax201,ax202,ax203,ax204,ax205,ax206,ax207,ax209,ax212) " & _
                                "Values('1','" & stA0202 & "', '" & stAx203 & "','" & adoQ.Fields("Ax204") & "','" & adoQ.Fields("Ax205") & "'," & _
                                    "0," & Val(adoQ.Fields("Ax206")) & ",'" & adoQ.Fields("Ax209") & "','" & adoQ.Fields("Ax212") & "')"
            Else
                strCmd = "Insert Into Acc021 (ax201,ax202,ax203,ax204,ax205,ax206,ax207,ax212,ax213) " & _
                                "Values('1','" & stA0202 & "', '" & stAx203 & "','" & adoQ.Fields("Ax204") & "','" & adoQ.Fields("Ax205") & "'," & _
                                    Val(adoQ.Fields("Ax207")) & ",0,'" & adoQ.Fields("Ax212") & "','" & adoQ.Fields("Ax213") & "')"
            End If
            adoTaie.Execute strCmd
            OldData = adoQ.Fields("Ax202")
            adoQ.MoveNext
        Loop
        strAxb(13) = stA0202: Lbl1(4) = strAxb(13)
    End If
     
    '*** 轉撥傳票 ***
    'Modify by Amy 2017/09/13 傳票日原做於系統日改「轉期末傳票日」, 判斷有修改或新增傳票才依轉出點數從小者扣至扣完,轉入人員部門由使用者輸入,若按「更正傳票」判斷SB TB修改日>傳票日才更正(避免非改轉撥需一直輸)
    strQ = GetPoint_SP(strAcDate, strAcDate, , , "SP40", False, Me.Name, , True)
    If adoQ.State = adStateOpen Then adoQ.Close
    adoQ.CursorLocation = adUseClient
    adoQ.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
    If adoQ.RecordCount > 0 And (strAxb(11) = MsgText(601) Or bolUpdAxb11 = True) Then
        bolTrans = True
       
        adoQ.MoveFirst
        '若產生傳票產生時並無結餘轉撥資料,但做更正傳票前又加了結餘轉撥資料則再多加一張傳票
        If intCmd = 2 Or (intCmd = 3 And strAxb(11) = MsgText(601)) Then
            stA0202 = AccAutoNo(MsgText(801), 4, Val(Left(strAcDate, 3)), Val(Mid(strAcDate, 4, 2)))
            strAxb(11) = stA0202
            strCmd = "Insert Into Acc020 (a0201,a0202,a0205,a0208,a0206,a0207) " & _
                            "Values('1','" & stA0202 & "', " & Val(strAxb02) & ",'" & strUserNum & "'," & Val(strSrvDate(2)) & "," & ServerTime & ")"
            adoTaie.Execute strCmd
            strCmd = AccSaveAutoNo(MsgText(801), Right(stA0202, 4), Val(Left(strAcDate, 3)), Val(Mid(strAcDate, 4, 2)))
        'bolUpdAxb11 = True 修改且需更新
        Else
            stA0202 = strAxb(11)
            strCmd = "Delete From Acc021 Where Ax201='1' And Ax202='" & stA0202 & "' "
            adoTaie.Execute strCmd
        End If
        'Modify by Amy 2017/12/05
        stAx203 = GetSeqNo("1", stA0202)
        Do While adoQ.EOF = False
            strTmp(0) = "": strTmp(1) = "": strTmp(2) = "": strTmp(3) = ""
            
            '結餘轉撥<0 借方
            If Val(adoQ.Fields("SP40")) < 0 Then
                strTmp(0) = Abs(adoQ.Fields("SP40")) '借方
                '轉出點數從小者扣至扣完
                strQ2 = "Select * From (" & _
                        GetBalanceSQL(3, strDate_S, strDate_E, adoQ.Fields("SP02")) & _
                        " Union All " & GetBalanceSQL(4, strDate_S, strDate_E, adoQ.Fields("SP02")) & _
                        " Union All " & GetBalanceSQL(5, strDate_S, strDate_E, adoQ.Fields("SP02")) & _
                        " Union All " & GetBalanceSQL(6, strDate_S, strDate_E, adoQ.Fields("SP02")) & _
                        ") Where stVal is not null Order by StVal Asc"
                If adoQ2.State = adStateOpen Then adoQ2.Close
                adoQ2.Open strQ2, adoTaie, adOpenStatic, adLockReadOnly
                If adoQ2.RecordCount <> 0 Then
                    adoQ2.MoveFirst
                    Do While Val(strTmp(0)) > 0
                        Select Case Val("" & adoQ2.Fields("Ord"))
                            Case 3
                                strTmp(2) = "T"
                            Case 4
                                strTmp(2) = "P"
                            Case 5
                                strTmp(2) = "CFT"
                            Case 6
                                strTmp(2) = "CFP"
                        End Select
                            
                        If Val(strTmp(0)) <= Val(adoQ2.Fields("StVal")) Then
                            strTmp(1) = strTmp(0)
                            strTmp(0) = 0
                        Else
                            strTmp(1) = adoQ2.Fields("StVal")
                            strTmp(0) = Val(strTmp(0)) - Val(strTmp(1))
                        End If
                        'Modify by Amy 2022/05/13 取代換行
                        strCmd = "Insert Into Acc021 (ax201,ax202,ax203,ax204,ax205,ax206,ax207,ax209,ax212) " & _
                                  "Values('1','" & stA0202 & "', '" & stAx203 & "','" & strTmp(2) & "','4194'," & Val(strTmp(1)) & ",0" & _
                                  ",'" & adoQ.Fields("SP02") & "','" & GetStaffName(adoQ.Fields("SP02"), True) & "：轉撥" & Replace("" & adoQ.Fields("SP41"), vbCrLf, "") & "')"
                        adoTaie.Execute strCmd
                        
                        stAx203 = GetSeqNo("1", stA0202)
                        If adoQ2.EOF = False Then
                            adoQ2.MoveNext
                        Else
                            strTmp(0) = 0
                        End If
                    Loop
                        
                End If
                adoQ2.Close
            '結餘轉撥>0 貸方
            Else
                strTmp(3) = adoQ.Fields("SP40") '貸方
                strCmd = "Insert Into Acc021 (ax201,ax202,ax203,ax204,ax205,ax206,ax207,ax209,ax212) " & _
                            "Values('1','" & stA0202 & "', '" & stAx203 & "','','4194',0," & Val(strTmp(3)) & _
                            ",'" & adoQ.Fields("SP02") & "','" & GetStaffName(adoQ.Fields("SP02"), True) & "：轉撥" & adoQ.Fields("SP41") & "')"
                adoTaie.Execute strCmd
                stAx203 = GetSeqNo("1", stA0202)
            End If
            adoQ.MoveNext
        Loop
        
        Lbl1(2) = strAxb(11)
    End If
    'end 2017/09/13
    
    '更新Acc01b對應傳票日期及號碼
    strExc(0) = "": strCmd = "": strField = ""
    
    'Modify by Amy 2017/09/13 拿掉Axb03(實績轉傳時存)
    If bol0b1HasDt = False Then
        For i = LBound(strAxb) To UBound(strAxb)
            strField = strField & ",axb" & IIf(i < 10, "0", "") & i
            strCmd = strCmd & ",'" & strAxb(i) & "'"
        Next i
        strCmd = "Insert Into Acc0b1 (axb01" & strField & ") " & _
                 "Values(" & strDate & strCmd & ")"
    Else
        For i = LBound(strAxb) To UBound(strAxb)
            strField = strField & ",axb" & IIf(i < 10, "0", "") & i & "=" & "'" & strAxb(i) & "'"
        Next i
        strCmd = "Update Acc0b1 Set " & Mid(strField, 2) & " Where axb01=" & Val(Left(strAcDate, 5))
    End If
    'end 2017/09/13
    
    If strCmd <> MsgText(601) Then
        adoTaie.Execute strCmd
    End If
    
    adoTaie.Execute "Update Acc0b0 Set a0b10= null Where a0b04 = '1'"
    adoTaie.CommitTrans
    
    SaveAcc = True
    Adodc1.Recordset.MoveFirst
    Exit Function
    
Checking:
    adoTaie.RollbackTrans
    If adoQ.State = adStateOpen Then adoQ.Close
    MsgBox Err.Description, , MsgText(5)
End Function

'組新增貸方語法
Private Function GetCreditSQL(ByVal stA0202 As String, ByVal stAx203 As String, ByVal stDept As String) As String
    Dim adoSum As New ADODB.Recordset
    Dim strSum As String
    
    GetCreditSQL = ""
    If stDept = "台北所" Then
        strSum = "Select '" & stDept & "' Zone,Sum(SB10)+Sum(SB11)+Sum(SB12)+Sum(SB13) Tol " & _
                        "From SalesBalance Where SB01=" & Left(strAcDate, 5) & " And SubStr(SB02,1,1)<>'S' "
    Else
        strSum = "Select Decode(SubStr(SB02,2,1),'1','台北所','2','台中所','3','台南所','高雄所') Zone,Sum(SB10)+Sum(SB11)+Sum(SB12)+Sum(SB13) Tol " & _
                        "From SalesBalance Where SB01=" & Left(strAcDate, 5) & " And SubStr(SB02,1,2)='" & Left(stDept, 2) & "' " & _
                        "Group by Decode(SubStr(SB02,2,1),'1','台北所','2','台中所','3','台南所','高雄所') "
    End If
    If adoSum.State = adStateOpen Then adoSum.Close
    adoSum.CursorLocation = adUseClient
    adoSum.Open strSum, adoTaie, adOpenStatic, adLockReadOnly
    If adoSum.RecordCount <> 0 Then
        With adoSum
            GetCreditSQL = "Insert Into Acc021 (ax201,ax202,ax203,ax204,ax205,ax206,ax207,ax212,ax213) " & _
                            "Values('1','" & stA0202 & "', '" & stAx203 & "','TOT','2493'," & _
                            "0 ," & Val(.Fields("Tol")) & ",'" & .Fields("Zone") & "/結餘保留','" & .Fields("Zone") & "')"
        End With
    End If
    adoSum.Close
End Function

Private Function GetValue(pFieldN As String) As Integer
   Dim jj As Integer
 
    For jj = 1 To UBound(strF)
       If UCase(strF(jj)) = UCase(pFieldN) Then
          GetValue = jj
          Exit For
       End If
    Next jj
End Function

Private Sub SetField(ByRef Wks As Worksheet, Optional ByVal bolLast As Boolean = False)
    Dim strTmp As String
    
    If bolLast = True Then
        Wks.Range(Chr(intField + GetValue("T")) & intTitle + 1 & ":" & Chr(intField + UBound(strF)) & intCounter).NumberFormatLocal = "#,##0.0"
        For i = GetValue("T") To GetValue("CFP")
            Select Case strF(i)
                Case "T"
                    strTmp = "Ｔ"
                Case "P"
                    strTmp = "Ｐ"
                Case "CFT"
                    strTmp = "ＣＦＴ"
                Case "CFP"
                    strTmp = "ＣＦＰ"
            End Select
            Wks.Range(Chr(i + intField) & "1").Value = strTmp
        Next i
        Wks.Range(Chr(intField) & intTitle + 1 & ":" & Chr(intField + UBound(strF)) & intCounter).Font.Size = 11
    Else
        For i = LBound(strF) To UBound(strF)
            Wks.Range(Chr(i + intField) & "1").Value = strF(i)
            Wks.Columns(Chr(i + intField) & ":" & Chr(i + intField)).ColumnWidth = intWidth(i)
            Wks.Range(Chr(i + intField) & "1").HorizontalAlignment = xlCenter
        Next i
    End If
End Sub

Private Sub ShowBt()
    CmdSaveTmp.Enabled = False
    CmdSaveAcc(2).Enabled = False
    CmdSaveAcc(3).Enabled = False
    CmdExcel.Enabled = False
    
    Erase strAxb
    ClearLabel
    'Modify by Amy 2017/09/13
    If Val(Replace(MaskEdBox1, "/", "")) < Val(業績自動轉傳票啟用年月) Then Exit Sub
    If Val(Replace(MaskEdBox1, "/", "")) > Left(strSrvDate(2), 5) Then Exit Sub
    strAxb02 = GetAxb0203(1, Replace(MaskEdBox1, "/", "")) '當月期末保留傳票日
    strAxb03 = GetAxb0203(2, Replace(MaskEdBox1, "/", "")) '隔月初轉回傳票日
 
    '抓智權點數傳票起始值
    bol0b1HasDt = bolAcc0b1(2, Replace(MaskEdBox1, "/", ""), strAxb())
    '畫面日期小於等於業績輸入關閉年月
    If Val(Replace(MaskEdBox1, "/", "")) <= Val(strA0b05) And strAxb02 <> MsgText(601) Then
        If Val(Left(strAxb02, 5)) <= Val(Replace(MaskEdBox1, "/", "")) Then
            If bol0b1HasDt = True Then
                bolHasAx210 = Pub_ChkAxbPost(strAxb(9), strAxb(10), strAxb(11))
                '傳票未過帳或當月傳票已產生,更正傳票鈕才可使用
                If strAxb(9) = MsgText(601) Then
                    CmdSaveTmp.Enabled = True
                    CmdSaveAcc(2).Enabled = True
                ElseIf bolHasAx210 = False Then
                    CmdSaveTmp.Enabled = True
                    CmdSaveAcc(3).Enabled = True
                End If
            Else
                CmdSaveTmp.Enabled = True
                CmdSaveAcc(2).Enabled = True
            End If
        End If
    End If
    If Val(Replace(MaskEdBox1, "/", "")) <= Val(strA0b05) And strAxb02 <> MsgText(601) Then
        CmdExcel.Enabled = True
    End If
    'end 2017/09/13
    Lbl1(0) = strAxb(9): Lbl1(1) = strAxb(10): Lbl1(2) = strAxb(11): Lbl1(3) = strAxb(12): Lbl1(4) = strAxb(13)
    'Add by Amy 2017/09/13 顯示 Axb02/03
    Lbl1(5) = ChangeTStringToTDateString(strAxb02): Lbl1(6) = ChangeTStringToTDateString(strAxb03)
    'Add by Amy 2018/06/13
    If CmdSaveTmp.Enabled = False And CmdSaveAcc(2).Enabled = False And CmdSaveAcc(3).Enabled = False Then
        DataGrid1.Enabled = False
    Else
        DataGrid1.Enabled = True
    End If
End Sub

Private Function SaveTmp() As Boolean

On Error GoTo ErrHand
 
    DataGrid1.Visible = False
    SaveTmp = False
    '抓取畫面資料更新至智權報出餘餘資料表
    strQ = "Select N.*,LT,LP,LCFT,LCFP From " & _
              "(Select R001 as NDep,R002 as NUser,R004 as NT,R005 as NP,R006 as NCFT,R007 as NCFP From Accrpt41H0 " & _
              "Where ID='" & strUserNum & "' And SubStr(R001,length(R001),1)<>'X' And R001<>'SZZZZ' And R003=3 ) N, " & _
              "(Select R001 as LDep,R002 as LUser,R004 as LT,R005 as LP,R006 as LCFT,R007 as LCFP From Accrpt41H0 " & _
              "Where ID='" & strUserNum & "' And SubStr(R001,length(R001),1)<>'X' And R001<>'SZZZZ' And R003=4 ) L " & _
              "Where NDep=LDep And NUser=LUser"
    If adoQ.State = adStateOpen Then adoQ.Close
    adoQ.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
    If adoQ.RecordCount <> 0 Then
        DataGrid1.Visible = False
        With adoQ
            .MoveFirst
            adoTaie.BeginTrans
            Do While .EOF = False
                strCmd = "Update SalesBalance Set SB04=" & Val(.Fields("NT")) & ",SB05=" & Val(.Fields("NP")) & ",SB06=" & Val(.Fields("NCFT")) & ",SB07=" & Val(.Fields("NCFP")) & _
                            ",SB08=" & Val(strSrvDate(2)) & ",SB09=" & ServerTime & _
                            ",SB10=" & Val(.Fields("LT")) & ",SB11=" & Val(.Fields("LP")) & ",SB12=" & Val(.Fields("LCFT")) & ",SB13=" & Val(.Fields("LCFP")) & _
                            " Where SB01=" & Left(strAcDate, 5) & " And SB02='" & .Fields("NDep") & "' And SB03='" & .Fields("NUser") & "' "
                adoTaie.Execute strCmd
                .MoveNext
            Loop
            adoTaie.CommitTrans
            bolNotUpd = False
            Adodc1.Recordset.MoveFirst
        End With
        DataGrid1.Visible = True
    End If
    SaveTmp = True
    Exit Function
    
ErrHand:
    adoTaie.RollbackTrans
    DataGrid1.Visible = True
   MsgBox Err.Description, , MsgText(5)
End Function

Private Function bolNotSave() As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim intQ As Integer
    Dim stSQL As String, stSBDT As String, stSPDT As String
    
    bolNotSave = False
    If strAcDate = MsgText(601) Or Adodc1.Recordset.RecordCount = 0 Then Exit Function

    '取得智權報出結餘(SalesBalance)修改時間
    strQ = "Select Max(SB08||Decode(length(SB09),5,'0'||SB09,SB09)) as DT From SalesBalance " & _
               "Where SB01=" & strAcDate & " Having Max(SB08||Decode(length(SB09),5,'0'||SB09,SB09)) is not null"
    intQ = 1
    Set rsTmp = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        stSBDT = "" & rsTmp.Fields("DT")
    End If
    rsTmp.Close
    '取得智權輸入結餘轉撥(SalesPoint)修改時間
    strQ = "Select Max(SP43-19110000||Decode(length(SP44),5,'0'||SP44,SP44)) as DT From SalesPoint " & _
               "Where SP01=" & Val(strAcDate) + 191100 & " Having Max(SP43-19110000||Decode(length(SP44),5,'0'||SP44,SP44)) is not null"
    intQ = 1
    Set rsTmp = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        stSPDT = "" & rsTmp.Fields("DT")
    End If
    rsTmp.Close
    
    If stSBDT = "" And stSPDT = "" Then Exit Function
       
    'Modify by Amy 2017/09/13 避免轉撥傳票為更新傳票時產生(傳票號未連號),故拆開判斷
    strQ = ""
    If strAxb(11) <> MsgText(601) Then
        stSQL = "Union Select A0206||Decode(length(A0207),5,'0'||A0207,A0207) as DT From Acc020 Where A0201='1' And A0202='" & strAxb(11) & "' " & _
                  "And A0206||Decode(length(A0207),5,'0'||A0207,A0207)<" & Val(stSPDT) & " And A0209||A0210 is null " & _
                "Union Select A0209||Decode(length(A0210),5,'0'||A0210,A0210) as DT From Acc020 Where A0201='1' And A0202='" & strAxb(11) & "' " & _
                  "And A0209||Decode(length(A0210),5,'0'||A0210,A0210)<" & Val(stSPDT) & " And A0209||A0210 is not null "

    End If
    strQ = "Select A0206||Decode(length(A0207),5,'0'||A0207,A0207) as DT From Acc020 Where A0201='1' And A0202>='" & strAxb(9) & "' And A0202<='" & strAxb(10) & "' " & _
              "And A0206||Decode(length(A0207),5,'0'||A0207,A0207)<" & Val(stSBDT) & " And A0209||A0210 is null " & _
     "Union Select A0209||Decode(length(A0210),5,'0'||A0210,A0210) as DT From Acc020 Where A0201='1' And A0202>='" & strAxb(9) & "' And A0202<='" & strAxb(10) & "' " & _
              "And A0209||Decode(length(A0210),5,'0'||A0210,A0210)<" & Val(stSBDT) & " And A0209||A0210 is not null " & _
            stSQL
    'end 2017/09/13
    intQ = 1
    Set rsTmp = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        bolNotSave = True
    End If
    rsTmp.Close
End Function

'Add by Amy 2017/09/13
Private Sub SetNoData()
    OpenTable
    Adodc1.Recordset.Requery
End Sub

'判斷SalesPoint轉撥已修改但轉票是否更新
Private Function bolNotUpdAxb11() As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim intQ As Integer
    Dim stSPDT As String
    
    bolNotUpdAxb11 = False
    strQ = "Select Max(SP43||SP44) as DT From SalesPoint " & _
            "Where SP01=" & Val(strAcDate) + 191100 & " Having Max(SP43||SP44) is not null"
    intQ = 1
    Set rsTmp = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        stSPDT = "" & rsTmp.Fields("DT")
    End If
    rsTmp.Close
    If stSPDT = "" Then Exit Function
    
    strQ = "Select A0206||A0207 as DT From Acc020 Where A0201='1' " & _
              "And A0202='" & strAxb(11) & "' " & _
              "And ((A0206||A0207 <" & Val(stSPDT) & " And A0209||A0210 is null) " & _
              "Or (A0209||A0210 <" & Val(stSPDT) & " And A0209||A0210 is not null)) "
    intQ = 1
    Set rsTmp = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        bolNotUpdAxb11 = True
    End If
    rsTmp.Close
End Function

Private Sub MaskEdBox1_Validate(Cancel As Boolean)
    Dim strMsg As String

    If MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Or bolShowMsg = False Then
        Exit Sub
    End If
    
    If bolBT = False Then SetBTAndLab
    strMsg = "轉期末傳票日期"
    If Val(Replace(MaskEdBox1, "/", "")) < Val(業績自動轉傳票啟用年月) Then
        MsgBox strMsg & "無舊資料！", , MsgText(5)
        SetNoData
        Cancel = True
        If Val(業績自動轉傳票啟用年月) > Left(strSrvDate(2), 5) Then MaskEdBox1.SetFocus
        Exit Sub
    End If
    If IsDate(ChangeTStringToWDateString(Replace(MaskEdBox1, "/", "") & "01")) = False Then
        MsgBox strMsg & "輸入錯誤！", , MsgText(5)
        SetNoData
        Cancel = True
        MaskEdBox1.SetFocus
        Exit Sub
    End If
    If Val(Replace(MaskEdBox1, "/", "")) >= Val(Left(strSrvDate(2), 5)) Then
        MsgBox strMsg & "需小於系統月份！", , MsgText(5)
        Cancel = True
        SetNoData
        MaskEdBox1.SetFocus
        Exit Sub
    End If
   
End Sub

Private Sub SetBTAndLab()
    CmdSaveTmp.Enabled = False
    CmdSaveAcc(2).Enabled = False
    CmdSaveAcc(3).Enabled = False
    CmdExcel.Enabled = False
    ClearLabel
End Sub

'Add by Amy 2019/01/15 智權部合計
Private Sub ReadSum()
    Dim strQ As String
    
    strQ = "Select Decode(R003,1,R009,'') Dept,Decode(R003,1,'合計','') StName,Decode(R003,1,'期初保留',Decode(R003,2,'本月放出',Decode(R003,3,'本月報出','期末保留'))) Type," & _
                "R004 T,R005 P,R006 CFT,R007 CFP,R008 Total,ID,R001,R002,R003 From Accrpt41h0 " & _
                "Where ID='" & strUserNum & "' And R002='SZZZZ' " & _
                "Order by R003"
    If RsSum.State = adStateOpen Then RsSum.Close
    RsSum.CursorLocation = adUseClient
    RsSum.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
    RsSum.MoveFirst
    Set Adodc2.Recordset = RsSum
    Adodc2.Recordset.Requery
End Sub
