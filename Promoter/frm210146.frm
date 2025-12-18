VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm210146 
   BorderStyle     =   1  '單線固定
   Caption         =   "請款作業及應收查詢"
   ClientHeight    =   5760
   ClientLeft      =   48
   ClientTop       =   348
   ClientWidth     =   9420
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   9420
   Begin VB.TextBox txtFormat 
      Height          =   300
      Left            =   5400
      MaxLength       =   1
      TabIndex        =   47
      Top             =   4776
      Width           =   348
   End
   Begin VB.CommandButton cmdCall 
      Caption         =   "應收帳款查詢"
      Height          =   465
      Left            =   7950
      TabIndex        =   43
      Top             =   780
      Width           =   1275
   End
   Begin VB.CheckBox chkPrintedOnly 
      Caption         =   "排除未列印收據"
      Height          =   195
      Left            =   6300
      TabIndex        =   8
      Top             =   960
      Width           =   1590
   End
   Begin VB.ComboBox cboComp 
      Height          =   300
      ItemData        =   "frm210146.frx":0000
      Left            =   4950
      List            =   "frm210146.frx":0002
      Style           =   2  '單純下拉式
      TabIndex        =   7
      Top             =   900
      Width           =   1125
   End
   Begin VB.TextBox txtDate 
      Height          =   285
      Index           =   0
      Left            =   1080
      MaxLength       =   7
      TabIndex        =   5
      Top             =   900
      Width           =   1095
   End
   Begin VB.TextBox txtDate 
      Height          =   285
      Index           =   1
      Left            =   2400
      MaxLength       =   7
      TabIndex        =   6
      Top             =   900
      Width           =   1095
   End
   Begin VB.ComboBox Combo2 
      Height          =   276
      Index           =   1
      Left            =   5250
      TabIndex        =   40
      Text            =   "Combo2"
      Top             =   4440
      Width           =   3000
   End
   Begin VB.ComboBox Combo2 
      Height          =   276
      Index           =   0
      Left            =   6108
      TabIndex        =   39
      Text            =   "Combo2"
      Top             =   5376
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "列印備註(&M)"
      Height          =   345
      Index           =   6
      Left            =   360
      TabIndex        =   35
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   120
      TabIndex        =   32
      Top             =   4440
      Width           =   4185
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   705
         Style           =   2  '單純下拉式
         TabIndex        =   33
         Top             =   180
         Width           =   3390
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   315
         Index           =   1
         Left            =   75
         TabIndex        =   34
         Top             =   240
         Width           =   765
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   8040
      Top             =   3000
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
   Begin VB.CommandButton cmdok 
      Caption         =   "選擇應收客戶(&C)"
      Height          =   345
      Index           =   5
      Left            =   5355
      TabIndex        =   10
      Top             =   90
      Width           =   1650
   End
   Begin VB.TextBox txtTot 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H8000000F&
      Height          =   270
      Index           =   1
      Left            =   8280
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   1275
      Width           =   1050
   End
   Begin VB.TextBox txtTot 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H8000000F&
      Height          =   270
      Index           =   0
      Left            =   6705
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   1275
      Width           =   1050
   End
   Begin VB.TextBox txtTot 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H8000000F&
      Height          =   270
      Index           =   5
      Left            =   8235
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   4080
      Width           =   1050
   End
   Begin VB.TextBox txtTot 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H8000000F&
      Height          =   270
      Index           =   4
      Left            =   6795
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   4080
      Width           =   915
   End
   Begin VB.TextBox txtTot 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H8000000F&
      Height          =   270
      Index           =   3
      Left            =   5220
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   4080
      Width           =   1050
   End
   Begin VB.TextBox txtTot 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H8000000F&
      Height          =   270
      Index           =   2
      Left            =   3645
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   4080
      Width           =   1050
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "查詢(&Q)"
      Height          =   345
      Index           =   1
      Left            =   4320
      TabIndex        =   9
      Top             =   90
      Width           =   975
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "列印(&P)"
      Enabled         =   0   'False
      Height          =   345
      Index           =   2
      Left            =   7065
      TabIndex        =   11
      Top             =   90
      Width           =   1155
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   345
      Index           =   0
      Left            =   8280
      TabIndex        =   12
      Top             =   90
      Width           =   975
   End
   Begin VB.TextBox txtCustNo 
      Height          =   285
      Index           =   0
      Left            =   5130
      MaxLength       =   9
      TabIndex        =   3
      Text            =   "X"
      Top             =   570
      Width           =   1095
   End
   Begin VB.TextBox txtCustNo 
      Height          =   285
      Index           =   1
      Left            =   6480
      MaxLength       =   9
      TabIndex        =   4
      Text            =   "X"
      Top             =   570
      Width           =   1095
   End
   Begin VB.TextBox txtSales 
      Height          =   285
      Left            =   1080
      MaxLength       =   6
      TabIndex        =   0
      Top             =   240
      Width           =   915
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frm210146.frx":0004
      Height          =   2475
      Left            =   90
      TabIndex        =   13
      Top             =   1560
      Width           =   9195
      _ExtentX        =   16214
      _ExtentY        =   4360
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      BackColor       =   -2147483624
      Enabled         =   -1  'True
      HeadLines       =   1
      RowHeight       =   14
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   23
      BeginProperty Column00 
         DataField       =   "選取"
         Caption         =   "選取"
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
         DataField       =   "a0k11C"
         Caption         =   "公司別"
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
         DataField       =   "收據號碼"
         Caption         =   "收據號碼"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "收據日期"
         Caption         =   "收據日期"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "ee/mm/dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "本所案號"
         Caption         =   "本所案號"
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
         DataField       =   "案件性質"
         Caption         =   "案件性質"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "申請國家"
         Caption         =   "申請國家"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "服務費"
         Caption         =   "服務費"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "規費"
         Caption         =   "規費"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "扣繳"
         Caption         =   "扣繳"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "扣繳金額"
         Caption         =   "扣繳金額"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "案件名稱"
         Caption         =   "案件名稱"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column12 
         DataField       =   "a0j01"
         Caption         =   "收文號"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column13 
         DataField       =   "amt1"
         Caption         =   "未收服務費"
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
      BeginProperty Column14 
         DataField       =   "amt2"
         Caption         =   "未收規費"
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
      BeginProperty Column15 
         DataField       =   "a0j07"
         Caption         =   "是否合併"
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
      BeginProperty Column16 
         DataField       =   "a0k19"
         Caption         =   "列印次數"
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
      BeginProperty Column17 
         DataField       =   "a0k03"
         Caption         =   "客戶編號"
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
      BeginProperty Column18 
         DataField       =   "a0k04"
         Caption         =   "收據抬頭"
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
      BeginProperty Column19 
         DataField       =   "CU04"
         Caption         =   "申請人名稱"
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
      BeginProperty Column20 
         DataField       =   "a0k11"
         Caption         =   "收據公司別代號"
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
      BeginProperty Column21 
         DataField       =   "a0k32"
         Caption         =   "a0k32"
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
      BeginProperty Column22 
         DataField       =   "a0k40"
         Caption         =   "a0k40"
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
      SplitCount      =   2
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         Locked          =   -1  'True
         Size            =   85
         BeginProperty Column00 
            Alignment       =   2
            ColumnAllowSizing=   -1  'True
            Locked          =   -1  'True
            ColumnWidth     =   432
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   599.811
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   1152
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   887.811
         EndProperty
         BeginProperty Column04 
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   1068.094
         EndProperty
         BeginProperty Column05 
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   840.189
         EndProperty
         BeginProperty Column06 
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   612.284
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            Object.Visible         =   -1  'True
            ColumnWidth     =   659.906
         EndProperty
         BeginProperty Column08 
            Alignment       =   1
            Object.Visible         =   -1  'True
            ColumnWidth     =   648
         EndProperty
         BeginProperty Column09 
            Alignment       =   2
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   432
         EndProperty
         BeginProperty Column10 
            Alignment       =   1
            ColumnAllowSizing=   -1  'True
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   720
         EndProperty
         BeginProperty Column11 
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   852.095
         EndProperty
         BeginProperty Column12 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
            ColumnWidth     =   828.284
         EndProperty
         BeginProperty Column13 
            Alignment       =   1
            Locked          =   -1  'True
            Object.Visible         =   0   'False
            ColumnWidth     =   792
         EndProperty
         BeginProperty Column14 
            Alignment       =   1
            Locked          =   -1  'True
            Object.Visible         =   0   'False
            ColumnWidth     =   959.811
         EndProperty
         BeginProperty Column15 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column16 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column17 
            Object.Visible         =   -1  'True
            ColumnWidth     =   959.811
         EndProperty
         BeginProperty Column18 
            Object.Visible         =   -1  'True
            ColumnWidth     =   2808
         EndProperty
         BeginProperty Column19 
            Object.Visible         =   -1  'True
            ColumnWidth     =   2232
         EndProperty
         BeginProperty Column20 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column21 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column22 
            Object.Visible         =   0   'False
         EndProperty
      EndProperty
      BeginProperty Split1 
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         Locked          =   -1  'True
         RecordSelectors =   0   'False
         Size            =   180
         BeginProperty Column00 
            Alignment       =   2
            ColumnAllowSizing=   -1  'True
            Locked          =   -1  'True
            Object.Visible         =   0   'False
            ColumnWidth     =   432
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
            ColumnWidth     =   599.811
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
            ColumnWidth     =   1152
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   887.811
         EndProperty
         BeginProperty Column04 
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   1068.094
         EndProperty
         BeginProperty Column05 
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   840.189
         EndProperty
         BeginProperty Column06 
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   612.284
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            Object.Visible         =   -1  'True
            ColumnWidth     =   659.906
         EndProperty
         BeginProperty Column08 
            Alignment       =   1
            Object.Visible         =   -1  'True
            ColumnWidth     =   648
         EndProperty
         BeginProperty Column09 
            Alignment       =   2
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   432
         EndProperty
         BeginProperty Column10 
            Alignment       =   1
            ColumnAllowSizing=   -1  'True
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   720
         EndProperty
         BeginProperty Column11 
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   852.095
         EndProperty
         BeginProperty Column12 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
            ColumnWidth     =   828.284
         EndProperty
         BeginProperty Column13 
            Alignment       =   1
            Locked          =   -1  'True
            Object.Visible         =   0   'False
            ColumnWidth     =   792
         EndProperty
         BeginProperty Column14 
            Alignment       =   1
            Locked          =   -1  'True
            Object.Visible         =   0   'False
            ColumnWidth     =   959.811
         EndProperty
         BeginProperty Column15 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column16 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column17 
            Object.Visible         =   -1  'True
            ColumnWidth     =   959.811
         EndProperty
         BeginProperty Column18 
            Object.Visible         =   -1  'True
            ColumnWidth     =   2808
         EndProperty
         BeginProperty Column19 
            Object.Visible         =   -1  'True
            ColumnWidth     =   2232
         EndProperty
         BeginProperty Column20 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column21 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column22 
            Object.Visible         =   0   'False
         EndProperty
      EndProperty
   End
   Begin VB.Label lblFormat 
      AutoSize        =   -1  'True
      Caption         =   "列印方向：　　 (1:橫印, 2:直印)"
      Height          =   180
      Left            =   4476
      TabIndex        =   46
      Top             =   4824
      Width           =   2508
   End
   Begin MSForms.ComboBox Combo3 
      Height          =   336
      Left            =   1080
      TabIndex        =   1
      Top             =   240
      Width           =   1920
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "3387;593"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.CheckBox ChkExcel 
      Height          =   320
      Left            =   7920
      TabIndex        =   45
      Top             =   450
      Visible         =   0   'False
      Width           =   1280
      BackColor       =   12648384
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "2258;564"
      Value           =   "1"
      Caption         =   "產生Excel檔"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtTitle 
      Height          =   300
      Left            =   1080
      TabIndex        =   2
      Top             =   562
      Width           =   2715
      VariousPropertyBits=   671105051
      Size            =   "4789;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblSalesName 
      Height          =   300
      Left            =   2070
      TabIndex        =   44
      Top             =   270
      Width           =   1470
      VariousPropertyBits=   27
      Size            =   "2593;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "公司別："
      Height          =   180
      Index           =   90
      Left            =   4185
      TabIndex        =   42
      Top             =   960
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收據日期："
      Height          =   180
      Index           =   23
      Left            =   135
      TabIndex        =   41
      Top             =   930
      Width           =   900
   End
   Begin VB.Line Line2 
      X1              =   2220
      X2              =   2400
      Y1              =   1050
      Y2              =   1050
   End
   Begin VB.Label lblOffice 
      Caption         =   "台一智慧"
      ForeColor       =   &H00FF0000&
      Height          =   252
      Index           =   1
      Left            =   3480
      TabIndex        =   38
      Top             =   5376
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label lblOffice 
      Caption         =   "1 專利商標"
      ForeColor       =   &H00FF0000&
      Height          =   252
      Index           =   0
      Left            =   5148
      TabIndex        =   37
      Top             =   5400
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "匯款帳號"
      Height          =   180
      Index           =   11
      Left            =   4476
      TabIndex        =   36
      Top             =   4488
      Width           =   720
   End
   Begin VB.Label lblMemo 
      AutoSize        =   -1  'True
      Caption         =   "符號說明： ◎未列印 △智權公司 ＃已開立INVOICE"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   0
      Left            =   132
      TabIndex        =   31
      Top             =   1320
      Width           =   4080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "總計"
      Height          =   180
      Index           =   10
      Left            =   7830
      TabIndex        =   27
      Top             =   4125
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "扣繳"
      Height          =   180
      Index           =   9
      Left            =   6390
      TabIndex        =   25
      Top             =   4125
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "規費"
      Height          =   180
      Index           =   8
      Left            =   4815
      TabIndex        =   23
      Top             =   4125
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "服務費"
      Height          =   180
      Index           =   7
      Left            =   3060
      TabIndex        =   22
      Top             =   4125
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "規費"
      Height          =   180
      Index           =   5
      Left            =   7875
      TabIndex        =   21
      Top             =   1320
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "點數"
      Height          =   180
      Index           =   4
      Left            =   6300
      TabIndex        =   20
      Top             =   1320
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "所有資料合計"
      Height          =   180
      Index           =   3
      Left            =   5130
      TabIndex        =   19
      Top             =   1320
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "點選資料合計"
      Height          =   180
      Index           =   6
      Left            =   1890
      TabIndex        =   17
      Top             =   4125
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收據抬頭："
      Height          =   180
      Index           =   1
      Left            =   135
      TabIndex        =   16
      Top             =   615
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "客戶編號："
      Height          =   180
      Index           =   2
      Left            =   4185
      TabIndex        =   15
      Top             =   615
      Width           =   900
   End
   Begin VB.Line Line1 
      X1              =   6255
      X2              =   6435
      Y1              =   705
      Y2              =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   180
      Index           =   0
      Left            =   135
      TabIndex        =   14
      Top             =   285
      Width           =   900
   End
End
Attribute VB_Name = "frm210146"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/05/04 為了方便主管查看不同客戶，應收帳款查詢改為只呼叫畫面不限制帶入客戶代號
'Memo by Lydia 2021/07/27 智權-調整財務系統(20200909) 'Memo by Lydia 2021/08/27 上線
'原「請款明細表」標題修改為「請款作業及應收查詢」。
'前三列為定格畫面。
'客戶代號均修改為”客戶編號”。
'客戶名稱均修改為”申請人名稱”。
'收據編號均修改為”收據號碼”。
'單據日期及收據開立日期均修改為”收據日期”。
'國別均修改為”申請國家”。
'加入”應收帳款查詢”功能鍵，按此功能鍵後，呈現的畫面如下：其中申請人改為”申請人名稱”，部門改為”業務區。
'end 2021/07/27
'Memo by Lydia 2021/07/16 DataGrid分割顯示的Split0必須要顯示全部要顯示的欄位，而Split1可以隱藏前面的欄位
'Memo by Lydia 2021/07/16 改成Form2.0 ; lblSalesName、txtTitle、DataGrid1改字型=新細明體-ExtB; Printer列印未改
'Memo by Lydia 2019/07/01 表單名稱:客戶請款明細表=>請款明細表
'Modify by Lydia 2014/9/25 (原frm210141) 新增"客戶請款明細表"
Option Explicit

Dim m_A4421 As String, m_A4427 As String
Dim adoTmp As New ADODB.Recordset
Dim stST05 As String
Dim txtSalesArea As String, txtSalesArea1 As String
Dim m_blnColOrderAsc As Boolean
Dim m_mouseRow As Integer, m_MouseCol As Integer
Dim m_A4425 As String '溢收款處理方式
'Add by Amy 2014/05/21
Dim bolSpecMan As Boolean  '是否為特殊設定檔人員
Dim strSpecCode As String '特殊設定檔設定代號

'Add by Lydia 列印報表用
Dim m_rs As New ADODB.Recordset
Dim strTemp(1 To 20) As String
'Modified by Lydia 2021/07/27 mSeqNo從Integer改成String
Dim mSeqNo As String
Dim minDate As String, maxDate As String, mFee1 As Double, mFee2 As Double, mdFee1 As Double
Dim t_A As String, t_B As String

Dim PLeft(1 To 13) As Integer, iPrint As Integer, iPage As Integer
Private Const ciTitleFontSize = 16, ciFontSize = 10
Private Const ciStartX = 500, ciStartY = 500, ciColGap = 200
Dim lngPageHeight As Long, lngPageWidth As Long, lngLineHeight As Long
'Added by Lydia 2015/06/17 列印備註
Public rptMemo As String
'Added by Lydia 2015/06/17 設定匯款資料
Dim sra(0 To 6) As String
Dim m_selA1 As String, m_selA2 As String, tA0k11 As String
Dim m_strListPer As String 'Add By Sindy 2020/7/28
Dim strPrinter As String 'Added by Morgan 2020/10/30
'Added by Lydia 2021/07/27
Dim mPrevForm As Form
'Added by Lydia 2022/02/17 Excel使用
Dim intRow As Integer, intField As Integer, intTitleR As Integer
Dim intUL As Integer  'U單號的資料列位置
Dim bolOpenXls As Boolean '已開啟Excel
Dim strFileN As String, strAllF As String, strWidth As String
Dim strField, intWidth
Dim iX As Integer
'Add By Sindy 2023/6/12
Dim arrID ', stST15 As String
'Dim bolAreaMan As Boolean '下拉選單有區主管
'2023/6/12 END
'Modified by Lydia 2024/09/16 增加長春人造X74310000,長春石油X74310010,大連化工X74310020
Private Const strSpecExcel As String = "X38805030,X82532010,X74310000,X74310010,X74310020"  'Added by Lydia 2023/11/08 可產生Excel的客戶:資策會(X38805030)、力成(X82532010)


Private Sub GetPleft()
   Printer.FontName = "新細明體"
   Printer.Font.Size = ciFontSize
   Printer.Font.Bold = False
   Printer.Font.Underline = False

   PLeft(1) = ciStartX '收據日期
   PLeft(2) = PLeft(1) + Printer.TextWidth(String(4, "　")) + ciColGap
   PLeft(3) = PLeft(2) + Printer.TextWidth(String(5, "　")) + ciColGap
   PLeft(4) = PLeft(3) + Printer.TextWidth(String(4, "　")) + ciColGap
   PLeft(5) = PLeft(4) + Printer.TextWidth(String(6, "　")) + ciColGap
   PLeft(6) = PLeft(5) + Printer.TextWidth(String(11, "　")) + ciColGap
   PLeft(7) = PLeft(6) + Printer.TextWidth(String(6, "　")) + ciColGap
   PLeft(8) = PLeft(7) + Printer.TextWidth(String(4, "　")) + ciColGap
   PLeft(9) = PLeft(8) + Printer.TextWidth(String(5, "　")) + ciColGap
   PLeft(10) = PLeft(9) + Printer.TextWidth(String(5, "　")) + ciColGap
   PLeft(11) = PLeft(10) + Printer.TextWidth(String(5, "　")) + ciColGap
   PLeft(12) = PLeft(11) + Printer.TextWidth(String(5, "　")) + ciColGap
   PLeft(13) = PLeft(12) + Printer.TextWidth(String(5, "　")) + ciColGap
End Sub
'Modified by Lydia 2017/03/23  iExtraLines As Integer = 3 => 2
Private Sub PrintNewLine(Optional ByVal bolSubtotal As Boolean = True, Optional ByVal iExtraLines As Integer = 2)

   iPrint = iPrint + lngLineHeight
   If iPrint >= (lngPageHeight - iExtraLines * lngLineHeight) Then
      Printer.CurrentX = ciStartX
      Printer.CurrentY = iPrint

      iPage = iPage + 1
      Printer.NewPage
      PrintHeader
   End If
    
End Sub

Private Sub cmdok_Click(Index As Integer)
Dim bolOK As Boolean 'Added by Lydia 2024/09/16

   Select Case Index
   Case 0 '結束
      Unload Me
   Case 1 '查詢
      'Added by Lydia 2015/06/17
      If rptMemo <> "" Then
         If MsgBox("是否保留列印備註?", vbYesNo + vbInformation) = vbNo Then rptMemo = ""
      End If
      
      Screen.MousePointer = vbHourglass
      doQuery
      Screen.MousePointer = vbDefault
   Case 2 '列印
      'Added by Lydia 2015/06/17 判斷公司別
      If m_selA1 = "Y" And Combo2(0).Text = "" Then
         'Modified by Lydia 2020/03/31
         'MsgBox "專利法律未選擇匯款帳號!!": Exit Sub
         MsgBox "專利商標未選擇匯款帳號!!": Exit Sub
      ElseIf m_selA2 = "Y" And Combo2(1).Text = "" Then
         'Modified by Lydia 2020/03/31
         'MsgBox "專利商標未選擇匯款帳號!!": Exit Sub
         MsgBox "台一智慧未選擇匯款帳號!!": Exit Sub
      End If
      
      If TxtValidate = True Then
          'Grid資料寫入RDataFactory
          Screen.MousePointer = vbHourglass
          'Modified by Lydia 2021/07/27 改在模組中取得序號 + mSeqNo
          Set Adodc1.Recordset = PUB_CreateRecordset(adoTmp, , , , Me.Name, mSeqNo)
      Else
          Exit Sub
      End If
      cmdok(2).Enabled = False 'Added by Lydia 2024/04/25
      'Added by Lydia 2022/02/17  資策會(X38805030)請款文件Excel
      'Modified by Lydia 2023/11/08 改判斷
      'If txtCustNo(0) = "X38805030" And ChkExcel.Visible = True And ChkExcel.Value = True Then
      If InStr(strSpecExcel, txtCustNo(0)) > 0 And txtCustNo(0) <> "" And ChkExcel.Visible = True And ChkExcel.Value = True Then
         'Modified by Lydia 2024/09/16 增加不同客戶不同格式
          ' If rptMemo = "" Then  '列印備註:預設第一筆案件資料
         '       rptMemo = GetGridSel
         '  End If
         '   bolOpenXls = False
         '   If SaveXLS_1 = True Then
         '       MsgBox "檔案已產生於" & vbCrLf & _
         '                     " [" & strExcelPath & strFileN & "] "
         '   End If
         bolOpenXls = False
         bolOK = False
         If InStr("資策會X38805030,力成X82532010", txtCustNo(0)) > 0 Then
            If rptMemo = "" Then  '列印備註:預設第一筆案件資料
               rptMemo = GetGridSel
            End If
            bolOK = SaveXLS_1
         ElseIf InStr("長春人造X74310000,長春石油X74310010,大連化工X74310020", txtCustNo(0)) > 0 Then
            bolOK = SaveXLS_2
         End If
         If bolOK = True Then
            MsgBox "檔案已產生於" & vbCrLf & _
                          " [" & strExcelPath & strFileN & "] "
         End If
         'end 2024/09/16
      Else
           ChkExcel.Visible = False
           ChkExcel.Value = False
      'end 2022/02/17
           'Added by Lydia 2024/04/25
           If txtFormat.Visible = True And txtFormat = "2" Then
               bolOpenXls = False
               PUB_SetOsDefaultPrinter Combo1 '切換Word/Excel印表機
               PUB_RestorePrinter Combo1
               Call PrintExcelMain
               PUB_SetOsDefaultPrinter strPrinter '切換Word/Excel印表機
               PUB_RestorePrinter strPrinter
           Else
           'end 2024/04/25
               Call PrtData
           End If
      End If 'Added by Lydia 2022/02/17
      cmdok(2).Enabled = True 'Added by Lydia 2024/04/25
      Screen.MousePointer = vbDefault

   Case 5 '選擇應收客戶
      ShowCustList
   'Added by Lydia 2015/06/17 +列印備註
   Case 6
      'Added by Lydia 2022/02/17  資策會(X38805030)請款文件:預設第一筆案件資料
      'Modified by Lydia 2023/11/08 改判斷
      'If rptMemo = "" And txtCustNo(0) = "X38805030" And ChkExcel.Visible = True And ChkExcel.Value = True Then
      If rptMemo = "" And InStr(strSpecExcel, txtCustNo(0)) > 0 And txtCustNo(0) <> "" And ChkExcel.Visible = True And ChkExcel.Value = True Then
          rptMemo = GetGridSel
      End If
      'end 2022/02/17
      
      Set frm880004.mPreForm = Me
      frm880004.iStiu = 2
      'Added by Lydia 2024/10/23 力成X82532010：增加申請案號、發明人 >>超過5行
      If InStr("力成X82532010", txtCustNo(0)) > 0 Then
         frm880004.m_TempList = 6
      End If
      'end 2024/10/23
      frm880004.Caption = "列印備註"
      frm880004.Show vbModal
   'end 2015/06/17
   End Select
   
End Sub

Private Sub ShowCustList()
   Dim strCon As String, bolOK As Boolean
   Dim stVTB1 As String, stVTB2 As String
   
   If txtSales = "" Then
      MsgBox "智權人員不可空白！", vbExclamation
      Exit Sub
   End If
   
   'index用a0k20較慢,略過
   '已收款
   'Modify By Sindy 2020/5/20 排除顯示案號TT-999999 : and a0j02<>'TT999999000'
   stVTB1 = "select a0j13 X1,a0j01 X2" & _
      ",nvl(sum(nvl(a1u04,0)+nvl(a1u07,0)-nvl(a1u08,0)),0) X3" & _
      ",nvl(sum(nvl(a1u05,0)+nvl(a1u09,0)-nvl(a1u10,0)),0) X4" & _
      ",nvl(sum(a1u06),0) X5 From acc0k0, acc0j0, acc1u0" & _
      " where a0k20='" & txtSales & "'" & strCon & _
      " and nvl(a0k09,0)=0 and (nvl(a0k06,0)+nvl(a0k07,0)) > (nvl(a0k17,0)+nvl(a0k18,0))" & _
      " and a0j13(+)=a0k01 and a1u02(+)=a0j13 and a1u03(+)=a0j01 and a0j02<>'TT999999000'" & _
      " group by a0j13,a0j01"
   'Modify By Sindy 2020/5/15 + 增加抓取法律所案源介紹人為此智權人員
   'Modified by Morgan 2025/3/17 調整語法
   'stVTB1 = stVTB1 & " union " & _
            "select a0j13 X1,a0j01 X2" & _
      ",nvl(sum(nvl(a1u04,0)+nvl(a1u07,0)-nvl(a1u08,0)),0) X3" & _
      ",nvl(sum(nvl(a1u05,0)+nvl(a1u09,0)-nvl(a1u10,0)),0) X4" & _
      ",nvl(sum(a1u06),0) X5" & _
      " From acc0k0, acc0j0, acc1u0, lawofficesource, caseprogress" & _
      " where instr(LOS04,'" & txtSales & "')>0" & strCon & _
      " and nvl(a0k09,0)=0 and (nvl(a0k06,0)+nvl(a0k07,0)) > (nvl(a0k17,0)+nvl(a0k18,0))" & _
      " and a0j13(+)=a0k01 and a1u02(+)=a0j13 and a1u03(+)=a0j01 and a0j02<>'TT999999000'" & _
      " and LOS15(+)=cp162 and cp162 is not null and cp09(+)=a0j01" & _
      " group by a0j13,a0j01"
   stVTB1 = stVTB1 & " union " & _
            "select a0j13 X1,a0j01 X2" & _
      ",nvl(sum(nvl(a1u04,0)+nvl(a1u07,0)-nvl(a1u08,0)),0) X3" & _
      ",nvl(sum(nvl(a1u05,0)+nvl(a1u09,0)-nvl(a1u10,0)),0) X4" & _
      ",nvl(sum(a1u06),0) X5" & _
      " From lawofficesource, caseprogress, acc0j0, acc0k0, acc1u0" & _
      " where instr(LOS04,'" & txtSales & "')>0" & strCon & _
      " and cp162(+)=LOS15 and cp162 is not null and a0j01(+)=cp09 and a0j02<>'TT999999000'" & _
      " and a0k01(+)=a0j13 and nvl(a0k09,0)=0 and (nvl(a0k06,0)+nvl(a0k07,0)) > (nvl(a0k17,0)+nvl(a0k18,0))" & _
      " and a1u02(+)=a0j13 and a1u03(+)=a0j01" & _
      " group by a0j13,a0j01"
   'end 2025/3/17
   '2020/5/15 END
   
   '已繳款未收款
   'Modify By Sindy 2020/5/20 排除顯示案號TT-999999 : and a0j02<>'TT999999000'
   stVTB2 = "select a0j13 Y1,a0j01 Y2" & _
      ",sum(AXD06) Y3,sum(AXD07)Y4,sum(AXD08) Y5" & _
      " From acc0k0, acc0j0, acc441,acc440" & _
      " where a0k20='" & txtSales & "'" & strCon & _
      " and nvl(a0k09,0)=0 and (nvl(a0k06,0)+nvl(a0k07,0)) > (nvl(a0k17,0)+nvl(a0k18,0))" & _
      " and a0j13(+)=a0k01 and axd04(+)=a0j13 and axd05(+)=a0j01" & _
      " and a4401(+)=axd01 and a4402(+)=axd02 and a4403(+)=axd03 and a4416 is null and a0j02<>'TT999999000'" & _
      " group by a0j13,a0j01"
   'Modify By Sindy 2020/5/15 + 增加抓取法律所案源介紹人為此智權人員
   'Modified by Morgan 2025/3/17 調整語法
   'stVTB2 = stVTB2 & " union " & _
            "select a0j13 Y1,a0j01 Y2" & _
      ",sum(AXD06) Y3,sum(AXD07)Y4,sum(AXD08) Y5" & _
      " From acc0k0, acc0j0, acc441,acc440, lawofficesource, caseprogress" & _
      " where instr(LOS04,'" & txtSales & "')>0" & strCon & _
      " and nvl(a0k09,0)=0 and (nvl(a0k06,0)+nvl(a0k07,0)) > (nvl(a0k17,0)+nvl(a0k18,0))" & _
      " and a0j13(+)=a0k01 and axd04(+)=a0j13 and axd05(+)=a0j01" & _
      " and a4401(+)=axd01 and a4402(+)=axd02 and a4403(+)=axd03 and a4416 is null and a0j02<>'TT999999000'" & _
      " and LOS15(+)=cp162 and cp162 is not null and cp09(+)=a0j01" & _
      " group by a0j13,a0j01"
   stVTB2 = stVTB2 & " union " & _
            "select a0j13 Y1,a0j01 Y2" & _
      ",sum(AXD06) Y3,sum(AXD07)Y4,sum(AXD08) Y5" & _
      " From lawofficesource, caseprogress, acc0j0, acc0k0, acc441, acc440" & _
      " where instr(LOS04,'" & txtSales & "')>0" & strCon & _
      " and cp162(+)=LOS15 and cp162 is not null and a0j01(+)=cp09 and a0j02<>'TT999999000'" & _
      " and a0k01(+)=a0j13 and nvl(a0k09,0)=0 and (nvl(a0k06,0)+nvl(a0k07,0)) > (nvl(a0k17,0)+nvl(a0k18,0))" & _
      " and axd04(+)=a0j13 and axd05(+)=a0j01" & _
      " and a4401(+)=axd01 and a4402(+)=axd02 and a4403(+)=axd03 and a4416 is null" & _
      " group by a0j13,a0j01"
   'end 2025/3/17
   '2020/5/15 END
   
   'Modified by Lydia 2024/04/01 改用acc0j0為基準
   'Modified by Morgan 2025/3/17 還原(因改後變慢且查詢結果應該還是相同)
   strExc(0) = "select distinct a0k04,a0k03,cu04" & _
      " from (" & stVTB1 & ") X,(" & stVTB2 & ") Y,acc0j0,acc0k0,caseprogress,customer" & _
      " where Y1(+)=X1 and Y2(+)=X2 and (a0j09-X3-nvl(Y3,0)>0 or a0j10-X4-nvl(Y4,0)>0) and a0j13(+)=X1 and a0j01(+)=X2 and a0k01(+)=a0j13" & _
      " and cp09(+)=a0j01 and cp79>0 and cu01(+)=substr(a0k03,1,8) and cu02(+)=substr(a0k03,9)" & _
      " order by 1,2,3"
   'strExc(0) = "select distinct a0k04,a0k03,cu04" & _
      " from (" & stVTB1 & ") X,(" & stVTB2 & ") Y,acc0j0,acc0k0,caseprogress,customer" & _
      " where a0j01=X2(+) and a0j13=x1(+) and a0j01=y2(+) and a0j13=y1(+) and (a0j09-X3-nvl(Y3,0)>0 or a0j10-X4-nvl(Y4,0)>0) and a0k01(+)=a0j13" & _
      " and cp09(+)=a0j01 and cp79>0 and cu01(+)=substr(a0k03,1,8) and cu02(+)=substr(a0k03,9)" & _
      " order by 1,2,3"
   'Modified by Lydia 2021/07/16
   'intI = 1
   intI = 0
   Screen.MousePointer = vbHourglass
   Set adoTmp = ClsLawReadRstMsg(intI, strExc(0))
   Screen.MousePointer = vbDefault
   
   If adoTmp.RecordCount > 0 Then 'Added by Lydia 2021/07/16 加判斷
       frm210141_4.m_Frm146_4 = True 'Add by Lydia 2014/9/25
    
       With frm210141_4
          'Modify by Amy 2014/06/16 +FormName 改暫存TB
          'Modified by Lydia 2021/07/27 改TableName
          'Set .Adodc1.Recordset = PUB_CreateRecordset(adoTmp, , , , .Name)
          Set .Adodc1.Recordset = PUB_CreateRecordset(adoTmp, , , , "frm210141_4")
         .Show vbModal
       End With
       If Me.Tag = "1" Then
          cmdok_Click 1
       End If
   End If 'Added by Lydia 2021/07/16 加判斷
End Sub

Private Function TxtValidate() As Boolean
   Dim bCancel As Boolean
   Dim tmpArr1 As Variant, tmpArr2 As Variant  'Added by Lydia 2023/11/13
   '至少勾選一筆
   If Val(Format(txtTot(2))) = 0 And Val(Format(txtTot(3))) = 0 And Val(Format(txtTot(4))) = 0 And Val(Format(txtTot(5))) = 0 Then
      MsgBox "請先選取資料！", vbExclamation
      Exit Function
   End If
   
   'Added by Lydia 2023/11/13 檢查同一INVOICE編號是否有勾選
   strExc(0) = "": strExc(1) = "": strExc(2) = "": strExc(3) = ""
   strExc(4) = "": strExc(5) = ""
   Set adoTmp = Adodc1.Recordset.Clone
   With adoTmp
      .MoveFirst
      Do While Not .EOF
         If "" & .Fields("a0k40") <> "" Then
            If "" & .Fields("選取") = "Y" Then
              If InStr(strExc(0) & ",", "" & .Fields("a0k40")) = 0 Then strExc(0) = strExc(0) & "," & .Fields("a0k40")
              If InStr(strExc(1) & ",", "" & .Fields("a0j01")) = 0 Then strExc(1) = strExc(1) & "," & .Fields("a0j01")
            Else
              strExc(2) = strExc(2) & "," & .Fields("a0k40")
              strExc(3) = strExc(3) & "," & .Fields("a0j01")
            End If
         End If
         .MoveNext
      Loop
   End With
   If strExc(0) <> "" And strExc(3) <> "" Then
      strExc(2) = Mid(strExc(2), 2)
      tmpArr1 = Split(strExc(2), ",")
      strExc(3) = Mid(strExc(3), 2)
      tmpArr2 = Split(strExc(3), ",")
      For intI = 0 To UBound(tmpArr1)
         If Trim(tmpArr1(intI)) <> "" Then
            If InStr(strExc(0), tmpArr1(intI)) > 0 Then
               strExc(4) = strExc(4) & "," & tmpArr1(intI)
               strExc(5) = strExc(5) & "," & tmpArr2(intI)
            End If
         End If
      Next intI
   End If
   If strExc(5) <> "" Then
      MsgBox "同一INVOICE編號有未勾選的收據！"
      Exit Function
   End If
   'end 2023/11/13
   
   'Added by Lydia 2024/04/25
   If txtFormat.Visible = True And Trim(txtFormat) = "" Then
      MsgBox "請輸入列印方向！"
      Exit Function
   End If
   'end 2024/04/25
   
   TxtValidate = True
   
End Function

Private Sub doQuery(Optional pNoMsg As Boolean = False, Optional pNoReset As Boolean = False)

   Dim strCon As String
   Dim stVTB1 As String, stVTB2 As String
'Added by Lydia 2016/03/24 列印特殊收據項目
   Dim lngAmount1 As Long '服務費
   Dim lngAmount2 As Long '規費
   Dim strA0K01 As String
   Dim strA0j22 As String
   Dim strCaseNo As String 'Added by Lydia 2016/11/01 案號
   Dim lngSFee As Long, lngOFee As Long 'Added by Lydia 2017/07/25
   Dim strMidSql As String 'Added by Lydia 2021/07/27
   Dim bolCancel As Boolean 'Add By Sindy 2023/6/12
   
   'Add By Sindy 2023/8/29 從接洽單做應收帳款查詢,不檢查此段權限
   If cmdok(5).Tag = "" Then
   '2023/8/29 END
      'Add By Sindy 2023/6/12
      If Combo3.Visible = True Then
         bolCancel = False
         Call Combo3_Validate(bolCancel)
         If bolCancel = True Then
            Exit Sub
         End If
      End If
      'Add by Amy 2020/03/25 +有下拉選單
      If Combo3.Visible = True Then
         Call Combo3_LostFocus 'Add By Sindy 2020/7/15 讓人員按Enter,須再啟動此函數,txtSales欄位值才會置換
         If Combo3 = MsgText(601) Then
             Call Combo3_Validate(bolCancel)
             If bolCancel = True Then
                 Combo3.SetFocus
                 Exit Sub
             End If
         ElseIf txtSales = MsgText(601) Then
             txtSales = Mid(Combo3, 1, Val(InStr(Combo3, " ")) - 1)
         End If
      End If
      Call txtSales_Validate(bolCancel)
      If bolCancel = True Then
         'Modify by Amy 2020/03/25 +有下拉選單
         If Combo3.Visible = True Then
            Combo3.SetFocus
         'Modified by Lydia 2021/05/20 排除隱藏
         'ElseIf txtSales.Enabled = True Then
         ElseIf txtSales.Enabled = True And txtSales.Visible = True Then
            txtSales.SetFocus
            txtSales_GotFocus
         End If
         Exit Sub
      End If
      '2023/6/12 END
   End If
   
   'Modify By Sindy 2023/8/17 + And cmdok(5).Tag = "": 從接洽單做應收帳款查詢,僅依客戶編號做查詢
   If txtSales = "" And cmdok(5).Tag = "" Then
      MsgBox "智權人員不可空白！", vbExclamation
      Exit Sub
   ElseIf cmdok(5).Tag <> "" Then
      If InStr(cmdok(5).Tag, Left(txtCustNo(0), 8)) = 0 Then
         MsgBox "客戶編號僅可輸入此接洽單的申請人做查詢！", vbExclamation
         Exit Sub
      End If
   End If
   
   If txtTitle <> "" Then
      strCon = strCon & " and a0k04 like '" & ChgSQL(txtTitle) & "%'"
   End If
   If txtCustNo(0) <> "" And txtCustNo(0) <> "X" Then
      strCon = strCon & " and a0k03>='" & txtCustNo(0) & "'"
   End If
   If txtCustNo(1) <> "" And txtCustNo(1) <> "X" Then
      strCon = strCon & " and a0k03<='" & txtCustNo(1) & "'"
   End If
   
   'Added by Morgan 2015/7/16
   If txtDate(0) <> "" Then
      strCon = strCon & " and a0k02>=" & Val(txtDate(0))
   End If
   If txtDate(1) <> "" Then
      strCon = strCon & " and a0k02<=" & Val(txtDate(1))
   End If
   If chkPrintedOnly.Value = 1 Then
      strCon = strCon & " and a0k19>0"
   End If
   If cboComp.ListIndex > 0 Then
      If cboComp.ListIndex = 1 Then
         strCon = strCon & " and a0k11='1'"
      ElseIf cboComp.ListIndex = 2 Then
         strCon = strCon & " and a0k11='2'"
      ElseIf cboComp.ListIndex = 3 Then
         strCon = strCon & " and a0k11='J'"
      'Added by Lydia 2020/03/31
      ElseIf cboComp.ListIndex = 4 Then
         strCon = strCon & " and a0k11='L'"
      'end 2020/03/31
      End If
   End If
   'end 2015/7/16
   'Modified by Lydia 2016/03/24 加上判斷是否為列印特殊收據項目
   '已收款
   'stVTB1 = "select a0j13 X1,a0j01 X2" & _
      ",nvl(sum(nvl(a1u04,0)+nvl(a1u07,0)-nvl(a1u08,0)),0) X3" & _
      ",nvl(sum(nvl(a1u05,0)+nvl(a1u09,0)-nvl(a1u10,0)),0) X4" & _
      ",nvl(sum(a1u06),0) X5,nvl(sum(a1u07),0) X6,nvl(sum(a1u09),0) X7" & _
      " From acc0k0, acc0j0, acc1u0 where a0k20='" & txtSales & "'" & strCon & _
      " and nvl(a0k09,0)=0 and (nvl(a0k06,0)+nvl(a0k07,0)) > (nvl(a0k17,0)+nvl(a0k18,0))" & _
      " and a0j13(+)=a0k01 and a1u02(+)=a0j13 and a1u03(+)=a0j01" & _
      " group by a0j13,a0j01"
      
   '已繳款未收款
   'stVTB2 = "select a0j13 Y1,a0j01 Y2" & _
      ",sum(AXD06) Y3,sum(AXD07)Y4,sum(AXD08) Y5" & _
      " From acc0k0, acc0j0, acc441,acc440" & _
      " where a0k20='" & txtSales & "'" & strCon & _
      " and nvl(a0k09,0)=0 and (nvl(a0k06,0)+nvl(a0k07,0)) > (nvl(a0k17,0)+nvl(a0k18,0))" & _
      " and a0j13(+)=a0k01 and axd04(+)=a0j13 and axd05(+)=a0j01" & _
      " and a4401(+)=axd01 and a4402(+)=axd02 and a4403(+)=axd03 and a4416 is null" & _
      " group by a0j13,a0j01"
   '已收款
   'Modify By Sindy 2020/5/20 排除顯示案號TT-999999 : and a0j02<>'TT999999000'
   'Modify By Sindy 2023/8/17 調整 1=1" & IIf(txtSales <> "", " and a0k20='" & txtSales & "'", "")
   stVTB1 = "select a0j13 X1,a0j01 X2,nvl(a0j22,'') X8,nvl(a0j25,'') X9" & _
      ",nvl(sum(nvl(a1u04,0)+nvl(a1u07,0)-nvl(a1u08,0)),0) X3" & _
      ",nvl(sum(nvl(a1u05,0)+nvl(a1u09,0)-nvl(a1u10,0)),0) X4" & _
      ",nvl(sum(a1u06),0) X5,nvl(sum(a1u07),0) X6,nvl(sum(a1u09),0) X7" & _
      " From acc0k0, acc0j0, acc1u0 where 1=1" & IIf(txtSales <> "", " and a0k20='" & txtSales & "'", "") & strCon & _
      " and nvl(a0k09,0)=0 and (nvl(a0k06,0)+nvl(a0k07,0)) > (nvl(a0k17,0)+nvl(a0k18,0))" & _
      " and a0j13(+)=a0k01 and a1u02(+)=a0j13 and a1u03(+)=a0j01 and a0j02<>'TT999999000'" & _
      " group by a0j13,a0j01,a0j22,a0j25"
   'Modify By Sindy 2020/5/15 + 增加抓取法律所案源介紹人為此智權人員
   'Modify By Sindy 2023/8/17 調整 1=1" & IIf(txtSales <> "", " and instr(LOS04,'" & txtSales & "')>0", "")
   stVTB1 = stVTB1 & " union " & _
            "select a0j13 X1,a0j01 X2,nvl(a0j22,'') X8,nvl(a0j25,'') X9" & _
      ",nvl(sum(nvl(a1u04,0)+nvl(a1u07,0)-nvl(a1u08,0)),0) X3" & _
      ",nvl(sum(nvl(a1u05,0)+nvl(a1u09,0)-nvl(a1u10,0)),0) X4" & _
      ",nvl(sum(a1u06),0) X5,nvl(sum(a1u07),0) X6,nvl(sum(a1u09),0) X7" & _
      " From acc0k0, acc0j0, acc1u0, lawofficesource, caseprogress" & _
      " where 1=1" & IIf(txtSales <> "", " and instr(LOS04,'" & txtSales & "')>0", "") & strCon & _
      " and nvl(a0k09,0)=0 and (nvl(a0k06,0)+nvl(a0k07,0)) > (nvl(a0k17,0)+nvl(a0k18,0))" & _
      " and a0j13(+)=a0k01 and a1u02(+)=a0j13 and a1u03(+)=a0j01 and a0j02<>'TT999999000'" & _
      " and LOS15(+)=cp162 and cp162 is not null and cp09(+)=a0j01" & _
      " group by a0j13,a0j01,a0j22,a0j25"
   '2020/5/15 END
      
   '已繳款未收款
   'Modify By Sindy 2020/5/20 排除顯示案號TT-999999 : and a0j02<>'TT999999000'
   'Modify By Sindy 2023/8/17 調整 1=1" & IIf(txtSales <> "", " and a0k20='" & txtSales & "'", "")
   stVTB2 = "select a0j13 Y1,a0j01 Y2,nvl(a0j22,'') Y6,nvl(a0j25,'') Y7" & _
      ",sum(AXD06) Y3,sum(AXD07)Y4,sum(AXD08) Y5" & _
      " From acc0k0, acc0j0, acc441,acc440" & _
      " where 1=1" & IIf(txtSales <> "", " and a0k20='" & txtSales & "'", "") & strCon & _
      " and nvl(a0k09,0)=0 and (nvl(a0k06,0)+nvl(a0k07,0)) > (nvl(a0k17,0)+nvl(a0k18,0))" & _
      " and a0j13(+)=a0k01 and axd04(+)=a0j13 and axd05(+)=a0j01" & _
      " and a4401(+)=axd01 and a4402(+)=axd02 and a4403(+)=axd03 and a4416 is null and a0j02<>'TT999999000'" & _
      " group by a0j13,a0j01,a0j22,a0j25"
   'Modify By Sindy 2020/5/15 + 增加抓取法律所案源介紹人為此智權人員
   'Modify By Sindy 2023/8/17 調整 1=1" & IIf(txtSales <> "", " and instr(LOS04,'" & txtSales & "')>0", "")
   stVTB2 = stVTB2 & " union " & _
            "select a0j13 Y1,a0j01 Y2,nvl(a0j22,'') Y6,nvl(a0j25,'') Y7" & _
      ",sum(AXD06) Y3,sum(AXD07)Y4,sum(AXD08) Y5" & _
      " From acc0k0, acc0j0, acc441,acc440, lawofficesource, caseprogress" & _
      " where 1=1" & IIf(txtSales <> "", " and instr(LOS04,'" & txtSales & "')>0", "") & strCon & _
      " and nvl(a0k09,0)=0 and (nvl(a0k06,0)+nvl(a0k07,0)) > (nvl(a0k17,0)+nvl(a0k18,0))" & _
      " and a0j13(+)=a0k01 and axd04(+)=a0j13 and axd05(+)=a0j01" & _
      " and a4401(+)=axd01 and a4402(+)=axd02 and a4403(+)=axd03 and a4416 is null and a0j02<>'TT999999000'" & _
      " and LOS15(+)=cp162 and cp162 is not null and cp09(+)=a0j01" & _
      " group by a0j13,a0j01,a0j22,a0j25"
   '2020/5/15 END
   
   'Modified by Lydia 2015/06/17 收據公司別改show名稱
   'strExc(0) = "select '' 選取,sqldatet(a0k02) 收據日期" & _
      ",cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) 本所案號" & _
      ",decode(a0j04,'020',cpm04,cpm03) 案件性質" & _
      ",na03 申請國家,a0j09-X3-nvl(Y3,0) 服務費,a0j10-X4-nvl(Y4,0) 規費,'' 扣繳,0 扣繳金額" & _
      ",a0k01||decode(nvl(a0k19,0),0,'◎')||decode(AXC01,null,'','＊')||decode(a0k11,'J','△') 收據號碼" & _
      ",nvl(tm05,nvl(pa05,nvl(lc05,nvl(sp05,hc06)))) 案件名稱" & _
      ",a0j01,a0j09-X3-nvl(Y3,0) amt1,a0j10-X4-nvl(Y4,0) amt2" & _
      ",X5+NVL(Y5,0) amt3" & _
      ",a0j07,a0k11,a0k19,a0k01,a0j09-X6 SFee,a0j10-X7 OFee,a0k03,a0k04,CU04,a0k05,RPAD(CP01||CP02||CP03||CP04,12,'0') as rCaseNo " & _
      " from (" & stVTB1 & ") X,(" & stVTB2 & ") Y,acc0j0,acc0k0,CUSTOMER,acc431,caseprogress" & _
      ",casepropertymap,nation,trademark,patent,lawcase,servicepractice,hirecase" & _
      " where Y1(+)=X1 and Y2(+)=X2 and (a0j09-X3-nvl(Y3,0)>0 or a0j10-X4-nvl(Y4,0)>0) and a0j13(+)=X1 and a0j01(+)=X2 and a0k01(+)=a0j13 and axc02(+)=a0j13" & _
      " AND CU01(+)=SUBSTR(A0K03,1,8) AND CU02(+)=SUBSTR(A0K03,9) and cp09(+)=a0j01 and cp79>0 and cpm01(+)=cp01 and cpm02(+)=cp10 and na01(+)=a0j04" & _
      " and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04" & _
      " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
      " and lc01(+)=cp01 and lc02(+)=cp02 and lc03(+)=cp03 and lc04(+)=cp04" & _
      " and sp01(+)=cp01 and sp02(+)=cp02 and sp03(+)=cp03 and sp04(+)=cp04" & _
      " and hc01(+)=cp01 and hc02(+)=cp02 and hc03(+)=cp03 and hc04(+)=cp04" & _
      " order by 2,10,12"
   'Added by Lydia 2016/03/24 處理特殊收據
   'strExc(0) = "select '' 選取,sqldatet(a0k02) 收據日期" & _
      ",cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) 本所案號" & _
      ",decode(a0j04,'020',cpm04,cpm03) 案件性質" & _
      ",na03 申請國家,a0j09-X3-nvl(Y3,0) 服務費,a0j10-X4-nvl(Y4,0) 規費,'' 扣繳,0 扣繳金額" & _
      ",a0k01||decode(nvl(a0k19,0),0,'◎')||decode(AXC01,null,'','＊')||decode(a0k11,'J','△') 收據號碼" & _
      ",nvl(tm05,nvl(pa05,nvl(lc05,nvl(sp05,hc06)))) 案件名稱" & _
      ",a0j01,a0j09-X3-nvl(Y3,0) amt1,a0j10-X4-nvl(Y4,0) amt2" & _
      ",X5+NVL(Y5,0) amt3" & _
      ",a0j07,a0k11,a0k19,a0k01,a0j09-X6 SFee,a0j10-X7 OFee,a0k03,a0k04,CU04,a0k05,RPAD(CP01||CP02||CP03||CP04,12,'0') as rCaseNo " & _
      ",decode(a0k11,'1','商標','2','專利','J','智權',a0k11) a0k11C from (" & stVTB1 & ") X,(" & stVTB2 & ") Y,acc0j0,acc0k0,CUSTOMER,acc431,caseprogress" & _
      ",casepropertymap,nation,trademark,patent,lawcase,servicepractice,hirecase" & _
      " where Y1(+)=X1 and Y2(+)=X2 and (a0j09-X3-nvl(Y3,0)>0 or a0j10-X4-nvl(Y4,0)>0) and a0j13(+)=X1 and a0j01(+)=X2 and a0k01(+)=a0j13 and axc02(+)=a0j13" & _
      " AND CU01(+)=SUBSTR(A0K03,1,8) AND CU02(+)=SUBSTR(A0K03,9) and cp09(+)=a0j01 and cp79>0 and cpm01(+)=cp01 and cpm02(+)=cp10 and na01(+)=a0j04" & _
      " and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04" & _
      " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
      " and lc01(+)=cp01 and lc02(+)=cp02 and lc03(+)=cp03 and lc04(+)=cp04" & _
      " and sp01(+)=cp01 and sp02(+)=cp02 and sp03(+)=cp03 and sp04(+)=cp04" & _
      " and hc01(+)=cp01 and hc02(+)=cp02 and hc03(+)=cp03 and hc04(+)=cp04" & _
      " order by 2,10,12"
   'intI = 1
   'Set adoTmp = ClsLawReadRstMsg(intI, strExc(0))
   'Modified by Lydia 2018/06/05 修改顯示案件性質 '020',CPM04,CPM03 => '000',CPM03,CPM04
   'Modified by Lydia 2020/03/31 +,'L','法律'
   'Modified by Lydia 2020/12/10 商標 改名 專利; decode(a0k11,'1','商標' => decode(a0k11,'1','專利'
   'Modified by Lydia 2021/09/06 公司別抓acc080簡稱a0820
   'strExc(0) = "select '' 選取,sqldatet(a0k02) 收據日期" & _
      ",cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) 本所案號" & _
      ",decode(a0k33,'Y',nvl(a0j22,decode(a0j04,'000',cpm03,cpm04)),decode(a0j04,'000',cpm03,cpm04)) 案件性質" & _
      ",na03 申請國家,a0j09-X3-nvl(Y3,0) 服務費,a0j10-X4-nvl(Y4,0) 規費,'' 扣繳,0 扣繳金額" & _
      ",a0k01||decode(nvl(a0k19,0),0,'◎')||decode(AXC01,null,'','＊')||decode(a0k11,'J','△') 收據號碼" & _
      ",nvl(tm05,nvl(pa05,nvl(lc05,nvl(sp05,hc06)))) 案件名稱" & _
      ",a0j01,a0j09-X3-nvl(Y3,0) amt1,a0j10-X4-nvl(Y4,0) amt2" & _
      ",X5+NVL(Y5,0) amt3" & _
      ",a0j07,a0k11,a0k19,a0k01,a0j09-X6 SFee,a0j10-X7 OFee,a0k03,a0k04,CU04,a0k05,RPAD(CP01||CP02||CP03||CP04,12,'0') as rCaseNo " & _
      ",decode(a0k11,'1','專利','2','專利','J','智權','L','法律',a0k11) a0k11C,a0k33 a0k33C ,a0j25 a0j25C  from (" & stVTB1 & ") X,(" & stVTB2 & ") Y,acc0j0,acc0k0,CUSTOMER,acc431,caseprogress" & _
      ",casepropertymap,nation,trademark,patent,lawcase,servicepractice,hirecase" & _
      " where Y1(+)=X1 and Y2(+)=X2 and (a0j09-X3-nvl(Y3,0)>0 or a0j10-X4-nvl(Y4,0)>0)" & _
      " and Y6(+)=X8 and Y7(+)=X9 and a0j13(+)=X1 and a0j01(+)=X2 and a0k01(+)=a0j13 and axc02(+)=a0j13" & _
      " AND CU01(+)=SUBSTR(A0K03,1,8) AND CU02(+)=SUBSTR(A0K03,9) and cp09(+)=a0j01 and cp79>0 and cpm01(+)=cp01 and cpm02(+)=cp10 and na01(+)=a0j04" & _
      " and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04" & _
      " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
      " and lc01(+)=cp01 and lc02(+)=cp02 and lc03(+)=cp03 and lc04(+)=cp04" & _
      " and sp01(+)=cp01 and sp02(+)=cp02 and sp03(+)=cp03 and sp04(+)=cp04" & _
      " and hc01(+)=cp01 and hc02(+)=cp02 and hc03(+)=cp03 and hc04(+)=cp04"
   'Modified by Lydia 2022/02/17 a0j25C後面+cp10 => R030;  客戶案件案號custcase=>R031
   'Modified by Morgan 2023/1/11 已開發票 '＊' --> '(已開發票)'
   'Modified by Lydia 2023/11/13 +A0K40開立INVOICE, a0k32=Z不列印收據
   'strExc(0) = "select '' 選取,sqldatet(a0k02) 收據日期" & _
      ",cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) 本所案號" & _
      ",decode(a0k33,'Y',nvl(a0j22,decode(a0j04,'000',cpm03,cpm04)),decode(a0j04,'000',cpm03,cpm04)) 案件性質" & _
      ",na03 申請國家,a0j09-X3-nvl(Y3,0) 服務費,a0j10-X4-nvl(Y4,0) 規費,'' 扣繳,0 扣繳金額" & _
      ",a0k01||decode(nvl(a0k19,0),0,'◎')||decode(AXC01,null,'','(已開發票)')||decode(a0k11,'J','△') 收據號碼" & _
      ",nvl(tm05,nvl(pa05,nvl(lc05,nvl(sp05,hc06)))) 案件名稱" & _
      ",a0j01,a0j09-X3-nvl(Y3,0) amt1,a0j10-X4-nvl(Y4,0) amt2" & _
      ",X5+NVL(Y5,0) amt3" & _
      ",a0j07,a0k11,a0k19,a0k01,a0j09-X6 SFee,a0j10-X7 OFee,a0k03,a0k04,CU04,a0k05,RPAD(CP01||CP02||CP03||CP04,12,'0') as rCaseNo " & _
      ",a0820 a0k11C,a0k33 a0k33C ,a0j25 a0j25C,CP10,nvl(tm35,nvl(pa48,nvl(lc17,nvl(sp29,'')))) custcase " & _
      "from (" & stVTB1 & ") X,(" & stVTB2 & ") Y,acc0j0,acc0k0,CUSTOMER,acc431,caseprogress" & _
      ",casepropertymap,nation,trademark,patent,lawcase,servicepractice,hirecase,Acc080 " & _
      " where Y1(+)=X1 and Y2(+)=X2 and (a0j09-X3-nvl(Y3,0)>0 or a0j10-X4-nvl(Y4,0)>0)" & _
      " and Y6(+)=X8 and Y7(+)=X9 and a0j13(+)=X1 and a0j01(+)=X2 and a0k01(+)=a0j13 and axc02(+)=a0j13" & _
      " AND CU01(+)=SUBSTR(A0K03,1,8) AND CU02(+)=SUBSTR(A0K03,9) and cp09(+)=a0j01 and cp79>0 and cpm01(+)=cp01 and cpm02(+)=cp10 and na01(+)=a0j04 and a0k11=a0801(+)" & _
      " and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04" & _
      " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
      " and lc01(+)=cp01 and lc02(+)=cp02 and lc03(+)=cp03 and lc04(+)=cp04" & _
      " and sp01(+)=cp01 and sp02(+)=cp02 and sp03(+)=cp03 and sp04(+)=cp04" & _
      " and hc01(+)=cp01 and hc02(+)=cp02 and hc03(+)=cp03 and hc04(+)=cp04"
   'Modified by Lydia 2024/04/01 比對frm210141，差別在沒有抓到Y值; 改用acc0j0為基準
   'strExc(0) = "select '' 選取,sqldatet(a0k02) 收據日期" & _
      ",cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) 本所案號" & _
      ",decode(a0k33,'Y',nvl(a0j22,decode(a0j04,'000',cpm03,cpm04)),decode(a0j04,'000',cpm03,cpm04)) 案件性質" & _
      ",na03 申請國家,a0j09-X3-nvl(Y3,0) 服務費,a0j10-X4-nvl(Y4,0) 規費,'' 扣繳,0 扣繳金額" & _
      ",a0k01||decode(a0k32,'Z','',decode(nvl(a0k19,0),0,'◎'))||decode(AXC01,null,'','(已開發票)')||decode(a0k11,'J','△')||decode(a0k40, null,'','＃') 收據號碼" & _
      ",nvl(tm05,nvl(pa05,nvl(lc05,nvl(sp05,hc06)))) 案件名稱" & _
      ",a0j01,a0j09-X3-nvl(Y3,0) amt1,a0j10-X4-nvl(Y4,0) amt2" & _
      ",X5+NVL(Y5,0) amt3" & _
      ",a0j07,a0k11,a0k19,a0k01,a0j09-X6 SFee,a0j10-X7 OFee,a0k03,a0k04,CU04,a0k05,RPAD(CP01||CP02||CP03||CP04,12,'0') as rCaseNo " & _
      ",a0820 a0k11C,a0k33 a0k33C ,a0j25 a0j25C,CP10,nvl(tm35,nvl(pa48,nvl(lc17,nvl(sp29,'')))) custcase,a0k32,a0k40 " & _
      "from (" & stVTB1 & ") X,(" & stVTB2 & ") Y,acc0j0,acc0k0,CUSTOMER,acc431,caseprogress" & _
      ",casepropertymap,nation,trademark,patent,lawcase,servicepractice,hirecase,Acc080 " & _
      " where Y1(+)=X1 and Y2(+)=X2 and (a0j09-X3-nvl(Y3,0)>0 or a0j10-X4-nvl(Y4,0)>0)" & _
      " and Y6(+)=X8 and Y7(+)=X9 and a0j13(+)=X1 and a0j01(+)=X2 and a0k01(+)=a0j13 and axc02(+)=a0j13" & _
      " AND CU01(+)=SUBSTR(A0K03,1,8) AND CU02(+)=SUBSTR(A0K03,9) and cp09(+)=a0j01 and cp79>0 and cpm01(+)=cp01 and cpm02(+)=cp10 and na01(+)=a0j04 and a0k11=a0801(+)" & _
      " and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04" & _
      " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
      " and lc01(+)=cp01 and lc02(+)=cp02 and lc03(+)=cp03 and lc04(+)=cp04" & _
      " and sp01(+)=cp01 and sp02(+)=cp02 and sp03(+)=cp03 and sp04(+)=cp04" & _
      " and hc01(+)=cp01 and hc02(+)=cp02 and hc03(+)=cp03 and hc04(+)=cp04"
   strExc(0) = "select '' 選取,sqldatet(a0k02) 收據日期" & _
      ",cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) 本所案號" & _
      ",decode(a0k33,'Y',nvl(a0j22,decode(a0j04,'000',cpm03,cpm04)),decode(a0j04,'000',cpm03,cpm04)) 案件性質" & _
      ",na03 申請國家,a0j09-X3-nvl(Y3,0) 服務費,a0j10-X4-nvl(Y4,0) 規費,'' 扣繳,0 扣繳金額" & _
      ",a0k01||decode(a0k32,'Z','',decode(nvl(a0k19,0),0,'◎'))||decode(AXC01,null,'','(已開發票)')||decode(a0k11,'J','△')||decode(a0k40, null,'','＃') 收據號碼" & _
      ",nvl(tm05,nvl(pa05,nvl(lc05,nvl(sp05,hc06)))) 案件名稱" & _
      ",a0j01,a0j09-X3-nvl(Y3,0) amt1,a0j10-X4-nvl(Y4,0) amt2" & _
      ",X5+NVL(Y5,0) amt3" & _
      ",a0j07,a0k11,a0k19,a0k01,a0j09-X6 SFee,a0j10-X7 OFee,a0k03,a0k04,CU04,a0k05,RPAD(CP01||CP02||CP03||CP04,12,'0') as rCaseNo " & _
      ",a0820 a0k11C,a0k33 a0k33C ,a0j25 a0j25C,CP10,nvl(tm35,nvl(pa48,nvl(lc17,nvl(sp29,'')))) custcase,a0k32,a0k40 " & _
      "from (" & stVTB1 & ") X,(" & stVTB2 & ") Y,acc0j0,acc0k0,CUSTOMER,acc431,caseprogress" & _
      ",casepropertymap,nation,trademark,patent,lawcase,servicepractice,hirecase,Acc080 " & _
      " where a0j01=X2(+) and a0j13=x1(+) and a0j01=y2(+) and a0j13=y1(+) and (a0j09-x3-nvl(y3,0)>0 or a0j10-x4-nvl(y4,0)>0) and a0k01(+)=a0j13 and axc02(+)=a0j13" & _
      " AND CU01(+)=SUBSTR(A0K03,1,8) AND CU02(+)=SUBSTR(A0K03,9) and cp09(+)=a0j01 and cp79>0 and cpm01(+)=cp01 and cpm02(+)=cp10 and na01(+)=a0j04 and a0k11=a0801(+)" & _
      " and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04" & _
      " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
      " and lc01(+)=cp01 and lc02(+)=cp02 and lc03(+)=cp03 and lc04(+)=cp04" & _
      " and sp01(+)=cp01 and sp02(+)=cp02 and sp03(+)=cp03 and sp04(+)=cp04" & _
      " and hc01(+)=cp01 and hc02(+)=cp02 and hc03(+)=cp03 and hc04(+)=cp04"
   'Modified by Lydia 2018/06/05 修改顯示案件性質 '020',CPM04,CPM03 => '000',CPM03,CPM04
   'Modified by Lydia 2020/03/31 +,'L','法律'
   'Modified by Lydia 2020/12/10 商標 改名 專利; decode(a0k11,'1','商標' => decode(a0k11,'1','專利'
   'Modified by Lydia 2021/09/06 公司別抓acc080簡稱a0820; decode(a0k11,'1','專利','2','專利','J','智權','L','法律',a0k11)=> a0820
   'Modified by Lydia 2022/02/17 a0j25C後面+cp10 => R030;  客戶案件案號custcase=>R031
   'Modified by Morgan 2023/1/11 已開發票 '＊' --> '(已開發票)'
   'Modified by Lydia 2023/11/13 +A0K40開立INVOICE, a0k32=Z不列印收據
   'strExc(0) = strExc(0) & " group by sqldatet(a0k02),cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04)," & _
      "decode(a0k33,'Y',nvl(a0j22,decode(a0j04,'000',cpm03,cpm04)),decode(a0j04,'000',cpm03,cpm04)) ,na03 ,a0j09-X3-nvl(Y3,0) ,a0j10-X4-nvl(Y4,0)" & _
      ",a0k01||decode(nvl(a0k19,0),0,'◎')||decode(AXC01,null,'','(已開發票)')||decode(a0k11,'J','△')," & _
      "nvl(tm05,nvl(pa05,nvl(lc05,nvl(sp05,hc06)))) ,a0j01,a0j09-X3-nvl(Y3,0) ,a0j10-X4-nvl(Y4,0) ,X5+NVL(Y5,0)," & _
      "a0j07,a0k11,a0k19,a0k01,a0j09-X6 ,a0j10-X7 ,a0k03,a0k04,CU04,a0k05,RPAD(CP01||CP02||CP03||CP04,12,'0')," & _
      "a0820,a0k33,a0j25,CP10,nvl(tm35,nvl(pa48,nvl(lc17,nvl(sp29,'')))) order by 2,10,12,a0j25 "
   strExc(0) = strExc(0) & " group by sqldatet(a0k02),cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04)," & _
      "decode(a0k33,'Y',nvl(a0j22,decode(a0j04,'000',cpm03,cpm04)),decode(a0j04,'000',cpm03,cpm04)) ,na03 ,a0j09-X3-nvl(Y3,0) ,a0j10-X4-nvl(Y4,0)" & _
      ",a0k01||decode(a0k32,'Z','',decode(nvl(a0k19,0),0,'◎'))||decode(AXC01,null,'','(已開發票)')||decode(a0k11,'J','△')||decode(a0k40, null,'','＃')," & _
      "nvl(tm05,nvl(pa05,nvl(lc05,nvl(sp05,hc06)))) ,a0j01,a0j09-X3-nvl(Y3,0) ,a0j10-X4-nvl(Y4,0) ,X5+NVL(Y5,0)," & _
      "a0j07,a0k11,a0k19,a0k01,a0j09-X6 ,a0j10-X7 ,a0k03,a0k04,CU04,a0k05,RPAD(CP01||CP02||CP03||CP04,12,'0')," & _
      "a0820,a0k33,a0j25,CP10,nvl(tm35,nvl(pa48,nvl(lc17,nvl(sp29,'')))),a0k32,a0k40 order by 2,10,12,a0j25 "
      
   '處理特殊收據
   If adoTmp.State <> adStateClosed Then adoTmp.Close
   Set adoTmp = Nothing
   adoTmp.CursorLocation = adUseClient
   adoTmp.Open strExc(0), cnnConnection, adOpenDynamic, adLockBatchOptimistic
   'Added by Lydia 2021/07/23 因為合併資料會發生"多重步驟發生錯誤",所以改寫法
   Set Adodc1.Recordset = PUB_CreateRecordset(adoTmp, , , , Me.Name, mSeqNo)
   'Modified by Lydia 2022/02/17 + R030 AS CP10, R031 AS CUSTCASE
   'Modified by Lydia 2023/11/13 +R032 AS A0K32, R033 AS A0K40
   strMidSql = " Select R001 as 選取,R002 as 收據日期,R003 as 本所案號,R004 as 案件性質," & _
                     "R005 as 申請國家,R006 as 服務費,R007 as 規費,R008 as 扣繳,NVL(R009,0) as 扣繳金額, " & _
                     "R010 as 收據號碼,R011 as 案件名稱,R012 as A0J01,R013 as AMT1,R014 as AMT2, " & _
                     "R015 As Amt3,R016 As A0j07,R017 As A0k11,R018 As A0k19,R019 As A0k01, " & _
                     "R020 As Sfee,R021 As Ofee,R022 As A0k03,R023 As A0k04,R024 As Cu04,R025 As A0k05,R026 As Rcaseno, " & _
                     "R027 As A0K11C, R028 As A0K33C, R029 As A0J25C, R030 AS CP10, R031 AS CUSTCASE, R032 AS A0K32, R033 AS A0K40 " & _
                     "From RDataFactory where FormName='" & Me.Name & "' And ID='" & strUserNum & "' and seqno = " & mSeqNo
   strMidSql = strMidSql & " order by r002,r010,r012,a0j25c"
   If adoTmp.State <> adStateClosed Then adoTmp.Close
   Set adoTmp = Nothing
   adoTmp.CursorLocation = adUseClient
   adoTmp.Open strMidSql, cnnConnection, adOpenDynamic, adLockBatchOptimistic
   'end 2021/07/23
   If adoTmp.RecordCount > 0 Then
        adoTmp.MoveFirst
        Do While Not adoTmp.EOF
           If IsNull(adoTmp.Fields("a0k33c")) = False Then '特殊收據
              '相鄰項目同案件性質,合併列印
              'Modified by Lydia 2016/11/01 +本所案號判斷
              'If strA0k01 = adoTmp.Fields("a0k01") And strA0j22 = "" & adoTmp.Fields("案件性質") Then
              If strA0K01 = adoTmp.Fields("a0k01") And strA0j22 = "" & adoTmp.Fields("案件性質") And strCaseNo = "" & adoTmp.Fields("本所案號") Then
                 strA0K01 = adoTmp.Fields("a0k01")
                 strA0j22 = "" & adoTmp.Fields("案件性質")
                 lngAmount1 = lngAmount1 + Val(adoTmp.Fields("服務費"))
                 lngAmount2 = lngAmount2 + Val(adoTmp.Fields("規費"))
                 'Added by Lydia 2017/07/25 計算扣繳總額
                 lngSFee = lngSFee + Val(adoTmp.Fields("SFee"))
                 lngOFee = lngOFee + Val(adoTmp.Fields("OFee"))
                 'end 2017/07/25
                 
                 '加總在第一筆
                 adoTmp.MovePrevious
                 adoTmp.Fields("服務費") = lngAmount1
                 adoTmp.Fields("規費") = lngAmount2
                 'Added by Lydia 2017/07/25 計算扣繳總額
                 adoTmp.Fields("SFee") = lngSFee
                 adoTmp.Fields("OFee") = lngOFee
                 'end 2017/07/25
                 'adoTmp.Fields("a0k32") = "" & adoTmp.Fields("a0k32") 'Added by Lydia 2023/11/13 收據暫不列印
                 adoTmp.UPDATE
                 '刪除被合併項目
                 adoTmp.MoveNext
                 adoTmp.Delete adAffectCurrent
              Else
                 strA0K01 = adoTmp.Fields("a0k01")
                 strA0j22 = "" & adoTmp.Fields("案件性質")
                 lngAmount1 = Val(adoTmp.Fields("服務費"))
                 lngAmount2 = Val(adoTmp.Fields("規費"))
                 'Added by Lydia 2017/07/25 計算扣繳總額
                 lngSFee = Val(adoTmp.Fields("SFee"))
                 lngOFee = Val(adoTmp.Fields("OFee"))
                 'end 2017/07/25
                 'adoTmp.Fields("a0k32") = "" & adoTmp.Fields("a0k32") 'Added by Lydia 2023/11/13 收據暫不列印
                 strCaseNo = "" & adoTmp.Fields("本所案號") 'Added by Lydia 2016/11/01
              End If
           End If
           If Not (adoTmp.EOF) Then
              adoTmp.MoveNext
           End If
        Loop
   End If
   'end 2016/03/24
   
   If pNoReset = False Then FormReset
   
   DataGrid1.Enabled = True
   'Modify by Amy 2014/06/16 +FormName 改暫存TB
   'Modified by Lydia 2021/07/27 改在模組中取得序號 + mSeqNo ;
   Set Adodc1.Recordset = PUB_CreateRecordset(adoTmp, , , , Me.Name, mSeqNo)
      
   txtSales.Tag = txtSales
   If adoTmp.RecordCount > 0 Then
      With adoTmp
      .MoveFirst
      Do While Not .EOF
         txtTot(0) = Val(txtTot(0)) + Format(Val("" & .Fields("服務費")) / 1000)
         txtTot(1) = Val(txtTot(1)) + Val("" & .Fields("規費"))
         .MoveNext
      Loop
      End With
      txtTot(0) = Format(txtTot(0), "0.00")
      txtTot(1) = Format(txtTot(1), DDollar2)
      cmdok(2).Enabled = True
    '  m_iCols = .Cols
   Else
      cmdok(2).Enabled = False
      If pNoMsg = False Then
         MsgBox "無符合資料！", vbExclamation
      End If
   End If
End Sub

Private Sub DataGrid1_AfterColUpdate(ByVal ColIndex As Integer)
   Dim bCancel As Boolean
   Select Case ColIndex
   'Modified by Morgan 2015/7/16 調欄位順序
   Case 7, 8
      DataGrid1.Columns(ColIndex).Value = Val(Format(DataGrid1.Columns(ColIndex)))
      If ColIndex = 7 Then
         If Val(DataGrid1.Columns(ColIndex).Value) > Val(Adodc1.Recordset.Fields("amt1")) Then
            MsgBox "收款服務費不可大於未收服務費！", vbInformation
            'Modify by Amy 2014/06/16 改暫存TB同一列選取三次修改三次會產生「找不到要更新的資料列」
            'bCancel = True
            DataGrid1.Columns(ColIndex).Value = Val(Adodc1.Recordset.Fields("amt1"))
         End If
      ElseIf ColIndex = 8 Then
         If Val(DataGrid1.Columns(ColIndex).Value) > Val(Adodc1.Recordset.Fields("amt2")) Then
            MsgBox "收款規費不可大於未收規費！", vbInformation
            'Modify by Amy 2014/06/16 改暫存TB同一列選取三次修改三次會產生「找不到要更新的資料列」
            'bCancel = True
            DataGrid1.Columns(ColIndex).Value = Val(Adodc1.Recordset.Fields("amt2"))
         End If
      End If
      
   Case Else
      bCancel = True
   End Select
   
   'Modify by Amy 2014/06/16 改暫存TB同一列選取三次修改三次會產生「找不到要更新的資料列」
   'If bCancel = True Then
   '   Adodc1.Recordset.CancelUpdate
   'Else
   '   Adodc1.Recordset.UpdateBatch
      Adodc1.Recordset.UPDATE
      SettxtTot
   'End If
End Sub

Private Sub DataGrid1_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
   Select Case ColIndex
   'Modified by Morgan 2015/7/16 調欄位
   Case 7, 8
      If DataGrid1.Columns(0).Text <> "Y" Then
         Cancel = 1
         DataGrid1.Columns(ColIndex).Value = Adodc1.Recordset.Fields(DataGrid1.Columns(ColIndex).DataField).Value
         MsgBox "請先選取！", vbExclamation
      End If
   Case Else
      Cancel = 1
      DataGrid1.Columns(ColIndex).Value = Adodc1.Recordset.Fields(DataGrid1.Columns(ColIndex).DataField)
   End Select
End Sub

Private Sub DataGrid1_Click()

   'Added by Lydia 2021/07/16 +判斷有資料才繼續
   If mSeqNo = "" Then Exit Sub
   If Adodc1.Recordset.RecordCount = 0 Then Exit Sub
   'end 2021/07/16
   
       SelectDataGrid1 m_MouseCol, m_mouseRow

End Sub

Private Sub SelectDataGrid1(pCol As Integer, pRow As Integer)
   Dim bUpdate As Boolean, stKey1 As String, stKey2 As String
   'Modified by Morgan 2015/7/14 調欄位
   If pRow >= 0 And (pCol = 0 Or pCol = 9) Then 'Lydia
      
      Set adoTmp = Adodc1.Recordset.Clone
      With adoTmp
      'Modified by Morgan 2014/3/25 重新排序後,Grid的rowid並沒有變
      'intI = DataGrid1.FirstRow + pRow - 1
      'adoTmp.Move intI, adBookmarkFirst
      .MoveFirst
      If DataGrid1.FirstRow > 1 Then
         .Move DataGrid1.FirstRow - 1
      End If
      If .Sort <> Adodc1.Recordset.Sort Then
         strExc(0) = .Fields("a0j01")
         .Sort = Adodc1.Recordset.Sort
         .MoveFirst
         .Find "a0j01='" & strExc(0) & "'"
      End If
      If pRow > 0 Then
         adoTmp.Move pRow
      End If
      'end 2014/3/25
      Select Case pCol
      Case 0 '選取
         If .Fields("選取") = "Y" Then
            'Modified by Lydia 2023/05/08 同收據扣繳要同步
            '.Fields("扣繳") = ""
            '.Fields("扣繳金額") = 0
            If .Fields("扣繳") = "Y" Then
               UpdateTax .Fields("a0k01"), False
            End If
            'end 2023/05/08
            .Fields("選取") = ""
            .Fields("服務費").Value = Val(.Fields("amt1"))
            .Fields("規費").Value = Val(.Fields("amt2"))
         Else
            'Modified by Lydia 2023/11/13 +判斷Z=確定不印
            If Val(.Fields("a0k19")) = 0 And "" & .Fields("a0k32") <> "Z" Then
               MsgBox "本收據尚未列印!", vbExclamation
            End If
            .Fields("選取") = "Y"
            UpdateSelected .Fields("a0k01") 'Added by Lydia 2023/05/08 同一收據勾選一筆收文自動預設勾其他收文
            'Added by Lydia 2023/11/13  同一INVOICE編號勾選一筆收文自動預設勾其他收文
            If "" & .Fields("a0k40") <> "" Then
               UpdateSelected .Fields("a0k40"), "1"
            End If
            'end 2023/11/13
         End If
         bUpdate = True
         
      Case 9 '扣繳
         If .Fields("選取") = "Y" Then
            If .Fields("扣繳") = "Y" Then
               'Modified by Lydia 2023/05/08 同收據扣繳要同步
               '.Fields("扣繳") = ""
               '.Fields("扣繳金額") = 0
               'bUpdate = True
               UpdateTax .Fields("a0k01"), False
               SettxtTot
               'end 2023/05/08
            Else
               If .Fields("a0k11") = "J" Then
                  MsgBox "智權公司不可勾選扣繳!", vbCritical
               ElseIf .Fields("a0k05") = "1" Then
                  MsgBox "個人不可勾選扣繳!", vbCritical
               ElseIf .Fields("amt3") > 0 Then
                  MsgBox "本收據已扣繳不可再勾選扣繳!", vbCritical
               Else
                  'Modified by Lydia 2023/05/08 同收據扣繳要同步
                  '.Fields("扣繳") = "Y"
                  ''是否合併
                  ''Modified by Morgan 2014/8/5 扣繳金額必須為整數(四捨五入)
                  'If .Fields("a0j07") = "Y" Then
                  '   '.Fields("扣繳金額").Value = 0.1 * (Val(.Fields("SFee")) + Val(.Fields("OFee")))
                  '   .Fields("扣繳金額").Value = Round(0.1 * (Val(.Fields("SFee")) + Val(.Fields("OFee"))))
                  'Else
                  '   '.Fields("扣繳金額").Value = 0.1 * Val(.Fields("SFee"))
                  '   .Fields("扣繳金額").Value = Round(0.1 * Val(.Fields("SFee")))
                  'End If
                  'bUpdate = True
                  UpdateTax .Fields("a0k01"), True
                  SettxtTot
                  'end 2023/05/08
               End If
            End If
         Else
            MsgBox "尚未選取不可勾選扣繳!", vbExclamation
         End If
      End Select
      
      If bUpdate = True Then
         'Modify by Amy 2014/06/16 改暫存TB同一列選取三次修改三次會產生「找不到要更新的資料列」
         '.UpdateBatch
         .UPDATE
         SettxtTot
      End If
      End With
   End If
End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
   Dim stValue As String
      
   If Adodc1.RecordSource = Empty Then Exit Sub 'Added by Lydia 2021/08/19 若尚未查詢，防止出錯
   
   If Adodc1.Recordset.RecordCount > 0 Then
      Select Case ColIndex
      'Modified by Morgan 2015/7/16 調欄位順序
      Case 0, 9
         Set adoTmp = Adodc1.Recordset.Clone
         With adoTmp
         .MoveFirst
         .Find "選取='Y'"
         If ColIndex = 0 Then
            If .EOF Then
               stValue = "Y"
            Else
               stValue = ""
            End If
         Else
            '無選取時不動作
            If .EOF Then
               Exit Sub
            Else
               .Find "扣繳='Y'"
               If .EOF Then
                  stValue = "Y"
               Else
                  stValue = ""
               End If
            End If
         End If
         
         .MoveFirst
         Do While Not .EOF
            '選取
            If ColIndex = 0 Then
               If "" & .Fields("選取") <> stValue Then
                  .Fields("選取") = stValue
                  If stValue = "" Then
                     .Fields("服務費") = .Fields("amt1")
                     .Fields("規費") = .Fields("amt2")
                     .Fields("扣繳") = ""
                     .Fields("扣繳金額") = 0
                  End If
               End If
            '扣繳
            Else
               If .Fields(0) = "Y" Then
                  If .Fields("a0k11") <> "J" And .Fields("a0k05") <> "1" And .Fields("amt3") = 0 And "" & .Fields("扣繳") <> stValue Then
                     If stValue = "Y" Then
                        'Modified by Lydia 2023/05/08 同收據扣繳要同步
                        '.Fields("扣繳") = stValue
                        ''是否合併
                        'If .Fields("a0j07") = "Y" Then
                        '   .Fields("扣繳金額") = 0.1 * (Val(.Fields("SFee")) + Val(.Fields("OFee")))
                        'Else
                        '   .Fields("扣繳金額") = 0.1 * Val(.Fields("SFee"))
                        'End If
                        UpdateTax .Fields("a0k01"), True
                        'end 2023/05/08
                     Else
                        .Fields("扣繳金額") = 0
                     End If
                  End If
               End If
            End If
        'Modify by Amy 2014/06/16 改暫存TB同一列選取三次修改三次會產生「找不到要更新的資料列」
            .UPDATE
            .MoveNext
         Loop
         '.UpdateBatch
         'end 2014/06/16
         SettxtTot
         End With
      'Added by Morgan 2013/12/18 排序
      Case Else
         If m_blnColOrderAsc Then
            Adodc1.Recordset.Sort = ""
            Adodc1.Recordset.Sort = DataGrid1.Columns(ColIndex).DataField & " desc"
            m_blnColOrderAsc = False
         Else
            Adodc1.Recordset.Sort = ""
            Adodc1.Recordset.Sort = DataGrid1.Columns(ColIndex).DataField & " asc"
            m_blnColOrderAsc = True
         End If
      End Select
   End If
   
End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
   'Modified by Morgan 2015/7/16 調欄位順序
   If DataGrid1.col <> 7 And DataGrid1.col <> 8 Then
      KeyCode = 0
   End If
End Sub

Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
   'Modified by Morgan 2015/7/16 調欄位順序
   Select Case DataGrid1.col
   Case 7
      If KeyCode = vbKeyReturn Then SendKeys "{RIGHT}"

   Case 8
      If KeyCode = vbKeyReturn Then
         If DataGrid1.row < Adodc1.Recordset.RecordCount - 1 Then
            SendKeys "{DOWN}"
            SendKeys "{LEFT}"
         End If
      End If
   End Select
End Sub

Private Sub DataGrid1_LostFocus()
On Error GoTo Checking
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   'Modified by Morgan 2015/7/16 調欄位順序
   If DataGrid1.col = 7 Or DataGrid1.col = 8 Then
      'DataGrid1.Columns(DataGrid1.col) = Val(DataGrid1.Columns(DataGrid1.col))
   End If
   'Modify by Amy 2014/06/16
   'Adodc1.Recordset.UpdateBatch
   Adodc1.Recordset.UPDATE
Checking:
   Exit Sub
End Sub

Private Sub DataGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   m_MouseCol = DataGrid1.ColContaining(x)
   m_mouseRow = DataGrid1.RowContaining(y)
End Sub

Private Sub Form_Load()
'Add by Lydia 2014/9/25 印表機列表
Dim SeekPrint As Integer, SeekPrintL As Integer
'Dim strSql As String, i As Integer, j As Integer 'Mark by Lydia 2024/11/05

   'Modified by Morgan 2020/10/30
   'strSql = Printer.DeviceName
   'SeekPrintL = Printer.Orientation
   'For i = 0 To Printers.Count - 1
   '   Set Printer = Printers(i)
   '   Combo1.AddItem Printer.DeviceName, j
   '   j = j + 1
   '   If Printer.DeviceName = strSql Then
   '      SeekPrint = i
   '   End If
   'Next i
   'Set Printer = Printers(SeekPrint)
   'Combo1.Text = Combo1.List(SeekPrint)
   PUB_SetPrinter Me.Name, Combo1, strPrinter, , , , , True
   'end 2020/10/30
   Me.Height = 5532 'Added by Lydia 2024/04/25
'---------Add by Lydia 2014/9/25 印表機列表
   
   MoveFormToCenter Me
   stST05 = PUB_GetST05(strUserNum)
   txtSales = strUserNum
   
   'bolAreaMan = False 'Add By Sindy 2023/6/12
   'Modify By Sindy 2020/7/28 設定員編,部門,所別權限
   'Modify By Sindy 2025/3/18 +Me.Name
   Call PUB_SetFormSaleDept(strUserNum, , , , txtSales, bolSpecMan, strSpecCode, , , , , , , , Me.Name)
   'Add By Sindy 2023/6/12
   '檢查當時是否需要為他人職代
   Combo3.Clear
   If txtSales <> strUserNum And txtSales <> "" Then
      Combo3.AddItem txtSales & " " & GetPrjSalesNM(txtSales)
   End If
   Combo3.AddItem strUserNum & " " & strUserName
   Call Pub_SetForOthersEmpCombo(strUserNum, Combo3, False, m_strListPer)
   If m_strListPer = "" Then
      Combo3.Visible = False
   Else
      'Add by Amy 2020/03/25 判斷下拉選單是否有區主管
'      If InStr(m_strListPer, GetDeptMan(stST15)) > 0 Then
'         bolAreaMan = True
'      End If
      Combo3.Visible = True
      Combo3.ListIndex = 0
      'Added by Lydia 2021/05/20 Form 2.0物件無法覆蓋Form 1.0
      txtSales.Visible = False
      lblSalesName.Visible = False
   End If
   '2023/6/12 END
   
'   txtSales.Enabled = False
'   Select Case strUserNum
'      '外商陳經理可看外商
'      Case "68005"
'         txtSales.Enabled = True
'         txtSalesArea = "F10"
'         txtSalesArea1 = "F19"
'
'      '小真可看全部
'      'modify by sonia 2014/6/9 +美珍77027
'      'Modify by Amy 2015/03/16 拿掉美珍 改寫至特殊人員(總經理業務工作代理人員)
'      Case "65001"
'         txtSales.Enabled = True
'
'      '杜燕文,劉大愛可看S31
'      'modify by sonia 2016/8/22劉大愛78007改為蘇嫄媛79053
'      Case "74018", "79053"
'         txtSales.Enabled = True
'         txtSalesArea = "S31"
'         txtSalesArea1 = "S31"
'
'
'      '王協理可看專利處
'      Case "71011"
'         txtSales.Enabled = True
'         txtSalesArea = "P10"
'         txtSalesArea1 = "P19"
'
'      '葉經理可看商標處
'      'modify by sonia 2016/2/24 +69008
'      Case "67002", "69008"
'         txtSales.Enabled = True
'         txtSalesArea = "P20"
'         txtSalesArea1 = "P29"
'
'      Case Else
'         Select Case stST05
'            '電腦中心,財務,總經理看全部
'            Case "00", "01"
'               txtSales.Enabled = True
'
'            '各區主管
'            Case "SM"
'               '簡協理可看北所全部
'               If strUserNum = "69005" Then
'                  txtSalesArea = "S10"
'                  txtSalesArea1 = "S19"
'
'               Else
'                  txtSalesArea = Pub_StrUserSt15
'                  txtSalesArea1 = Pub_StrUserSt15
'               End If
'
'               txtSales.Enabled = True
'            '外商主管  王宗珮、洪琬姿、葉易雲
'            Case "21", "26", "28"
'               txtSales.Enabled = True
'               txtSalesArea = Pub_StrUserSt15
'               txtSalesArea1 = Pub_StrUserSt15
'
'            '其他只能看自己
'            Case Else
'               'Add By Sindy 2020/7/1 柄佑不能查其他智權人員,只開放ST52 ex:75007
'               txtSalesArea = Pub_StrUserSt15
'               txtSalesArea1 = Pub_StrUserSt15
'               '2020/7/1 END
'         End Select
'   End Select
'
'   '若操作人員的ST05=SA且在職員工的ST52有該編號存在,此類人員稱之為帶人主管,開放智權人員欄位可輸入
'   If PUB_GetST05Limits(strUserNum) = True Then
'      txtSales.Enabled = True
'   End If
'   '中三區人員可單獨下68096看期限
'   If PUB_GetStaffST15(strUserNum, "1") = "S23" Then
'      txtSales.Enabled = True
'   End If
'
'   '北五區人員可單獨下10051看期限
'   If PUB_GetStaffST15(strUserNum, "1") = "S15" Then
'      txtSales.Enabled = True
'   End If
'
'   'Add by Amy 2015/03/16 +總經理業務工作代理人員
'   If CheckLevel(strUserNum, "總經理業務工作代理人員") = True Then
'       bolSpecMan = True
'       strSpecCode = "總經理業務工作代理人員"
'   'Modify  by Amy 2014/05/21 開放專利處部份智權同仁資料給彥葶代為處理
'   ElseIf CheckLevel(strUserNum, "A8") = True Then
'        bolSpecMan = True
'        strSpecCode = "A8"
'   End If
'
'   If bolSpecMan = True Then
'        'Add by Amy 2015/03/16 +總經理業務工作代理人員
'        If InStr(strSpecCode, "總經理業務工作代理人員") > 0 Then txtSales.Enabled = True
'        If InStr(strSpecCode, "A8") > 0 Then txtSales.Enabled = True
'   End If
'   'end 2014/05/21
   
   'Added by Morgan 2015/7/16
   'Modified by Lydia 2021/09/09 比照frm210141的列表
   cboComp.Clear
   cboComp.AddItem "", 0
   'cboComp.AddItem "商標", 1 'Remove by Lydia 2020/12/10 取銷 1 專利商標
   'Modified by Lydia 2020/12/10 Index  2~4 改 Index 1~3
   'Modified by Lydia 2021/09/09 比照frm210141的列表
   'cboComp.AddItem "專利", 1
   'cboComp.AddItem "智權", 2
   'cboComp.AddItem "法律", 3 'Added by Lydia 2020/03/31
   'end 2015/7/16
   cboComp.AddItem "商標", 1
   cboComp.AddItem "智慧所", 2
   cboComp.AddItem "智權", 3
   cboComp.AddItem "法律所", 4
   'end 2021/09/09
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set adoTmp = Nothing
   Set frm210146 = Nothing
End Sub

'Add By Sindy 2022/12/14 外部呼叫應收帳款查詢
Public Function CallQueryData(strSales As String, strCustNo_1 As String, strCustNo_2 As String, _
   strCustNo_3 As String, strCustNo_4 As String, strCustNo_5 As String) As Boolean
   
   CallQueryData = False
'   'Add By Sindy 2023/6/19
'   If Me.Combo3.Visible = True Then
'      Combo3.Text = strSales
'      Call Combo3_Validate(True)
'      Combo3.Enabled = False
'   Else
'      txtSales = strSales
'      txtSales.Enabled = False
'   End If
'   'Add By Sindy 2023/8/17
   Combo3.Visible = False
   Combo3.Enabled = False
   Combo3.Text = ""
   txtSales.Visible = True
   txtSales.Enabled = False
   txtSales.Text = ""
   txtTitle.Enabled = False
   txtDate(0).Enabled = False
   txtDate(1).Enabled = False
   cmdok(5).Enabled = False
   cmdok(5).Tag = strCustNo_1 & "," & strCustNo_2 & "," & strCustNo_3 & "," & strCustNo_4 & "," & strCustNo_5 '記錄接洽單申請人編號
   '2023/8/17 END
   txtCustNo(0) = strCustNo_1
   txtCustNo_Validate 0, False
   'txtCustNo_Validate 1, False
   'Modify By Sindy 2023/2/16 接洽單直接開應收帳款時，客戶編號的起迄值都預設相同的，迄號不要ZZZ
   'Call txtCustNo_GotFocus(1)
   txtCustNo(1) = txtCustNo(0)
   '2023/2/16 END
   Call cmdok_Click(1)
   CallQueryData = True
End Function

Private Function SetA4425() As Boolean
   Do
      strExc(1) = InputBox("請輸入溢收款處理方式：" & vbCrLf & vbCrLf & "1. 列暫收" & vbCrLf & "2. 退客戶 ( 須特別寄出者，請於備註欄註明 )" & vbCrLf & vbCrLf & "請輸入 1 或 2")
      'Added by Morgan 2014/3/31 金額可能輸錯
      If strExc(1) = "" Then
         Exit Do
      'end 2014/3/31
      ElseIf strExc(1) <> "1" And strExc(1) <> "2" Then
         MsgBox "請輸入 1 或 2 !!", vbExclamation
      Else
         m_A4425 = strExc(1)
         SetA4425 = True
         Exit Do
      End If
   Loop
End Function

Private Sub txtCustNo_Change(Index As Integer)
   If txtSales.Tag <> "" Then  '判斷非正常操作
      FormReset
   End If
End Sub

Private Sub txtCustNo_GotFocus(Index As Integer)
   If Left(txtCustNo(Index), 1) = "X" Then
      txtCustNo(Index).SelStart = 1
      txtCustNo(Index).SelLength = Len(txtCustNo(Index)) - 1
   Else
      TextInverse txtCustNo(Index)
   End If
   
   CloseIme
   If Index = 1 And Len(txtCustNo(0)) = 9 Then
      'Added by Lydia 2022/02/17  資策會(X38805030)請款文件
      'Modified by Lydia 2023/11/08 改判斷
      'If txtCustNo(0) = "X38805030" Then
      If InStr(strSpecExcel, txtCustNo(0)) > 0 Then
        ChkExcel.Visible = True
        ChkExcel.Value = True
        txtCustNo(Index) = Left(txtCustNo(0), 8) & "Z"
        txtCustNo(Index).SelStart = 9
        txtCustNo(Index).SelLength = 1
        'Added by Lydia 2024/04/25
        lblFormat.Visible = False
        txtFormat.Visible = False
        'end 2024/04/25
      Else
        ChkExcel.Visible = False
        ChkExcel.Value = False
      'end 2022/02/17
        txtCustNo(Index) = Left(txtCustNo(0), 6) & "ZZZ"
        txtCustNo(Index).SelStart = 6
        txtCustNo(Index).SelLength = 3
        'Added by Lydia 2024/04/25
        lblFormat.Visible = True
        txtFormat.Visible = True
        'end 2024/04/25
      End If  'Added by Lydia 2022/02/17
   End If
End Sub

Private Sub txtCustNo_Validate(Index As Integer, Cancel As Boolean)
   If Index = 0 And Len(txtCustNo(Index)) > 5 Then
      txtCustNo(Index) = Left(txtCustNo(Index) & "000", 9)
   End If
   'Added by Lydia 2022/02/17
   'Modified by Lydia 2023/11/08 改判斷
   'If Index = 0 And ChkExcel.Visible = True And txtCustNo(0) <> "X38805030" Then
   If Index = 0 And ChkExcel.Visible = True And InStr(strSpecExcel, txtCustNo(0)) = 0 And Len(txtCustNo(0)) = 9 Then
        ChkExcel.Visible = False
        ChkExcel.Value = False
   End If
   'Added by Lydia 2023/11/08
   If Len(txtCustNo(0)) = 0 And Len(txtCustNo(1)) = 0 Then
      ChkExcel.Visible = False
      ChkExcel.Value = False
   End If

End Sub

'Added by Lydia 2024/04/25
Private Sub txtFormat_GotFocus()
   TextInverse txtFormat
End Sub

'Added by Lydia 2024/04/25
Private Sub txtFormat_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtSales_Change()
   If Len(txtSales) > 4 Then
      lblSalesName = GetStaffName(txtSales, True)
      Call SetCombo2 'Added by Lydia 2015/06/17
   Else
      lblSalesName = ""
   End If
   
   If txtSales.Tag <> "" Then
      FormReset
   End If
End Sub

Private Sub txtSales_GotFocus()
   TextInverse txtSales
   CloseIme
   If Combo3.Enabled = True And Combo3.Visible = True Then Combo3.SetFocus 'Add By Sindy 2023/6/12
End Sub

Private Sub txtSales_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtSales_LostFocus()
   If Trim(txtSales) = "" Then
       lblSalesName = ""
   End If
End Sub

Private Sub txtSales_Validate(Cancel As Boolean)
   'Modify By Sindy 2024/9/18 智權部:使用智權人員ID查詢權限; 改為共用函數
   'Modify By Sindy 2025/3/17 +Me.Name
   If PUB_txtSales_Limit(txtSales, m_strListPer, , , , _
                         bolSpecMan, strSpecCode, lblSalesName, Me.Name) = False Then
      If txtSales.Visible = True Then 'Added by Lydia 2021/05/20 排除隱藏
        txtSales.SetFocus
        txtSales_GotFocus
      End If 'Added by Lydia 2021/05/20
      Cancel = True
      Exit Sub
   Else
      txtSales.Tag = txtSales.Text
   End If
   '2024/9/18 END
End Sub

Private Sub FormReset()
   Dim oText As TextBox
   
   For Each oText In txtTot
      oText.Text = "0"
   Next
   
   m_A4421 = "": m_A4427 = ""
   
   cmdok(2).Enabled = False
   
   txtSales.Tag = ""
   
   If Not Adodc1.Recordset Is Nothing Then
      If Adodc1.Recordset.State = 1 Then
         Adodc1.Recordset.Close
         DataGrid1.Refresh
      End If
   End If
   
   'Added by Morgan 2014/1/9
   m_A4425 = ""
End Sub

Private Sub SettxtTot()
   Set adoTmp = Adodc1.Recordset.Clone
   With adoTmp
   .MoveFirst
   txtTot(2) = 0
   txtTot(3) = 0
   txtTot(4) = 0
   txtTot(5) = 0
   'Added by Lydia 2015/06/17
   m_selA1 = "": m_selA2 = ""
   
   Do While Not .EOF
      If .Fields(0) = "Y" Then
         txtTot(2) = Val(txtTot(2)) + Val("" & .Fields("服務費"))
         txtTot(3) = Val(txtTot(3)) + Val("" & .Fields("規費"))
         txtTot(4) = Val(txtTot(4)) + Val("" & .Fields("扣繳金額"))
         txtTot(5) = Val(txtTot(2)) + Val(txtTot(3)) - Val(txtTot(4))
          'Added by Lydia 2015/06/17 判斷公司別
         If .Fields("a0k11") = "1" Then
            'Modified by Lydia 2020/12/10 取銷 1 專利商標
            'm_selA1 = "Y"
            m_selA2 = "Y"
         ElseIf .Fields("a0k11") = "2" Then
            m_selA2 = "Y"
         End If
         
      End If
      .MoveNext
   Loop
   
   txtTot(2) = Format(txtTot(2), DDollar2)
   txtTot(3) = Format(txtTot(3), DDollar2)
   txtTot(4) = Format(txtTot(4), DDollar2)
   txtTot(5) = Format(txtTot(5), DDollar2)
   End With

End Sub

Private Sub txtTitle_Change()
   If txtSales.Tag <> "" Then  '判斷非正常操作
      FormReset
   End If
End Sub

Private Sub txtTitle_GotFocus()
   TextInverse txtTitle
   OpenIme
End Sub

Sub PrtData()

Dim strA As String
Dim rsA As New ADODB.Recordset
Dim RmFee1 As Double, RmFee2 As Double, RmDFee1 As Double '小計金額

'Modified by Morgan 2020/10/30
'Set Printer = Printers(Combo1.ListIndex)
PUB_RestorePrinter Combo1
'end 2020/10/30

Printer.EndDoc
Printer.Orientation = 2 '1.直印 2.橫印
Printer.PaperSize = 9  'A4
lngPageHeight = Printer.ScaleHeight
lngPageWidth = Printer.ScaleWidth
lngLineHeight = 300
   
'Remove by Lydia 2021/07/27 改在模組中取得
'strA = "select max(SeqNo) as mno from RDataFactory Where FormName='" & Me.Name & "' And ID='" & strUserNum & "' "
'If rsA.State = 1 Then rsA.Close
'rsA.CursorLocation = adUseClient
'rsA.Open strA, cnnConnection, adOpenStatic, adLockReadOnly
'
'mSeqNo = 1
'If Not (IsNull(rsA!mNo)) Then
'   mSeqNo = rsA!mNo
'End If
'end 2021/07/27

'抬頭資料和總計資料
strA = " select min(R002) as mindate,max(R002) as maxdate," & _
       " NVL(sum(decode (R016,'Y',R006+R007,R006)),0) as fee1," & _
       " NVL(sum(decode (R016,'Y',0,R007)),0) as fee2, NVL(sum(R009),0) as dfee1 " & _
       " From RDataFactory where R001 = 'Y' and FormName='" & Me.Name & "' And ID='" & strUserNum & "' and seqno = " & mSeqNo & " "
Set rsA = Nothing
rsA.Open strA, cnnConnection, adOpenStatic, adLockReadOnly

minDate = "" & rsA!minDate '帳款日期
maxDate = "" & rsA!maxDate
mFee1 = "" & rsA!fee1  '總計:應收-服務費
mFee2 = "" & rsA!fee2  '總計:應收-規費
mdFee1 = "" & rsA!dfee1 '總計:扣繳


strA = " Select R001 as 選取,R002 as 收據日期,R003 as 本所案號,R004 as 案件性質, " & _
       " R005 as 申請國家,R006 as 服務費,R007 as 規費,R008 as 扣繳,NVL(R009,0) as 扣繳金額, " & _
       " R010 as 收據號碼,R011 as 案件名稱,R012 as A0J01,R013 as AMT1,R014 as AMT2, " & _
       " R015 as AMT3,R016 as A0J07,R017 as A0K11,R018 as A0K19,R019 as A0K01, " & _
       " R020 as SFEE,R021 as OFEE,R022 as A0K03,R023 as A0K04,R024 as CU04,R025 as A0K05,A0802,R026 as rCaseNo  " & _
       " From RDataFactory, acc080 where R017=A0801 and R001 = 'Y' and FormName='" & Me.Name & "' " & _
       " And ID='" & strUserNum & "' and seqno = " & mSeqNo & " Order by A0K11,A0K04,RowSeq"

If m_rs.State = 1 Then m_rs.Close
m_rs.CursorLocation = adUseClient
m_rs.Open strA, cnnConnection, adOpenStatic, adLockReadOnly
iPage = 0

GetPleft

If Not m_rs.EOF And Not m_rs.BOF Then
    With m_rs
       m_rs.MoveFirst

       t_A = LTrim(RTrim(m_rs!A0802)) '公司別
       t_B = LTrim(RTrim(m_rs!A0k04)) '收據抬頭
       tA0k11 = Trim(m_rs!A0k11) 'Added by Lydia 2015/06/17
       
       If m_rs!A0j07 = "Y" Then
         RmFee1 = Val(m_rs!服務費) + Val(m_rs!規費)
         RmFee2 = 0
       Else
         RmFee1 = Val(m_rs!服務費)
         RmFee2 = Val(m_rs!規費)
       End If
       RmDFee1 = Val(m_rs!扣繳金額)
       
       PrintHeader '列印表頭
       
       Do While Not m_rs.EOF
            iPage = iPage + 1

           '明細
             strTemp(1) = m_rs!收據日期
             strTemp(2) = LTrim(RTrim(m_rs!a0k01))
             strTemp(3) = LTrim(RTrim(lblSalesName))
             strTemp(4) = m_rs!Rcaseno '本所案號
             'Modified by Morgan 2020/3/3 修正案件名稱Null問題
             'Modified by Morgan 2024/9/18 修正太長最後變?號問題
             'strTemp(5) = StrConv(LeftB(StrConv(LTrim(RTrim("" & m_rs!案件名稱)), vbFromUnicode), 22), vbUnicode) '取中文混雜指定長度(中文2,英文1 byte) 'Mid(LTrim(RTrim(m_rs!案件名稱)), 1, 13)
             strTemp(5) = convForm(("" & m_rs!案件名稱), 23, "")
             'end 2024/9/18
             strTemp(6) = Mid(m_rs!案件性質, 1, 6)
             strTemp(7) = m_rs!申請國家
             If m_rs!A0j07 = "Y" Then
                strTemp(8) = Val(m_rs!服務費) + Val(m_rs!規費)
                strTemp(9) = 0
             Else
                strTemp(8) = Val(m_rs!服務費)
                strTemp(9) = Val(m_rs!規費)
             End If
             strTemp(10) = Val(strTemp(8)) + Val(strTemp(9))
             
             strTemp(11) = Val(m_rs!扣繳金額)
             strTemp(12) = Val(strTemp(10)) - Val(strTemp(11))
       
            If t_A <> LTrim(RTrim(m_rs!A0802)) Or t_B <> LTrim(RTrim(m_rs!A0k04)) Then  '公司別+收據抬頭=>組別(換頁)

                Call PrintSum(1, RmFee1, RmFee2, RmDFee1)
                'Added by Lydia 2015/06/17 列印匯款帳號和列印備註
                    PrintNewLine
                    Printer.Line (ciStartX, iPrint)-(PLeft(13) - 200, iPrint), vbBlack
                    'Added by Lydia 2017/05/25 非智權加印備註
                    If InStr(t_A, "智權") = 0 Then
                        PrintNewLine
                        Printer.CurrentX = PLeft(6)
                        Printer.CurrentY = iPrint
                        Printer.Print "請直接由款項中扣除本所服務費10%所得稅(惟每次稅款如低於2000元，則依法請勿代扣繳)。"
                    End If
                    'end 2017/05/25
                    
                    Call PrintRptAccount(tA0k11)
                    tA0k11 = Trim(m_rs!A0k11)
                'end 2015/06/17
                
               t_A = LTrim(RTrim(m_rs!A0802)) '公司別
               t_B = LTrim(RTrim(m_rs!A0k04)) '收據抬頭
               
                If m_rs!A0j07 = "Y" Then
                  RmFee1 = Val(m_rs!服務費) + Val(m_rs!規費)
                  RmFee2 = 0
                Else
                  RmFee1 = Val(m_rs!服務費)
                  RmFee2 = Val(m_rs!規費)
                End If
                RmDFee1 = Val(m_rs!扣繳金額)

               Printer.NewPage
               PrintHeader

            Else
              If .AbsolutePosition > 1 Then
                RmFee1 = RmFee1 + strTemp(8)
                RmFee2 = RmFee2 + strTemp(9)
                RmDFee1 = RmDFee1 + strTemp(11)
              End If
            End If
            
            PrintDetail '列印明細
                    
            If .AbsolutePosition = .RecordCount Then
               Call PrintSum(1, RmFee1, RmFee2, RmDFee1)
            End If
              
            m_rs.MoveNext
        Loop

    End With


'列印總計和表尾

Call PrintSum(2, mFee1, mFee2, mdFee1)

Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print String(29, "＝")
PrintNewLine

Printer.Line (ciStartX, iPrint)-(PLeft(13) - 200, iPrint), vbBlack
'Added by Lydia 2017/05/25 非智權加印備註
If InStr(t_A, "智權") = 0 Then
    PrintNewLine
    Printer.CurrentX = PLeft(6)
    Printer.CurrentY = iPrint
    Printer.Print "請直接由款項中扣除本所服務費10%所得稅(惟每次稅款如低於2000元，則依法請勿代扣繳)。"
End If
'end 2017/05/25

'Added by Lydia 2015/06/17 列印匯款帳號和列印備註
Call PrintRptAccount(tA0k11)

Else
   MsgBox "無符合列印的資料!!!", vbExclamation + vbOKOnly
   Exit Sub
End If


Printer.EndDoc
PUB_RestorePrinter strPrinter 'Added by Morgan 2020/10/30

ShowPrintOk
End Sub


Sub PrintHeader()
Dim paStr As String

Dim strPTmp As String
iPrint = ciStartY
Printer.FontName = "新細明體" 'Added by Lydia 2016/09/14 Windows 7 字型會變
Printer.Font.Size = ciTitleFontSize
Printer.Font.Bold = True
Printer.Font.Underline = False
'Modified by Lydia 2019/07/01
'strPTmp = "客戶請款明細表"
strPTmp = "請款明細表"
'Added by Lydia 2015/06/17 改表頭
'Printer.CurrentX = (lngPageWidth - Printer.TextWidth(strPTmp)) / 2
Printer.CurrentX = (lngPageWidth - Printer.TextWidth(t_A)) / 2
Printer.CurrentY = iPrint
Printer.Print t_A  '公司別--置頂
iPrint = iPrint + 420
Printer.FontName = "新細明體" 'Added by Lydia 2016/09/14 Windows 7 字型會變
Printer.Font.Size = ciFontSize + 2
Printer.CurrentX = (lngPageWidth - Printer.TextWidth(strPTmp)) / 2
Printer.CurrentY = iPrint
'end 2015/06/17
Printer.Print strPTmp

iPrint = iPrint + 300
Printer.FontName = "新細明體" 'Added by Lydia 2016/09/14 Windows 7 字型會變
Printer.Font.Size = ciFontSize
Printer.Font.Bold = False
Printer.Font.Underline = False
'Remove by Lydia 2015/06/17 改表頭

    'PrintNewLine  '換行,判斷Y座標
    'paStr = "客戶編號：" & LTrim(RTrim(txtCustNo(0))) & "~" & LTrim(RTrim(txtCustNo(1)))
    'Printer.CurrentX = 7200
    'Printer.CurrentY = iPrint
    'Printer.Print paStr

    'PrintNewLine
    'paStr = "收據抬頭：" & LTrim(RTrim(txtTitle))
    'Printer.CurrentX = 7200
    'Printer.CurrentY = iPrint
    'Printer.Print paStr
'end 2015/06/17


PrintNewLine
Printer.CurrentX = ciStartX
Printer.CurrentY = iPrint
'Modified by Morgan 2024/9/25
'Printer.Print "列印人員：" & strUserName
PUB_PrintUnicodeText "列印人員：" & strUserName, Printer.CurrentX, Printer.CurrentY, 0
'end 2024/9/25

paStr = "帳款日期：" & LTrim(RTrim(minDate)) & "~" & LTrim(RTrim(maxDate))
Printer.CurrentX = 7200
Printer.CurrentY = iPrint
Printer.Print paStr

Printer.CurrentX = lngPageWidth - Printer.TextWidth(String(20, "　"))
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(strSrvDate(2), "##/##/##")


PrintNewLine
Printer.CurrentX = ciStartX
Printer.CurrentY = iPrint
'Move by Lydia 2015/06/17
'Printer.Print "公  司  別：" & t_A
'Modified by Morgan 2024/9/25
'Printer.Print "收據抬頭：" & t_B
PUB_PrintUnicodeText "收據抬頭：" & t_B, Printer.CurrentX, Printer.CurrentY, 0
'end 2024/9/25

Printer.CurrentX = lngPageWidth - Printer.TextWidth(String(20, "　"))
Printer.CurrentY = iPrint
Printer.Print "頁　　次：" & Printer.Page

'Move by Lydia 2015/06/17 上移一行
    'PrintNewLine
    'Printer.CurrentX = ciStartX
    'Printer.CurrentY = iPrint
    'Printer.Print "收據抬頭：" & t_B


PrintNewLine
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "收據日期"

Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "收據號碼"

Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "智權人員"

Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "本所案號"

Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "案件名稱"

Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "案件性質"

Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "申請國家"

Printer.CurrentX = PLeft(9) - Printer.TextWidth("服務費") - ciColGap
Printer.CurrentY = iPrint
Printer.Print "服務費"

Printer.CurrentX = PLeft(10) - Printer.TextWidth("規費") - ciColGap
Printer.CurrentY = iPrint
Printer.Print "規費"

Printer.CurrentX = PLeft(11) - Printer.TextWidth("請款金額") - ciColGap
Printer.CurrentY = iPrint
Printer.Print "請款金額"

Printer.CurrentX = PLeft(12) - Printer.TextWidth("扣繳金額") - ciColGap
Printer.CurrentY = iPrint
Printer.Print "扣繳金額"

Printer.CurrentX = PLeft(13) - Printer.TextWidth("應付金額") - ciColGap
Printer.CurrentY = iPrint
Printer.Print "應付金額"

PrintNewLine

Printer.Line (ciStartX, iPrint)-(PLeft(13) - 200, iPrint), vbBlack

iPrint = iPrint + 150

End Sub

Sub PrintDetail()
Dim pB As String ', dSpace As Integer
'收據日期
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print strTemp(1)
'收據號碼
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print strTemp(2)
'智權人員
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
'Modified by Morgan 2024/9/25
'Printer.Print strTemp(3)
PUB_PrintUnicodeText strTemp(3), Printer.CurrentX, Printer.CurrentY, 0
'end 2024/9/25
'本所案號
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print strTemp(4)
'案件名稱
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
'Modified by Morgan 2024/9/25
'Printer.Print strTemp(5)
PUB_PrintUnicodeText strTemp(5), Printer.CurrentX, Printer.CurrentY, 0
'end 2024/9/25
'案件性質
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print strTemp(6)
'申請國家
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print strTemp(7)

'服務費
pB = Format(strTemp(8), DDollar2)
Printer.CurrentX = PLeft(9) - Printer.TextWidth(pB) - ciColGap
Printer.CurrentY = iPrint
Printer.Print pB
'規費
pB = Format(strTemp(9), DDollar2)
Printer.CurrentX = PLeft(10) - Printer.TextWidth(pB) - ciColGap
Printer.CurrentY = iPrint
Printer.Print pB
'請款金額
pB = Format(strTemp(10), DDollar2)
Printer.CurrentX = PLeft(11) - Printer.TextWidth(pB) - ciColGap
Printer.CurrentY = iPrint
Printer.Print pB
'扣繳
pB = Format(strTemp(11), DDollar2)
Printer.CurrentX = PLeft(12) - Printer.TextWidth(pB) - ciColGap
Printer.CurrentY = iPrint
Printer.Print pB
'應付金額
pB = Format(strTemp(12), DDollar2)
Printer.CurrentX = PLeft(13) - Printer.TextWidth(pB) - ciColGap
Printer.CurrentY = iPrint
Printer.Print pB

PrintNewLine
End Sub


Private Sub PrintSum(m_T As Integer, m_M1 As Double, m_M2 As Double, m_D1 As Double)
Dim pB2 As String ', dSpace2 As Integer

Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print String(29, "－")

PrintNewLine
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
If m_T = 1 Then
  Printer.Print "小計："
Else
  Printer.Print "合計："
End If

'服務費
pB2 = Format(m_M1, DDollar2)
Printer.CurrentX = PLeft(9) - Printer.TextWidth(pB2) - ciColGap
Printer.CurrentY = iPrint
Printer.Print pB2
'規費
pB2 = Format(m_M2, DDollar2)
Printer.CurrentX = PLeft(10) - Printer.TextWidth(pB2) - ciColGap
Printer.CurrentY = iPrint
Printer.Print pB2
'請款金額
pB2 = Format(m_M1 + m_M2, DDollar2)
Printer.CurrentX = PLeft(11) - Printer.TextWidth(pB2) - ciColGap
Printer.CurrentY = iPrint
Printer.Print pB2
'扣繳
pB2 = Format(m_D1, DDollar2)
Printer.CurrentX = PLeft(12) - Printer.TextWidth(pB2) - ciColGap
Printer.CurrentY = iPrint
Printer.Print pB2
'應付金額
pB2 = Format(m_M1 + m_M2 - m_D1, DDollar2)
Printer.CurrentX = PLeft(13) - Printer.TextWidth(pB2) - ciColGap
Printer.CurrentY = iPrint
Printer.Print pB2

PrintNewLine

End Sub
'END by Lydia 2014/9/25 (原frm210141) 新增"客戶請款明細列印"

'Added by Lydia 2015/06/17 設定匯款資料
Private Sub SetRptAccount(pAcctX As String, Optional ByVal bolAddr As Boolean = False)
Dim pAcctX2 As String 'Memo by Lydia 2020/03/31 銀行帳號：寄到那一個所
Dim stAccPerson As String, stTxtPerson As String 'Add by Amy 2024/05/17

    sra(0) = ""  'Modified by Lydia 2015/07/09 帳戶+人名
    Select Case pAcctX
        Case "1"
             'Modified by Lydia 2022/03/21 增加代碼
             sra(1) = "瑞興商業銀行長安分行(代碼101)": sra(2) = "台一國際專利商標事務所"
             sra(3) = "0075 21 0149980": pAcctX2 = "1"
        Case "2"
             'Modified by Lydia 2020/12/15 戶名「台一國際專利法律事務所」=>台一國際智慧財產事務所
              'Modified by Lydia 2022/03/21 增加代碼
             sra(1) = "華南商業銀行長安分行(代碼008)": sra(2) = "台一國際智慧財產事務所"
             sra(3) = "145-10-020233-0": pAcctX2 = "1"
        Case "3"
             'Modified by Lydia 2020/07/24 戶名「台一國際專利法律事務所」=>台一國際智慧財產事務所
             'Modified by Lydia 2020/08/07 行名「中國信託商業銀行」=>中國信託銀行
             'Modified by Lydia 2022/03/21 增加代碼
             sra(1) = "中國信託銀行科博館分行(代碼822)": sra(2) = "台一國際智慧財產事務所"
             sra(3) = "159540272896": pAcctX2 = "2"
        'Remove by Lydia 2020/08/18
        'Case "4"
        '     sra(1) = "日盛國際商業銀行台南分行": sra(2) = "台一國際專利商標事務所"
        '     sra(3) = "007-01022093-900": pAcctX2 = "3"
        'Case "5"
        '     sra(1) = "彰化銀行建興分行"
        '     'Modified by Lydia 2015/07/09 帳戶+人名
        '     'sra(2) = "台一國際專利商標事務所林景郁"
        '     sra(2) = "台一國際專利商標事務所": sra(0) = "林景郁"
        '     sra(3) = "9634 01 003460 00":  pAcctX2 = "4"
        ''Added by Lydia 2017/03/06 高所 +合庫
        'Case "6"
        '     sra(1) = "合作金庫銀行七賢分行": sra(2) = "台一國際專利商標事務所"
        '     'Modified by Lydia 2017/11/27 因為客戶ATM轉帳有006會造成轉帳失敗,拿掉006
        '     'sra(3) = "006 5252 717 319173": pAcctX2 = "4"
        '     sra(3) = "5252 717 319173": pAcctX2 = "4"
        'end 2020/08/18
        'Added by Lydia 2020/03/31
        Case "7"
             'Modified by Lydia 2022/03/21 增加代碼
             sra(1) = "合作金庫銀行七賢分行(代碼006)": sra(2) = "台一國際智慧財產事務所"
             sra(3) = "5252 717 321283": pAcctX2 = "4"
        'Added by Lydia 2020/03/31
        Case "8"
             'Modified by Lydia 2022/03/21 增加代碼
             'Modified by Lydia 2023/04/10 被合併
             'sra(1) = "日盛國際商業銀行台南分行(代碼815)": sra(2) = "台一國際智慧財產事務所"
             'sra(3) = "107-30304658-888": pAcctX2 = "3"
             sra(1) = "台北富邦銀行東城分行(代碼012)": sra(2) = "台一國際智慧財產事務所"
             sra(3) = "05730304658888": pAcctX2 = "3"
        'Added by Lydia 2020/03/31
        Case "9"
             'Modified by Lydia 2022/03/21 增加代碼
             sra(1) = "瑞興商業銀行長安分行(代碼101)": sra(2) = "台一國際智慧財產事務所"
             sra(3) = "0075 21 1756680": pAcctX2 = "1"
        Case "J"
             'Modified by Lydia 2022/03/21 增加代碼
             sra(1) = "瑞興商業銀行長安分行(代碼101)": sra(2) = "台一智權股份有限公司"
             sra(3) = "0075 21 1607750": pAcctX2 = "1"
        'Added by Lydia 2020/03/31
        Case "L"
             'Modified by Lydia 2022/03/21 增加代碼
             sra(1) = "瑞興商業銀行長安分行(代碼101)": sra(2) = "台一國際法律事務所"
             sra(3) = "0075 21 1756890": pAcctX2 = "1"
    End Select
If bolAddr = True Then
    Select Case pAcctX2
        Case "1" '台北主事務所
             'Modify by Amy 2024/05/15 財務2個特殊設定拆成3個
            If Val(strSrvDate(1)) >= Val(財務拆總帳出納國內應收啟用日) Then
                stAccPerson = Pub_GetSpecMan("財務處應收處理人員")
            Else
               stAccPerson = Pub_GetSpecMan("財務處出納人員")
            End If
            stTxtPerson = stAccPerson '取第一個人
            If InStr(stTxtPerson, ";") > 0 Then stTxtPerson = Mid(stTxtPerson, 1, Val(InStr(stTxtPerson, ";")) - 1)
             sra(4) = "10491 台北市中山區長安東路二段112號9樓"    '地址
             sra(5) = "(02)2506-8147"                      '傳真
             sra(6) = Replace(stTxtPerson, ";", "@taie.com.tw;")  'E-mail
             'end 2024/05/17
        Case "2" '台中辦公室
             sra(4) = "40353 台中市西區臺灣大道二段300號10樓"
             sra(5) = "(04)23227483"
             sra(6) = "" & Replace(Pub_GetSpecMan("出納人員-中所"), ";", "@taie.com.tw;")
        Case "3" '台南辦公室
             sra(4) = "70141 台南市東區府連路364號4樓"
             sra(5) = "(06)274-4030"
             sra(6) = Replace(Pub_GetSpecMan("出納人員-南所"), ";", "@taie.com.tw;")
        Case "4" '高雄辦公室
             sra(4) = "80750 高雄市三民區建國二路36號8樓"
             sra(5) = "(07)236-4360"
             sra(6) = Replace(Pub_GetSpecMan("出納人員-高所"), ";", "@taie.com.tw;")
    End Select
    sra(6) = IIf(Right(sra(6), 1) = ";", Mid(sra(6), 1, Len(sra(6)) - 1), sra(6))
End If

End Sub
'Added by Lydia 2015/06/17 設定匯款帳號選項(依USER的所別有所不同)
Private Sub SetCombo2()
     
    Combo2(0).Clear: Combo2(1).Clear
    Select Case PUB_GetST06(txtSales) 'pub_strUserOffice
       'Memo by Lydia 2020/03/31 重整編號
       Case "1" '北所
            'Memo by Lydia 2020/12/15 統一將.Text最前面的數字改放到ItemData
            Combo2(0).AddItem "瑞興－專利商標"
            Combo2(0).ItemData(0) = 1
            
            'Modified by Lydia 2020/12/10 更名
            'Combo2(1).AddItem "2 華南－專利法律"
            Combo2(1).AddItem "華南－台一智慧"
            Combo2(1).ItemData(0) = 2
            Combo2(1).AddItem "瑞興－台一智慧"
            Combo2(1).ItemData(1) = 9
       Case "2" '中所
            Combo2(0).AddItem "1 瑞興－專利商標"
            Combo2(0).ItemData(0) = 1
            
            'Modified by Lydia 2020/07/24 更名
            'Combo2(1).AddItem "3 中國信託－專利法律"
            Combo2(1).AddItem "中國信託－台一智慧"
            Combo2(1).ItemData(0) = 3
            'Modified by Lydia 2020/12/10 更名
            'Combo2(1).AddItem "2 華南－專利法律"
            Combo2(1).AddItem "華南－台一智慧"
            Combo2(1).ItemData(1) = 2
            Combo2(1).AddItem "瑞興－台一智慧"
            Combo2(1).ItemData(2) = 9
       Case "3" '南所
            'Combo2(0).AddItem "4 日盛－專利商標" 'Remove by Lydia 2020/08/18

            'Modified by Lydia 2020/12/10 更名
            'Combo2(1).AddItem "2 華南－專利法律"
            Combo2(1).AddItem "華南－台一智慧"
            Combo2(1).ItemData(0) = 2
            'Combo2(1).AddItem "4 日盛－專利商標"  'Remove by Lydia 2020/08/18
            'Modified by Lydia 2023/04/10 被合併
            'Combo2(1).AddItem "日盛－台一智慧"
            Combo2(1).AddItem "富邦－台一智慧"
            Combo2(1).ItemData(1) = 8
            Combo2(1).AddItem "瑞興－台一智慧" 'Added by Lydia 2020/04/07
            Combo2(1).ItemData(2) = 9
       Case "4" '高所
            'Combo2(0).AddItem "5 彰銀－專利商標" 'Remove by Lydia 2020/08/18
            'Combo2(0).AddItem "6 合庫－專利商標" 'Remove by Lydia 2020/08/18

            'Modified by Lydia 2020/12/10 更名
            'Combo2(1).AddItem "2 華南－專利法律"
            Combo2(1).AddItem "華南－台一智慧"
            Combo2(1).ItemData(0) = 2
            'Combo2(1).AddItem "6 合庫－專利商標" 'Added by Lydia 2017/03/06 高所 +合庫 'Remove by Lydia 2020/12/10
            Combo2(1).AddItem "合庫－台一智慧" 'Added by Lydia 2020/03/31
            Combo2(1).ItemData(1) = 7
            Combo2(1).AddItem "瑞興－台一智慧" 'Added by Lydia 2020/04/07
            Combo2(1).ItemData(2) = 9
       Case Else
            Combo2(0).AddItem "瑞興－專利商標"
            Combo2(0).ItemData(0) = 1
            
            'Modified by Lydia 2020/12/10 更名
            'Combo2(1).AddItem "2 華南－專利法律"
            Combo2(1).AddItem "華南－台一智慧"
            Combo2(1).ItemData(0) = 2
            Combo2(1).AddItem "瑞興－台一智慧"
            Combo2(1).ItemData(1) = 9
    End Select
End Sub

Private Sub Combo2_Click(Index As Integer)
    If Combo2(Index).Text <> "" Then
       'Modified by Lydia 2020/12/15 統一將.Text最前面的數字改放到ItemData
       'Call SetRptAccount(Left(Combo2(Index).Text, 1))
       strExc(0) = Combo2(Index).ItemData(Combo2(Index).ListIndex)
       SetRptAccount strExc(0), True
    End If
End Sub
'Added by Lydia 2015/06/17 列印匯款帳號和列印備註
Private Sub PrintRptAccount(stAcnt As String)
Dim StrS As String
Dim posX As Integer, bolMemo As Boolean, idx As Integer
Dim arrM
Dim iTmp As Integer 'Added by Lydia 2018/03/31
Dim intMove As Integer 'Added by Lydia 2022/03/21

    Call PrintNewLine(True, 10) '+保留7行
    iPrint = lngPageHeight - 10 * lngLineHeight '固定在頁尾
    intMove = 150 'Added by Lydia 2022/03/21 因為增加行庫代碼,所以備註往右移
    
    'Modified by Lydia 2020/03/31 重整公司
'    If stAcnt <> "J" Then
'        StrS = Combo2(Val(stAcnt) - 1).Text
'        '抓指定匯款帳號
'        SetRptAccount Left(Combo2(Val(stAcnt) - 1).Text, 1), True
'    Else
'        SetRptAccount "J", True
'    End If
    If stAcnt = "J" Then
        SetRptAccount "J", True
    ElseIf stAcnt = "L" Then
        SetRptAccount "L", True
    Else
        stAcnt = IIf(stAcnt = "1", "2", "2") 'Added by Lydia 2020/12/10 取銷 1 專利商標
        'Modified by Lydia 2020/12/15 統一將.Text最前面的數字改放到ItemData
        'StrS = Combo2(Val(stAcnt) - 1).Text
        ''抓指定匯款帳號
        'SetRptAccount Left(Combo2(Val(stAcnt) - 1).Text, 1), True
        StrS = Combo2(Val(stAcnt) - 1).Text
        strExc(0) = Combo2(Val(stAcnt) - 1).ItemData(Combo2(Val(stAcnt) - 1).ListIndex)
        SetRptAccount strExc(0), True  '抓指定匯款帳號
    End If
    'end 2020/03/31
    
    If Len(rptMemo) > 0 Then
       arrM = Split(rptMemo, vbCrLf)
       bolMemo = True:   idx = UBound(arrM)
    End If
    
    Printer.Font = ciFontSize
    Printer.CurrentX = ciStartX
    Printer.CurrentY = iPrint
    Printer.Print "付款方式："
    If bolMemo Then
        'Modified by Lydia 2022/03/21 + intMove
        Printer.CurrentX = PLeft(9) + intMove
        Printer.CurrentY = iPrint
        Printer.Print "列印備註："
    End If
    'Line 1
    PrintNewLine
    Printer.CurrentX = ciStartX
    Printer.CurrentY = iPrint
    'Modified by Lydia 2022/03/23 本公司=>本所
    Printer.Print "1. 請將上述款項匯入本所以下帳戶"
    If bolMemo Then
        'Modified by Lydia 2022/03/21 + intMove
        Printer.CurrentX = PLeft(9) + intMove
        Printer.CurrentY = iPrint
        Printer.Print arrM(0)
    End If
    'Line 2
    PrintNewLine
    posX = ciStartX + Printer.TextWidth(String(1, "　"))
    Printer.CurrentX = posX
    Printer.CurrentY = iPrint
    Printer.Print "帳戶：" & sra(1)
    'Modified by Lydia 2016/09/14
    'Printer.FontBold = True
    'Added by Lydia 2018/03/31
    iTmp = posX + Printer.TextWidth("帳戶：" & sra(1)) + Printer.TextWidth(String(2, "　"))
    Printer.Font.Size = 13 '放大字體
    'end 2081/03/31
    Printer.Font.Bold = True
    'Modified by Lydia 2018/03/31
    'Printer.CurrentX = posX + Printer.TextWidth("帳戶：" & sra(1)) + Printer.TextWidth(String(2, "　"))
    'Printer.CurrentY = iPrint
    Printer.CurrentX = iTmp
    Printer.CurrentY = iPrint - 50
    'end 2018/03/31
    'Modified by Lydia 2015/07/09 帳戶+人名
    'Modified by Lydia 2018/03/31 少一個全形空白
    Printer.Print "戶名：" & sra(2) & sra(0) & "　" & "帳號：" & sra(3)
    Printer.Font.Size = ciFontSize 'Added by Lydia 2018/03/31 還原
    If bolMemo And idx >= 1 Then
        'Modified by Lydia 2016/09/14
        'Printer.FontBold = False
        Printer.Font.Bold = False
        'Modified by Lydia 2022/03/21 + intMove
        Printer.CurrentX = PLeft(9) + intMove
        Printer.CurrentY = iPrint
        Printer.Print arrM(1)
    End If
    'Line 3
    PrintNewLine
    'Modified by Lydia 2016/09/14
    'Printer.FontBold = False
    Printer.Font.Bold = False
    Printer.CurrentX = ciStartX
    Printer.CurrentY = iPrint
    Printer.Print "2. 如以支票支付，支票抬頭請開「" & sra(2) & "」"
    If bolMemo And idx >= 2 Then
        'Modified by Lydia 2022/03/21 + intMove
        Printer.CurrentX = PLeft(9) + intMove
        Printer.CurrentY = iPrint
        Printer.Print arrM(2)
    End If
    'Line 4
    PrintNewLine
    Printer.CurrentX = posX
    Printer.CurrentY = iPrint
    Printer.Print "並請以掛號逕寄 " & sra(4) & " 會計室收"
    If bolMemo And idx >= 3 Then
        'Modified by Lydia 2022/03/21 + intMove
        Printer.CurrentX = PLeft(9) + intMove
        Printer.CurrentY = iPrint
        Printer.Print arrM(3)
    End If
    'Added by Lydia 2017/03/23 增加備註
    'Line 5
    PrintNewLine
    Printer.CurrentX = ciStartX
    Printer.CurrentY = iPrint
    'Modified by Lydia 2017/05/25 備註移到報表尾端
    'Printer.Print "備註：1.請直接由款項中扣除本所服務費10%所得稅(惟每次稅款如低於2000元，則依法請勿代扣繳)。"
    'Remove by Lydia 2021/01/14  不列印”備註：”
    'Printer.Print "備註：付款時請將相關收據以E-mail通知 " & Trim(sra(6)) & "@taie.com.tw"
    If bolMemo And idx >= 4 Then
        'Modified by Lydia 2022/03/21 + intMove
        Printer.CurrentX = PLeft(9) + intMove
        Printer.CurrentY = iPrint
        Printer.Print arrM(4)
    End If
    'end 2017/03/23
    'Modified by Lydia 2017/03/23 5->6
    'Line 6
    PrintNewLine
    'Modified by Lydia 2017/05/25
    'Printer.CurrentX = ciStartX
    Printer.CurrentX = ciStartX + Printer.TextWidth(String(3, "　"))
    Printer.CurrentY = iPrint
    'Modified by Lydia 2017/03/23
    'Printer.Print "備註：付款時請將相關收據以E-mail通知 " & Trim(sra(6)) & "@taie.com.tw"
    'If bolMemo And idx >= 4 Then
    'Modified by Lydia 2017/05/25
    'Printer.Print "　　　2.付款時請將相關收據以E-mail通知 " & Trim(sra(6)) & "@taie.com.tw"
    'Remove by Lydia 2021/01/14  不列印”備註：”
    'Printer.Print "或請傳真 " & sra(5) & " 會計室收，以利沖帳，謝謝您的支持與合作。"
    If bolMemo And idx >= 5 Then
    'end 2017/03/23
        'Modified by Lydia 2022/03/21 + intMove
        Printer.CurrentX = PLeft(9) + intMove
        Printer.CurrentY = iPrint
        'Modified by Lydia 2017/03/23 4->5
        Printer.Print arrM(5)
    End If
    'Modified by Lydia 2017/03/23 6->7
    'Line 7
    'Remove by Lydia 2017/05/25
    'PrintNewLine
    'Printer.CurrentX = ciStartX + Printer.TextWidth(String(3, "　"))
    'Printer.CurrentY = iPrint
    'Printer.Print "或請傳真 " & sra(5) & " 會計室收，以利沖帳，謝謝您的支持與合作。"
    'end 2017/05/25
    'Remove by Lydia 2017/03/23
    'If bolMemo And idx >= 5 Then
    '    Printer.CurrentX = PLeft(9)
    '    Printer.CurrentY = iPrint
    '    Printer.Print arrM(5)
    'End If

End Sub

'Added by Lydia 2021/07/27 呼叫”應收帳款查詢”
Private Sub cmdCall_Click()
    
    If PUB_CheckFormExist("frm210122") Then
        MsgBox "請先關閉〔應收帳款查詢〕畫面！"
        Exit Sub
    End If
    'Added by Lydia 2021/07/16 +判斷有資料才繼續
    'Modified by Lydia 2023/05/04
    'If mSeqNo = "" Then Exit Sub
    'If Adodc1.Recordset.RecordCount = 0 Then Exit Sub
    If mSeqNo = "" Then
        Call CallFrm210122
        Exit Sub
    End If
    If Adodc1.Recordset.RecordCount = 0 Then
        Call CallFrm210122
        Exit Sub
    End If
    
    'Grid資料寫入RDataFactory
    Screen.MousePointer = vbHourglass
    Set Adodc1.Recordset = PUB_CreateRecordset(adoTmp, , , , Me.Name, mSeqNo)
    Screen.MousePointer = vbDefault
    
    strExc(0) = "SELECT R022 FROM RDataFactory where R001 = 'Y' and FormName='" & Me.Name & "' And ID='" & strUserNum & "' and seqno = " & mSeqNo & " " & _
                 "order by rowseq "
    intI = 1
    strExc(1) = ""
    Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
    If intI = 1 Then
       '依順序讀取
       strExc(1) = ""
       RsTemp.MoveFirst
       Do While Not RsTemp.EOF
            If InStr(strExc(1) & ",", "" & RsTemp.Fields("R022")) = 0 Then
                strExc(1) = strExc(1) & "," & RsTemp.Fields("R022")
            End If
            RsTemp.MoveNext
       Loop
       strExc(1) = Mid(strExc(1), 2)
    End If
    
    'Mark by Lydia 2023/05/04 為了方便主管查看不同客戶，應收帳款查詢改為只呼叫畫面不限制帶入客戶代號
    'If strExc(1) = "" Then
    '    MsgBox "請選取記錄!!", vbCritical
    '    Exit Sub
    'Else
    'end 2023/05/04
        '從請款單明細表frm210146：先檢查 ”請款作業及應收查詢”是否有勾選記錄，有多筆則逐一帶入勾選的客戶編號進行查詢。
        Call frm210122.SetParent(Me, strExc(1), IIf(txtSales.Text = "", strUserNum, txtSales.Text))
        frm210122.Show
        Me.Hide
    'End If 'Mark by Lydia 2023/05/04
End Sub

'Added by Lydia 2023/05/04
Private Sub CallFrm210122()
    Call frm210122.SetParent(Me, "")
    frm210122.Show
    Me.Hide
End Sub

'Added by Lydia 2022/02/17 取得勾選的第一筆案件資料
Private Function GetGridSel() As String
Dim strCase(0 To 4) As String  'Added by Lydia 2024/10/23

   GetGridSel = ""
   '判斷有資料才繼續
   If mSeqNo = "" Then Exit Function
   If Adodc1.Recordset.RecordCount = 0 Then Exit Function

      Set adoTmp = Adodc1.Recordset.Clone
      With adoTmp
           .MoveFirst
           Do While Not .EOF
               If .Fields("選取") = "Y" Then
                   If .Fields("CUSTCASE") <> "" Then
                       GetGridSel = "貴會案件案號：" & .Fields("CUSTCASE") & vbCrLf
                   End If
                   GetGridSel = GetGridSel & "本所案號：" & .Fields("本所案號") & vbCrLf & _
                                      "案件名稱：" & .Fields("案件名稱") & vbCrLf & _
                                      "申請國家：" & .Fields("申請國家") & vbCrLf
                   'Added by Lydia 2024/10/23 力成X82532010：增加申請案號、發明人
                   If InStr("力成X82532010", txtCustNo(0)) > 0 Then
                      strCase(0) = Replace(.Fields("本所案號"), "-", "")
                      Call ChgCaseNo(strCase(0), strCase)
                      If strCase(1) <> "" And strCase(2) <> "" Then
                         strCase(0) = "select pa11,getinventorc(pa01,pa02,pa03,pa04) as pidata from patent where pa01='" & strCase(1) & "' and pa02='" & strCase(2) & "' and pa03='" & strCase(3) & "' and pa04='" & strCase(4) & "' "
                         intI = 1
                         Set RsTemp = ClsLawReadRstMsg(intI, strCase(0))
                         If intI = 1 Then
                            GetGridSel = GetGridSel & "申請案號：" & RsTemp.Fields("pa11") & vbCrLf & _
                                         "發明人：" & RsTemp.Fields("pidata") & vbCrLf
                         End If
                      End If
                   End If
                   'end 2024/10/23
                   Exit Do
               End If
               .MoveNext
           Loop
      End With

End Function

'Added by Lydia 2022/02/17  使用Excel產生報表:資策會X38805030,力成X82532010
Private Function SaveXLS_1() As Boolean
Dim xlsRpt As New Excel.Application
Dim wksrpt As New Worksheet
Dim strWkName As String, strFormat As String
Dim strGrpNo As String, strGrpDate As String
Dim strFAno1 As String, strFAmt As String, strFArate As String 'U單號的金額和匯率
Dim intAlign As Integer, intPage As Integer
Dim strTmp(1) As String

Dim strA As String, intA As Integer
Dim rsA As New ADODB.Recordset

strA = " Select R001 as 選取,R002 as 收據日期,R003 as 本所案號,R004 as 案件性質, " & _
       " R005 as 申請國家,R006 as 服務費,R007 as 代收代付款,R008 as 扣繳,NVL(R009,0) as 扣繳金額, " & _
       " R010 as 收據號碼,R011 as 案件名稱,R012 as A0J01,R013 as AMT1,R014 as AMT2, " & _
       " R015 as AMT3,R016 as A0J07,R017 as A0K11,R018 as A0K19,R019 as A0K01, " & _
       " R020 as SFEE,R021 as OFEE,R022 as A0K03,R023 as A0K04,R024 as CU04,R025 as A0K05,A0802,R026 as rCaseNo,R030 as CP10  " & _
       " From RDataFactory, acc080 where R017=A0801 and R001 = 'Y' and FormName='" & Me.Name & "' " & _
       " And ID='" & strUserNum & "' and seqno = " & mSeqNo & " Order by A0K11,A0K04,RowSeq"
If m_rs.State = 1 Then m_rs.Close
m_rs.CursorLocation = adUseClient
m_rs.Open strA, cnnConnection, adOpenStatic, adLockReadOnly

If Not m_rs.EOF And Not m_rs.BOF Then

    '-------預設Excel
    SaveXLS_1 = False
    intAlign = 0 '0-置中/1-靠左/2-靠右
    '起始位置intField=65=>A
    intField = 65:  intRow = 1: intTitleR = 1: intPage = 1
    strFileN = txtCustNo(0).Text & "請款明細表" & ServerDate & ServerTime & MsgText(43)
    If Dir(strExcelPath & strFileN) = MsgText(601) Then
        If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
            MkDir strExcelPath
        End If
    Else
        Kill strExcelPath & strFileN
    End If
    xlsRpt.SheetsInNewWorkbook = 3
    xlsRpt.Workbooks.add
    '工作表名稱改為中文
    If strWkName = MsgText(601) Then strWkName = Left(xlsRpt.Worksheets(1).Name, Len(xlsRpt.Worksheets(1).Name) - 1)
    Set wksrpt = xlsRpt.Worksheets(strWkName & intPage)
    wksrpt.Activate
    bolOpenXls = True
    strAllF = "收據日期,收據號碼,案件性質,原幣別,匯率,服務費,代收代付款,請款金額,扣繳金額,應付金額"
    strWidth = "10,10,40,10,10,10,10,10,10,10"
    strField = Split(strAllF, ",")
    intWidth = Split(strWidth, ",")
    '-------預設Excel
    
    With m_rs
       .MoveFirst
       Do While Not .EOF
            '不同收據換頁
            If strGrpNo <> "" & .Fields("A0K11") & .Fields("收據號碼") Then
                 If strGrpNo <> "" Then
                     Call SetXlsEnd_1(xlsRpt, wksrpt, Left(strGrpNo, 1), Mid(strGrpNo, 2), strGrpDate)
                     wksrpt.Name = Mid(strGrpNo, 2)
                      intPage = intPage + 1
                      If intPage > 3 Then
                          xlsRpt.Worksheets.add
                      End If
                      Set wksrpt = xlsRpt.Worksheets(strWkName & intPage)
                      wksrpt.Activate
                 End If
                 intRow = 1: intTitleR = 1
                 wksrpt.Range("A:A").Font.Size = 11
                 Call SetXlsTitle_1(wksrpt, "" & m_rs.Fields("A0802"), "" & m_rs.Fields("A0k04"))
                 intUL = 0
            End If
            For iX = LBound(strField) To UBound(strField)
               strFormat = "": intAlign = 0: strTmp(1) = ""
               strTmp(0) = Replace(strField(iX), "<br>", "")
               Select Case strTmp(0)
                    Case "收據日期"
                        strFormat = "@"
                        strTmp(1) = "" & m_rs.Fields(strTmp(0))
                        If intUL = 0 And "" & m_rs.Fields("CP10") = "1908" Then intUL = intRow
                        strGrpDate = strTmp(1)
                    Case "收據號碼", "案件性質"
                        strFormat = "@"
                        intAlign = 1
                        strTmp(1) = "" & m_rs.Fields(strTmp(0))
                    Case "原幣別"
                        strFormat = "@"
                        intAlign = 1
                        strTmp(1) = ""  '鎖定性質1908代理人請款
                    Case "匯率"
                        strFormat = "@"
                        intAlign = 2
                        strTmp(1) = ""  '鎖定性質1908代理人請款
                    Case "服務費", "代收代付款", "扣繳金額"
                        intAlign = 2
                        strFormat = "##,##0"
                        strTmp(1) = "" & m_rs.Fields(strTmp(0))
                    Case "請款金額"
                        intAlign = 2
                        strFormat = "##,##0"
                        strTmp(1) = "=(F" & intRow & "+G" & intRow & ")"
                    Case "應付金額"
                        intAlign = 2
                        strFormat = "##,##0"
                        strTmp(1) = "=(H" & intRow & "-I" & intRow & ")"
                End Select
                '設定儲存格格式
                If strFormat <> MsgText(601) Then
                    wksrpt.Range(Chr(intField + iX) & intRow).NumberFormatLocal = strFormat
                End If
                wksrpt.Range(Chr(intField + iX) & intRow).Value = strTmp(1)
                Select Case intAlign
                    Case 0 '置中
                        wksrpt.Range(Chr(intField + iX) & intRow).HorizontalAlignment = xlCenter
                    Case 1 '靠左
                        wksrpt.Range(Chr(intField + iX) & intRow).HorizontalAlignment = xlLeft
                    Case 2 '靠右
                        wksrpt.Range(Chr(intField + iX) & intRow).HorizontalAlignment = xlRight
                End Select
            Next iX
            strGrpNo = "" & .Fields("A0K11") & .Fields("收據號碼")
            intRow = intRow + 1
            .MoveNext
       Loop
    End With
    
    Call SetXlsEnd_1(xlsRpt, wksrpt, Left(strGrpNo, 1), Mid(strGrpNo, 2), strGrpDate)
    wksrpt.Name = Mid(strGrpNo, 2)
    '判斷若版本2007以上改變存格式
    If Val(xlsRpt.Version) < 12 Then
        xlsRpt.Workbooks(1).SaveAs FileName:=strExcelPath & strFileN, FileFormat:=-4143
    Else
        xlsRpt.Workbooks(1).SaveAs FileName:=strExcelPath & strFileN, FileFormat:=56
    End If
    xlsRpt.Workbooks.Close
    xlsRpt.Quit
    SaveXLS_1 = True
    Set xlsRpt = Nothing
    Set wksrpt = Nothing
    Exit Function
Else
    MsgBox "無資料列印！"
End If

ErrHnd1:
    If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
    If bolOpenXls = True Then
        '判斷若版本2007以上改變存格式
        If Val(xlsRpt.Version) < 12 Then
            xlsRpt.Workbooks(1).SaveAs FileName:=strExcelPath & strFileN, FileFormat:=-4143
        Else
            xlsRpt.Workbooks(1).SaveAs FileName:=strExcelPath & strFileN, FileFormat:=56
        End If
        xlsRpt.Workbooks.Close
        xlsRpt.Quit
    End If
    
End Function

'使用Excel產生報表:資策會X38805030,力成X82532010
'***Excel頁首***
Private Sub SetXlsTitle_1(ByRef Wks As Worksheet, ByVal stCompName As String, ByVal stBillTitle As String)
    With Wks
        '***表頭設定***
        .Range(Chr(intField) & intRow).Value = LTrim(RTrim(stCompName)) '公司別
        .Range(Chr(intField) & intRow).Font.Size = 18
        .Range(Chr(intField) & intRow).Font.Bold = True
        .Range(Chr(intField) & intRow & ":" & Chr(intField + UBound(strField)) & intRow).MergeCells = True
        .Range(Chr(intField) & intRow & ":" & Chr(intField + UBound(strField)) & intRow).HorizontalAlignment = xlCenter
        .Range(Chr(intField) & intRow & ":" & Chr(intField + UBound(strField)) & intRow).VerticalAlignment = xlCenter
        intRow = intRow + 1
        .Range(Chr(intField) & intRow).Value = "請款明細表"
        .Range(Chr(intField) & intRow).Font.Size = 14
        .Range(Chr(intField) & intRow).Font.Bold = True
        .Range(Chr(intField) & intRow & ":" & Chr(intField + UBound(strField)) & intRow).MergeCells = True
        .Range(Chr(intField) & intRow & ":" & Chr(intField + UBound(strField)) & intRow).HorizontalAlignment = xlCenter
        .Range(Chr(intField) & intRow & ":" & Chr(intField + UBound(strField)) & intRow).VerticalAlignment = xlCenter
        intRow = intRow + 2
                
        .Range(Chr(intField) & intRow).Value = "收據抬頭：" & LTrim(RTrim(stBillTitle))
        .Range(Chr(intField) & intRow).Font.Size = 11
        .Range(Chr(intField) & intRow).HorizontalAlignment = xlLeft
        .Range(Chr(intField + UBound(strField) - 2) & intRow).Font.Size = 11
        .Range(Chr(intField + UBound(strField) - 2) & intRow).Value = "列印日期：" & CFDate(ACDate(ServerDate))
        intRow = intRow + 1
        
        For iX = LBound(strField) To UBound(strField)
            .Columns(Chr(intField + iX) & ":" & Chr(intField + iX)).ColumnWidth = intWidth(iX)
            .Range(Chr(intField + iX) & intRow).Value = Replace(strField(iX), "<br>", vbCrLf)
            If iX < 5 Then
                .Range(Chr(intField + iX) & intRow).HorizontalAlignment = xlLeft
            Else
                .Range(Chr(intField + iX) & intRow).HorizontalAlignment = xlCenter
            End If
        Next iX
        intTitleR = intRow
        .Range(Chr(intField) & intRow & ":" & Chr(intField + UBound(strField)) & intRow).Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range(Chr(intField) & intRow & ":" & Chr(intField + UBound(strField)) & intRow).Borders(xlEdgeBottom).Weight = xlThin
        intRow = intRow + 1
    End With
End Sub

'使用Excel產生報表:資策會X38805030,力成X82532010
'***Excel頁尾***
Private Sub SetXlsEnd_1(ByRef Xls As Excel.Application, ByRef Wks As Worksheet, ByVal stAcnt As String, ByVal stA0k11 As String, ByVal strGDate As String)
'stAcnt : 選擇匯款帳號
'stA0k11 : 收據號碼
Dim arrM
Dim idx As Integer
Dim strFormat As String, intAlign As Integer
Dim strTmp(1) As String

    '增加匯款手續費(空白列)
    For idx = LBound(strField) To UBound(strField)
        strFormat = "": intAlign = 0: strTmp(1) = ""
        strTmp(0) = Replace(strField(idx), "<br>", "")
        Select Case strTmp(0)
             Case "收據日期"
                 strFormat = "@"
                 strTmp(1) = strGDate
             Case "收據號碼", "案件性質" ', "費用別", "原幣別", "匯率"
                 strFormat = "@"
                 intAlign = 1
                 If strTmp(0) = "收據號碼" Then
                     strTmp(1) = stA0k11
                 ElseIf strTmp(0) = "案件性質" Then
                     strTmp(1) = "匯款手續費"
                 Else
                     strTmp(1) = ""
                 End If
             Case "原幣別"
                 strFormat = "@"
                 intAlign = 1
                 strTmp(1) = ""  '鎖定性質1908代理人請款
             Case "匯率"
                 strFormat = "@"
                 intAlign = 2
                 strTmp(1) = ""  '鎖定性質1908代理人請款
             Case "服務費", "代收代付款", "扣繳金額"
                 intAlign = 2
                 strFormat = "##,##0"
                 strTmp(1) = 0
             Case "請款金額"
                 intAlign = 2
                 strFormat = "##,##0"
                 strTmp(1) = "=(F" & intRow & "+G" & intRow & ")"
             Case "應付金額"
                 intAlign = 2
                 strFormat = "##,##0"
                 strTmp(1) = "=(H" & intRow & "-I" & intRow & ")"
         End Select
         '設定儲存格格式
         If strFormat <> MsgText(601) Then
             Wks.Range(Chr(intField + idx) & intRow).NumberFormatLocal = strFormat
         End If
         Wks.Range(Chr(intField + idx) & intRow).Value = strTmp(1)
         Select Case intAlign
             Case 0 '置中
                 Wks.Range(Chr(intField + idx) & intRow).HorizontalAlignment = xlCenter
             Case 1 '靠左
                 Wks.Range(Chr(intField + idx) & intRow).HorizontalAlignment = xlLeft
             Case 2 '靠右
                 Wks.Range(Chr(intField + idx) & intRow).HorizontalAlignment = xlRight
         End Select
    Next idx

    '取得U單號的資料
    'Modified by Lydia 2025/07/25 因為同一張收據分不同帳單的匯率差距過大，改成平均匯率；ex.CFP-035084收據號E11413510
    'strExc(0) = "select a1903,sum(a1904) amt1,avg(a1906) arate from acc150,acc190 where a1501 in (select cp61 from caseprogress where cp60='" & stA0k11 & "')  and a1501=a1902(+) group by a1903 "
    strExc(0) = "select a1903,sum(a1904) amt1,round(sum(a1904*a1906)/sum(a1904),6) as arate from acc150,acc190 where a1501 in (select cp61 from caseprogress where cp60='" & stA0k11 & "')  and a1501=a1902(+) group by a1903 "
    intI = 1
    strExc(1) = "": strExc(2) = "": strExc(3) = ""
    Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
    If intI = 1 Then
       RsTemp.MoveFirst
       Do While Not RsTemp.EOF
           '若有不同幣別用串接
           strExc(2) = strExc(2) & IIf(strExc(2) <> "", vbCrLf, "") & RsTemp.Fields("a1903") & " " & RsTemp.Fields("amt1")
           strExc(3) = strExc(3) & IIf(strExc(3) <> "", vbCrLf, "") & RsTemp.Fields("arate")
           RsTemp.MoveNext
       Loop
    End If
    If strExc(2) <> "" Then
        If intUL = 0 Then intUL = intRow  '如果沒有勾選1908代理人請款,改放在最後一列
        Wks.Range("D" & intUL).Value = strExc(2)
        Wks.Range("E" & intUL).Value = strExc(3)
    End If
    intRow = intRow + 1
    
    '合計
    Wks.Range("C" & intRow).Value = "合　　計："
    Wks.Range("C" & intRow & ":" & Chr(intField + UBound(strField)) & intRow).HorizontalAlignment = xlRight
    For iX = 5 To UBound(strField)
         Wks.Range(Chr(intField + iX) & intRow).Value = "=SUM(" & Chr(intField + iX) & intTitleR + 1 & ":" & Chr(intField + iX) & intRow - 1 & ")"
         Wks.Range(Chr(intField + iX) & intRow).NumberFormatLocal = "##,##0"
    Next iX
    Wks.Range("D" & intRow & ":" & Chr(intField + UBound(strField)) & intRow).Borders(xlEdgeTop).LineStyle = xlDot
    Wks.Range("D" & intRow & ":" & Chr(intField + UBound(strField)) & intRow).Borders(xlEdgeTop).Weight = xlThin
    intRow = intRow + 1
    Wks.Range(Chr(intField) & intRow & ":" & Chr(intField + UBound(strField)) & intRow).Borders(xlEdgeTop).LineStyle = xlContinuous
    Wks.Range(Chr(intField) & intRow & ":" & Chr(intField + UBound(strField)) & intRow).Borders(xlEdgeTop).Weight = xlThin
    If stAcnt <> "J" Then
      Wks.Range("D" & intRow).Value = "請直接由款項中扣除本所服務費10%所得稅(惟每次稅款如低於2000元，則依法請勿代扣繳)。"
    End If
    
    If intRow < 27 Then
        intRow = 27 '固定在頁尾
    Else
        intRow = intRow + 7 '空7行
    End If
    
    Wks.Range("3:" & intRow + 10).Font.Size = 11 '預設字體大小
    
    If stAcnt = "J" Then
        SetRptAccount "J", True
    ElseIf stAcnt = "L" Then
        SetRptAccount "L", True
    Else
        stAcnt = IIf(stAcnt = "1", "2", "2")
        strExc(0) = Combo2(Val(stAcnt) - 1).ItemData(Combo2(Val(stAcnt) - 1).ListIndex)
        SetRptAccount strExc(0), True  '抓指定匯款帳號
    End If
    
    Wks.Range(Chr(intField) & intRow).Value = "付款方式："
    'Modified by Lydia 2022/03/23 本公司=>本所
    Wks.Range(Chr(intField) & intRow + 1).Value = "1. 請將上述款項匯入本所以下帳戶"
    'Modified by Lydia 2022/03/21 分成兩行
    'strExc(4) = "　帳戶：" & sra(1) & "　戶名：" & sra(2) & sra(0) & "　" & "帳號：" & sra(3)
    'Wks.Range(Chr(intField) & intRow + 2).Value = strExc(4)
    Wks.Range(Chr(intField) & intRow + 2).Value = "　帳戶：" & sra(1)
    Wks.Range(Chr(intField) & intRow + 3).Value = "　戶名：" & sra(2) & sra(0) & "　" & "帳號：" & sra(3)
    Wks.Range(Chr(intField) & intRow + 3).Cells.Font.Size = 14
    Wks.Range(Chr(intField) & intRow + 3).Cells.Font.Bold = True
    'end 2022/03/21
    Wks.Range(Chr(intField) & intRow + 4).Value = "2. 如以支票支付，支票抬頭請開「" & sra(2) & "」"
    Wks.Range(Chr(intField) & intRow + 5).Value = "　並請以掛號逕寄 " & sra(4) & " 會計室收"
    '帳戶字體加粗
    'Mark by Lydia 2022/03/21 分成兩行
    'idx = InStr(strExc(4), "戶名")
    'Wks.Range(Chr(intField) & intRow + 2).Cells.Characters(idx).Font.Size = 14
    'Wks.Range(Chr(intField) & intRow + 2).Cells.Characters(idx).Font.Bold = True
    'end 2022/03/21
    
    If Len(rptMemo) > 0 Then
        arrM = Split(rptMemo, vbCrLf)
        'Modified by Lydia 2022/03/21 6=>5 (G欄)
        Wks.Range(Chr(intField + 5) & intRow).Value = "列印備註："
        For idx = 0 To UBound(arrM)
            'Modified by Lydia 2022/03/21 6=>5 (G欄)
            Wks.Range(Chr(intField + 5) & intRow + 1 + idx).Value = arrM(idx)
        Next idx
    End If

    '設定
    Wks.PageSetup.PaperSize = 9 'A4
    Wks.PageSetup.PrintTitleRows = "$1:$" & intTitleR
    Wks.PageSetup.Orientation = xlLandscape '橫印
    Wks.PageSetup.LeftMargin = 0.4 '邊界
    Wks.PageSetup.RightMargin = 0.4
    Wks.PageSetup.TopMargin = Xls.InchesToPoints(0.4)
    Wks.PageSetup.BottomMargin = Xls.InchesToPoints(0.4)
    Wks.PageSetup.CenterHorizontally = True '版面設定->邊界->水平置中
End Sub

'Added by Lydia 2023/05/08 同收據扣繳要同步 ---- 比照frm210141
Private Sub UpdateTax(pReceiptNo As String, Optional pChecked As Boolean = True)
   Dim rsUpdate As ADODB.Recordset
   
   Set rsUpdate = Adodc1.Recordset.Clone
   With rsUpdate
   
   If pChecked Then
      .MoveFirst
      .Find "a0k01='" & pReceiptNo & "'"
      Do While Not .EOF
         If .Fields("選取") = "Y" Then
            .Find "a0k01='" & pReceiptNo & "'", 1
         Else
            MsgBox "收據(" & pReceiptNo & ")有多個收文號，請全部選取後再扣繳！", vbExclamation
            GoTo NoAct
         End If
      Loop
   End If
   
   
   .MoveFirst
   .Find "a0k01='" & pReceiptNo & "'"
   Do While Not .EOF
      If pChecked Then
         .Fields("扣繳") = "Y"
         '是否合併
         If .Fields("a0j07") = "Y" Then
            .Fields("扣繳金額") = 0.1 * (Val(.Fields("SFee")) + Val(.Fields("OFee")))
         Else
            .Fields("扣繳金額") = 0.1 * Val(.Fields("SFee"))
         End If
         .Fields("扣繳金額") = Round(.Fields("扣繳金額")) 'Added by Morgan 2023/2/14 不要有小數否因為財務收款只會顯示整數導致不知要調整哪一筆
      Else
         .Fields("扣繳") = ""
         .Fields("扣繳金額") = 0
      End If
      .UPDATE
      .Find "a0k01='" & pReceiptNo & "'", 1
   Loop
   
NoAct:
   End With
   
   Set rsUpdate = Nothing
End Sub

'Added by Lydia 2023/05/08 同一收據勾選一筆收文自動預設勾其他收文
'Modified by Lydia 2023/11/13 增加判斷種類pKind
Private Sub UpdateSelected(pReceiptNo As String, Optional ByVal pKind As String = "0")
   Dim rsUpdate As ADODB.Recordset
   
   Set rsUpdate = Adodc1.Recordset.Clone
   With rsUpdate
      .MoveFirst
      If pKind = "0" Then 'Added by Lydia 2023/11/13
         .Find "a0k01='" & pReceiptNo & "'"
      'Added by Lydia 2023/11/13
      ElseIf pKind = "1" Then 'INVOICE編號
         .Find "a0k40='" & pReceiptNo & "'"
      End If
      'end 2023/11/13
      Do While Not .EOF
         If "" & .Fields("選取") <> "Y" Then
            .Fields("選取") = "Y"
            .UPDATE
         End If
         If pKind = "0" Then 'Added by Lydia 2023/11/13
            .Find "a0k01='" & pReceiptNo & "'", 1
         'Added by Lydia 2023/11/13
         ElseIf pKind = "1" Then 'INVOICE編號
            .Find "a0k40='" & pReceiptNo & "'", 1
         End If
         'end 2023/11/13
      Loop
   End With
   
   Set rsUpdate = Nothing
End Sub

'Add By Sindy 2023/6/12
Private Sub Combo3_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub Combo3_LostFocus()
   If Trim(Combo3) <> "" And Trim(Combo3) <> "全部" Then
      arrID = Split(Combo3, " ")
      txtSales = arrID(0)
      Combo3 = txtSales & " " & GetPrjSalesNM(txtSales)
   ElseIf Trim(Combo3) <> "全部" Then
      txtSales = ""
   End If
End Sub
Private Sub Combo3_Validate(Cancel As Boolean)
Dim strEmp As String
Dim stTmp As String 'Add by Amy 2020/03/25
   
   If Combo3 <> "" And Trim(Combo3) <> "全部" Then
      'Add by Amy 2020/03/25 只能輸入下拉選單中已有的人員
      stTmp = Combo3
      '直接輸員編未串名字會錯
      If InStr(stTmp, " ") > 0 Then
        stTmp = Mid(stTmp, 1, Val(InStr(stTmp, " ")) - 1)
      Else
        stTmp = Combo3
      End If
      'Modify By Sindy 2020/6/15 Mark
'      If InStr(m_strListPer, stTmp) = 0 And stTmp <> strUserNum And Pub_StrUserSt03 <> "M51" Then
'         MsgBox "不可輸入下拉選單以外的人員！"
'         Cancel = True
'         Combo3.SetFocus
'         Exit Sub
'      End If
      'end 2020/03/25
      arrID = Split(Combo3, " ")
      txtSales = arrID(0)
      lblSalesName.Caption = GetStaffName(txtSales, True)
      If lblSalesName.Caption = "" Then
         MsgBox "智權人員輸入錯誤！", vbCritical
         Combo3.SetFocus
         Cancel = True
      End If
      Combo3 = txtSales & " " & GetPrjSalesNM(txtSales)
   'Modify By Sindy 2024/8/5 mark; 因 txtSales_Validate 會檢查相關的權限
   Else
      txtSales = ""
'   'Modify by Amy 2023/05/09 +st05
'   ElseIf Combo3 = MsgText(601) And stST05 <> "00" And stST05 <> "01" And stST05 <> "08" Then
'        'Add by Amy 2020/03/25 下拉選單無區主管智權人員不可為空
'        'Modify By Sindy 2020/7/14
'        'If bolAreaMan = False And Pub_StrUserSt03 <> "M51" Then
'        'Modify By Sindy 2023/9/21 開放杜協理權限 + And Not (Mid(txtSalesArea, 1, 1) = "S" And Mid(txtSalesArea1, 1, 1) = "S" And InStr(Pub_GetSpecMan("全所智權部主管"), strUserNum) > 0)
'        If (GetDeptMan(txtSalesArea) <> strUserNum Or GetDeptMan(txtSalesArea1) <> strUserNum) _
'            And Pub_StrUserSt03 <> "M51" _
'            And Not (Mid(txtSalesArea, 1, 1) = "S" And Mid(txtSalesArea1, 1, 1) = "S" And InStr(Pub_GetSpecMan("全所智權部主管"), strUserNum) > 0) Then
'        '2020/7/14 END
'           MsgBox "非區主管職代智權人員不可空白！"
'           Cancel = True
'           Combo3.SetFocus
'           Exit Sub
'        End If
'        'end 2020/03/25
   '2024/8/5 END
   End If
   'end 2016/6/7
End Sub
'2023/6/12 END

'Added by Lydia 2024/04/25  使用Excel產生報表
Private Function PrintExcelMain() As Boolean
Dim xlsPoint As New Excel.Application
Dim WksPoint As New Worksheet
Dim bolOpenxlsPoint As Boolean
Dim strA As String
Dim strA0802 As String, strA0K04 As String, strA0K11 As String
Dim rsA As New ADODB.Recordset
Dim RmFee1 As Double, RmFee2 As Double, RmDFee1 As Double '小計金額
Dim strTemp2(0 To 6) As String

'抬頭資料和總計資料
strA = " select min(R002) as mindate,max(R002) as maxdate," & _
       " NVL(sum(decode (R016,'Y',R006+R007,R006)),0) as fee1," & _
       " NVL(sum(decode (R016,'Y',0,R007)),0) as fee2, NVL(sum(R009),0) as dfee1 " & _
       " From RDataFactory where R001 = 'Y' and FormName='" & Me.Name & "' And ID='" & strUserNum & "' and seqno = " & mSeqNo & " "
Set rsA = Nothing
rsA.Open strA, cnnConnection, adOpenStatic, adLockReadOnly

minDate = "" & rsA!minDate '帳款日期
maxDate = "" & rsA!maxDate
mFee1 = "" & rsA!fee1  '總計:應收-服務費
mFee2 = "" & rsA!fee2  '總計:應收-規費
mdFee1 = "" & rsA!dfee1 '總計:扣繳

strA = " Select R001 as 選取,R002 as 收據日期,R003 as 本所案號,R004 as 案件性質, " & _
       " R005 as 申請國家,R006 as 服務費,R007 as 規費,R008 as 扣繳,NVL(R009,0) as 扣繳金額, " & _
       " R010 as 收據號碼,R011 as 案件名稱,R012 as A0J01,R013 as AMT1,R014 as AMT2, " & _
       " R015 as AMT3,R016 as A0J07,R017 as A0K11,R018 as A0K19,R019 as A0K01, " & _
       " R020 as SFEE,R021 as OFEE,R022 as A0K03,R023 as A0K04,R024 as CU04,R025 as A0K05,A0802,R026 as rCaseNo  " & _
       " From RDataFactory, acc080 where R017=A0801 and R001 = 'Y' and FormName='" & Me.Name & "' " & _
       " And ID='" & strUserNum & "' and seqno = " & mSeqNo & " Order by A0K11,A0K04,RowSeq"

If m_rs.State = 1 Then m_rs.Close
m_rs.CursorLocation = adUseClient
m_rs.Open strA, cnnConnection, adOpenStatic, adLockReadOnly

iPage = 0
If Not m_rs.EOF And Not m_rs.BOF Then
    '-------預設Excel
    PrintExcelMain = False
    intField = 65
    strFileN = "$$請款明細表" & MsgText(43)
    If Dir(strExcelPath & strFileN) = MsgText(601) Then
        If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
            MkDir strExcelPath
        End If
    Else
        Kill strExcelPath & strFileN
    End If
    xlsPoint.SheetsInNewWorkbook = 1
    xlsPoint.Workbooks.add
    xlsPoint.Application.Visible = False
    Set WksPoint = xlsPoint.Worksheets(1)
    bolOpenxlsPoint = True
    '-------預設Excel
    
    With m_rs
       m_rs.MoveFirst
       Do While Not m_rs.EOF
          '不同收據換頁
          If strA0802 & strA0K04 = "" Or (strA0802 & strA0K04 <> "" & m_rs.Fields("A0802") & m_rs.Fields("A0k04")) Then  '公司別+收據抬頭=>組別(換頁)
              If strA0802 & strA0K04 <> "" Then
                 '小計
                 Call PrintExcelEnd("1", xlsPoint, WksPoint, strA0K11, strA0802, strA0K04, RmFee1, RmFee2, RmDFee1)
                 WksPoint.Range(iX & ":" & iX, intRow & ":" & intRow).Delete 'Added by Lydia 2024/11/05
              End If
              Call PrintExcelTitle(WksPoint, "" & m_rs.Fields("A0802"), "" & m_rs.Fields("A0k04"))
              RmFee1 = 0: RmFee2 = 0: RmDFee1 = 0
              If iPage = 1 Then iX = intRow 'Added by Lydia 2024/11/05
          End If
          'Modified by Lydia 2024/11/05　一次只印一頁，所以換頁先處理
          'If intRow > 47 Then '換頁列印
          If m_rs.AbsolutePosition Mod 20 = 0 Then
              If iPage = 1 Then
                 Call PrintExcelMainSet(xlsPoint, WksPoint)
              End If
              WksPoint.PrintOut Copies:=1, Collate:=True
              WksPoint.Range(iX & ":" & iX, intRow & ":" & intRow).Delete 'Added by Lydia 2024/11/05
              Call PrintExcelTitle(WksPoint, "" & m_rs.Fields("A0802"), "" & m_rs.Fields("A0k04"))
          End If
          '資料第1列
          strTemp2(0) = "" & m_rs.Fields("收據日期")
          strTemp2(1) = "" & m_rs.Fields("收據號碼")
          If Combo3 <> "" And Trim(Combo3) <> "全部" Then
             strTemp2(2) = Trim(Mid(Combo3, 7))
          Else
             strTemp2(2) = Trim(lblSalesName)
          End If
          strTemp2(3) = "" & m_rs.Fields("rcaseno")
          strTemp2(4) = "" & m_rs.Fields("案件名稱")
          strTemp2(5) = "" & m_rs.Fields("案件性質")
          strTemp2(6) = "" & m_rs.Fields("申請國家")
          WksPoint.Range(Chr(intField) & intRow & ":" & Chr(intField + UBound(strTemp2)) & intRow).Value = strTemp2
          WksPoint.Range(Chr(intField) & intRow & ":" & Chr(intField + UBound(strTemp2)) & intRow).HorizontalAlignment = xlLeft
          WksPoint.Range(Chr(intField) & intRow & ":" & Chr(intField + UBound(strTemp2)) & intRow).NumberFormatLocal = "@"
          '資料第2列
          intRow = intRow + 1
          strTemp2(0) = "":  strTemp2(1) = ""
          If "" & m_rs.Fields("a0j07") = "Y" Then
             strTemp2(2) = Val("" & m_rs.Fields("服務費")) + Val("" & m_rs.Fields("規費"))
             strTemp2(3) = 0
          Else
             strTemp2(2) = Val("" & m_rs.Fields("服務費"))
             strTemp2(3) = Val("" & m_rs.Fields("規費"))
          End If
          '請款金額
          strTemp2(4) = Val(strTemp2(2)) + Val(strTemp2(3))
          strTemp2(5) = Val("" & m_rs.Fields("扣繳金額"))
          '應付金額
          strTemp2(6) = Val(strTemp2(4)) - Val(strTemp2(5))
          WksPoint.Range(Chr(intField) & intRow & ":" & Chr(intField + UBound(strTemp2)) & intRow).Value = strTemp2
          WksPoint.Range(Chr(intField) & intRow & ":" & Chr(intField + UBound(strTemp2)) & intRow).HorizontalAlignment = xlRight
          WksPoint.Range(Chr(intField + 2) & intRow & ":" & Chr(intField + UBound(strTemp2)) & intRow).NumberFormatLocal = "##,##0"
          WksPoint.Range(Chr(intField) & intRow & ":" & Chr(intField + UBound(strTemp2)) & intRow).Value = WksPoint.Range(Chr(intField) & intRow & ":" & Chr(intField + UBound(strTemp2)) & intRow).Value
          WksPoint.Range(Chr(intField + 1) & intRow & ":" & Chr(intField + UBound(strTemp2)) & intRow).Borders(xlEdgeBottom).LineStyle = xlContinuous
          '--------------------
          RmFee1 = RmFee1 + Val(strTemp2(2)) '服務費
          RmFee2 = RmFee2 + Val(strTemp2(3)) '規費
          RmDFee1 = RmDFee1 + Val(strTemp2(5))  '扣繳金額
          strA0802 = "" & m_rs.Fields("A0802")   '公司名稱
          strA0K04 = "" & m_rs.Fields("A0k04")   '收據抬頭
          strA0K11 = "" & m_rs.Fields("a0k11")  '公司別
          intRow = intRow + 1
          m_rs.MoveNext
       Loop
    End With
    
    '合計
    Call PrintExcelEnd("2", xlsPoint, WksPoint, strA0K11, strA0802, strA0K04, RmFee1, RmFee2, RmDFee1)
    
    WksPoint.Application.DisplayAlerts = False
    xlsPoint.Quit
    PrintExcelMain = True
    Set xlsPoint = Nothing
    Set WksPoint = Nothing
    ShowPrintOk
    Exit Function
Else
    MsgBox "無資料列印！"
End If

ErrHnd1:
    If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
    If bolOpenxlsPoint = True Then
        xlsPoint.Workbooks(1).Close xlDoNotSaveChanges
        xlsPoint.Quit
    End If
    
End Function

'Added by Lydia 2024/04/25  使用Excel產生報表----表頭
Private Sub PrintExcelTitle(ByRef Wks As Worksheet, ByVal stCompName As String, ByVal stBillTitle As String)
Dim intS As Integer, stFields As String, stFieldE As String

   With Wks
      stFields = "A": stFieldE = "G"
      intS = intS + 1
      iPage = iPage + 1
      If iPage = 1 Then
         .Range("A:A").ColumnWidth = 9
         .Range("B:B").ColumnWidth = 9
         .Range("C:C").ColumnWidth = 10
         .Range("D:D").ColumnWidth = 12
         .Range("E:E").ColumnWidth = 21
         .Range("F:F").ColumnWidth = 15
         .Range("G:G").ColumnWidth = 15
      End If

      '***表頭設定***
      .Range(stFields & intS).Value = LTrim(RTrim(stCompName))  '公司別
      .Range(stFields & intS).Font.Size = 16
      .Range(stFields & intS).Font.Bold = True
      If iPage = 1 Then
         .Range(intS & ":" & intS).RowHeight = 20
         .Range(stFields & intS & ":" & stFieldE & intS).MergeCells = True
         .Range(stFields & intS & ":" & stFieldE & intS).HorizontalAlignment = xlCenter
         .Range(stFields & intS & ":" & stFieldE & intS).VerticalAlignment = xlCenter
      End If
      
      intS = intS + 1
      If iPage = 1 Then
         .Range(intS & ":" & intS).RowHeight = 20
         .Range(stFields & intS).Value = "請款明細表"
         .Range(stFields & intS).Font.Size = 14
         .Range(stFields & intS).Font.Bold = True
         .Range(stFields & intS & ":" & stFieldE & intS).MergeCells = True
         .Range(stFields & intS & ":" & stFieldE & intS).HorizontalAlignment = xlCenter
         .Range(stFields & intS & ":" & stFieldE & intS).VerticalAlignment = xlCenter
         '其餘統一列高
         For intI = intS + 1 To 47
            .Range(intI & ":" & intI).RowHeight = 16.5
            .Range(intI & ":" & intI).Font.Size = 10
            .Range(intI & ":" & intI).Font.Bold = False
         Next intI
      End If
      intS = intS + 2
              
      If iPage = 1 Then
         .Range(stFields & intS).Value = "列印人員：" & strUserName
         .Range("D" & intS).Value = "帳款日期：" & LTrim(RTrim(minDate)) & "~" & LTrim(RTrim(maxDate))
         .Range(Chr(Asc(stFieldE) - 1) & intS).Value = "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
         .Range("D" & intS & ":E" & intS).MergeCells = True
         .Range("D" & intS & ":E" & intS).HorizontalAlignment = xlCenter
         .Range("D" & intS & ":E" & intS).VerticalAlignment = xlCenter
         '列印日期會超出右邊界
         .Range(Chr(Asc(stFieldE) - 1) & intS & ":" & stFieldE & intS).MergeCells = True
         .Range(Chr(Asc(stFieldE) - 1) & intS & ":" & stFieldE & intS).HorizontalAlignment = xlRight
         .Range(Chr(Asc(stFieldE) - 1) & intS & ":" & stFieldE & intS).VerticalAlignment = xlCenter
      End If
      
      intS = intS + 1
      .Range(stFields & intS).Value = "收據抬頭：" & stBillTitle
      .Range(stFieldE & intS).Value = "頁　　次：" & iPage
      
      '欄位抬頭1
      intS = intS + 1
      If iPage = 1 Then
         .Range(stFields & intS).Value = "收據日期"
         .Range(Chr(Asc(stFields) + 1) & intS).Value = "收據號碼"
         .Range(Chr(Asc(stFields) + 2) & intS).Value = "智權人員"
         .Range(Chr(Asc(stFields) + 3) & intS).Value = "本所案號"
         .Range(Chr(Asc(stFields) + 4) & intS).Value = "案件名稱"
         .Range(Chr(Asc(stFields) + 5) & intS).Value = "案件性質"
         .Range(stFieldE & intS).Value = "申請國家"
         .Range(stFields & intS & ":" & stFieldE & intS).HorizontalAlignment = xlLeft
         intS = intS + 1
         .Range(Chr(Asc(stFields) + 2) & intS).Value = "服務費"
         .Range(Chr(Asc(stFields) + 3) & intS).Value = "規　費"
         .Range(Chr(Asc(stFields) + 4) & intS).Value = "請款金額"
         .Range(Chr(Asc(stFields) + 5) & intS).Value = "扣繳金額"
         .Range(stFieldE & intS).Value = "應付金額"
         .Range(stFields & intS & ":" & stFieldE & intS).HorizontalAlignment = xlRight
         .Range(stFields & intS & ":" & stFieldE & intS).Borders(xlEdgeBottom).LineStyle = xlContinuous
         .Range(stFields & intS & ":" & stFieldE & intS).Borders(xlEdgeBottom).Weight = xlThin
      Else
         intS = intS + 1
         '清除資料
         .Range(stFields & intS + 1 & ":" & stFieldE & "52").ClearContents
         For intI = intS + 1 To intRow
           .Range(stFields & intI & ":" & stFieldE & intI).Borders(xlEdgeBottom).LineStyle = xlNone
           .Range(intI & ":" & intI).Font.Size = 10
           .Range(intI & ":" & intI).Font.Bold = False
         Next intI
      End If
      
      intS = intS + 1
      intRow = intS
      
   End With
End Sub

'Added by Lydia 2024/04/25  使用Excel產生報表----報表尾端
Private Sub PrintExcelEnd(ByVal pType As String, ByRef Xls As Excel.Application, ByRef Wks As Worksheet, ByVal stAcnt As String, ByVal stCompName As String, ByVal stBillTitle As String, ByVal pFee01 As Double, ByVal pFee02 As Double, ByVal pDFee01 As Double)
'stAcnt : 選擇匯款帳號
Dim arrM
Dim idx As Integer
Dim strTemp2(0 To 6) As String
    
    '小計
    strTemp2(0) = ""
    strTemp2(1) = "小　計："
    strTemp2(2) = pFee01 '服務費
    strTemp2(3) = pFee02 '規費
    strTemp2(4) = pFee01 + pFee02 '請款金額
    strTemp2(5) = pDFee01 '扣繳金額
    strTemp2(6) = pFee01 + pFee02 - pDFee01 '應付金額
    Wks.Range(Chr(intField) & intRow & ":" & Chr(intField + UBound(strTemp2)) & intRow).Value = strTemp2
    Wks.Range(Chr(intField) & intRow & ":" & Chr(intField + UBound(strTemp2)) & intRow).HorizontalAlignment = xlRight
    Wks.Range(Chr(intField + 2) & intRow & ":" & Chr(intField + UBound(strTemp2)) & intRow).NumberFormatLocal = "##,##0"
    Wks.Range(Chr(intField) & intRow & ":" & Chr(intField + UBound(strTemp2)) & intRow).Value = Wks.Range(Chr(intField) & intRow & ":" & Chr(intField + UBound(strTemp2)) & intRow).Value
    Wks.Range(Chr(intField + 1) & intRow & ":" & Chr(intField + UBound(strTemp2)) & intRow).Borders(xlEdgeBottom).LineStyle = xlContinuous
    intRow = intRow + 1
    '合計
    If pType = "2" Then
       strTemp2(0) = ""
       strTemp2(1) = "合　計："
       strTemp2(2) = Val(mFee1) '服務費
       strTemp2(3) = Val(mFee2) '規費
       strTemp2(4) = Val(mFee1) + Val(mFee2) '請款金額
       strTemp2(5) = Val(mdFee1) '扣繳金額
       strTemp2(6) = Val(mFee1) + Val(mFee2) - Val(mdFee1) '應付金額
       Wks.Range(Chr(intField) & intRow & ":" & Chr(intField + UBound(strTemp2)) & intRow).Value = strTemp2
       Wks.Range(Chr(intField) & intRow & ":" & Chr(intField + UBound(strTemp2)) & intRow).HorizontalAlignment = xlRight
       Wks.Range(Chr(intField + 2) & intRow & ":" & Chr(intField + UBound(strTemp2)) & intRow).NumberFormatLocal = "##,##0"
       Wks.Range(Chr(intField) & intRow & ":" & Chr(intField + UBound(strTemp2)) & intRow).Value = Wks.Range(Chr(intField) & intRow & ":" & Chr(intField + UBound(strTemp2)) & intRow).Value
       Wks.Range(Chr(intField + 1) & intRow & ":" & Chr(intField + UBound(strTemp2)) & intRow).Borders(xlEdgeBottom).LineStyle = xlDouble
       intRow = intRow + 1
    End If
    If InStr(stCompName, "智權") = 0 Then
       intRow = intRow + 1
       Wks.Range(Chr(intField) & intRow & ":" & Chr(intField + UBound(strTemp2)) & intRow).NumberFormatLocal = "@"
       Wks.Range(Chr(intField) & intRow & ":" & Chr(intField + UBound(strTemp2)) & intRow).HorizontalAlignment = xlLeft
       Wks.Range(Chr(intField + 1) & intRow).Value = "請直接由款項中扣除本所服務費10%所得稅(惟每次稅款如低於2000元，則依法請勿代扣繳)。"
       intRow = intRow + 1
    End If
    
    intI = 0
    If Len(rptMemo) > 0 Then
       arrM = Split(rptMemo, vbCrLf)
       intI = UBound(arrM)
    End If
    
    If intRow < 35 Then
       intRow = 35 '固定在頁尾
    Else
       If intRow + 5 + intI <= 48 Then
          intRow = intRow + 1
       Else
          'Added by Lydia 2024/11/05 一次只印一頁
          If iPage = 1 Then
             Call PrintExcelMainSet(Xls, Wks)
          End If
          Wks.PrintOut Copies:=1, Collate:=True
          Wks.Range(iX & ":" & iX, intRow & ":" & intRow).Delete
          'end 2024/11/05
          '換頁
          Call PrintExcelTitle(Wks, stCompName, stBillTitle)
       End If
    End If
    
    If stAcnt = "J" Then
        SetRptAccount "J", True
    ElseIf stAcnt = "L" Then
        SetRptAccount "L", True
    Else
        stAcnt = IIf(stAcnt = "1", "2", "2")
        strExc(0) = Combo2(Val(stAcnt) - 1).ItemData(Combo2(Val(stAcnt) - 1).ListIndex)
        SetRptAccount strExc(0), True  '抓指定匯款帳號
    End If
    
    Wks.Range(Chr(intField) & intRow).Value = "付款方式："
    Wks.Range(Chr(intField) & intRow + 1).Value = "1. 請將上述款項匯入本所以下帳戶"
    Wks.Range(Chr(intField) & intRow + 2).Value = "　帳戶：" & sra(1)
    Wks.Range(Chr(intField) & intRow + 3).Value = "　戶名：" & sra(2) & sra(0) & "　" & "帳號：" & sra(3)
    Wks.Range(Chr(intField) & intRow + 3).Cells.Font.Size = 14
    Wks.Range(Chr(intField) & intRow + 3).Cells.Font.Bold = True
    Wks.Range(Chr(intField) & intRow + 4).Value = "2. 如以支票支付，支票抬頭請開「" & sra(2) & "」"
    Wks.Range(Chr(intField) & intRow + 5).Value = "　並請以掛號逕寄 " & sra(4) & " 會計室收"
    For intI = intRow To 48
       Wks.Range(Chr(intField) & intI & ":" & Chr(intField + UBound(strTemp2)) & intI).NumberFormatLocal = "@"
       Wks.Range(Chr(intField) & intI & ":" & Chr(intField + UBound(strTemp2)) & intI).HorizontalAlignment = xlLeft
    Next intI
    intRow = intRow + 7
    
    If Len(rptMemo) > 0 Then
        Wks.Range(Chr(intField) & intRow).Value = "列印備註："
        For idx = 0 To UBound(arrM)
           Wks.Range(Chr(intField + 1) & intRow + idx).Value = arrM(idx)
        Next idx
    End If

    '列印頁面設定
    If iPage = 1 Then
       Call PrintExcelMainSet(Xls, Wks)
    End If
    Wks.PrintOut Copies:=1, Collate:=True
End Sub

Private Sub PrintExcelMainSet(ByRef pXLS As Excel.Application, ByRef pWks As Excel.Worksheet)
   pWks.PageSetup.PaperSize = 9 'A4
   pWks.PageSetup.PrintTitleRows = "$1:$7"
   pWks.PageSetup.Orientation = xlPortrait '直印
   pWks.PageSetup.LeftMargin = pXLS.CentimetersToPoints(0.8) '邊界
   pWks.PageSetup.RightMargin = pXLS.CentimetersToPoints(0.8)
   pWks.PageSetup.TopMargin = pXLS.CentimetersToPoints(1)
   pWks.PageSetup.BottomMargin = pXLS.CentimetersToPoints(1)
   pWks.PageSetup.CenterHorizontally = True '版面設定->邊界->水平置中
End Sub

'Added by Lydia 2024/09/16 使用Excel產生報表:長春人造X74310000,長春石油X74310010,大連化工X74310020
Private Function SaveXLS_2() As Boolean
Dim XlsRpt2 As New Excel.Application
Dim wksRpt2 As New Worksheet
Dim tmpArr As Variant
Dim strA As String, intA As Integer

'Modified by Lydia 2024/12/17 台灣案分「服務費」和「規費」顯示
'strA = "SELECT 格式,收據日期,收據號碼,A0807,CU11,SUM(AMT) TAMT,PA48,LISTAGG(CPM0304,'_') WITHIN GROUP (ORDER BY R012) AS CPM03,PA11,PA149N " & _
       "FROM (SELECT DECODE(AXC01,NULL,'收據','發票') AS 格式,R002 AS 收據日期,NVL(AXC01,R010)AS 收據號碼,A0807,CU11 " & _
       ",R006+R007-NVL(R009,0) AS AMT,PA48,DECODE(PA09,'000',NVL(CPM03,CPM04),NVL(CPM04,CPM03)) AS CPM0304,PA11,NVL(PCC04,NVL(PCC03,PCC05)) AS PA149N,R012 " & _
       "From RDATAFACTORY, ACC080, ACC431, CUSTOMER, CASEPROGRESS, PATENT, CASEPROPERTYMAP, POTCUSTCONT " & _
       "WHERE R017=A0801 AND R001 = 'Y' and FormName='" & Me.Name & "' And ID='" & strUserNum & "' and seqno = " & mSeqNo & " " & _
       "AND R010=AXC02(+) AND SUBSTR(R022,1,8)=CU01(+) AND SUBSTR(R022,9,1)=CU02(+) " & _
       "AND R012=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) " & _
       "AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(PA26,1,8)=PCC01(+) AND PA149=PCC02(+) " & _
       ") GROUP BY 格式,收據日期,收據號碼,A0807,CU11,PA48,PA11,PA149N "
'大陸案
strA = "SELECT DECODE(AXC01,NULL,'收據','發票') AS 格式,R002 AS 收據日期,NVL(AXC01,R010)AS 收據號碼,A0807,CU11 " & _
       ",R006+R007-NVL(R009,0) AS AMT,PA48,DECODE(PA09,'000',NVL(CPM03,CPM04),NVL(CPM04,CPM03)) AS CPM0304,PA11,NVL(PCC04,NVL(PCC03,PCC05)) AS PA149N,R012 " & _
       ",NA01,NA03,'服務費' AS CTITLE From RDATAFACTORY, ACC080, ACC431, CUSTOMER, CASEPROGRESS, PATENT, CASEPROPERTYMAP, POTCUSTCONT,NATION " & _
       "WHERE R017=A0801 AND R001 = 'Y' and FormName='" & Me.Name & "' And ID='" & strUserNum & "' and seqno = " & mSeqNo & " " & _
       "AND R010=AXC02(+) AND SUBSTR(R022,1,8)=CU01(+) AND SUBSTR(R022,9,1)=CU02(+) " & _
       "AND R012=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) " & _
       "AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(PA26,1,8)=PCC01(+) AND PA149=PCC02(+) AND PA09=NA01(+) AND PA09<>'000' "
'台灣案
strA = strA & "UNION SELECT DECODE(AXC01,NULL,'收據','發票') AS 格式,R002 AS 收據日期,NVL(AXC01,R010)AS 收據號碼,A0807,CU11 " & _
       ",R006-NVL(R009,0) AS AMT,PA48,DECODE(PA09,'000',NVL(CPM03,CPM04),NVL(CPM04,CPM03)) AS CPM0304,PA11,NVL(PCC04,NVL(PCC03,PCC05)) AS PA149N,R012 " & _
       ",NA01,NA03,'服務費' AS CTITLE From RDATAFACTORY, ACC080, ACC431, CUSTOMER, CASEPROGRESS, PATENT, CASEPROPERTYMAP, POTCUSTCONT,NATION " & _
       "WHERE R017=A0801 AND R001 = 'Y' and FormName='" & Me.Name & "' And ID='" & strUserNum & "' and seqno = " & mSeqNo & " " & _
       "AND R010=AXC02(+) AND SUBSTR(R022,1,8)=CU01(+) AND SUBSTR(R022,9,1)=CU02(+) " & _
       "AND R012=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) " & _
       "AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(PA26,1,8)=PCC01(+) AND PA149=PCC02(+) AND PA09=NA01(+) AND PA09='000' "
strA = strA & "UNION SELECT DECODE(AXC01,NULL,'收據','發票') AS 格式,R002 AS 收據日期,NVL(AXC01,R010)AS 收據號碼,A0807,CU11 " & _
       ",TO_NUMBER(R007) AS AMT,PA48,DECODE(PA09,'000',NVL(CPM03,CPM04),NVL(CPM04,CPM03)) AS CPM0304,PA11,NVL(PCC04,NVL(PCC03,PCC05)) AS PA149N,R012 " & _
       ",NA01,NA03,'規費' AS CTITLE From RDATAFACTORY, ACC080, ACC431, CUSTOMER, CASEPROGRESS, PATENT, CASEPROPERTYMAP, POTCUSTCONT,NATION " & _
       "WHERE R017=A0801 AND R001 = 'Y' and FormName='" & Me.Name & "' And ID='" & strUserNum & "' and seqno = " & mSeqNo & " " & _
       "AND R010=AXC02(+) AND SUBSTR(R022,1,8)=CU01(+) AND SUBSTR(R022,9,1)=CU02(+) " & _
       "AND R012=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) " & _
       "AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(PA26,1,8)=PCC01(+) AND PA149=PCC02(+) AND PA09=NA01(+) AND PA09='000' "
strA = "SELECT 格式,收據日期,收據號碼,A0807,CU11,SUM(AMT) TAMT,PA48,LISTAGG(CPM0304,'') WITHIN GROUP (ORDER BY R012) AS CPM03,PA11,PA149N,NA01,NA03,CTITLE " & _
       "FROM (" & strA & ") GROUP BY 格式,收據日期,收據號碼,A0807,CU11,PA48,PA11,PA149N,NA01,NA03,CTITLE HAVING SUM(AMT) <> 0 "
'end 2024/12/17
strA = strA & " Order by 2,3 "

If m_rs.State = 1 Then m_rs.Close
m_rs.CursorLocation = adUseClient
m_rs.Open strA, cnnConnection, adOpenStatic, adLockReadOnly

If Not m_rs.EOF And Not m_rs.BOF Then

    '-------預設Excel
    SaveXLS_2 = False
    '起始位置intField=65=>A
    intField = 65:  intRow = 1
    strFileN = txtCustNo(0).Text & "請款明細表" & ServerDate & ServerTime & MsgText(43)
    If Dir(strExcelPath & strFileN) = MsgText(601) Then
        If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
            MkDir strExcelPath
        End If
    Else
        Kill strExcelPath & strFileN
    End If
    XlsRpt2.SheetsInNewWorkbook = 1
    XlsRpt2.Workbooks.add
    Set wksRpt2 = XlsRpt2.Worksheets(1)
    wksRpt2.Activate
    bolOpenXls = True
    strAllF = "發票格式(收據/手寫三聯發票/收銀機發票/電子發票),收據日期,收據號碼,賣方統編,買方統編人造11384708石油11384806大連12233428,台幣分攤金額,長春案號,請款事由(服務費or代收代付+具體事由),專利申請號,長春承辦人"
    strWidth = "9,12,15,13,13,14,12,48,18,11"
    strField = Split(strAllF, ",")
    intWidth = Split(strWidth, ",")
    ReDim tmpArr(0 To UBound(strField))
    
    For intA = 0 To UBound(strField)
       wksRpt2.Range(Chr(intField + intA) & intRow).Value = strField(intA)
       wksRpt2.Range(Chr(intField + intA) & intRow).ColumnWidth = Val(intWidth(intA))
       wksRpt2.Range(Chr(intField + intA) & ":" & Chr(intField + intA)).HorizontalAlignment = xlCenter
       If intA <> 5 Then
          wksRpt2.Range(intRow & ":" & intRow).NumberFormat = "@"
       End If
       If intA = 5 Then '金額
          'Modified by Lydia 2024/10/07+m_rs.RecordCount
          wksRpt2.Range(Chr(intField + intA) & intRow + 1 & ":" & Chr(intField + intA) & intRow + 7 + m_rs.RecordCount).HorizontalAlignment = xlRight
          wksRpt2.Range(Chr(intField + intA) & intRow + 1 & ":" & Chr(intField + intA) & intRow + 7 + m_rs.RecordCount).NumberFormat = "##,##0"
       ElseIf intA = 7 Then '案件性質
          wksRpt2.Range(Chr(intField + intA) & intRow + 1 & ":" & Chr(intField + intA) & intRow + 7 + m_rs.RecordCount).HorizontalAlignment = xlLeft
       Else
          wksRpt2.Range(Chr(intField + intA) & intRow + 1 & ":" & Chr(intField + intA) & intRow + 7 + m_rs.RecordCount).HorizontalAlignment = xlCenter
          wksRpt2.Range(Chr(intField + intA) & intRow + 1 & ":" & Chr(intField + intA) & intRow + 7 + m_rs.RecordCount).NumberFormat = "@"
       End If
    Next intA
    wksRpt2.Range(intRow & ":" & intRow).RowHeight = 105
    wksRpt2.Range(Chr(intField) & intRow & ":" & Chr(intField + UBound(strField)) & intRow).WrapText = True
    wksRpt2.Range(Chr(intField) & intRow & ":" & Chr(intField + UBound(strField)) & intRow).Interior.ColorIndex = 15 '底色:灰色
    intRow = 2
    '-------預設Excel
    m_rs.MoveFirst
    Do While Not m_rs.EOF
       For intA = 0 To UBound(strField)
          If intA = 1 Then
             tmpArr(intA) = ChangeTStringToWDateString(Replace("" & m_rs.Fields("收據日期"), "/", ""))
          'Added by Lydia 2024/12/17
          ElseIf intA = 2 Then '收據號碼：台灣案規費的收據號碼=空白，由使用者輸入
             If "" & m_rs.Fields("na01") = "000" And "" & m_rs.Fields("ctitle") = "規費" Then
                tmpArr(intA) = ""
             Else
                tmpArr(intA) = "" & m_rs.Fields("收據號碼")
             End If
          ElseIf intA = 7 Then '請款事由=服務費_[國家]案_[案件性質]
             tmpArr(intA) = "" & m_rs.Fields("ctitle") & "_" & IIf("" & m_rs.Fields("na01") = "020", "大陸", "" & m_rs.Fields("na03")) & "案_" & m_rs.Fields("CPM03")
          'end 2024/12/17
          ElseIf intA = 5 Then
             tmpArr(intA) = Val("" & m_rs.Fields("TAMT"))
          Else
             tmpArr(intA) = "" & m_rs.Fields(intA)
          End If
       Next intA
       wksRpt2.Range(Chr(intField) & intRow & ":" & Chr(intField + UBound(strField)) & intRow).Value = tmpArr
       intRow = intRow + 1
       m_rs.MoveNext
    Loop
    
    '加框線
    intRow = intRow + 6
    wksRpt2.Range(Chr(intField) & "1:" & Chr(intField + UBound(strField)) & intRow).Borders.LineStyle = xlContinuous
    wksRpt2.Range(Chr(intField) & "1:" & Chr(intField + UBound(strField)) & intRow).Borders.Weight = xlThin
    intRow = intRow + 2
    wksRpt2.Range(Chr(intField) & intRow).Value = "製表日期：" & ChangeTStringToTDateString(strSrvDate(1))
    wksRpt2.Range(Chr(intField) & intRow).HorizontalAlignment = xlLeft
    wksRpt2.Range(Chr(intField) & "1:" & Chr(intField + UBound(strField)) & intRow).Font.Name = "Arial"
    wksRpt2.Range(Chr(intField) & "1:" & Chr(intField + UBound(strField)) & intRow).Font.Size = 11
    
    wksRpt2.Columns("I:I").EntireColumn.AutoFit 'Added by Lydia 2024/10/07 調整為能使文字全部顯示之欄寬
    
    '判斷若版本2007以上改變存格式
    If Val(XlsRpt2.Version) < 12 Then
        XlsRpt2.Workbooks(1).SaveAs FileName:=strExcelPath & strFileN, FileFormat:=-4143
    Else
        XlsRpt2.Workbooks(1).SaveAs FileName:=strExcelPath & strFileN, FileFormat:=56
    End If
    XlsRpt2.Workbooks.Close
    XlsRpt2.Quit
    SaveXLS_2 = True
    Set XlsRpt2 = Nothing
    Set wksRpt2 = Nothing
    Exit Function
Else
    MsgBox "無資料列印！"
End If

ErrHnd1:
    If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
    If bolOpenXls = True Then
        '判斷若版本2007以上改變存格式
        If Val(XlsRpt2.Version) < 12 Then
            XlsRpt2.Workbooks(1).SaveAs FileName:=strExcelPath & strFileN, FileFormat:=-4143
        Else
            XlsRpt2.Workbooks(1).SaveAs FileName:=strExcelPath & strFileN, FileFormat:=56
        End If
        XlsRpt2.Workbooks.Close
        XlsRpt2.Quit
    End If
    
End Function
