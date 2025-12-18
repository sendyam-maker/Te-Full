VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm210141 
   BorderStyle     =   1  '單線固定
   Caption         =   "繳款作業及收據PDF"
   ClientHeight    =   5832
   ClientLeft      =   48
   ClientTop       =   348
   ClientWidth     =   9384
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5832
   ScaleWidth      =   9384
   Begin VB.FileListBox File2 
      Height          =   180
      Left            =   0
      TabIndex        =   71
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtNote 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   585
      Left            =   1980
      Locked          =   -1  'True
      TabIndex        =   70
      Text            =   "產生收據中，暫時不要使用Word..."
      Top             =   2010
      Visible         =   0   'False
      Width           =   5610
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "收據明細PDF"
      Enabled         =   0   'False
      Height          =   345
      Index           =   6
      Left            =   7740
      TabIndex        =   24
      Top             =   810
      Width           =   1515
   End
   Begin VB.ComboBox cboComp 
      Height          =   300
      ItemData        =   "frm210141.frx":0000
      Left            =   4725
      List            =   "frm210141.frx":0002
      Style           =   2  '單純下拉式
      TabIndex        =   7
      Top             =   840
      Width           =   1125
   End
   Begin VB.CheckBox chkPrintedOnly 
      Caption         =   "排除未列印收據"
      Height          =   195
      Left            =   6075
      TabIndex        =   8
      Top             =   893
      Width           =   1590
   End
   Begin VB.TextBox txtDate 
      Height          =   300
      Index           =   1
      Left            =   2280
      MaxLength       =   7
      TabIndex        =   6
      Top             =   840
      Width           =   1050
   End
   Begin VB.TextBox txtDate 
      Height          =   300
      Index           =   0
      Left            =   1080
      MaxLength       =   7
      TabIndex        =   5
      Top             =   840
      Width           =   1050
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   8325
      Top             =   3390
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
      Left            =   5310
      TabIndex        =   10
      Top             =   30
      Width           =   1650
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "帶入簽收金額(&I)"
      Height          =   345
      Index           =   4
      Left            =   90
      TabIndex        =   11
      Top             =   4050
      Width           =   1560
   End
   Begin VB.TextBox txtTot 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H8000000F&
      Height          =   270
      Index           =   1
      Left            =   8235
      Locked          =   -1  'True
      TabIndex        =   43
      Top             =   1185
      Width           =   1050
   End
   Begin VB.TextBox txtTot 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H8000000F&
      Height          =   270
      Index           =   0
      Left            =   6660
      Locked          =   -1  'True
      TabIndex        =   42
      Top             =   1185
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
      TabIndex        =   41
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
      TabIndex        =   39
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
      TabIndex        =   37
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
      TabIndex        =   31
      Top             =   4080
      Width           =   1050
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "查詢(&Q)"
      Height          =   345
      Index           =   1
      Left            =   4185
      TabIndex        =   9
      Top             =   30
      Width           =   1065
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "繳款記錄刪除(&E)"
      Height          =   345
      Index           =   3
      Left            =   7740
      TabIndex        =   23
      Top             =   420
      Width           =   1515
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "繳款存檔(&S)"
      Enabled         =   0   'False
      Height          =   345
      Index           =   2
      Left            =   7020
      TabIndex        =   22
      Top             =   30
      Width           =   1155
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   345
      Index           =   0
      Left            =   8235
      TabIndex        =   25
      Top             =   30
      Width           =   1020
   End
   Begin VB.TextBox txtCustNo 
      Height          =   300
      Index           =   0
      Left            =   4905
      MaxLength       =   9
      TabIndex        =   3
      Text            =   "X"
      Top             =   506
      Width           =   1230
   End
   Begin VB.TextBox txtCustNo 
      Height          =   300
      Index           =   1
      Left            =   6345
      MaxLength       =   9
      TabIndex        =   4
      Text            =   "X"
      Top             =   506
      Width           =   1230
   End
   Begin VB.TextBox txtSales 
      Height          =   300
      Left            =   1080
      MaxLength       =   6
      TabIndex        =   0
      Top             =   173
      Width           =   915
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frm210141.frx":0004
      Height          =   2565
      Left            =   90
      TabIndex        =   26
      Top             =   1470
      Width           =   9195
      _ExtentX        =   16214
      _ExtentY        =   4530
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
         DataField       =   "公司別"
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
         DataField       =   "最早收文日"
         Caption         =   "最早收文日"
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
      BeginProperty Column06 
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
      BeginProperty Column07 
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
      BeginProperty Column08 
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
      BeginProperty Column09 
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
      BeginProperty Column10 
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
      BeginProperty Column11 
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
      BeginProperty Column12 
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
      BeginProperty Column13 
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
      BeginProperty Column14 
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
      BeginProperty Column15 
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
      BeginProperty Column16 
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
      BeginProperty Column17 
         DataField       =   "a0k11"
         Caption         =   "a0k11"
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
      BeginProperty Column19 
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
      BeginProperty Column20 
         DataField       =   "cp27t"
         Caption         =   "發文日期"
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
      BeginProperty Column22 
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
      SplitCount      =   2
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         Size            =   166
         BeginProperty Column00 
            Alignment       =   2
            ColumnAllowSizing=   -1  'True
            Locked          =   -1  'True
            ColumnWidth     =   432
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   468.283
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   1247.811
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   11.906
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   1068.094
         EndProperty
         BeginProperty Column06 
            ColumnAllowSizing=   -1  'True
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   840.189
         EndProperty
         BeginProperty Column07 
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   852.095
         EndProperty
         BeginProperty Column08 
            Alignment       =   1
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   659.906
         EndProperty
         BeginProperty Column09 
            Alignment       =   1
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   648
         EndProperty
         BeginProperty Column10 
            Alignment       =   2
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   432
         EndProperty
         BeginProperty Column11 
            Alignment       =   1
            ColumnAllowSizing=   -1  'True
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   659.906
         EndProperty
         BeginProperty Column12 
            ColumnAllowSizing=   -1  'True
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   2628.284
         EndProperty
         BeginProperty Column13 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
            ColumnWidth     =   828.284
         EndProperty
         BeginProperty Column14 
            Alignment       =   1
            Locked          =   -1  'True
            Object.Visible         =   0   'False
            ColumnWidth     =   792
         EndProperty
         BeginProperty Column15 
            Alignment       =   1
            Locked          =   -1  'True
            Object.Visible         =   0   'False
            ColumnWidth     =   959.811
         EndProperty
         BeginProperty Column16 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column17 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column18 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column19 
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   959.811
         EndProperty
         BeginProperty Column20 
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   887.811
         EndProperty
         BeginProperty Column21 
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   2808
         EndProperty
         BeginProperty Column22 
            Object.Visible         =   0   'False
            ColumnWidth     =   2232
         EndProperty
      EndProperty
      BeginProperty Split1 
         AllowRowSizing  =   0   'False
         RecordSelectors =   0   'False
         Size            =   456
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
            ColumnWidth     =   552.189
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
            ColumnWidth     =   1247.811
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   840.189
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   828.284
         EndProperty
         BeginProperty Column05 
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   1068.094
         EndProperty
         BeginProperty Column06 
            ColumnAllowSizing=   -1  'True
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   840.189
         EndProperty
         BeginProperty Column07 
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   480.189
         EndProperty
         BeginProperty Column08 
            Alignment       =   1
            Object.Visible         =   -1  'True
            ColumnWidth     =   659.906
         EndProperty
         BeginProperty Column09 
            Alignment       =   1
            Object.Visible         =   -1  'True
            ColumnWidth     =   648
         EndProperty
         BeginProperty Column10 
            Alignment       =   2
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   432
         EndProperty
         BeginProperty Column11 
            Alignment       =   1
            ColumnAllowSizing=   -1  'True
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   659.906
         EndProperty
         BeginProperty Column12 
            ColumnAllowSizing=   -1  'True
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   2628.284
         EndProperty
         BeginProperty Column13 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
            ColumnWidth     =   828.284
         EndProperty
         BeginProperty Column14 
            Alignment       =   1
            Locked          =   -1  'True
            Object.Visible         =   0   'False
            ColumnWidth     =   792
         EndProperty
         BeginProperty Column15 
            Alignment       =   1
            Locked          =   -1  'True
            Object.Visible         =   0   'False
            ColumnWidth     =   959.811
         EndProperty
         BeginProperty Column16 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column17 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column18 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column19 
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   959.811
         EndProperty
         BeginProperty Column20 
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   887.811
         EndProperty
         BeginProperty Column21 
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   2808
         EndProperty
         BeginProperty Column22 
            Object.Visible         =   0   'False
            ColumnWidth     =   2232
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   1305
      Left            =   45
      TabIndex        =   45
      Top             =   4410
      Width           =   9330
      Begin VB.TextBox Text1 
         Height          =   525
         Index           =   11
         Left            =   5130
         MaxLength       =   50
         MultiLine       =   -1  'True
         ScrollBars      =   2  '垂直捲軸
         TabIndex        =   64
         Top             =   720
         Width           =   2310
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   0
         Left            =   4680
         MaxLength       =   13
         TabIndex        =   62
         Top             =   390
         Width           =   870
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   10
         Left            =   8280
         MaxLength       =   13
         TabIndex        =   20
         Top             =   390
         Width           =   870
      End
      Begin VB.TextBox txtTot 
         Alignment       =   1  '靠右對齊
         Appearance      =   0  '平面
         BackColor       =   &H8000000F&
         Height          =   270
         Index           =   7
         Left            =   8190
         Locked          =   -1  'True
         TabIndex        =   57
         Top             =   990
         Width           =   1050
      End
      Begin VB.TextBox txtTot 
         Alignment       =   1  '靠右對齊
         Appearance      =   0  '平面
         BackColor       =   &H8000000F&
         Height          =   270
         Index           =   6
         Left            =   8190
         Locked          =   -1  'True
         TabIndex        =   55
         Top             =   690
         Width           =   1050
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   9
         Left            =   7380
         MaxLength       =   13
         TabIndex        =   19
         Top             =   390
         Width           =   870
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   1
         Left            =   180
         MaxLength       =   13
         TabIndex        =   12
         Top             =   390
         Width           =   870
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '靠右對齊
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   1080
         MaxLength       =   13
         TabIndex        =   13
         Top             =   390
         Width           =   870
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '靠右對齊
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   1980
         MaxLength       =   13
         TabIndex        =   14
         Top             =   390
         Width           =   870
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   4
         Left            =   2880
         MaxLength       =   13
         TabIndex        =   15
         Top             =   390
         Width           =   870
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   5
         Left            =   3780
         MaxLength       =   13
         TabIndex        =   16
         Top             =   390
         Width           =   870
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   6
         Left            =   5580
         MaxLength       =   13
         TabIndex        =   17
         Top             =   390
         Width           =   870
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   7
         Left            =   6480
         MaxLength       =   13
         TabIndex        =   18
         Top             =   390
         Width           =   870
      End
      Begin VB.TextBox Text1 
         Height          =   525
         Index           =   8
         Left            =   585
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  '垂直捲軸
         TabIndex        =   21
         Top             =   720
         Width           =   4110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "其他備註"
         Height          =   360
         Index           =   25
         Left            =   4770
         TabIndex        =   63
         Top             =   750
         Width           =   405
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "其　他"
         Height          =   180
         Index           =   24
         Left            =   4815
         TabIndex        =   61
         Top             =   180
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "外　幣"
         Height          =   180
         Index           =   11
         Left            =   8415
         TabIndex        =   59
         Top             =   180
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "差額："
         Height          =   180
         Index           =   22
         Left            =   7560
         TabIndex        =   58
         Top             =   1035
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "合計："
         Height          =   180
         Index           =   21
         Left            =   7560
         TabIndex        =   56
         Top             =   735
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "補扣繳"
         Height          =   180
         Index           =   20
         Left            =   7590
         TabIndex        =   54
         Top             =   180
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "票據金額"
         Height          =   180
         Index           =   12
         Left            =   255
         TabIndex        =   53
         Top             =   180
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "台北電匯"
         Height          =   180
         Index           =   13
         Left            =   1155
         TabIndex        =   52
         Top             =   180
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "分所電匯"
         Height          =   180
         Index           =   14
         Left            =   2055
         TabIndex        =   51
         Top             =   180
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "現　金"
         Height          =   180
         Index           =   15
         Left            =   3000
         TabIndex        =   50
         Top             =   180
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "抵暫收款"
         Height          =   180
         Index           =   16
         Left            =   3855
         TabIndex        =   49
         Top             =   180
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "溢 收 款"
         Height          =   180
         Index           =   17
         Left            =   5655
         TabIndex        =   48
         Top             =   180
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "手 續 費"
         Height          =   180
         Index           =   18
         Left            =   6555
         TabIndex        =   47
         Top             =   180
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "備註"
         Height          =   180
         Index           =   19
         Left            =   225
         TabIndex        =   46
         Top             =   750
         Width           =   360
      End
   End
   Begin VB.Frame FrameDept 
      Height          =   315
      Left            =   1600
      TabIndex        =   66
      Top             =   30
      Visible         =   0   'False
      Width           =   2865
      Begin VB.TextBox txtSalesArea1 
         Height          =   285
         Left            =   1860
         TabIndex        =   68
         Top             =   0
         Width           =   915
      End
      Begin VB.TextBox txtSalesArea 
         Height          =   285
         Left            =   855
         TabIndex        =   67
         Top             =   0
         Width           =   915
      End
      Begin VB.Line Line3 
         X1              =   1770
         X2              =   2040
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "業務區："
         Height          =   180
         Index           =   26
         Left            =   0
         TabIndex        =   69
         Top             =   45
         Width           =   720
      End
   End
   Begin MSForms.ComboBox Combo3 
      Height          =   336
      Left            =   1080
      TabIndex        =   1
      Top             =   180
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
   Begin MSForms.TextBox txtTitle 
      Height          =   300
      Left            =   1080
      TabIndex        =   2
      Top             =   506
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
      TabIndex        =   72
      Top             =   240
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
      Left            =   3960
      TabIndex        =   65
      Top             =   960
      Width           =   720
   End
   Begin VB.Line Line2 
      X1              =   2100
      X2              =   2280
      Y1              =   990
      Y2              =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收據日期："
      Height          =   180
      Index           =   23
      Left            =   135
      TabIndex        =   60
      Top             =   900
      Width           =   900
   End
   Begin VB.Label lblMemo 
      AutoSize        =   -1  'True
      Caption         =   "符號說明： ◎未列印 △智權公司 ＃已開立INVOICE"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   0
      Left            =   96
      TabIndex        =   44
      Top             =   1236
      Width           =   4080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "總計"
      Height          =   180
      Index           =   10
      Left            =   7830
      TabIndex        =   40
      Top             =   4125
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "扣繳"
      Height          =   180
      Index           =   9
      Left            =   6390
      TabIndex        =   38
      Top             =   4125
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "規費"
      Height          =   180
      Index           =   8
      Left            =   4815
      TabIndex        =   36
      Top             =   4125
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "服務費"
      Height          =   180
      Index           =   7
      Left            =   3060
      TabIndex        =   35
      Top             =   4125
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "規費"
      Height          =   180
      Index           =   5
      Left            =   7830
      TabIndex        =   34
      Top             =   1230
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "點數"
      Height          =   180
      Index           =   4
      Left            =   6255
      TabIndex        =   33
      Top             =   1230
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "所有資料合計"
      Height          =   180
      Index           =   3
      Left            =   5070
      TabIndex        =   32
      Top             =   1230
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "點選資料合計"
      Height          =   180
      Index           =   6
      Left            =   1890
      TabIndex        =   30
      Top             =   4125
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收據抬頭："
      Height          =   180
      Index           =   1
      Left            =   135
      TabIndex        =   29
      Top             =   555
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "客戶編號："
      Height          =   180
      Index           =   2
      Left            =   3960
      TabIndex        =   28
      Top             =   555
      Width           =   900
   End
   Begin VB.Line Line1 
      X1              =   6165
      X2              =   6345
      Y1              =   645
      Y2              =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   180
      Index           =   0
      Left            =   135
      TabIndex        =   27
      Top             =   225
      Width           =   900
   End
End
Attribute VB_Name = "frm210141"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/07/27 智權-調整財務系統(20200909) 'Memo by Lydia 2021/08/27 上線
'原「繳款輸入」標題修改為「繳款作業及收據PDF」。
'前三列為定格畫面。
'客戶代號均修改為”客戶編號”。
'收據編號均修改為”收據號碼”。
'單據日期及收據開立日期均修改為”收據日期”。
'國別均修改為”申請國家”。
'客戶名稱均修改為”申請人名稱”。
'end 2021/07/27
'Memo by Lydia 2021/07/16 DataGrid分割顯示的Split0必須要顯示全部要顯示的欄位，而Split1可以隱藏前面的欄位
'Memo by Lydia 2021/07/16 改成Form2.0 ; lblSalesName、txtTitle、DataGrid1改字型=新細明體-ExtB
'Memo by Lydia 2019/07/01 表單名稱:智權人員繳款資料輸入=>繳款輸入
'Created by Morgan 2013/11/28
Option Explicit

Dim m_A4421 As String, m_A4427 As String
Dim adoTmp As ADODB.Recordset
Dim stST05 As String
Dim m_blnColOrderAsc As Boolean
Dim m_mouseRow As Integer, m_MouseCol As Integer
Dim m_A4425 As String '溢收款處理方式
'Add by Amy 2014/05/21
Dim bolSpecMan As Boolean  '是否為特殊設定檔人員
Dim strSpecCode As String '特殊設定檔設定代號
Dim m_strListPer As String 'Add By Sindy 2020/7/28
Dim m_FileName2 As String 'Add By Sindy 2020/10/27
Dim m_FileNameL As String 'Add By Sindy 2020/11/20
Dim m_FileNameJ As String 'Add By Sindy 2020/11/23
Dim m_FileNameJdoc As String 'Add By Sindy 2022/1/17
Dim mSeqNo As String 'Added by Lydia 2021/07/27
'Add By Sindy 2023/6/12
Dim arrID ', stST15 As String
'Dim bolAreaMan As Boolean '下拉選單有區主管
'2023/6/12 END


'Modify By Sindy 2020/8/31
'Private Sub cmdok_Click(Index As Integer)
Public Sub cmdok_Click(Index As Integer)
'2020/8/31 END
   Dim bolChkTT As Boolean
   Dim strNoArr As String
   Dim i As Long
   Dim PrinterIndex As Integer
   
   Select Case Index
   Case 0 '結束
      Unload Me
   Case 1 '查詢
      Screen.MousePointer = vbHourglass
      doQuery
      Screen.MousePointer = vbDefault
   Case 2 '繳款存檔
      Screen.MousePointer = vbHourglass
      If TxtValidate = True Then
         If FormSave = False Then
            MsgBox "存檔失敗，請洽系統管理員 !", vbCritical
         Else
            MsgBox "存檔成功!", vbInformation
            'Modified by Morgan 2015/7/15
            'doQuery True
            Screen.MousePointer = vbDefault
            If m_A4421 <> "" Then
               bolChkTT = True
            End If
            txtCustNo(0) = "X": txtCustNo(1) = "X"
            txtTitle = ""
            txtDate(0) = "": txtDate(1) = ""
            chkPrintedOnly.Value = 0
            cboComp.ListIndex = 0
            Call PUB_SendMailCache 'Added by Lydia 2025/04/18
            FormReset
            If bolChkTT Then
               cmdOK(4).Value = True
            End If
            'end 2015/7/15
         End If
      End If
      Screen.MousePointer = vbDefault
   Case 3 '繳款記錄刪除
      ShowDelete
   Case 4 '帶入電匯金額
      ShowTT
   Case 5 '選擇應收客戶
      ShowCustList
   Case 6 '收據明細PDF Add By Sindy 2020/10/27
      'Added by Lydia 2021/07/16 +判斷有資料才繼續
      If mSeqNo = "" Then Exit Sub
      If Adodc1.Recordset.RecordCount = 0 Then Exit Sub
      'end 2021/07/16
      PrinterIndex = -1
      For i = 0 To Printers.Count - 1
       If UCase(Printers(i).DeviceName) = UCase$("PDFCreator") Then
        PrinterIndex = i
        Exit For
       End If
      Next i
      If PrinterIndex < 0 Then
         MsgBox "請通知電腦中心安裝PDFCreator !!!" & vbCrLf & vbCrLf & _
                "（因為J公司產生收據PDF檔，會使用到。）"
         'Exit Function
      End If
      
      Set adoTmp = Adodc1.Recordset.Clone
      strExc(1) = ""
      With adoTmp
         .MoveFirst
         Do While Not .EOF
            If .Fields(0) = "Y" Then
               'Added by Lydia 2023/11/13
               If "" & .Fields("a0k32") = "Z" Then
                  If InStr(strNoArr, .Fields("a0k01")) = 0 Then
                     MsgBox .Fields("a0k01") & " 本收據確定不印!"
                     strNoArr = strNoArr & "," & .Fields("a0k01")
                  End If
               'end 2023/11/13
               'Add By Sindy 2020/11/19
               'Modified by Lydia 2023/11/13 +else
               ElseIf Val(.Fields("a0k19")) = 0 Then
                  .Fields(0) = ""
                  If InStr(strNoArr, .Fields("a0k01")) = 0 Then
                     MsgBox .Fields("a0k01") & " 本收據尚未列印!"
                     strNoArr = strNoArr & "," & .Fields("a0k01")
                  End If
'               'ElseIf .Fields(16) = "J" Or .Fields(16) = "L" Then
'               ElseIf .Fields(16) = "J" Then
'                  .Fields(0) = ""
'                  If InStr(strNoArr, .Fields("a0k01")) = 0 Then
'                     MsgBox .Fields("a0k01") & " (" & .Fields(26) & ")收據範本製作中,尚無法產生PDF檔,抱歉!"
'                     strNoArr = strNoArr & "," & .Fields("a0k01")
'                  End If
               Else
               '2020/11/19 END
                  strExc(1) = "Y"
               End If
            End If
            .MoveNext
         Loop
      End With
      If strExc(1) = "Y" Then
         strExc(1) = InputBox("請輸入產出收據的方式：" & vbCrLf & vbCrLf & _
         "1. 一收據一檔案" & vbCrLf & _
         "2. 同一客戶之收據同一檔案，且一收據一檔案的資訊也出現" & vbCrLf & _
         "3. 全部點選的收據同一檔案，且一收據一檔案的資訊也出現" & vbCrLf & vbCrLf & _
         "請輸入 1 或 2 或 3", Me.Caption, 1)
         If strExc(1) = "" Then
            Exit Sub
         ElseIf strExc(1) <> "1" And strExc(1) <> "2" And strExc(1) <> "3" Then
            MsgBox "請輸入 1 或 2 或 3 !!", vbExclamation
         Else
            Call PrintReceiptPDF(Val(strExc(1)))
            Exit Sub
         End If
      Else
         MsgBox "請選取欲產生收據的資料列！", vbExclamation
         Exit Sub
      End If
   End Select
End Sub

Private Sub PrintReceiptPDF(intPDFKind As Integer)
Dim strNo As String
Dim TmpDirNm As String
Dim AdoRs As ADODB.Recordset
Dim strCUID As String, strSaveFileName As String
Dim strNoArr As String, strCUIDArr As String
Dim strMergeFile As String, strCmd As String
Dim process_id As Long
Dim process_handle As Long
Dim iTimes As Integer
Dim arrPer As Variant, idx As Integer
Dim strMergeFN As String
Dim dblFCnt As Double
Dim hLocalFile As Long
   
On Error GoTo ErrHand
   
   '暫時工作區
   TmpDirNm = App.path & "\TempRev"
'   If Dir(TmpDirNm, vbDirectory) <> "" Then
'      Kill TmpDirNm & "\*.*"
'      RmDir TmpDirNm '刪除一個現有的目錄或檔案夾。
'   End If
   Call PUB_KillTempFolder("TempRev", App.path)
   Sleep 6000
   If Dir(TmpDirNm, vbDirectory) = "" Then
      MkDir TmpDirNm
   End If
   
   Screen.MousePointer = vbHourglass
   txtNote.Visible = True
   
   Set adoTmp = Adodc1.Recordset.Clone
   With adoTmp
   '       18.a0k03                                    2.收據號碼
   'Modify By Sindy 2025/6/16
   '.Sort = DataGrid1.Columns(18).DataField & " asc," & DataGrid1.Columns(2).DataField & " asc"
   .Sort = "a0k03 asc,收據號碼 asc"
   '2025/6/16 END
   .MoveFirst
   strNoArr = "": strCUIDArr = ""
   Do While Not .EOF
      If .Fields("選取") = "Y" Then
         strNo = Left(.Fields("收據號碼"), 9)
         strCUID = .Fields("a0k03")
         If InStr(strNoArr, strNo) = 0 Then
            strNoArr = strNoArr & "," & strNo
         Else
            GoTo ReadNextRow
         End If
         If InStr(strCUIDArr, strCUID) = 0 Then
            strCUIDArr = strCUIDArr & "," & strCUID
         End If
         
         'Add By Sindy 2020/11/20 法律所收據/智權公司收據
         'Modify By Sindy 2025/6/16
         'If .Fields(16) = "L" Or .Fields(16) = "J" Then
         If .Fields("a0k11") = "L" Or .Fields("a0k11") = "J" Then
         '2025/6/16 END
            'Modify By Sindy 2022/1/17
            ',acc0j0,caseprogress
            'and a0j13=a0k01 and a0j01=cp09
            strExc(0) = "select * from acc0k0,customer,acc0j0,caseprogress" & _
                        " where substr(a0k03, 1, 8) = cu01(+)" & _
                        " and substr(a0k03, 9, 1) = cu02(+) and a0k01 = '" & strNo & "' and a0j13=a0k01 and a0j01=cp09" & _
                        " and ((to_number(substr(a0k01, 5, 5)) > 2000) or to_number(substr(a0k01, 5, 5)) <= 2000 and a0k02 >= 920101)" & _
                        " and (a0k09 is null or a0k09 = 0) and (a0k37 is null or a0k37<>'N')" & _
                        " and a0k01 not in (select a0m02 from acc0m0 where a0m02 = '" & strNo & "' and a0m03 is not null)"
            intI = 1
            Set AdoRs = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               'PDF的檔名 : 客戶編號+收據編號.pdf
               strSaveFileName = strCUID & "-" & strNo
               'Modify By Sindy 2025/6/16
               'If .Fields(16) = "L" Then
               If .Fields("a0k11") = "L" Then
               '2025/6/16 END
                  'Modify By Sindy 2021/8/20 + & "\" & strUserNum
                  PUB_PrintCaseReceipt_L App.path & "\" & strUserNum & "\" & m_FileNameL, AdoRs, 0, 0, , , , , False, TmpDirNm & "\" & strSaveFileName & ".pdf", True
               Else
                  'Modify By Sindy 2022/1/17 Form2.0無法用Printer物件
                  '改呼叫共用(與補印請款單統一)
                  PUB_PrintCaseReceipt_J_Doc App.path & "\" & strUserNum & "\" & m_FileNameJdoc, AdoRs, 0, 0, , , , , , False, True, TmpDirNm & "\" & strSaveFileName & ".pdf"
                  '2022/1/17 END
                  
'                  'Add By Sindy 2020/11/23
'                  '檢查是否有安裝PDFCreator
'                  Load frmPDF
'                  frmPDF.StartProcess TmpDirNm, strSaveFileName
'                  '列印J公司請款單
'                  Printer.FontSize = 12
'                  Printer.Font = "標楷體"
'                  PUB_PrintCaseReceipt_J AdoRs, 0, 0, , , , , , False, True
'                  Printer.EndDoc
'                  Printer.Font = "新細明體"
'                  'END
'                  frmPDF.EndtProcess
'                  Unload frmPDF
               End If
            End If
            AdoRs.Close
         '智慧所收據
         Else
            'PDF的檔名 : 客戶編號+收據編號.pdf
            strSaveFileName = strCUID & "-" & strNo
            strExc(0) = "select * from acc0k0,customer" & _
                        " where substr(a0k03, 1, 8) = cu01(+) and substr(a0k03, 9, 1) = cu02(+)" & _
                        " and a0k11<>'J' and a0k01 = '" & strNo & "'" & _
                        " and ((to_number(substr(a0k01, 5, 5)) > 2000) or to_number(substr(a0k01, 5, 5)) <= 2000 and a0k02 >= 920101)" & _
                        " and (a0k09 is null or a0k09 = 0) and (a0k37 is null or a0k37<>'N')" & _
                        " and a0k01 not in (select a0m02 from acc0m0 where a0m02 = '" & strNo & "' and a0m03 is not null)"
            intI = 1
            Set AdoRs = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               'Modify By Sindy 2022/1/26 改共用函數
               PUB_PrintCaseReceipt_Doc App.path & "\" & strUserNum & "\" & m_FileName2, AdoRs, 0, 0, , , , , , False, True, TmpDirNm & "\" & strSaveFileName & ".pdf"
            End If
            AdoRs.Close
         End If
      End If
ReadNextRow:
      .MoveNext
   Loop
   End With

   Set AdoRs = Nothing
   Screen.MousePointer = vbDefault
   txtNote.Visible = False
   
   If strNoArr <> "" Then
      strNoArr = Mid(strNoArr, 2)
      strCUIDArr = Mid(strCUIDArr, 2)
      
      '切換至來源目錄
      ChDir TmpDirNm
      If intPDFKind <> "1" Then
         arrPer = Split(strCUIDArr, ",")
         For idx = 0 To UBound(arrPer)
            strMergeFN = "" '組欲合併的檔案
            '執行合併動作
            If intPDFKind = "2" Then '2.同一客戶之收據同一檔案
'               strMergeFile = TmpDirNm & "\" & arrPer(idx) & ".pdf"
'               strCmd = "pdftk.exe " & TmpDirNm & "\" & arrPer(idx) & "*.pdf" & " cat output " & strMergeFile
               strMergeFile = arrPer(idx) & ".pdf"
               File2.path = TmpDirNm '& "\" & arrPer(idx) & "*.pdf"
               File2.Refresh
               For dblFCnt = 0 To File2.ListCount - 1
                  If UCase(Right(File2.List(dblFCnt), 4)) = ".PDF" And UCase(Left(File2.List(dblFCnt), Len(arrPer(idx)))) = UCase(arrPer(idx)) Then
                     strMergeFN = strMergeFN & IIf(strMergeFN <> "", " ", "") & ".\" & File2.List(dblFCnt)
                  End If
               Next dblFCnt
               strCmd = pub_PdftkEXE & " " & strMergeFN & " cat output .\" & strMergeFile
               
            ElseIf intPDFKind = "3" Then '3.全部點選的收據同一檔案
'               strMergeFile = TmpDirNm & "\ReceiptAll.pdf"
'               strCmd = "pdftk.exe " & TmpDirNm & "\*.pdf" & " cat output " & strMergeFile
               strMergeFile = "ReceiptAll.pdf"
               File2.path = TmpDirNm '& "\*.pdf"
               File2.Refresh
               For dblFCnt = 0 To File2.ListCount - 1
                  If UCase(Right(File2.List(dblFCnt), 4)) = ".PDF" Then
                     strMergeFN = strMergeFN & IIf(strMergeFN <> "", " ", "") & ".\" & File2.List(dblFCnt)
                  End If
               Next dblFCnt
               strCmd = pub_PdftkEXE & " " & strMergeFN & " cat output .\" & strMergeFile
            End If
            DoEvents
            process_id = SHELL(strCmd, vbHide)
            process_handle = OpenProcess(PROCESS_TERMINATE, 0, process_id)
            If process_handle <> 0 Then
               For iTimes = 1 To 10
                  If PUB_CheckIsRunning("pdftk.exe") = True Then
                     Sleep 2000
                  Else
                     Exit For
                  End If
               Next
               If iTimes > 10 Then
                  TerminateProcess process_handle, 0&
                  CloseHandle process_handle
               End If
            End If
            '檢查是否有產生合併PDF檔
            DoEvents
            If Dir(strMergeFile) = "" Then
               MsgBox "合併失敗！位置於:" & TmpDirNm, vbExclamation
               Exit Sub
            End If
            If intPDFKind = "3" Then Exit For
         Next idx
      End If
      
      ChDir App.path '目錄切回
      MsgBox "收據已產生完畢。位置於:" & TmpDirNm, vbExclamation
      '直接開啟視窗
      'SHELL "Explorer.exe " & TmpDirNm, vbNormalFocus
      'Lydia:用檔案總管開啟放置1~2分鐘後,檔案總管會出錯(ex. A2037, A4041)
      ShellExecute hLocalFile, "explore", TmpDirNm, vbNullString, vbNullString, 1
   End If
   
   Exit Sub
   
ErrHand:
   ChDir App.path '目錄切回
   Screen.MousePointer = vbDefault
   txtNote.Visible = False
   If Err.Number <> 0 Then MsgBox Err.Number & vbCrLf & Err.Description
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
   'Modify By Sindy 2020/5/8 + 增加抓取法律所案源介紹人為此智權人員
   'Modified by Morgan 2025/3/17 調整語法
   'stVTB1 = stVTB1 & " union " & _
            "select a0j13 X1,a0j01 X2" & _
      ",nvl(sum(nvl(a1u04,0)+nvl(a1u07,0)-nvl(a1u08,0)),0) X3" & _
      ",nvl(sum(nvl(a1u05,0)+nvl(a1u09,0)-nvl(a1u10,0)),0) X4" & _
      ",nvl(sum(a1u06),0) X5 From acc0k0, acc0j0, acc1u0, lawofficesource, caseprogress" & _
      " where instr(LOS04,'" & txtSales & "')>0" & strCon & _
      " and nvl(a0k09,0)=0 and (nvl(a0k06,0)+nvl(a0k07,0)) > (nvl(a0k17,0)+nvl(a0k18,0))" & _
      " and a0j13(+)=a0k01 and a1u02(+)=a0j13 and a1u03(+)=a0j01 and a0j02<>'TT999999000'" & _
      " and LOS15(+)=cp162 and cp162 is not null and cp09(+)=a0j01" & _
      " group by a0j13,a0j01"
   stVTB1 = stVTB1 & " union " & _
            "select a0j13 X1,a0j01 X2" & _
      ",nvl(sum(nvl(a1u04,0)+nvl(a1u07,0)-nvl(a1u08,0)),0) X3" & _
      ",nvl(sum(nvl(a1u05,0)+nvl(a1u09,0)-nvl(a1u10,0)),0) X4" & _
      ",nvl(sum(a1u06),0) X5 From lawofficesource, caseprogress, acc0j0, acc0k0, acc1u0 " & _
      " where instr(LOS04,'" & txtSales & "')>0" & strCon & _
      " and cp162(+)=LOS15 and cp162 is not null and a0j01(+)=cp09 and a0j02<>'TT999999000'" & _
      " and a0k01(+)=a0j13 and nvl(a0k09,0)=0 and (nvl(a0k06,0)+nvl(a0k07,0)) > (nvl(a0k17,0)+nvl(a0k18,0))" & _
      " and a1u02(+)=a0j13 and a1u03(+)=a0j01" & _
      " group by a0j13,a0j01"
   'end 2025/3/17
   '2020/5/8 END
   
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
   'Modify By Sindy 2020/5/8 + 增加抓取法律所案源介紹人為此智權人員
   stVTB2 = stVTB2 & " union " & _
            "select a0j13 Y1,a0j01 Y2" & _
      ",sum(AXD06) Y3,sum(AXD07)Y4,sum(AXD08) Y5" & _
      " From acc0k0, acc0j0, acc441,acc440, lawofficesource, caseprogress" & _
      " where instr(LOS04,'" & txtSales & "')>0" & strCon & _
      " and nvl(a0k09,0)=0 and (nvl(a0k06,0)+nvl(a0k07,0)) > (nvl(a0k17,0)+nvl(a0k18,0))" & _
      " and a0j13(+)=a0k01 and axd04(+)=a0j13 and axd05(+)=a0j01" & _
      " and a4401(+)=axd01 and a4402(+)=axd02 and a4403(+)=axd03 and a4416 is null and a0j02<>'TT999999000'" & _
      " and LOS15(+)=cp162 and cp162 is not null and cp09(+)=a0j01" & _
      " group by a0j13,a0j01"
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
      
   '2020/5/8 END
   'Modified by Lydia 2024/04/01 改用acc0j0為基準
   'Modified by Morgan 2024/5/22 還原(因改後變慢且查詢結果應該還是相同)
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
   'end 2024/5/22
   'Modified by Lydia 2021/07/16
   'intI = 1
   intI = 0
   Screen.MousePointer = vbHourglass
   Set adoTmp = ClsLawReadRstMsg(intI, strExc(0))
   Screen.MousePointer = vbDefault
   
   If adoTmp.RecordCount > 0 Then 'Added by Lydia 2021/07/16 加判斷
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
   
   'Added by Morgan 2014/1/9
   If Val(Text1(6)) > 0 Then
      If m_A4425 = "" Then SetA4425
   Else
      m_A4425 = ""
   End If
      
   '手續費
   Text1_Validate 7, bCancel
   If bCancel = True Then
      Text1(7).SetFocus
      Text1_GotFocus 7
      Exit Function
   End If
   
   '至少勾選一筆
   If Val(Format(txtTot(5))) = 0 Then
      MsgBox "請先選取資料！", vbExclamation
      Exit Function
   ElseIf Val(Format(txtTot(7))) <> 0 Then
      MsgBox "點選資料總計與收款金額不符！", vbCritical
      Exit Function
   End If
   
   'Added by Morgan 2015/7/15
   '有"其他"時也要輸"其他備註"
   If Val(Format(Text1(0))) > 0 And Text1(11) = "" Then
      MsgBox "請輸入其他被備註!!" & vbCrLf & vbCrLf & vbCrLf & "(例:客戶付款預扣尾牙或交際禮金等...)", vbExclamation
      If Text1(11).Enabled = True Then Text1(11).SetFocus
      Exit Function
   End If
   'end 2015/7/15
   
   'Added by Morgan 2023/10/25
   '溢收款轉暫收款時備註欄要有值
   If Val(Format(Text1(6))) > 0 And m_A4425 = "1" And Text1(8) = "" Then
      MsgBox "請於備註欄說明暫收原因,僅能是沖抵客戶之後案件之費用！", vbExclamation
      If Text1(8).Enabled = True Then Text1(8).SetFocus
      Exit Function
   End If
   'end 2023/10/25
   
   'Added by Morgan 2023/7/24
   '只開放財務處可部分收款
   'Removed by Morgan 2023/8/1 先取消--婉莘
   'Modified by Morgan 2023/10/19 改可部分收款但必須先繳規費--財務處
   'Modified by Morgan 2023/10/20 財務處人員除外--婉莘
   If strSrvDate(1) >= "20231023" And Pub_StrUserSt03 <> "M31" Then
      Set adoTmp = Adodc1.Recordset.Clone
      With adoTmp
      .MoveFirst
      Do While Not .EOF
         If .Fields("選取") = "Y" Then
            If .Fields("服務費") > 0 Then
               If CheckFee(.Fields("收據號碼")) = False Then
                  MsgBox "部分繳款必須先繳規費！【" & .Fields("收據號碼") & "】", vbCritical
                  Exit Function
               End If
            End If
         End If
         .MoveNext
      Loop
      End With
   End If
   'end 2023/7/24
   
   
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
      If MsgBox("同一INVOICE編號有未勾選的收據，是否繼續作業？", vbYesNo + vbInformation + vbDefaultButton2) = vbNo Then
         Exit Function
      End If
   End If
   'end 2023/11/13
   
   TxtValidate = True
End Function

'Added by Morgan 2023/10/19
'檢查收據號是否有未繳規費
Private Function CheckFee(pRcpNo As String) As Boolean
   Dim rsTmp As ADODB.Recordset
   Set rsTmp = Adodc1.Recordset.Clone
   
   CheckFee = True
   With rsTmp
   .MoveFirst
   Do While Not .EOF
      If .Fields("收據號碼") = pRcpNo Then
         If .Fields("規費") <> .Fields("amt2") Or .Fields("選取") <> "Y" Then
            CheckFee = False
            Exit Do
         End If
      End If
      .MoveNext
   Loop
   End With
   Set rsTmp = Nothing
End Function


Private Function FormSave() As Boolean
   Dim stDate As String, stTime As String, stNums As String
   Dim A2309 As String
   Dim strQuery As String, intQ As Integer, rsQD As New ADODB.Recordset 'Added by Lydia 2025/04/18
   
   cnnConnection.BeginTrans
On Error GoTo ErrHnd
   
   stDate = strSrvDate(1)
   stTime = ServerTime
   'Modified by Morgan 2015/7/14 +A4430,A4431
   strSql = "insert into ACC440(A4401,A4402,A4403,A4405,A4406,A4407,A4408,A4409,A4410,A4411,A4412,A4417,A4421,A4422,A4425,A4426,A4427,A4430,A4431)" & _
      " values('" & txtSales & "'," & stDate & "," & stTime & "," & Val(Text1(1)) & _
      "," & Val(Text1(2)) & "," & Val(Text1(3)) & "," & Val(Text1(4)) & "," & Val(Text1(5)) & "," & Val(Text1(6)) & _
      "," & Val(Text1(7)) & ",'" & ChgSQL(Text1(8)) & "','" & strUserNum & "','" & m_A4421 & "'," & Val(Text1(9)) & _
      ",'" & m_A4425 & "'," & Val(Text1(10)) & ",'" & m_A4427 & "'," & Val(Text1(0)) & ",'" & ChgSQL(Text1(11)) & "')"
   cnnConnection.Execute strSql, intI
   
   Set adoTmp = Adodc1.Recordset.Clone
   With adoTmp
   Do While Not .EOF
      If .Fields("選取") = "Y" Then
         strSql = "insert into ACC441(AXD01,AXD02,AXD03,AXD04,AXD05,AXD06,AXD07,AXD08)" & _
            "values('" & txtSales & "'," & stDate & "," & stTime & ",'" & .Fields("a0k01") & "'" & _
            ",'" & .Fields("a0j01") & "'," & Val(.Fields("服務費")) & "," & Val(.Fields("規費")) & _
            "," & Val(.Fields("扣繳金額")) & ")"
         cnnConnection.Execute strSql, intI
         
         '收據號碼
         If A2309 = "" Then
            A2309 = .Fields("a0k01")
         ElseIf InStr(A2309, .Fields("a0k01")) = 0 Then
            A2309 = A2309 & "," & .Fields("a0k01")
         End If
         'Added by Lydia 2025/04/18 TIPS分配比例管制：智權人員繳款第一階段款項，需通知顧服組主管
         If InStr("" & .Fields("本所案號"), "ACS-") > 0 And "" & .Fields("a0k01") <> "" Then
            strExc(0) = Pub_RplStr("" & .Fields("本所案號"))
            If InStr(strExc(0), "-") > 0 Then
               If InStrRev(strExc(0), "-") < 6 Then strExc(0) = strExc(0) & "-0-00"
            End If
            strExc(0) = Replace(strExc(0), "-", "")
            Call ChgCaseNo(strExc(0), strExc)
            strExc(5) = Pub_GetSpecMan("TIPS分配比例不適用案件")
            If strExc(1) = "ACS" And Len(strExc(2)) = 6 And InStr(strExc(5) & ";", strExc(1) & strExc(2) & strExc(3) & strExc(4)) = 0 Then
                strQuery = "select cp01,cp02,cp03,cp04,cp09,cp158,cp10,nvl(cpm03,cpm04) as cpm0304,cp156,cp144,atr06,atr07,atr08 " & _
                           "From caseprogress, casepropertymap, acs_tips_rate " & _
                           "where cp01='" & strExc(1) & "' and cp02='" & strExc(2) & "' and cp03='" & strExc(3) & "' and cp04='" & strExc(4) & "' and cp60='" & .Fields("a0k01") & "' and nvl(cp156,0)=1 " & _
                           "and cp01=cpm01(+) and cp10=cpm02(+) and cp01=atr01(+) and cp02=atr02(+) and cp03=atr03(+) and cp04=atr04(+) and cp156=atr05(+) and cp09=atr06(+) "
                intQ = 1
                Set rsQD = ClsLawReadRstMsg(intQ, strQuery)
                If intQ = 1 Then
                   '第一次繳款：先產生空白的智權業務獎金比例
                   If "" & rsQD.Fields("atr06") = "" Then
                       strSql = "insert into acs_tips_rate (atr01,atr02,atr03,atr04,atr05,atr06,atr07,atr13,atr14) " & _
                                "values ('" & rsQD.Fields("cp01") & "','" & rsQD.Fields("cp02") & "','" & rsQD.Fields("cp03") & "','" & rsQD.Fields("cp04") & "','1','" & rsQD.Fields("cp09") & "','" & txtSales & "','" & strUserNum & "',sysdate)"
                       cnnConnection.Execute strSql
                   End If
                   strExc(5) = "請顧服組主管於輸入智權比例後通知財務處!" & vbCrLf & _
                               "收據號碼：" & .Fields("a0k01") & vbCrLf & _
                               "第一階段款項：繳款日期" & ChangeTStringToTDateString(strSrvDate(2)) & "，金額" & Format(Val(.Fields("服務費")) + Val(.Fields("規費")), "###,###") & "元，" & vbCrLf & _
                               vbCrLf & "依據智權人員過往TIPS案件至今協作情況評比，項目(權重)如下：" & vbCrLf & _
                               "報價/簽約/請款(25)、向客戶說明(30)、專案過程會議參與(20)、專案執行過程溝通協調(20)、IP案件開發(5)，各權重加總大於70智權比例為「40%」，小於為「35%」" & vbCrLf & _
                               vbCrLf & vbCrLf & "主管請至【法務系統->ACS->資料處理->TIPS案請款階段分配比例維護作業】進行設定。"
                   strExc(6) = Pub_GetSpecMan("ACS郵件通知主管")
                   If strExc(6) <> "" Then
                      strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09) " & _
                                 "values('" & strUserNum & "','" & strExc(6) & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss') " & _
                                 ",'" & strExc(1) & "-" & strExc(2) & IIf(strExc(3) & strExc(4) = "000", "", "-" & strExc(3) & "-" & strExc(4)) & "第一階段款項已繳款，顧服組主管請於系統輸入智權業務獎金分配比例','" & ChgSQL(strExc(5)) & "',null) "
                      cnnConnection.Execute strSql
                   End If
                'Added by Lydia 2025/08/20 ACS案針對其他案件性質增加請款發文定稿通知=>僅發mail提醒顧服組主管，提供智權比例給財務處(以mail方式)
                Else
                   strQuery = "select c1.cp14 From caseprogress c1,caseprogress c2 " & _
                              "where c1.cp01='" & strExc(1) & "' and c1.cp02='" & strExc(2) & "' and c1.cp03='" & strExc(3) & "' and c1.cp04='" & strExc(4) & "' " & _
                              "and c1.cp60='" & .Fields("a0k01") & "' and c1.cp43=c2.cp09(+) and (c1.cp10 in (" & ACSforLetter & ") or c2.cp10 in (" & ACSforLetter & ")) "
                   strQuery = strQuery & " order by c1.cp05,c1.cp09 "
                   intQ = 1
                   Set rsQD = ClsLawReadRstMsg(intQ, strQuery)
                   If intQ = 1 Then
                     strExc(7) = rsQD.GetString(adClipString, , , ";")
                     strExc(5) = "請顧服組主管與承辦人確認智權比例後，以E-MAIL通知財務處!" & vbCrLf & _
                                 "收據號碼：" & .Fields("a0k01") & vbCrLf & _
                                 "繳款日期" & ChangeTStringToTDateString(strSrvDate(2)) & "，金額" & Format(Val(.Fields("服務費")) + Val(.Fields("規費")), "###,###") & "元，" & vbCrLf & _
                                 vbCrLf & "依據智權人員過往案件至今協作情況評比，項目(權重)如下：" & vbCrLf & _
                                 "報價/簽約/請款(25)、向客戶說明(30)、專案過程會議參與(20)、專案執行過程溝通協調(20)、IP案件開發(5)，各權重加總大於70智權比例為「40%」，小於為「35%」" & vbCrLf
                     strExc(6) = Pub_GetSpecMan("ACS郵件通知主管")
                     strExc(9) = Pub_GetSpecMan("財務處出納人員") 'CC:財務處的負責人員
                     If strExc(6) <> "" Then
                        strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09) " & _
                                  "values('" & strUserNum & "','" & strExc(6) & ";" & Mid(strExc(7), 1, Len(strExc(7)) - 1) & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss') " & _
                                  ",'" & strExc(1) & "-" & strExc(2) & IIf(strExc(3) & strExc(4) = "000", "", "-" & strExc(3) & "-" & strExc(4)) & "已繳款，請確認智權業務獎金分配比例後E-MAIL通知財務處','" & ChgSQL(strExc(5)) & "','" & ChgSQL(strExc(9)) & "') "
                        cnnConnection.Execute strSql
                     End If
                   End If
                'end 2025/08/20
                End If
                Set rsQD = Nothing
            End If
         End If
         'end 2025/04/18
      End If
      .MoveNext
   Loop
   End With
   
   '更新簽收紀錄
   If m_A4421 <> "" Then
      strSql = "update ACC230 set A2308=" & strSrvDate(2) & ",A2309='" & A2309 & "',A2321=sysdate where A2301='" & m_A4421 & "'"
      cnnConnection.Execute strSql, intI
      If m_A4427 <> "" Then
         strSql = "update ACC230 set A2308=" & strSrvDate(2) & ",A2309='" & A2309 & "',A2321=sysdate where A2301 in ('" & Replace(m_A4427, ";", "','") & "')"
         cnnConnection.Execute strSql, intI
      End If
   End If
   
   cnnConnection.CommitTrans
   FormSave = True
   Exit Function
   
ErrHnd:
   cnnConnection.RollbackTrans
   MsgBox Err.Description, vbCritical
   Set rsQD = Nothing 'Added by Lydia 2025/04/18
   
End Function

Private Sub doQuery(Optional pNoMsg As Boolean = False, Optional pNoReset As Boolean = False)
   Dim strCon As String
   Dim strConA0k As String, strConLoS As String 'Add By Sindy 2020/9/1
   Dim stVTB1 As String, stVTB2 As String
   Dim bolCancel As Boolean 'Add By Sindy 2023/6/12
   
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
   
'注意:業務區是用在外部呼叫此作業時, 如:frm210107
'   'Add By Sindy 2020/9/1
'   If txtSalesArea.Text <> "" Then
'      If txtSalesArea1.Text = "" Then txtSalesArea1.Text = txtSalesArea.Text
'      strConA0k = " and st15 >='" & txtSalesArea & "' and st15 <='" & txtSalesArea1 & "'"
'      strConLoS = " and st15 >='" & txtSalesArea & "' and st15 <='" & txtSalesArea1 & "'"
'   Else
'   '2020/9/1 END
      If txtSales = "" Then
         MsgBox "智權人員不可空白！", vbExclamation
         Exit Sub
      End If
      strConA0k = " and a0k20='" & txtSales & "'" 'Add By Sindy 2020/9/1
      strConLoS = " and instr(LOS04,'" & txtSales & "')>0" 'Add By Sindy 2020/9/1
'   End If
   
   If txtTitle <> "" Then
      strCon = strCon & " and a0k04 like '" & ChgSQL(txtTitle) & "%'"
   End If
   'Modified by Morgan 2023/4/12
   '改抓客戶編號的所有抬頭--瑞婷
'   If txtCustNo(0) <> "" And txtCustNo(0) <> "X" Then
'      strCon = strCon & " and a0k03>='" & txtCustNo(0) & "'"
'   End If
'   If txtCustNo(1) <> "" And txtCustNo(1) <> "X" Then
'      strCon = strCon & " and a0k03<='" & txtCustNo(1) & "'"
'   End If
   If (txtCustNo(0) <> "" And txtCustNo(0) <> "X") Or (txtCustNo(1) <> "" And txtCustNo(1) <> "X") Then
      If Not (txtCustNo(0) <> "" And txtCustNo(0) <> "X" And txtCustNo(1) <> "" And txtCustNo(1) <> "X") Then
         MsgBox "若以客戶編號查詢必須起迄都要輸入！", vbExclamation
         If txtCustNo(0) = "" Or txtCustNo(0) <> "X" Then
            txtCustNo(0).SetFocus
         Else
            txtCustNo(1).SetFocus
         End If
         Exit Sub
      Else
         'Modified by Morgan 2023/4/27 排除已收款
         strCon = strCon & " and a0k04 in (select distinct a0k04 from acc0k0 where a0k03>='" & txtCustNo(0) & "' and a0k03<='" & txtCustNo(1) & "' and nvl(a0k09,0)=0 and (nvl(a0k06,0)+nvl(a0k07,0)) > (nvl(a0k17,0)+nvl(a0k18,0)))"
      End If
   End If
   'end 2023/4/12
   
   'Added by Morgan 2015/7/14
   If txtDate(0) <> "" Then
      strCon = strCon & " and a0k02>=" & Val(txtDate(0))
   End If
   If txtDate(1) <> "" Then
      strCon = strCon & " and a0k02<=" & Val(txtDate(1))
   End If
   If chkPrintedOnly.Value = 1 Then
      strCon = strCon & " and a0k19>0"
   End If
   'end 2015/7/14
   
   'Added by Morgan 2015/7/16
   If cboComp.ListIndex > 0 Then
      If cboComp.ListIndex = 1 Then
         strCon = strCon & " and a0k11='1'"
      ElseIf cboComp.ListIndex = 2 Then
         strCon = strCon & " and a0k11='2'"
      ElseIf cboComp.ListIndex = 3 Then
         strCon = strCon & " and a0k11='J'"
      'Add by Amy 2020/04/09
      ElseIf cboComp.ListIndex = 4 Then
         strCon = strCon & " and a0k11='L'"
      End If
   End If
   'end 2015/7/16
   
   '已收款
   'Modify By Sindy 2020/5/20 排除顯示案號TT-999999 : and a0j02<>'TT999999000'
   'Modify By Sindy 2020/9/1 + staff, a0k20=st01(+)
   '                         a0k20='" & txtSales & "'" => strConA0k
   stVTB1 = "select a0j13 X1,a0j01 X2" & _
      ",nvl(sum(nvl(a1u04,0)+nvl(a1u07,0)-nvl(a1u08,0)),0) X3" & _
      ",nvl(sum(nvl(a1u05,0)+nvl(a1u09,0)-nvl(a1u10,0)),0) X4" & _
      ",nvl(sum(a1u06),0) X5,nvl(sum(a1u07),0) X6,nvl(sum(a1u09),0) X7" & _
      " From acc0k0, acc0j0, acc1u0, staff where a0k20=st01(+)" & strConA0k & strCon & _
      " and nvl(a0k09,0)=0 and (nvl(a0k06,0)+nvl(a0k07,0)) > (nvl(a0k17,0)+nvl(a0k18,0))" & _
      " and a0j13(+)=a0k01 and a1u02(+)=a0j13 and a1u03(+)=a0j01 and a0j02<>'TT999999000'" & _
      " group by a0j13,a0j01"
   'Modify By Sindy 2020/5/8 + 增加抓取法律所案源介紹人為此智權人員
   'Modify By Sindy 2020/9/1 + staff, substr(LOS04,1,5)=st01(+)
   '                         instr(LOS04,'" & txtSales & "')>0 => strConLos
   stVTB1 = stVTB1 & " union " & _
            "select a0j13 X1,a0j01 X2" & _
      ",nvl(sum(nvl(a1u04,0)+nvl(a1u07,0)-nvl(a1u08,0)),0) X3" & _
      ",nvl(sum(nvl(a1u05,0)+nvl(a1u09,0)-nvl(a1u10,0)),0) X4" & _
      ",nvl(sum(a1u06),0) X5,nvl(sum(a1u07),0) X6,nvl(sum(a1u09),0) X7" & _
      " From acc0k0, acc0j0, acc1u0, lawofficesource, caseprogress, staff" & _
      " where substr(LOS04,1,5)=st01(+)" & strConLoS & strCon & _
      " and nvl(a0k09,0)=0 and (nvl(a0k06,0)+nvl(a0k07,0)) > (nvl(a0k17,0)+nvl(a0k18,0))" & _
      " and a0j13(+)=a0k01 and a1u02(+)=a0j13 and a1u03(+)=a0j01 and a0j02<>'TT999999000'" & _
      " and LOS15(+)=cp162 and cp162 is not null and cp09(+)=a0j01" & _
      " group by a0j13,a0j01"
   '2020/5/8 END
   
   '已繳款未收款
   'Modify By Sindy 2020/5/20 排除顯示案號TT-999999 : and a0j02<>'TT999999000'
   'Modify By Sindy 2020/9/1 + staff, a0k20=st01(+)
   '                         a0k20='" & txtSales & "'" => strConA0k
   stVTB2 = "select a0j13 Y1,a0j01 Y2" & _
      ",sum(AXD06) Y3,sum(AXD07)Y4,sum(AXD08) Y5" & _
      " From acc0k0, acc0j0, acc441,acc440, staff" & _
      " where a0k20=st01(+)" & strConA0k & strCon & _
      " and nvl(a0k09,0)=0 and (nvl(a0k06,0)+nvl(a0k07,0)) > (nvl(a0k17,0)+nvl(a0k18,0))" & _
      " and a0j13(+)=a0k01 and axd04(+)=a0j13 and axd05(+)=a0j01" & _
      " and a4401(+)=axd01 and a4402(+)=axd02 and a4403(+)=axd03 and a4416 is null and a0j02<>'TT999999000'" & _
      " group by a0j13,a0j01"
   'Modify By Sindy 2020/5/8 + 增加抓取法律所案源介紹人為此智權人員
   'Modify By Sindy 2020/9/1 + staff, substr(LOS04,1,5)=st01(+)
   '                         instr(LOS04,'" & txtSales & "')>0 => strConLos
   stVTB2 = stVTB2 & " union " & _
            "select a0j13 Y1,a0j01 Y2" & _
      ",sum(AXD06) Y3,sum(AXD07)Y4,sum(AXD08) Y5" & _
      " From acc0k0, acc0j0, acc441,acc440, lawofficesource, caseprogress, staff" & _
      " where substr(LOS04,1,5)=st01(+)" & strConLoS & strCon & _
      " and nvl(a0k09,0)=0 and (nvl(a0k06,0)+nvl(a0k07,0)) > (nvl(a0k17,0)+nvl(a0k18,0))" & _
      " and a0j13(+)=a0k01 and axd04(+)=a0j13 and axd05(+)=a0j01" & _
      " and a4401(+)=axd01 and a4402(+)=axd02 and a4403(+)=axd03 and a4416 is null and a0j02<>'TT999999000'" & _
      " and LOS15(+)=cp162 and cp162 is not null and cp09(+)=a0j01" & _
      " group by a0j13,a0j01"
   '2020/5/8 END
   
   'Modified by Lydia 2019/11/06 +增加發文日cp27t於收據抬頭之前
   'Modify by Amy 2020/04/09 公司別抓acc080簡稱 原:decode(a0k11,'1','商標','2','專利','J','智權',a0k11)
   'Modified by Morgan 2023/1/11 已開發票 '＊' --> '(已開發票)'
   'Modified by Morgan 2023/5/3 +最早收文日, order by 2,10,12 -> order by 2,11,13
   'Modified by Lydia 2023/11/13 +A0K40開立INVOICE --> ||decode(a0k40, null,'','＃'),||decode(a0k32,'Z','',decode(nvl(a0k19,0),0,'◎'))
   'strExc(0) = "select '' 選取,sqldatet(a0k02) 收據日期" & _
      ",sqldatet(b.cp05) 最早收文日" & _
      ",a.cp01||'-'||a.cp02||decode(a.cp03||a.cp04,'000','','-'||a.cp03||'-'||a.cp04) 本所案號" & _
      ",decode(a0j04,'000',cpm03,cpm04) 案件性質" & _
      ",na03 申請國家,a0j09-X3-nvl(Y3,0) 服務費,a0j10-X4-nvl(Y4,0) 規費,'' 扣繳,0 扣繳金額" & _
      ",a0k01||decode(nvl(a0k19,0),0,'◎')||decode(AXC01,null,'','(已開發票)')||decode(a0k11,'J','△') 收據號碼" & _
      ",nvl(tm05,nvl(pa05,nvl(lc05,nvl(sp05,hc06)))) 案件名稱" & _
      ",a0j01,a0j09-X3-nvl(Y3,0) amt1,a0j10-X4-nvl(Y4,0) amt2" & _
      ",X5+NVL(Y5,0) amt3" & _
      ",a0j07,a0k11,a0k19,a0k01,a0j09-X6 SFee,a0j10-X7 OFee,a0k03,sqldatet(a.cp27) cp27t,a0k04,CU04,a0k05,a0820 公司別" & _
      " from (" & stVTB1 & ") X,(" & stVTB2 & ") Y,acc0j0,acc0k0,CUSTOMER,acc431,caseprogress a,Acc080" & _
      ",casepropertymap,nation,trademark,patent,lawcase,servicepractice,hirecase,acc1m0 m,caseprogress b" & _
      " where Y1(+)=X1 and Y2(+)=X2 and (a0j09-X3-nvl(Y3,0)>0 or a0j10-X4-nvl(Y4,0)>0) and a0j13(+)=X1 and a0j01(+)=X2 and a0k01(+)=a0j13 and axc02(+)=a0j13" & _
      " And a0k11=a0801(+) AND CU01(+)=SUBSTR(A0K03,1,8) AND CU02(+)=SUBSTR(A0K03,9) and a.cp09(+)=a0j01 and a.cp79>0 and cpm01(+)=a.cp01 and cpm02(+)=a.cp10 and na01(+)=a0j04" & _
      " and tm01(+)=a.cp01 and tm02(+)=a.cp02 and tm03(+)=a.cp03 and tm04(+)=a.cp04" & _
      " and pa01(+)=a.cp01 and pa02(+)=a.cp02 and pa03(+)=a.cp03 and pa04(+)=a.cp04" & _
      " and lc01(+)=a.cp01 and lc02(+)=a.cp02 and lc03(+)=a.cp03 and lc04(+)=a.cp04" & _
      " and sp01(+)=a.cp01 and sp02(+)=a.cp02 and sp03(+)=a.cp03 and sp04(+)=a.cp04" & _
      " and hc01(+)=a.cp01 and hc02(+)=a.cp02 and hc03(+)=a.cp03 and hc04(+)=a.cp04" & _
      " and a1m01(+)=a0k01 and b.cp09(+)=a1m02 and not exists(select * from acc1m0 x,caseprogress y where a1m01=a0k01 and cp09(+)=a1m02 and cp05||cp09<b.cp05||b.cp09)" & _
      " order by 2,11,13"
   'Modified by Lydia 2024/04/01 改用acc0j0為基準
   'Modified by Morgan 2024/5/22 還原(因改後變慢且查詢結果應該還是相同)
   strExc(0) = "select '' 選取,sqldatet(a0k02) 收據日期" & _
      ",sqldatet(b.cp05) 最早收文日" & _
      ",a.cp01||'-'||a.cp02||decode(a.cp03||a.cp04,'000','','-'||a.cp03||'-'||a.cp04) 本所案號" & _
      ",decode(a0j04,'000',cpm03,cpm04) 案件性質" & _
      ",na03 申請國家,a0j09-X3-nvl(Y3,0) 服務費,a0j10-X4-nvl(Y4,0) 規費,'' 扣繳,0 扣繳金額" & _
      ",a0k01||decode(a0k32,'Z','',decode(nvl(a0k19,0),0,'◎'))||decode(AXC01,null,'','(已開發票)')||decode(a0k11,'J','△')||decode(a0k40, null,'','＃') 收據號碼" & _
      ",nvl(tm05,nvl(pa05,nvl(lc05,nvl(sp05,hc06)))) 案件名稱" & _
      ",a0j01,a0j09-X3-nvl(Y3,0) amt1,a0j10-X4-nvl(Y4,0) amt2" & _
      ",X5+NVL(Y5,0) amt3,a0j07,a0k11,a0k19,a0k01,a0j09-X6 SFee,a0j10-X7 OFee,a0k03" & _
      ",sqldatet(a.cp27) cp27t,a0k04,CU04,a0k05,a0820 公司別, a0k32, a0k40 " & _
      " from (" & stVTB1 & ") X,(" & stVTB2 & ") Y,acc0j0,acc0k0,CUSTOMER,acc431,caseprogress a,Acc080" & _
      ",casepropertymap,nation,trademark,patent,lawcase,servicepractice,hirecase,acc1m0 m,caseprogress b" & _
      " where Y1(+)=X1 and Y2(+)=X2 and (a0j09-X3-nvl(Y3,0)>0 or a0j10-X4-nvl(Y4,0)>0) and a0j13(+)=X1 and a0j01(+)=X2 and a0k01(+)=a0j13 and axc02(+)=a0j13" & _
      " And a0k11=a0801(+) AND CU01(+)=SUBSTR(A0K03,1,8) AND CU02(+)=SUBSTR(A0K03,9) and a.cp09(+)=a0j01 and a.cp79>0 and cpm01(+)=a.cp01 and cpm02(+)=a.cp10 and na01(+)=a0j04" & _
      " and tm01(+)=a.cp01 and tm02(+)=a.cp02 and tm03(+)=a.cp03 and tm04(+)=a.cp04" & _
      " and pa01(+)=a.cp01 and pa02(+)=a.cp02 and pa03(+)=a.cp03 and pa04(+)=a.cp04" & _
      " and lc01(+)=a.cp01 and lc02(+)=a.cp02 and lc03(+)=a.cp03 and lc04(+)=a.cp04" & _
      " and sp01(+)=a.cp01 and sp02(+)=a.cp02 and sp03(+)=a.cp03 and sp04(+)=a.cp04" & _
      " and hc01(+)=a.cp01 and hc02(+)=a.cp02 and hc03(+)=a.cp03 and hc04(+)=a.cp04" & _
      " and a1m01(+)=a0k01 and b.cp09(+)=a1m02 and not exists(select * from acc1m0 x,caseprogress y where a1m01=a0k01 and cp09(+)=a1m02 and cp05||cp09<b.cp05||b.cp09)" & _
      " order by 2,11,13"
   'strExc(0) = "select '' 選取,sqldatet(a0k02) 收據日期" & _
      ",sqldatet(b.cp05) 最早收文日" & _
      ",a.cp01||'-'||a.cp02||decode(a.cp03||a.cp04,'000','','-'||a.cp03||'-'||a.cp04) 本所案號" & _
      ",decode(a0j04,'000',cpm03,cpm04) 案件性質" & _
      ",na03 申請國家,a0j09-X3-nvl(Y3,0) 服務費,a0j10-X4-nvl(Y4,0) 規費,'' 扣繳,0 扣繳金額" & _
      ",a0k01||decode(a0k32,'Z','',decode(nvl(a0k19,0),0,'◎'))||decode(AXC01,null,'','(已開發票)')||decode(a0k11,'J','△')||decode(a0k40, null,'','＃') 收據號碼" & _
      ",nvl(tm05,nvl(pa05,nvl(lc05,nvl(sp05,hc06)))) 案件名稱" & _
      ",a0j01,a0j09-X3-nvl(Y3,0) amt1,a0j10-X4-nvl(Y4,0) amt2" & _
      ",X5+NVL(Y5,0) amt3,a0j07,a0k11,a0k19,a0k01,a0j09-X6 SFee,a0j10-X7 OFee,a0k03" & _
      ",sqldatet(a.cp27) cp27t,a0k04,CU04,a0k05,a0820 公司別, a0k32, a0k40 " & _
      " from (" & stVTB1 & ") X,(" & stVTB2 & ") Y,acc0j0,acc0k0,CUSTOMER,acc431,caseprogress a,Acc080" & _
      ",casepropertymap,nation,trademark,patent,lawcase,servicepractice,hirecase,acc1m0 m,caseprogress b" & _
      " where a0j01=X2(+) and a0j13=x1(+) and a0j01=y2(+) and a0j13=y1(+) and (a0j09-X3-nvl(Y3,0)>0 or a0j10-X4-nvl(Y4,0)>0) and a0k01(+)=a0j13 and axc02(+)=a0j13" & _
      " And a0k11=a0801(+) AND CU01(+)=SUBSTR(A0K03,1,8) AND CU02(+)=SUBSTR(A0K03,9) and a.cp09(+)=a0j01 and a.cp79>0 and cpm01(+)=a.cp01 and cpm02(+)=a.cp10 and na01(+)=a0j04" & _
      " and tm01(+)=a.cp01 and tm02(+)=a.cp02 and tm03(+)=a.cp03 and tm04(+)=a.cp04" & _
      " and pa01(+)=a.cp01 and pa02(+)=a.cp02 and pa03(+)=a.cp03 and pa04(+)=a.cp04" & _
      " and lc01(+)=a.cp01 and lc02(+)=a.cp02 and lc03(+)=a.cp03 and lc04(+)=a.cp04" & _
      " and sp01(+)=a.cp01 and sp02(+)=a.cp02 and sp03(+)=a.cp03 and sp04(+)=a.cp04" & _
      " and hc01(+)=a.cp01 and hc02(+)=a.cp02 and hc03(+)=a.cp03 and hc04(+)=a.cp04" & _
      " and a1m01(+)=a0k01 and b.cp09(+)=a1m02 and not exists(select * from acc1m0 x,caseprogress y where a1m01=a0k01 and cp09(+)=a1m02 and cp05||cp09<b.cp05||b.cp09)" & _
      " order by 2,11,13"
   'end 2024/5/22
   intI = 1
   Set adoTmp = ClsLawReadRstMsg(intI, strExc(0))
   
   If pNoReset = False Then FormReset
   
   DataGrid1.Enabled = True
   'Modify by Amy 2014/06/16 +FormName 改暫存TB
   'Modified by Lydia 2021/07/27 改在模組中取得序號 + mSeqNo
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
      txtTot(1) = Format(txtTot(1), "#,##0")
      
      'If strCon <> "" Then
         Frame1.Enabled = True
         'cmdOK(4).Enabled = ShowTT(True)
         cmdOK(2).Enabled = True
         cmdOK(6).Enabled = True 'Add By Sindy 2020/10/28
      'End If
   Else
      cmdOK(2).Enabled = False
      cmdOK(6).Enabled = False 'Add By Sindy 2020/10/28
      If pNoMsg = False Then
         If Me.Visible = True Then MsgBox "無符合資料！", vbExclamation
      End If
   End If
End Sub

Private Sub DataGrid1_AfterColUpdate(ByVal ColIndex As Integer)
   Dim bCancel As Boolean
   Select Case ColIndex
   'Modified by Morgan 2015/7/14 調欄位順序
   'Modified by Morgan 2023/5/4 插入最大收文日調整欄位順序
   'Case 7, 8
   Case 8, 9
      'Added by Morgan 2015/7/6
      If Val(Format(DataGrid1.Columns(ColIndex))) < 0 Then
         MsgBox "輸入金額不可小於0！", vbInformation
         If ColIndex = 7 Then
            DataGrid1.Columns(ColIndex).Value = Val(Adodc1.Recordset.Fields("amt1"))
         ElseIf ColIndex = 8 Then
            DataGrid1.Columns(ColIndex).Value = Val(Adodc1.Recordset.Fields("amt2"))
         End If
      Else
      'end 2015/7/6
         DataGrid1.Columns(ColIndex).Value = Val(Format(DataGrid1.Columns(ColIndex)))
         'Modified by Morgan 2023/5/4 插入最大收文日調整欄位順序
         'If ColIndex = 7 Then
         If ColIndex = 8 Then
            If Val(DataGrid1.Columns(ColIndex).Value) > Val(Adodc1.Recordset.Fields("amt1")) Then
               MsgBox "收款服務費不可大於未收服務費！", vbInformation
               'Modify by Amy 2014/06/16 改暫存TB同一列選取三次修改三次會產生「找不到要更新的資料列」
               'bCancel = True
               DataGrid1.Columns(ColIndex).Value = Val(Adodc1.Recordset.Fields("amt1"))
            End If
         'Modified by Morgan 2023/5/4 插入最大收文日調整欄位順序
         'ElseIf ColIndex = 8 Then
         ElseIf ColIndex = 9 Then
            If Val(DataGrid1.Columns(ColIndex).Value) > Val(Adodc1.Recordset.Fields("amt2")) Then
               MsgBox "收款規費不可大於未收規費！", vbInformation
               'Modify by Amy 2014/06/16 改暫存TB同一列選取三次修改三次會產生「找不到要更新的資料列」
               'bCancel = True
               DataGrid1.Columns(ColIndex).Value = Val(Adodc1.Recordset.Fields("amt2"))
            End If
         End If
      End If 'Added by Morgan 2015/7/6
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
   'Modified by Morgan 2015/7/14 調欄位
   'Modified by Morgan 2023/5/4 插入最大收文日調整欄位順序
   'Case 7, 8
   Case 8, 9
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
   'Modified by Morgan 2023/5/4 插入最大收文日調整欄位順序
   'If pRow >= 0 And (pCol = 0 Or pCol = 9) Then
   If pRow >= 0 And (pCol = 0 Or pCol = 10) Then
      
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
            'Modified by Morgan 2017/4/5 同收據扣繳要同步
            '.Fields("扣繳") = ""
            '.Fields("扣繳金額") = 0
            If .Fields("扣繳") = "Y" Then
               UpdateTax .Fields("a0k01"), False
            End If
            'end 2017/4/5
            .Fields("服務費").Value = Val(.Fields("amt1"))
            .Fields("規費").Value = Val(.Fields("amt2"))
            .Fields("選取") = ""
         Else
            'Modified by Lydia 2023/11/13 判斷Z=確定不印
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
      'Modified by Morgan 2023/5/4 插入最大收文日調整欄位順序
      'Case 9 '扣繳
      Case 10 '扣繳
         If .Fields("選取") = "Y" Then
            If .Fields("扣繳") = "Y" Then
               'Modified by Morgan 2017/4/5 同收據扣繳要同步
               '.Fields("扣繳") = ""
               '.Fields("扣繳金額") = 0
               'bUpdate = True
               UpdateTax .Fields("a0k01"), False
               SettxtTot
               'end 2017/4/5
            Else
               If .Fields("a0k11") = "J" Then
                  MsgBox "智權公司不可勾選扣繳!", vbCritical
               ElseIf .Fields("a0k05") = "1" Then
                  MsgBox "個人不可勾選扣繳!", vbCritical
               ElseIf .Fields("amt3") > 0 Then
                  MsgBox "本收據已扣繳不可再勾選扣繳!", vbCritical
               Else
                  
                  'Modified by Morgan 2017/4/5 同收據扣繳要同步
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
                  'end 2017/4/5
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

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
   Dim stValue As String
      
   If Adodc1.RecordSource = Empty Then Exit Sub 'Added by Lydia 2021/08/19 若尚未查詢，防止出錯
   
   If Adodc1.Recordset.RecordCount > 0 Then
      Select Case ColIndex
      'Modified by Morgan 2015/7/14 調欄位順序
      'Modified by Morgan 2023/5/4 插入最大收文日調整欄位順序
      'Case 0, 9
      Case 0, 10
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
                        'Modified by Morgan 2017/4/5
                        '.Fields("扣繳") = stValue
                        ''是否合併
                        'If .Fields("a0j07") = "Y" Then
                        '   .Fields("扣繳金額") = 0.1 * (Val(.Fields("SFee")) + Val(.Fields("OFee")))
                        'Else
                        '   .Fields("扣繳金額") = 0.1 * Val(.Fields("SFee"))
                        'End If
                        UpdateTax .Fields("a0k01"), True
                        'end 2017/4/5
                     Else
                        .Fields("扣繳金額") = 0
                        .Fields("扣繳") = stValue
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
   'Modified by Morgan 2015/7/14 調欄位順序
   'Modified by Morgan 2023/5/4 插入最大收文日調整欄位順序
   'If DataGrid1.col <> 7 And DataGrid1.col <> 8 Then
   If DataGrid1.col <> 8 And DataGrid1.col <> 9 Then
      KeyCode = 0
   End If
End Sub

Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
   'Modified by Morgan 2015/7/14 調欄位順序
   'Modified by Morgan 2023/5/4 插入最大收文日調整欄位順序
   Select Case DataGrid1.col
   'Case 7
   Case 8
      If KeyCode = vbKeyReturn Then SendKeys "{RIGHT}"

   'Case 8
   Case 9
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

Private Sub Form_Load()
   MoveFormToCenter Me
   'stST15 = PUB_GetStaffST15(strUserNum, 1)
   'bolAreaMan = False 'Add By Sindy 2023/6/12
   stST05 = PUB_GetST05(strUserNum)
   txtSales = strUserNum
   
   'Modify By Sindy 2020/7/28 設定員編,部門,所別權限
   Call PUB_SetFormSaleDept(strUserNum, , , , txtSales, bolSpecMan, strSpecCode)
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
   
'   If strUserNum = "74018" Then
'      cmdok(6).Visible = True
'   Else
'      cmdok(6).Visible = False
'   End If
   
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
'      'Modify by Amy 2015/02/04 拿掉美珍 改寫至特殊人員(總經理業務工作代理人員)
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
'   'Add by Amy 2015/02/04 +總經理業務工作代理人員
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
'        'Add by Amy 2015/02/04 +總經理業務工作代理人員
'        If InStr(strSpecCode, "總經理業務工作代理人員") > 0 Then txtSales.Enabled = True
'        If InStr(strSpecCode, "A8") > 0 Then txtSales.Enabled = True
'   End If
'   'end 2014/05/21
   
   'Added by Morgan 2015/7/16
   cboComp.Clear
   cboComp.AddItem "", 0
   cboComp.AddItem "商標", 1
   cboComp.AddItem "智慧所", 2 'Modify by Amy 2020/04/09 原:專利
   cboComp.AddItem "智權", 3
   cboComp.AddItem "法律所", 4 'Add by Amy 2020/04/09
   'end 2015/7/16
   
   'Modify By Sindy 2021/8/20 + & "\" & strUserNum
   'Add By Sindy 2020/10/27
   m_FileName2 = "$$智慧所收據非正式.doc"
   If Dir(App.path & "\" & strUserNum & "\" & m_FileName2) <> "" Then
      Kill App.path & "\" & strUserNum & "\" & m_FileName2
   End If
   Call PUB_GetSampleFile(m_FileName2, "M31-000011-0-00", , App.path & "\" & strUserNum & "\")
   '2020/10/27 END
   'Add By Sindy 2020/11/20
   m_FileNameL = "$$法律所收據.doc"
   If Dir(App.path & "\" & strUserNum & "\" & m_FileNameL) <> "" Then
      Kill App.path & "\" & strUserNum & "\" & m_FileNameL
   End If
   Call PUB_GetSampleFile(m_FileNameL, "M31-000007-0-00", , App.path & "\" & strUserNum & "\")
   '2020/11/20 END
   'Add By Sindy 2020/11/23
   m_FileNameJ = "$$M51000089000.jpg" '非正式收據
   If Dir(App.path & "\" & strUserNum & "\" & m_FileNameJ) <> "" Then
      Kill App.path & "\" & strUserNum & "\" & m_FileNameJ
   End If
   Call PUB_GetSampleFile(m_FileNameJ, "M51-000089-0-00", , App.path & "\" & strUserNum & "\")
   '2020/11/23 END
   '2021/8/20 END
   'Add By Sindy 2022/1/17
   m_FileNameJdoc = "$$J公司請款單.doc"
   If Dir(App.path & "\" & strUserNum & "\" & m_FileNameJdoc) <> "" Then
      Kill App.path & "\" & strUserNum & "\" & m_FileNameJdoc
   End If
   Call PUB_GetSampleFile(m_FileNameJdoc, "M31-000012-0-00", , App.path & "\" & strUserNum & "\")
   '2022/1/17 END
   
   'Added by Morgan 2023/9/25 現金及支票不開放輸入--楊瑞婷/杜協理同意
   Text1(1).Enabled = False
   Text1(4).Enabled = False
   'end 2023/9/25
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set adoTmp = Nothing
   Set frm210141 = Nothing
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   If Index = 8 Or Index = 11 Then
      OpenIme
   Else
      CloseIme
   End If
   TextInverse Text1(Index)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   If Index <> 8 And Index <> 11 Then
      If KeyAscii = vbKeyReturn Then
         CheckSum
      ElseIf KeyAscii <> 8 And IsNumeric(Chr(KeyAscii)) = False Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Function SetA4425() As Boolean
'Modified by Morgan 2023/12/6
'   Do
'      'Modified by Morgan 2023/10/25
'      'strExc(1) = InputBox("請輸入溢收款處理方式：" & vbCrLf & vbCrLf & "1. 列暫收" & vbCrLf & "2. 退客戶 ( 須特別寄出者，請於備註欄註明 )" & vbCrLf & vbCrLf & "請輸入 1 或 2")
'      strExc(1) = InputBox("請輸入溢收款處理方式：" & vbCrLf & vbCrLf & "1. 列暫收 ( 請於備註欄說明暫收原因, 僅能是沖抵客戶　　　　　之後案件之費用 )" & vbCrLf & "2. 退客戶 ( 請二個月內提供客戶帳號資料交財務退費 )" & vbCrLf & vbCrLf & "請輸入 1 或 2")
'      'Added by Morgan 2014/3/31 金額可能輸錯
'      If strExc(1) = "" Then
'         Exit Do
'      'end 2014/3/31
'      ElseIf strExc(1) <> "1" And strExc(1) <> "2" Then
'         MsgBox "請輸入 1 或 2 !!", vbExclamation
'      Else
'         m_A4425 = strExc(1)
'         SetA4425 = True
'         Exit Do
'      End If
'   Loop
   frm880023.p_iChoice = 1
   frm880023.Show vbModal
   m_A4425 = frm880023.p_sReturn
   If m_A4425 <> "" Then
      SetA4425 = True
   End If
   Set frm880023 = Nothing
'end 2023/12/6
End Function

Private Sub Text1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   If Index <> 8 And Index <> 6 And Index <> 7 And Index <> 9 And Text1(Index).Enabled = True Then
      CheckSum '欄位值可能已修改,先算一次
      If Text1(Index) = "" And Val(Format(txtTot(7))) < 0 Then
         Text1(Index) = Abs(Format(txtTot(7)))
         TextInverse Text1(Index)
         CheckSum
      End If
   End If
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
   If Index = 6 Then
      If Val(Text1(6)) > 0 Then
         If SetA4425 = False Then
            Text1_GotFocus 6
            Cancel = True
         End If
      End If
   ElseIf Index = 7 Then
      If Val(Text1(7)) > 100 Then
         MsgBox "手續費不可超過 100 元！", vbCritical
         Text1_GotFocus 7
         Cancel = True
      End If
   End If
   CheckSum Index
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
      txtCustNo(Index) = Left(txtCustNo(0), 6) & "ZZZ"
      txtCustNo(Index).SelStart = 6
      txtCustNo(Index).SelLength = 3
   End If
End Sub

Private Sub txtCustNo_Validate(Index As Integer, Cancel As Boolean)
   If Index = 0 And Len(txtCustNo(Index)) > 5 Then
      txtCustNo(Index) = Left(txtCustNo(Index) & "000", 9)
   End If
End Sub

Private Sub txtDate_GotFocus(Index As Integer)
   TextInverse txtDate(Index)
   CloseIme
End Sub

Private Sub txtDate_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
   End If
End Sub

Private Sub txtDate_Validate(Index As Integer, Cancel As Boolean)
   If txtDate(Index).Text <> "" Then
      If CheckIsTaiwanDate(txtDate(Index).Text, True) = False Then
         Cancel = True
      End If
   End If
End Sub

Private Sub txtSales_Change()
   If Len(txtSales) > 4 Then
      lblSalesName = GetStaffName(txtSales, True)
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
   If PUB_txtSales_Limit(txtSales, m_strListPer, , txtSalesArea, txtSalesArea1, _
                         bolSpecMan, strSpecCode, lblSalesName) = False Then
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
   
   For Each oText In Text1
      oText.Text = ""
   Next
   
   m_A4421 = "": m_A4427 = ""
   
   cmdOK(2).Enabled = False
   cmdOK(6).Enabled = False 'Add By Sindy 2020/10/28
   'cmdOK(4).Enabled = False
   Frame1.Enabled = False
   
   txtSales.Tag = ""
   
   If Not Adodc1.Recordset Is Nothing Then
      If Adodc1.Recordset.State = 1 Then
         Adodc1.Recordset.Close
         DataGrid1.Refresh
      End If
   End If
   
   'Added by Morgan 2014/1/9
   m_A4425 = ""
   
   'Added by Morgan 2015/7/17
   'Removed by Morgan 2023/9/25 現金及支票不開放輸入--楊瑞婷/杜協理同意
   'Text1(1).Enabled = True
   'Text1(4).Enabled = True
   'end 2023/9/25
End Sub

Private Sub SettxtTot()
   Set adoTmp = Adodc1.Recordset.Clone
   With adoTmp
   .MoveFirst
   txtTot(2) = 0
   txtTot(3) = 0
   txtTot(4) = 0
   txtTot(5) = 0
   Do While Not .EOF
      If .Fields(0) = "Y" Then
         txtTot(2) = Val(txtTot(2)) + Val("" & .Fields("服務費"))
         txtTot(3) = Val(txtTot(3)) + Val("" & .Fields("規費"))
         txtTot(4) = Val(txtTot(4)) + Val("" & .Fields("扣繳金額"))
         txtTot(5) = Val(txtTot(2)) + Val(txtTot(3)) - Val(txtTot(4))
      End If
      .MoveNext
   Loop
   
   txtTot(2) = Format(txtTot(2), "#,##0")
   txtTot(3) = Format(txtTot(3), "#,##0")
   txtTot(4) = Format(txtTot(4), "#,##0")
   txtTot(5) = Format(txtTot(5), "#,##0")
   End With
   CheckSum
End Sub

Private Sub CheckSum(Optional idx As Integer)
   Dim dblSum As Double, dblDiff As Double
   
   'Added by Morgan 2014/1/22
   '若有電匯或票據時50元以內差額自動調整到手續費(要剔除手續費欄位離開時,否則該欄位會無法修改)
   If Val(Text1(1)) + Val(Text1(2)) + Val(Text1(3)) > 0 And idx <> 7 Then
      'Modified by Morgan 2015/7/14 +其他
      dblSum = Val(Text1(1)) + Val(Text1(2)) + Val(Text1(3)) + Val(Text1(4)) + Val(Text1(5)) + Val(Text1(0)) - Val(Text1(6)) + Val(Text1(9)) + Val(Text1(10))
      dblDiff = Val(Format(txtTot(5))) - dblSum
      If dblDiff > 0 And dblDiff <= 50 Then
         Text1(7) = dblDiff
      End If
   End If
   'end 2014/1/22
   'Modified by Morgan 2015/7/14 +其他
   txtTot(6) = Format(Val(Text1(1)) + Val(Text1(2)) + Val(Text1(3)) + Val(Text1(4)) + Val(Text1(5)) + Val(Text1(0)) - Val(Text1(6)) + Val(Text1(7)) + Val(Text1(9)) + Val(Text1(10)), "#,##0")
   txtTot(7) = Format(Val(Format(txtTot(6))) - Val(Format(txtTot(5))), "#,##0")
   If Format(txtTot(7)) < 0 Then
      txtTot(7).ForeColor = vbRed
   Else
      txtTot(7).ForeColor = vbBlack
   End If
End Sub

Private Function ShowTT(Optional pbCheck As Boolean) As Boolean
   Dim strCon230 As String
   Dim rsQuery As ADODB.Recordset
   Dim strCustNo As String, strCustName As String
   'Added by Morgan 2015/7/16
   Dim bolOneRec As Boolean
   Dim dblAmt1 As Double, dblAmt3 As Double
   
   If txtCustNo(0) <> "" And txtCustNo(0) <> "X" Then
      strCon230 = strCon230 & " and a2304>='" & txtCustNo(0) & "'"
   End If
   
   If txtCustNo(1) <> "" And txtCustNo(1) <> "X" Then
      strCon230 = strCon230 & " and a2304<='" & txtCustNo(1) & "'"
   End If
   
   'Modified by Morgan 2015/7/16
   'If strCon230 = "" And txtTitle <> "" Then
   ''Modified by Morgan 2014/1/28 抓關係企業
   '   'strCon230 = " and exists(select a0k03 from acc0k0 where a0k03=a2304 and a0k04 like '" & txtTitle & "%')"
   '   strCon230 = " and exists(select a0k03 from acc0k0 where substr(a0k03,1,6)=substr(a2304,1,6) and a0k04 like '" & txtTitle & "%')"
   If txtTitle <> "" Or txtDate(0) & txtDate(1) <> "" Or cboComp.ListIndex > 0 Or chkPrintedOnly.Value = 1 Then
      'Modified by Morgan 2023/5/12 客戶名稱符合也要列(簽收的客戶編號沒有開過收據Ex:X87431000謝淞名)
      'strCon230 = strCon230 & " and exists(select a0k03 from acc0k0 where substr(a0k03,1,6)=substr(a2304,1,6)"
      If txtTitle <> "" Then
         strCon230 = strCon230 & " and (cu04 like '" & txtTitle & "%' or exists(select a0k03 from acc0k0 where substr(a0k03,1,6)=substr(a2304,1,6)"
      Else
         strCon230 = strCon230 & " and exists(select a0k03 from acc0k0 where substr(a0k03,1,6)=substr(a2304,1,6)"
      End If
      'end 2023/5/12
      
      If txtTitle <> "" Then
         strCon230 = strCon230 & " and a0k04 like '" & txtTitle & "%'"
      End If
      If txtDate(0) <> "" Then
         strCon230 = strCon230 & " and a0k02>=" & Val(txtDate(0))
      End If
      If txtDate(1) <> "" Then
         strCon230 = strCon230 & " and a0k02<=" & Val(txtDate(1))
      End If
      If chkPrintedOnly.Value = 1 Then
         strCon230 = strCon230 & " and a0k19>0"
      End If
      If cboComp.ListIndex > 0 Then
         If cboComp.ListIndex = 1 Then
            strCon230 = strCon230 & " and a0k11='1'"
         ElseIf cboComp.ListIndex = 2 Then
            strCon230 = strCon230 & " and a0k11='2'"
         ElseIf cboComp.ListIndex = 3 Then
            strCon230 = strCon230 & " and a0k11='J'"
         'Added by Lydia 2021/09/09
         ElseIf cboComp.ListIndex = 4 Then
            strCon230 = strCon230 & " and a0k11='L'"
         'end 2021/09/09
         End If
      End If
      strCon230 = strCon230 & ")"
      If txtTitle <> "" Then strCon230 = strCon230 & ")" 'Added by Morgan 2023/5/12
      
   'end 2015/7/16
   End If
   'Modified by Morgan 2023/7/12 所別改抓 A2305
   '銀存
   strExc(0) = "select decode(instr('" & m_A4421 & ";" & m_A4427 & "',A2301),0,'','Y') 選取,A2301,sqldatet(A2302) A2302,A2318,cu04,A2305,a0102 Memo,'電匯' Type,A2304,'2' Src" & _
      " from acc230,customer,acc010 where a2303='" & txtSales & "' and a2321 is null and a2304 is not null and a2318>0" & strCon230 & _
      " and cu01(+)=substr(a2304,1,8) and cu02(+)=substr(a2304,9) and a0101(+)=A2322"
   'Modified by Morgan 2015/7/16
   'strExc(0) = strExc(0) & " order by A2302,A2301"
   '現金
   strExc(0) = strExc(0) & " union all" & _
      " select decode(instr('" & m_A4421 & ";" & m_A4427 & "',A2301),0,'','Y') 選取,A2301,sqldatet(A2302) A2302,A2317,cu04,A2305,'' Memo,'現金' Type,A2304,'3' Src" & _
      " from acc230,customer where a2303='" & txtSales & "' and a2321 is null and a2304 is not null and a2317>0" & strCon230 & _
      " and cu01(+)=substr(a2304,1,8) and cu02(+)=substr(a2304,9)"
   '票據
   strExc(0) = strExc(0) & " union all" & _
      " select decode(instr('" & m_A4421 & ";" & m_A4427 & "',A2301),0,'','Y') 選取,A2301,sqldatet(A2302) A2302,A2306,cu04,A2305,decode(a2325,'','',sqldatet(a2325)||' '||a0g02) Memo,'票據'  Type,A2304,'1' Src" & _
      " from acc230,customer,acc0g0 where a2303='" & txtSales & "' and a2321 is null and a2304 is not null and a2306>0" & strCon230 & _
      " and cu01(+)=substr(a2304,1,8) and cu02(+)=substr(a2304,9) and a0g01(+)=a2327"
   strExc(0) = strExc(0) & " order by A2302,A2301,Src"
   'end 2015/7/16
   
   m_A4421 = "": m_A4427 = "": Text1(2) = "": Text1(3) = ""
   intI = 1
   Set adoTmp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      ShowTT = True
      If pbCheck = False Then
         '單筆有輸入條件
         'Added by Morgan 2015/7/16
         'If adoTmp.RecordCount = 1 And strCon230 <> "" Then
         bolOneRec = True
         With adoTmp
         .MoveFirst
         strExc(0) = .Fields("a2301")
         Do While Not .EOF
            If .Fields("a2301") <> strExc(0) Then
               bolOneRec = False
            End If
            .MoveNext
         Loop
         .MoveFirst
         End With
         If bolOneRec = True Then
         'end 2015/7/16
            With adoTmp
            m_A4421 = .Fields("A2301")
            strCustNo = Left(.Fields("A2304"), 6)
            strCustName = "" & .Fields("cu04") 'Added by Morgan 2023/5/12
            'Modified by Morgan 2015/7/16
            'MsgBox "本次電匯銀行帳戶為【" & .Fields("a0102") & "-" & .Fields("cu04") & "】", vbInformation
            strExc(1) = "本次簽收客戶為" & .Fields("cu04") & vbCrLf
            Do While Not .EOF
               strExc(1) = strExc(1) & vbCrLf & "【" & .Fields("Type") & " " & Format(.Fields("A2318"), "#,##0") & " " & .Fields("Memo") & "】"
               If .Fields("Src") = "1" Then
                  dblAmt1 = dblAmt1 + .Fields("A2318")
               ElseIf .Fields("Src") = "2" Then
               'end 2015/7/16
               
                  If .Fields("A2305") = "1" Then
                     Text1(2) = .Fields("A2318")
                  Else
                     Text1(3) = .Fields("A2318")
                  End If
                  
               'Added by Morgan 2015/7/16
               ElseIf .Fields("Src") = "3" Then
                  dblAmt3 = dblAmt3 + .Fields("A2318")
               End If
               .MoveNext
            Loop
            'end 2015/7/16
            End With
         Else
            With frm210141_1
                .txtSales = Me.txtSales
                .lblSalesName = Me.lblSalesName
                'Modify by Amy 2014/06/16 +FormName 改暫存TB
                'Modified by Lydia 2021/07/16 改TableName
                'Set rsQuery = PUB_CreateRecordset(adoTmp, , , , .Name)
                Set rsQuery = PUB_CreateRecordset(adoTmp, , , , "frm210141_1")
                Set .Adodc1.Recordset = rsQuery
                .Show vbModal
            End With

            With rsQuery
            .MoveFirst
            Do While Not .EOF
               If .Fields("選取") = "Y" Then
                  If m_A4421 = "" Then
                     m_A4421 = .Fields("A2301")
                  'Modified by Morgan 2023/2/4 同一簽收可能會有兩筆 Ex:S112020066(票+現金)
                  'Else
                  ElseIf m_A4421 <> .Fields("A2301") Then
                  'end 2023/2/4
                     m_A4427 = m_A4427 & ";" & .Fields("A2301")
                  End If
                  strCustNo = Left(.Fields("A2304"), 6)
                  strCustName = "" & .Fields("cu04") 'Added by Morgan 2023/5/12
                  'Added by Morgan 2015/7/16
                  If .Fields("Src") = "1" Then
                     dblAmt1 = dblAmt1 + .Fields("A2318")
                  ElseIf .Fields("Src") = "2" Then
                  'end 2015/7/16
                     If .Fields("A2305") = "1" Then
                        Text1(2) = Val(Text1(2)) + .Fields("A2318")
                     Else
                        Text1(3) = Val(Text1(3)) + .Fields("A2318")
                     End If
                  'Added by Morgan 2015/7/16
                  ElseIf .Fields("Src") = "3" Then
                     dblAmt3 = dblAmt3 + .Fields("A2318")
                  End If
                  'end 2015/7/16
               End If
               .MoveNext
            Loop
            If m_A4427 <> "" Then m_A4427 = Mid(m_A4427, 2) '去掉字首的分號
            End With
            
         End If
         
         'Added by Morgan 2015/7/17 若從簽收來的不可修改--辜
         'Modified by Morgan 2023/9/25 現金及支票不開放輸入--楊瑞婷/杜協理同意
         'If Text1(1).Enabled = False Then Text1(1) = "": Text1(1).Enabled = True
         'If Text1(4).Enabled = False Then Text1(4) = "": Text1(4).Enabled = True
         Text1(1) = ""
         Text1(4) = ""
         'end 2023/9/25
         If dblAmt1 > 0 Then Text1(1) = dblAmt1: Text1(1).Enabled = False
         If dblAmt3 > 0 Then Text1(4) = dblAmt3: Text1(4).Enabled = False
         If bolOneRec = True Then
            MsgBox strExc(1), vbInformation
         End If
         'end 2015/7/17
         
         'Added by Morgan 2014/1/22
         '若收據總計為0時,依電匯的關係企業編號重新帶收據資料
         If Val(Format(txtTot(5))) = 0 And strCustNo <> "" Then
            'Modified by Morgan 2015/7/16 抬頭條件不要清除--余政興
            'txtCustNo(0) = strCustNo & "000"
            'txtCustNo(1) = strCustNo & "ZZZ"
            'txtTitle = ""
            If txtCustNo(0) < strCustNo & "000" Then txtCustNo(0) = strCustNo & "000"
            If txtCustNo(1) = "X" Or txtCustNo(1) = "" Or txtCustNo(1) > strCustNo & "ZZZ" Then txtCustNo(1) = strCustNo & "ZZZ"
            'end 2015/7/16
            
            'Modified by Morgan 2023/5/12 若編號找不到資料時再用客戶名稱找一次
            'doQuery , True
            doQuery True, True
            If adoTmp.RecordCount = 0 Then
               txtCustNo(0) = "X": txtCustNo(1) = "X"
               txtTitle = strCustName
               doQuery , True
            End If
            'end 2023/5/12
            
         End If
         'end 2014/1/22
            
         CheckSum
      End If
      
   ElseIf intI = 0 Then
      If pbCheck = False Then
         MsgBox "目前已無未確認之簽收記錄！", vbExclamation
         'cmdOK(4).Enabled = False
      End If
   End If
   Set rsQuery = Nothing
End Function

Private Function ShowDelete()
   strExc(0) = "select sqldatet(a4402) rdate,sqltime(a4403) rtime,x.* from acc440 x" & _
      " where a4401='" & txtSales & "' and a4413||a4416 is null"
   intI = 1
   Set adoTmp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With frm210141_2
         .txtSales = Me.txtSales
         .lblSalesName = Me.lblSalesName
         'Modify by Amy 2014/06/16 +FormName 改暫存TB
         'Modified by Lydia 2021/07/16 改TableName
         'Set .Adodc1.Recordset = PUB_CreateRecordset(adoTmp, , , , .Name)
         Set .Adodc1.Recordset = PUB_CreateRecordset(adoTmp, , , , "frm210141_2")
        .Show vbModal
      End With
      
   ElseIf intI = 0 Then
      MsgBox "無可刪除之繳款記錄！", vbExclamation
   End If

End Function

Private Sub txtSalesArea_GotFocus()
   TextInverse txtSalesArea
   CloseIme
End Sub

Private Sub txtSalesArea_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtSalesArea1_GotFocus()
   TextInverse txtSalesArea1
   CloseIme
End Sub

Private Sub txtSalesArea1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtSalesArea1_Validate(Cancel As Boolean)
   If Trim(txtSalesArea1) <> "" Then
      If RunNick(txtSalesArea, txtSalesArea1) = True Then
         Cancel = True
         Exit Sub
      End If
   End If
End Sub

Private Sub txtTitle_GotFocus()
   TextInverse txtTitle
   OpenIme
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

