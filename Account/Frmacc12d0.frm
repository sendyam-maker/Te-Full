VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc12d0 
   AutoRedraw      =   -1  'True
   Caption         =   "應收帳款綜合查詢"
   ClientHeight    =   5950
   ClientLeft      =   50
   ClientTop       =   340
   ClientWidth     =   9200
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5950
   ScaleWidth      =   9200
   Begin VB.CheckBox ChkBillDate 
      Caption         =   "排除未達客戶付款週期之應收帳款"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5100
      TabIndex        =   22
      Top             =   2352
      Width           =   3984
   End
   Begin VB.CheckBox Check2 
      Caption         =   "不含智權部同仁"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5100
      TabIndex        =   21
      Top             =   2040
      Width           =   2148
   End
   Begin VB.TextBox Text12 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Index           =   1
      Left            =   2430
      MaxLength       =   6
      TabIndex        =   20
      Top             =   2016
      Width           =   880
   End
   Begin VB.TextBox Text12 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Index           =   0
      Left            =   1340
      MaxLength       =   6
      TabIndex        =   19
      Top             =   2016
      Width           =   880
   End
   Begin VB.TextBox TxtDeptS 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   4440
      MaxLength       =   3
      TabIndex        =   6
      Top             =   696
      Width           =   660
   End
   Begin VB.TextBox TxtDeptE 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   5460
      MaxLength       =   3
      TabIndex        =   7
      Top             =   696
      Width           =   660
   End
   Begin VB.CommandButton Command4 
      Caption         =   "E-Mail"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8230
      TabIndex        =   26
      Top             =   1110
      Width           =   800
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   960
      MaxLength       =   3
      TabIndex        =   14
      Top             =   1680
      Width           =   592
   End
   Begin VB.TextBox Text9 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1545
      MaxLength       =   6
      TabIndex        =   15
      Top             =   1680
      Width           =   852
   End
   Begin VB.TextBox Text10 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2385
      MaxLength       =   1
      TabIndex        =   16
      Top             =   1680
      Width           =   252
   End
   Begin VB.TextBox Text11 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2625
      MaxLength       =   2
      TabIndex        =   17
      Top             =   1680
      Width           =   372
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   4440
      TabIndex        =   13
      Text            =   "1"
      Top             =   1350
      Width           =   400
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   960
      TabIndex        =   11
      Top             =   1350
      Width           =   1200
   End
   Begin VB.CommandButton cmdLikeSearch 
      Caption         =   "搜尋"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7920
      TabIndex        =   1
      Top             =   45
      Width           =   675
   End
   Begin VB.CommandButton Command1 
      Caption         =   "收據抬頭修改"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7250
      TabIndex        =   25
      Top             =   750
      Width           =   1785
   End
   Begin VB.CommandButton cmdChoice 
      Caption         =   "取消"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   8230
      TabIndex        =   28
      Top             =   1710
      Visible         =   0   'False
      Width           =   800
   End
   Begin VB.CommandButton cmdChoice 
      Caption         =   "選擇"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   7320
      TabIndex        =   27
      Top             =   1710
      Visible         =   0   'False
      Width           =   800
   End
   Begin VB.CommandButton Command3 
      Caption         =   "請款單開立發票"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7250
      TabIndex        =   24
      Top             =   390
      Width           =   1785
   End
   Begin VB.CheckBox Check1 
      Caption         =   "含未列印收據"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5100
      TabIndex        =   18
      Top             =   1710
      Width           =   1815
   End
   Begin VB.TextBox txtAmt 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   5
      Left            =   8124
      TabIndex        =   49
      Top             =   5580
      Width           =   1000
   End
   Begin VB.TextBox txtAmt 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   6600
      TabIndex        =   47
      Top             =   5580
      Width           =   1000
   End
   Begin VB.TextBox txtAmt 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   5064
      TabIndex        =   45
      Top             =   5580
      Width           =   1000
   End
   Begin VB.TextBox txtAmt 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   3540
      TabIndex        =   43
      Top             =   5580
      Width           =   1000
   End
   Begin VB.TextBox txtAmt 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   2004
      TabIndex        =   41
      Top             =   5580
      Width           =   1000
   End
   Begin VB.TextBox txtAmt 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   480
      TabIndex        =   39
      Top             =   5580
      Width           =   1000
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc12d0.frx":0000
      Height          =   2892
      Left            =   84
      TabIndex        =   38
      Top             =   2664
      Width           =   9000
      _ExtentX        =   15875
      _ExtentY        =   5098
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   18
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "客戶帳款查詢"
      ColumnCount     =   22
      BeginProperty Column00 
         DataField       =   "NotSend"
         Caption         =   "送件"
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
         DataField       =   "a0k11"
         Caption         =   "公司"
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
         DataField       =   "RDate"
         Caption         =   "單據日期"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "###/##/##"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "st02"
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
      BeginProperty Column04 
         DataField       =   "RNo"
         Caption         =   "單據編號"
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
         DataField       =   "a0k32"
         Caption         =   "列"
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
         DataField       =   "axc01"
         Caption         =   "發"
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
      BeginProperty Column07 
         DataField       =   "a0j02"
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
      BeginProperty Column08 
         DataField       =   "na03"
         Caption         =   "申請國家"
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
         DataField       =   "cp10N"
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
      BeginProperty Column10 
         DataField       =   "RAmount"
         Caption         =   "應收金額"
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
      BeginProperty Column11 
         DataField       =   "EAmount"
         Caption         =   "已收金額"
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
         DataField       =   "DAmount"
         Caption         =   "扣繳額"
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
         DataField       =   "CAmount"
         Caption         =   "銷帳金額"
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
      BeginProperty Column14 
         DataField       =   "BAmount"
         Caption         =   "退費金額"
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
      BeginProperty Column15 
         DataField       =   "NAmount"
         Caption         =   "未收金額"
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
      BeginProperty Column16 
         DataField       =   "AccNo"
         Caption         =   "傳票編號"
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
         DataField       =   "a0j01"
         Caption         =   "總收文號"
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
         DataField       =   "a0k03"
         Caption         =   "客戶代號"
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
         DataField       =   "CusName"
         Caption         =   "客戶名稱"
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
         DataField       =   "a0k40"
         Caption         =   "INVOICE編號"
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
            ColumnWidth     =   290.268
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   530.079
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   970.016
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   849.827
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   319.748
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   319.748
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1319.811
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   879.874
         EndProperty
         BeginProperty Column09 
         EndProperty
         BeginProperty Column10 
            Alignment       =   1
         EndProperty
         BeginProperty Column11 
            Alignment       =   1
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column12 
            Alignment       =   1
            ColumnWidth     =   1250.079
         EndProperty
         BeginProperty Column13 
            Alignment       =   1
         EndProperty
         BeginProperty Column14 
            Alignment       =   1
         EndProperty
         BeginProperty Column15 
            Alignment       =   1
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   2810.268
         EndProperty
         BeginProperty Column17 
            ColumnWidth     =   1310.173
         EndProperty
         BeginProperty Column18 
            ColumnWidth     =   2569.89
         EndProperty
         BeginProperty Column19 
            ColumnWidth     =   1069.795
         EndProperty
         BeginProperty Column20 
            ColumnWidth     =   2269.984
         EndProperty
         BeginProperty Column21 
            ColumnWidth     =   1360.063
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   312
      Left            =   60
      Top             =   2580
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   2117
      _ExtentY        =   564
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
   Begin VB.CommandButton Command2 
      Caption         =   "單據內容"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5970
      TabIndex        =   23
      Top             =   390
      Width           =   1212
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   4440
      TabIndex        =   10
      Text            =   "1"
      Top             =   1020
      Width           =   400
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   1740
      MaxLength       =   1
      TabIndex        =   9
      Top             =   1020
      Width           =   400
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   960
      MaxLength       =   1
      TabIndex        =   8
      Top             =   1020
      Width           =   400
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   960
      TabIndex        =   4
      Top             =   690
      Width           =   1100
      _ExtentX        =   1940
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   2280
      MaxLength       =   9
      TabIndex        =   3
      Top             =   360
      Width           =   1100
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   960
      MaxLength       =   9
      TabIndex        =   2
      Top             =   360
      Width           =   1100
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   300
      Left            =   2280
      TabIndex        =   5
      Top             =   690
      Width           =   1100
      _ExtentX        =   1940
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Label lblPS 
      Caption         =   "列：N 未列印收據，Z 不列印收據，# 已開INVOICE"
      ForeColor       =   &H000000C0&
      Height          =   216
      Left            =   72
      TabIndex        =   60
      Top             =   2424
      Width           =   4260
   End
   Begin VB.Label Label18 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   250
      Left            =   2250
      TabIndex        =   59
      Top             =   2060
      Width           =   280
   End
   Begin VB.Label Label17 
      BackStyle       =   0  '透明
      Caption         =   "INVOICE編號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   60
      TabIndex        =   58
      Top             =   2040
      Width           =   1380
   End
   Begin VB.Label Label16 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   260
      Left            =   5220
      TabIndex        =   57
      Top             =   690
      Width           =   260
   End
   Begin VB.Label Label15 
      BackStyle       =   0  '透明
      Caption         =   "部門"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   56
      Top             =   690
      Width           =   600
   End
   Begin MSForms.TextBox Text13 
      Height          =   300
      Left            =   2190
      TabIndex        =   12
      Top             =   1350
      Width           =   1200
      VariousPropertyBits=   671105055
      BackColor       =   -2147483633
      MaxLength       =   30
      Size            =   "2117;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cboTitle 
      Height          =   300
      Left            =   960
      TabIndex        =   0
      Top             =   30
      Width           =   6780
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "11959;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label14 
      Caption         =   "註: 送件 o  未送件 x"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   3420
      TabIndex        =   55
      Top             =   1770
      Width           =   1605
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '透明
      Caption         =   "本所案號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   60
      TabIndex        =   54
      Top             =   1680
      Width           =   972
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '透明
      Caption         =   "排序"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   53
      Top             =   1380
      Width           =   600
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '透明
      Caption         =   "(1.抬頭+日期 2.收據編號)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5100
      TabIndex        =   52
      Top             =   1380
      Width           =   2655
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "智權人員"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   60
      TabIndex        =   51
      Top             =   1356
      Width           =   972
   End
   Begin VB.Label lblAmt 
      BackStyle       =   0  '透明
      Caption         =   "未收"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   5
      Left            =   7692
      TabIndex        =   50
      Top             =   5604
      Width           =   492
   End
   Begin VB.Label lblAmt 
      BackStyle       =   0  '透明
      Caption         =   "退費"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   4
      Left            =   6156
      TabIndex        =   48
      Top             =   5604
      Width           =   492
   End
   Begin VB.Label lblAmt 
      BackStyle       =   0  '透明
      Caption         =   "銷帳"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   3
      Left            =   4632
      TabIndex        =   46
      Top             =   5604
      Width           =   492
   End
   Begin VB.Label lblAmt 
      BackStyle       =   0  '透明
      Caption         =   "扣繳"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   2
      Left            =   3096
      TabIndex        =   44
      Top             =   5604
      Width           =   492
   End
   Begin VB.Label lblAmt 
      BackStyle       =   0  '透明
      Caption         =   "已收"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   1
      Left            =   1572
      TabIndex        =   42
      Top             =   5604
      Width           =   492
   End
   Begin VB.Label lblAmt 
      BackStyle       =   0  '透明
      Caption         =   "應收"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   0
      Left            =   36
      TabIndex        =   40
      Top             =   5604
      Width           =   492
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   -24
      Top             =   96
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "(1.未收 2.收款 3.往來)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5100
      TabIndex        =   37
      Top             =   1080
      Width           =   2655
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "查詢資料"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3570
      TabIndex        =   36
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   260
      Left            =   1500
      TabIndex        =   35
      Top             =   1020
      Width           =   260
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "公司別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   60
      TabIndex        =   34
      Top             =   1020
      Width           =   972
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2100
      TabIndex        =   33
      Top             =   690
      Width           =   255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "往來日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   60
      TabIndex        =   32
      Top             =   696
      Width           =   972
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "收據抬頭"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   60
      TabIndex        =   31
      Top             =   60
      Width           =   972
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2100
      TabIndex        =   30
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "客戶代號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   60
      TabIndex        =   29
      Top             =   360
      Width           =   972
   End
End
Attribute VB_Name = "Frmacc12d0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/16 Form2.0已修改
'Memo by Sindy 2016/6/6 依據Frmacc1220客戶帳款查詢修改部分功能
Option Explicit

Public adoacc0m0 As New ADODB.Recordset
Public adocustomer As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Dim strSql As String
Dim strSQL1 As String
Dim strSQL2 As String
Dim strType As String
Public frmCall As Form 'Added by Morgan 2015/7/23


'Add By Sindy 2016/6/7
Private Sub cboTitle_Click()
   If cboTitle.ListIndex > 0 Then
      'Modify By Sindy 2024/12/31 + Text1.Text = "X" Or
      '                             Text2.Text = "X" Or
      If Text1.Text = "X" Or Text1.Text = "" Then
         Text1.Text = Right(cboTitle.Text, 9)
      ElseIf Text2.Text = "X" Or Text2.Text = "" Then
         Text2.Text = Right(cboTitle.Text, 9)
      End If
      strExc(1) = cboTitle.List(cboTitle.ListIndex)
      cboTitle.List(0) = RTrim(Left(strExc(1), Len(strExc(1)) - 9))
   End If
   cboTitle.ListIndex = 0
End Sub
Private Sub cboTitle_GotFocus()
   OpenIme
End Sub
Private Sub cboTitle_KeyPress(KeyAscii As MSForms.ReturnInteger)
   If Text1 <> "" Or Text2 <> "" Or cboTitle.ListCount > 0 Then
      Text1 = "": Text2 = ""
      Text7 = "": Text13 = ""
      cboTitle.Clear
   End If
End Sub
Private Sub cboTitle_Validate(Cancel As Boolean)
   If CheckLen(Label1, cboTitle, 100) = MsgText(603) Then
      Cancel = True
      Exit Sub
   End If
   '切換輸入法改用API
   If Cancel = False Then CloseIme
End Sub
'2016/6/7 END

'Added by Morgan 2015/7/23
Private Sub cmdChoice_Click(Index As Integer)
   If Index = 1 Then
      frmCall.GetSelect
   End If
   Unload Me
End Sub

'Add By Sindy 2016/6/7
Private Sub cmdLikeSearch_Click()
Dim strText1 As String, strText2 As String
Dim strTitle As String 'Add By Sindy 2025/1/3
   
   strTitle = cboTitle.Text 'Add By Sindy 2025/1/3
   If cboTitle.Text = "" Then
      MsgBox "請輸入收據抬頭！", vbCritical
   Else
      'Add By Sindy 2025/1/3
      If Text1 <> "" Or Text2 <> "" Or cboTitle.ListCount > 0 Then
         Text1 = "": Text2 = ""
         Text7 = "": Text13 = ""
         cboTitle.Clear
         cboTitle.Text = strTitle
      End If
      '2025/1/3 END
      'ADD BY Sindy 2024/12/31 若客戶編號為X~X則略過
      strText1 = Trim(Text1)
      strText2 = Trim(Text2)
      If Trim(Text1) = Trim(Text2) And (Trim(Text1) = "X" Or Trim(Text1) = "") Then
         strText1 = ""
         strText2 = ""
      End If
      '2024/12/31 END
      'PUB_AddItem2CboTitle cboTitle, Text1, Text2, "" ', True
      PUB_AddItem2CboTitle cboTitle, strText1, strText2, "" ', True
   End If
End Sub

'Added by Lydia 2016/01/19 點選呼叫"收據抬頭修改"
Private Sub Command1_Click()
   If Adodc1.Recordset.State = 1 Then
      If Adodc1.Recordset.RecordCount = 0 Then
         Exit Sub
      End If
      strItemNo = Adodc1.Recordset.Fields("RNo").Value
      strTitle = Me.Name
      If Mid(strItemNo, 1, 1) = "E" Then
         tool14_enabled
         Frmacc1140.Show
         Me.Enabled = False
      Else
         MsgBox "請點選收據/請款單資料..."
         strItemNo = ""
         strTitle = ""
      End If
   Else
      MsgBox "請先按 F12 查詢並點選單據資料..."
   End If
End Sub
'end 2016/01/19

Private Sub Command2_Click()
   'Modified by Morgan 2012/9/14 有資料並點選才可顯示單據內容
   If Adodc1.Recordset.State = 1 Then
      If Adodc1.Recordset.RecordCount = 0 Then
         Exit Sub
      End If
      strCon1 = Adodc1.Recordset.Fields("RNo").Value
      strFormLink = Name
      Select Case Mid(strCon1, 1, 1)
         Case "F"
            Frmacc1221.Show
            Me.Enabled = False
         Case "E"
            'Added by Lydia 2016/04/11 舊單據無明細檔
            If IsNull(Adodc1.Recordset.Fields("a0j02")) Then
               MsgBox "無本所案號!", vbCritical
               Exit Sub
            End If
            Frmacc1222.Show
            Me.Enabled = False
         Case "I"
            Frmacc1224.Show
            Me.Enabled = False
      End Select
   Else
      MsgBox "請先按 F12 查詢並點選單據資料..."
   End If
End Sub

'Add By Sindy 2013/12/31
'請款單開立發票
Private Sub Command3_Click()
   If Adodc1.Recordset.State = 1 Then
      If Adodc1.Recordset.RecordCount = 0 Then
         Exit Sub
      End If
      'Modify By Sindy 2016/6/8
      'If Adodc1.Recordset.Fields("a0k11").Value = "J" And
      If Replace(Adodc1.Recordset.Fields("a0k11").Value, "x", "") = "J" And _
         Left(Adodc1.Recordset.Fields("RNo").Value, 1) = "E" Then
      '2016/6/8 END
         strItemNo = Adodc1.Recordset.Fields("RNo").Value
         strTitle = Me.Name
         Me.Enabled = False
         Screen.MousePointer = vbHourglass
         Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
         Frmacc1127.Text1.Enabled = False
         Frmacc1127.Show
         Screen.MousePointer = vbDefault
         Frmacc0000.StatusBar1.Panels(1).Text = MsgText(601)
      Else
         MsgBox "J 公司的請款單,才可開立發票!!"
      End If
   Else
      MsgBox "請先按 F12 查詢並點選單據資料..."
   End If
End Sub

'Added by Morgan 2015/7/23
Public Sub SetForm()
   Dim ii As Integer
   
   Command3.Visible = False
   cmdChoice(0).Visible = True
   cmdChoice(1).Visible = True
   For ii = 0 To DataGrid1.Columns.Count - 1
      If DataGrid1.Columns(ii).Caption = "收據抬頭" Then
         DataGrid1.Columns(ii).Visible = True
      ElseIf DataGrid1.Columns(ii).Caption = "公司" Then
         DataGrid1.Columns(ii).Visible = True
      ElseIf DataGrid1.Columns(ii).Caption = "客戶代號" Then
         DataGrid1.Columns(ii).Visible = True
      ElseIf DataGrid1.Columns(ii).Caption = "客戶名稱" Then
         DataGrid1.Columns(ii).Visible = True
      ElseIf DataGrid1.Columns(ii).Caption = "智權人員" Then
         DataGrid1.Columns(ii).Visible = True
      ElseIf DataGrid1.Columns(ii).Caption = "未收金額" Then
         DataGrid1.Columns(ii).Visible = True
      Else
         DataGrid1.Columns(ii).Visible = False
      End If
   Next
   Forms(0).Toolbar1.Enabled = False
End Sub

'Add by Sindy 2016/6/13 點選E-Mail
Private Sub Command4_Click()
   If Adodc1.Recordset.State = 1 Then
      If Adodc1.Recordset.RecordCount = 0 Then
         Exit Sub
      End If
      strFormLink = Name
      strItemNo = Adodc1.Recordset.Fields("RNo").Value
      strTitle = Me.Name
      If Mid(strItemNo, 1, 1) = "E" Then
         If Val(Adodc1.Recordset.Fields("NAmount").Value) > 0 Then
            tool14_enabled
            Frmacc12d1.Text7 = Adodc1.Recordset.Fields("a0j02").Value '本所案號
            Frmacc12d1.Text8 = Adodc1.Recordset.Fields("cp10N").Value '案件性質
            Frmacc12d1.Text9 = Format(Val(Adodc1.Recordset.Fields("NAmount").Value), DDollar2) '未收金額
            Frmacc12d1.m_Appl = Adodc1.Recordset.Fields("CusName").Value '申請人
            Frmacc12d1.m_A0K04 = Adodc1.Recordset.Fields("a0k04").Value '收據抬頭
            Frmacc12d1.Show
            Me.Enabled = False
         Else
            MsgBox strItemNo & "無未收金額!"
            strItemNo = ""
            strTitle = ""
         End If
      Else
         MsgBox "請點選收據/請款單資料..."
         strItemNo = ""
         strTitle = ""
      End If
   Else
      MsgBox "請先按 F12 查詢並點選單據資料..."
   End If
End Sub
'2016/6/13 END

Private Sub Form_Activate()
   strFormName = Name
   strFormLink = ""
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   
'Modified by Lydia 2023/11/13 表單初始化
'Modified by Lydia 2025/05/23 H: 6100=> 6400 ; 順便刪除舊Code
   PUB_InitForm Me, 9300, 6400, strBackPicPath2, lngWidth, lngHeight
'end 2023/11/13

   Text1 = "X"
   Text2 = "X"
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat
   '2013/1/8 ADD BY SONIA 瑞婷說直接預設系統日前一天
   MaskEdBox2.Text = CFDate(ACDate(PUB_GetWorkDay1(strSrvDate(1) - 1, True)))
   OpenTable
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
   
   'Add By Sindy 2014/1/3
   If PUB_GetST06(strUserNum) = "1" Then
      Command3.Visible = True
      'Added by Lydia 2016/01/20
      'Added by Lydia 2016/01/25 +判斷,辜從簽收作業(新增或修改)->客戶帳款查詢,有時連按2下會造成跳到收據抬頭修改
      If frmCall Is Nothing Then
         Command1.Visible = True
      Else
         Command1.Visible = False
      End If
   Else
      Command3.Visible = False
      'Added by Lydia 2016/01/20 分所不可使用"收據抬頭修改"
      Command1.Visible = False
   End If
   '2014/1/3 END
End Sub

Private Sub Form_Unload(Cancel As Integer)
   StatusClear
   strFormName = MsgText(601)
   'Added by Morgan 2015/7/23
   If Not frmCall Is Nothing Then
      Forms(0).Toolbar1.Enabled = True
      frmCall.Enabled = True
      frmCall.SetFocus
   Else
   'end 2015/7/23
      KeyEnter vbKeyEscape
      MenuEnabled
   End If 'Added by Morgan 2015/7/23
   
   Set Frmacc12d0 = Nothing
End Sub

'2013/1/8 CANCEL BY SONIA 瑞婷說直接預設故取消
'Private Sub MaskEdBox1_Validate(Cancel As Boolean)
'   MaskEdBox2.Mask = ""
'   'MaskEdBox2.Text = MaskEdBox1.Text
'   MaskEdBox2.Mask = DFormat
'End Sub
'2013/1/8 END

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
'on error GoTo Checking
   adoadodc1.CursorLocation = adUseClient
   'Modify by Morgan 2011/8/23 改從 0j0 抓 cp
   'Modified by Morgan 2011/12/27 取消 a0j20,a0j21
   'adoadodc1.Open "select a0k03, a0k04, a0k11, a0k02, a0k01, a0k20, a0j02, cp09, a0j21, a0j20, (a0j09 + a0j10) as RAmount, cp75 as EAmount, cp76, cp77, cp78 as BAmount, cp79 as NAmount from acc0k0, acc0j0, caseprogress where a0j13(+)= a0k01 and cp09(+) = a0j01 and a0k03 = '1'", adoTaie, adOpenStatic, adLockReadOnly
   adoadodc1.Open "select a0k03, a0k04, a0k11, a0k02, a0k01, a0k20, a0j02, cp09, na03,getcp10desc(cp01,cp10,a0j04) cp10N, (a0j09 + a0j10) as RAmount, cp75 as EAmount, cp76, cp77, cp78 as BAmount, cp79 as NAmount from acc0k0, acc0j0, caseprogress,nation where a0j13(+)= a0k01 and cp09(+) = a0j01 and a0k03 = '1' and na01(+)=a0j04", adoTaie, adOpenStatic, adLockReadOnly
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   Select Case Len(Text1)
      Case 6
         Text1 = AfterZero(Text1)
      Case 7
         Text1 = Text1 & "00"
      Case 8
         Text1 = Text1 & "0"
   End Select
   Text2 = Text1
End Sub

Private Sub Text2_GotFocus()
   If Text1.Text <> "" Then
      'Modify By Sindy 2014/8/11 999=>ZZZ
      'Text2.Text = Left(Text1.Text, 6) & "999"
      Text2.Text = Left(Text1.Text, 6) & "ZZZ"
   End If
   TextInverse Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
   Select Case Len(Text2)
      Case 6
         Text2 = AfterZero(Text2)
      Case 7
         Text2 = Text2 & "00"
      Case 8
         Text2 = Text2 & "0"
   End Select
End Sub

'Add By Sindy 2016/6/7
Private Sub Text3_GotFocus()
   TextInverse Text3
End Sub

Private Sub Text4_GotFocus()
   TextInverse Text4
   CloseIme
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text5_GotFocus()
   TextInverse Text5
   CloseIme
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text6_GotFocus()
   TextInverse Text6
End Sub

'*************************************************
'  重新整理 Adodc 之資料
'
'*************************************************
Public Sub AdodcRefresh()
Dim strUnion As String
   
On Error GoTo Checking
   
   strSql = ""
   strSQL1 = ""
   strSQL2 = ""
   strType = ""
   'ADD BY SONIA 2014/6/20 若客戶編號為X~X則略過
   If Text1 = Text2 And Text1 = "X" Then
   Else
   'END 2014/6/20
      If Text1 <> MsgText(601) Then
         strSql = " and a0k03 >= '" & Text1 & "'"
         strSQL1 = " and a0k03 >= '" & Text1 & "'"
         strSQL2 = " and a0k03 >= '" & Text1 & "'"
      End If
      If Text2 <> MsgText(601) Then
         strSql = strSql & " and a0k03 <= '" & Text2 & "'"
         strSQL1 = strSQL1 & " and a0k03 <= '" & Text2 & "'"
         strSQL2 = strSQL2 & " and a0k03 <= '" & Text2 & "'"
      End If
   End If 'ADD BY SONIA 2014/6/20
      
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      strSql = strSql & " and a0k02 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
      strSQL1 = strSQL1 & " and a0l02 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
      strSQL2 = strSQL2 & " and a0s03 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
   End If
   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      strSql = strSql & " and a0k02 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
      strSQL1 = strSQL1 & " and a0l02 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
      strSQL2 = strSQL2 & " and a0s03 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
   End If
   If Text4 <> MsgText(601) Then
      strSql = strSql & " and a0k11 >= '" & Text4 & "'"
      strSQL1 = strSQL1 & " and a0k11 >= '" & Text4 & "'"
      strSQL2 = strSQL2 & " and a0k11 >= '" & Text4 & "'"
   End If
   If Text5 <> MsgText(601) Then
      strSql = strSql & " and a0k11 <= '" & Text5 & "'"
      strSQL1 = strSQL1 & " and a0k11 <= '" & Text5 & "'"
      strSQL2 = strSQL2 & " and a0k11 <= '" & Text5 & "'"
   End If
   
   '2012/10/15 add by sonia +不包含未列印收據故加 and a0k32 is null條件
   If Me.Check1.Value = 0 Then
       'Modified by Lydia 2023/11/13 含Z.確定不印
       'strSql = strSql & " and A0K32 is null"
       'strSQL1 = strSQL1 & " and A0K32 is null"
       'strSQL2 = strSQL2 & " and A0K32 is null"
       strSql = strSql & " and nvl(A0K32,'Z') = 'Z'"
       strSQL1 = strSQL1 & " and nvl(A0K32,'Z') = 'Z'"
       strSQL2 = strSQL2 & " and nvl(A0K32,'Z') = 'Z'"
       'end 2023/11/13
   End If
   
   Select Case Text6
      Case "2"
         'Modify by Morgan 2006/4/13 有收過款的都要 --瑞婷,秀玲
         'strType = " and (a0k06+a0k07) <= (nvl(a0k17, 0)+nvl(a0k18, 0))"
         strType = " and (nvl(a0k17, 0)+nvl(a0k18, 0))>0"
      Case "1"
         strType = " and (a0k06+a0k07) > (nvl(a0k17, 0)+nvl(a0k18, 0)) and nvl(cp79, 0) <> 0"
   End Select
   'Add By Cheng 2004/01/12
   '若非北所員工, 只能列印該所資料
   If pub_strUserOffice <> "1" Then
       strSql = strSql & " And CU13=ST01(+) And ''||ST06='" & pub_strUserOffice & "' "
       strSQL1 = strSQL1 & " And CU13=ST01(+) And ''||ST06='" & pub_strUserOffice & "' "
       strSQL2 = strSQL2 & " And CU13=ST01(+) And ''||ST06='" & pub_strUserOffice & "' "
   Else
       strSql = strSql & " And A0K20=ST01(+) "
       strSQL1 = strSQL1 & " And A0K20=ST01(+) "
       strSQL2 = strSQL2 & " And A0K20=ST01(+) "
   End If
   'End
   
   'Add by Amy 2023/08/14 +部門
   If TxtDeptS <> MsgText(601) Then
      strSql = strSql & " And ST15>='" & TxtDeptS & "' "
      strSQL1 = strSQL1 & " And ST15>='" & TxtDeptS & "' "
      strSQL2 = strSQL2 & " And ST15>='" & TxtDeptS & "' "
   End If
   If TxtDeptE <> MsgText(601) Then
      strSql = strSql & " And ST15<='" & TxtDeptE & "' "
      strSQL1 = strSQL1 & " And ST15<='" & TxtDeptE & "' "
      strSQL2 = strSQL2 & " And ST15<='" & TxtDeptE & "' "
   End If
   'end 2023/08/14
   'Add By Sindy 2025/4/24
   If Me.Check2.Value = 1 Then '不含智權部同仁
      strSql = strSql & " And substr(ST15,1,1)<>'S' "
      strSQL1 = strSQL1 & " And substr(ST15,1,1)<>'S' "
      strSQL2 = strSQL2 & " And substr(ST15,1,1)<>'S' "
   End If
   '2025/4/24 END
   
   'Add By Sindy 2016/6/7
   '收據抬頭
   If cboTitle.Text <> MsgText(601) Then
      '2011/10/20 MODIFY BY SONIA E10023515
      'strSql = strSql & " and instr(a0k04, '" & cboTitle.Text & "') > 0"
      'strSQL1 = strSQL1 & " and instr(a0k04, '" & cboTitle.Text & "') > 0"
      'strSQL2 = strSQL2 & " and instr(a0k04, '" & cboTitle.Text & "') > 0"
      strSql = strSql & " and instr(UPPER(a0k04), UPPER('" & cboTitle.Text & "')) > 0"
      strSQL1 = strSQL1 & " and instr(UPPER(a0k04), UPPER('" & cboTitle.Text & "')) > 0"
      strSQL2 = strSQL2 & " and instr(UPPER(a0k04), UPPER('" & cboTitle.Text & "')) > 0"
   End If
   '智權人員
   If Text7 <> MsgText(601) Then
      strSql = strSql & " and a0k20 = '" & Text7 & "'"
      strSQL1 = strSQL1 & " and a0k20 = '" & Text7 & "'"
      strSQL2 = strSQL2 & " and a0k20 = '" & Text7 & "'"
   End If
   
   '本所案號
   If Text8 <> "" And Text9 <> "" And Text10 <> "" And Text11 <> "" Then
      strSql = strSql & " and a0j02 = '" & Text8 & Text9 & Text10 & Text11 & "'"
      strSQL1 = strSQL1 & " and a0j02 = '" & Text8 & Text9 & Text10 & Text11 & "'"
      strSQL2 = strSQL2 & " and a0j02 = '" & Text8 & Text9 & Text10 & Text11 & "'"
      
   'Added by Morgan 2024/4/19 +只輸入系統別也可查詢 --瑞婷
   ElseIf Text8 <> "" Then
      strSql = strSql & " and a0j02 like '" & Text8 & "%'"
      strSQL1 = strSQL1 & " and a0j02 like '" & Text8 & "%'"
      strSQL2 = strSQL2 & " and a0j02 like '" & Text8 & "%'"
   'end 2024/4/19
   End If
   '2016/6/7 END
   
   'Added by Lydia 2023/11/13 INVOICE No.
   If Text12(0) <> MsgText(601) Then
      strSql = strSql & " and a0k40 >= '" & Text12(0) & "'"
      strSQL1 = strSQL1 & " and a0k40 >= '" & Text12(0) & "'"
      strSQL2 = strSQL2 & " and a0k40 >= '" & Text12(0) & "'"
   End If
   If Text12(1) <> MsgText(601) Then
      strSql = strSql & " and a0k40 <= '" & Text12(1) & "'"
      strSQL1 = strSQL1 & " and a0k40 <= '" & Text12(1) & "'"
      strSQL2 = strSQL2 & " and a0k40 <= '" & Text12(1) & "'"
   End If
   If InStr(UCase(strSQL1), "A0K40") > 0 Then
      strSql = strSql & " and a0k40 is not null"
      strSQL1 = strSQL1 & " and a0k40 is not null"
      strSQL2 = strSQL2 & " anda0k40 is not null"
   End If
   'end 2023/11/13
   
   If adoadodc1.State = adStateOpen Then
      adoadodc1.Close
   End If
   adoadodc1.CursorLocation = adUseClient
   
   'Modify by Morgan 2008/10/24 銷帳金額要抓a1u0資料,否則金額會不對且資料會重複
   'Add by Morgan 2006/4/13  不必抓收據的資料 --瑞婷,秀玲
   
   'Modify by Morgan 2011/10/27 考慮拆收據情形改先寫暫存
   'If Text6 = "1" Then
      'Modify by Morgan 2011/8/23 改從 0j0 抓 cp
      'strUnion = "select a0k03, a0k04, a0k11, a0k02 as RDate, a0k01 as RNo, a0k20||' '||ST02 As a0k20, (cp01||cp02||cp03||cp04) as a0j02, cp09 as a0j01, nvl(a0j21, a0k23) as a0j21, a0j20, nvl(cp16, 0) as RAmount, nvl(cp75, 0) as EAmount, nvl(cp76, 0)  as DAmount, nvl(cp77, 0) as CAmount, nvl(cp78, 0) as BAmount, nvl(cp79, 0) as NAmount, '' as AccNo from acc0k0, acc0j0, caseprogress, Staff, Customer where a0k01 = a0j13 (+) and cp09(+) = a0j01  and (a0k09 is null or a0k09 = 0) and a0k01 is not null And substr(A0K03,1,8)=CU01(+) And substr(A0K03,9,1)=CU02(+) and nvl(cp79, 0) <> 0 " & strSql & strType
      'strUnion = strUnion & " union select a0k03, a0k04, a0k11, a0l02 as RDate, a0l01 as RNo, a0k20||' '||ST02 As a0k20, (cp01||cp02||cp03||cp04) as a0j02, cp09 as a0j01, nvl(a0j21, a0k23) as a0j21, a0j20, 0 as RAmount, (nvl(a1u04, 0) + nvl(a1u05, 0)) as EAmount, a1u06 as DAmount, 0 as CAmount, 0 as BAmount, 0 as NAmount, a1p22 as AccNo from acc0k0, acc1u0, acc0l0, acc0j0, caseprogress, (select distinct a1p04, a1p22 from acc1p0 where a1p01 = '1' and a1p02 = 'A') new, Staff, Customer where a0k01 = a1u02 (+) and a1u01 = a0l01 (+) and a1u02 = a0j13 (+) and a1u03 = a0j01 (+) and a1u03 = cp09 (+) and a0l01 = a1p04 (+) and (a0k09 is null or a0k09 = 0) and a0l01 is not null And substr(A0K03,1,8)=CU01(+) And substr(A0K03,9,1)=CU02(+) and nvl(cp79, 0) <> 0" & strSQL1 & strType
      'If Text6 = "1" Or Text6 = "3" Or Text6 = "" Then
      '   strUnion = strUnion & " union select a0k03, a0k04, a0k11, a0s03 as RDate, a0s01 as RNo, a0k20||' '||ST02 As a0k20, a0j02, a0j01, a0j21, a0j20, 0 as Ramount, 0 as EAmount, 0 as Damount, nvl(a1u07, 0)+nvl(a1u09, 0) as CAmount, nvl(a1u08, 0)+nvl(a1u10, 0) as BAmount, 0 as NAmount, a1p22 as AccNo from acc0s0, acc1u0, acc0k0, (select distinct a1p04, a1p22 from acc1p0 where a1p01 = '1' and a1p02 = 'Z') new, acc0j0, caseprogress, Staff, Customer where a0s02 = a0k01 (+) and a0s01 = a1p04 (+) and a1u01(+)=a0s01 and a0j01(+)=a1u03 and a0j01 = cp09 and (a0k09 is null or a0k09 = 0) and a0s01 is not null And substr(A0K03,1,8)=CU01(+) And substr(A0K03,9,1)=CU02(+) and (nvl(cp77, 0) <> 0 or nvl(cp78, 0) <> 0) and nvl(cp79, 0) <> 0" & strSQL2
      'End If
   'ElseIf Text6 = "2" Then
      'strUnion = "select a0k03, a0k04, a0k11, a0l02 as RDate, a0l01 as RNo, a0k20||' '||ST02 As a0k20, (cp01||cp02||cp03||cp04) as a0j02, cp09 as a0j01, nvl(a0j21, a0k23) as a0j21, a0j20, 0 as RAmount, (nvl(a1u04, 0) + nvl(a1u05, 0)) as EAmount, a1u06 as DAmount, 0 as CAmount, 0 as BAmount, 0 as NAmount, a1p22 as AccNo from acc0k0, acc1u0, acc0l0, acc0j0, caseprogress, (select distinct a1p04, a1p22 from acc1p0 where a1p01 = '1' and a1p02 = 'A') new, Staff, Customer where a0k01 = a1u02 (+) and a1u01 = a0l01 (+) and a1u02 = a0j13 (+) and a1u03 = a0j01 (+) and a1u03 = cp09 (+) and a0l01 = a1p04 (+) and (a0k09 is null or a0k09 = 0) and a0l01 is not null And substr(A0K03,1,8)=CU01(+) And substr(A0K03,9,1)=CU02(+) " & strSQL1 & strType
   'Else
      ''Modify by Morgan 2011/8/23 改從 0j0 抓 cp
      'strUnion = "select a0k03, a0k04, a0k11, a0k02 as RDate, a0k01 as RNo, a0k20||' '||ST02 As a0k20, (cp01||cp02||cp03||cp04) as a0j02, cp09 as a0j01, nvl(a0j21, a0k23) as a0j21, a0j20, nvl(cp16, 0) as RAmount, nvl(cp75, 0) as EAmount, nvl(cp76, 0)  as DAmount, nvl(cp77, 0) as CAmount, nvl(cp78, 0) as BAmount, nvl(cp79, 0) as NAmount, '' as AccNo from acc0k0, caseprogress, acc0j0, Staff, Customer where a0k01 = a0j13 (+) and cp09(+) = a0j01  and (a0k09 is null or a0k09 = 0) and a0k01 is not null And substr(A0K03,1,8)=CU01(+) And substr(A0K03,9,1)=CU02(+) " & strSql & strType
      'strUnion = strUnion & " union select a0k03, a0k04, a0k11, a0l02 as RDate, a0l01 as RNo, a0k20||' '||ST02 As a0k20, (cp01||cp02||cp03||cp04) as a0j02, cp09 as a0j01, nvl(a0j21, a0k23) as a0j21, a0j20, 0 as RAmount, (nvl(a1u04, 0) + nvl(a1u05, 0)) as EAmount, a1u06 as DAmount, 0 as CAmount, 0 as BAmount, 0 as NAmount, a1p22 as AccNo from acc0k0, acc1u0, acc0l0, acc0j0, caseprogress, (select distinct a1p04, a1p22 from acc1p0 where a1p01 = '1' and a1p02 = 'A') new, Staff, Customer where a0k01 = a1u02 (+) and a1u01 = a0l01 (+) and a1u02 = a0j13 (+) and a1u03 = a0j01 (+) and a1u03 = cp09 (+) and a0l01 = a1p04 (+) and (a0k09 is null or a0k09 = 0) and a0l01 is not null And substr(A0K03,1,8)=CU01(+) And substr(A0K03,9,1)=CU02(+) " & strSQL1 & strType
      'If Text6 = "1" Or Text6 = "3" Or Text6 = "" Then
      '   strUnion = strUnion & " union select a0k03, a0k04, a0k11, a0s03 as RDate, a0s01 as RNo, a0k20||' '||ST02 As a0k20, a0j02, a0j01, a0j21, a0j20, 0 as Ramount, 0 as EAmount, 0 as Damount, nvl(a1u07, 0)+nvl(a1u09, 0) as CAmount, nvl(a1u08, 0)+nvl(a1u10, 0) as BAmount, 0 as NAmount, a1p22 as AccNo from acc0s0,acc1u0, acc0k0, (select distinct a1p04, a1p22 from acc1p0 where a1p01 = '1' and a1p02 = 'Z') new, acc0j0, caseprogress, Staff, Customer where a0s02 = a0k01 (+) and a0s01 = a1p04 (+) and a1u01(+)=a0s01 and a0j01(+)=a1u03 and a0j01 = cp09 and (a0k09 is null or a0k09 = 0) and a0s01 is not null And substr(A0K03,1,8)=CU01(+) And substr(A0K03,9,1)=CU02(+) and (nvl(cp77, 0) <> 0 or nvl(cp78, 0) <> 0)" & strSQL2
      'End If
   'End If
   adoTaie.Execute "delete ACCTMP08 where T05='" & Me.Name & "' and T14='" & strUserNum & "'"
   If Text6 = "2" Then
      'Modified by Lydia 2016/04/11 舊收據無a0j01
      strUnion = "select a0k01, NVL(a0j01,' ') a0j01,'" & Me.Name & "',a0l01,'" & strUserNum & "' T14 from acc0k0, acc0j0, caseprogress, acc1u0, acc0l0, Customer, Staff"
      strUnion = strUnion & " where (a0k09 is null or a0k09 = 0) and a0j13(+)=a0k01 and cp09(+)=a0j01"
      strUnion = strUnion & " and a1u02(+)=a0j13 and a1u03(+)=a0j01 and substr(a1u01,1,1)='F' and a0l01(+)=a1u01"
      strUnion = strUnion & " And CU01(+)=substr(A0K03,1,8) And substr(A0K03,9,1)=CU02(+)" & strSQL1 & strType
      'end 2016/04/11
      adoTaie.Execute "insert into ACCTMP08(T01,T02,T05,T06,T14) " & strUnion, intI
      
      'Added by Lydia 2025/07/25 排除未達客戶付款週期之應收帳款
      If ChkBillDate.Value = 1 Then
         Call PUB_ProcAcctmp08(Me.Name, strUserNum)
      End If
      'end 2025/07/25
      
      '更新收款資料傳票號
      'Modified by Morgan 2016/3/4 收款可能會有兩家公司別不能只抓1公司
      strSql = "update ACCTMP08 set T07=(select max(a1p22||'('||a1p01||')')||decode(max(a1p01||a1p22),min(a1p01||a1p22),'',','||min(a1p22||'('||a1p01||')')) from acc1p0 where a1p02 = 'A' and a1p04=T06)" & _
         " where T05='" & Me.Name & "' and T14='" & strUserNum & "' and substr(T06,1,1)='F'"
      adoTaie.Execute strSql, intI
            
      'Modified by Morgan 2011/12/27 取消 a0j20,a0j21
      'Modify By Sindy 2016/6/8 + decode(cp27,null,'x','o') NotSend
      'Modify By Sindy 2021/8/2 ST02 => ST02||getlos04(a0j01,1) as ST02
      strUnion = " select decode(cp27,null,'x','o') NotSend,a0k03, a0k04, a0k11, a0l02 RDate, a0l01 RNo,'' a0k32,'' axc01, ST02||getlos04(a0j01,1) as ST02, a0k20"
      strUnion = strUnion & ", a0j02, a0j01, na03, getcp10desc(cp01,cp10,a0j04) cp10N, 0 RAmount"
      strUnion = strUnion & ", nvl(a1u04, 0) + nvl(a1u05, 0) EAmount, nvl(a1u06,0) DAmount, 0 CAmount, 0 BAmount, 0 NAmount, T07 AccNo,nvl(CU04,nvl(rtrim(CU05||' '||CU88||' '||CU89||' '||CU90),CU06)) CusName"
      strUnion = strUnion & " from ACCTMP08, acc0j0, acc0k0, Staff, acc1u0, acc0l0,caseprogress,nation,customer"
      strUnion = strUnion & " where T05='" & Me.Name & "' and T14='" & strUserNum & "' and substr(T06,1,1)='F' and a0j13(+)=T01 and a0j01(+)=T02"
      strUnion = strUnion & " and a0k01(+)=T01 and st01(+)=a0k20 and a1u01(+)=T06 and a1u02(+)=T01 and a1u03(+)=T02 and a0l01(+)=T06 and cp09(+)=a0j01 and na01(+)=a0j04 and cu01(+)=substr(a0k03,1,8) and cu02(+)=substr(a0k03,9)"
      'end 2011/10/27
    Else
      '先抓進度檔有未收的寫暫存然後再過濾掉拆收據已收齊的
      'Modified by Lydia 2016/04/11 舊收據無a0j01
      strUnion = "select a0k01, NVL(a0j01,' ') a0j01,'" & Me.Name & "',a0k01,'" & strUserNum & "' T14 from acc0k0, acc0j0, caseprogress, Customer, Staff"
      strUnion = strUnion & " where (a0k09 is null or a0k09 = 0) and a0j13(+)=a0k01 and cp09(+)=a0j01"
      strUnion = strUnion & " And CU01(+)=substr(A0K03,1,8) And substr(A0K03,9,1)=CU02(+) " & strSql & strType
            
      strUnion = strUnion & " union select a0k01, NVL(a0j01,' ') a0j01,'" & Me.Name & "',a0l01,'" & strUserNum & "' T14 from acc0k0, acc0j0, caseprogress, acc1u0, acc0l0, Customer, Staff"
      strUnion = strUnion & " where (a0k09 is null or a0k09 = 0) and nvl(a0k17,0)+nvl(a0k18,0)>0 and a0j13(+)=a0k01"
      strUnion = strUnion & " and cp09(+)=a0j01 and a1u02(+)=a0j13 and a1u03(+)=a0j01 and substr(a1u01,1,1)='F'"
      strUnion = strUnion & " and a0l01(+)=a1u01 And CU01(+)=substr(A0K03,1,8) And substr(A0K03,9,1)=CU02(+)" & strSQL1 & strType
      
      If Text6 = "3" Then 'Added by Morgan 2016/4/12 查往來才抓銷退
         strUnion = strUnion & " union select a0k01, NVL(a0j01,' ') a0j01,'" & Me.Name & "',a0s01,'" & strUserNum & "' T14 from acc0k0, acc0j0, caseprogress, acc1u0, acc0s0, Customer, Staff"
         strUnion = strUnion & " where (a0k09 is null or a0k09 = 0) and a0k10 is not null and a0j13(+)=a0k01"
         strUnion = strUnion & " and cp09(+)=a0j01 and (nvl(cp77, 0) <> 0 or nvl(cp78, 0) <> 0)"
         strUnion = strUnion & " and a1u02(+)=a0j13 and a1u03(+)=a0j01 and substr(a1u01,1,1)='I' and a0s01(+)=a1u01  and a0s02(+)=a1u02"
         strUnion = strUnion & " And CU01(+)=substr(A0K03,1,8) And substr(A0K03,9,1)=CU02(+)" & strSQL2
      End If 'Added by Morgan 2016/4/12
      
      'end 2016/04/11
      adoTaie.Execute "insert into ACCTMP08(T01,T02,T05,T06,T14) " & strUnion, intI
            
      'Added by Lydia 2025/07/25 排除未達客戶付款週期之應收帳款
      If ChkBillDate.Value = 1 Then
         Call PUB_ProcAcctmp08(Me.Name, strUserNum)
      End If
      'end 2025/07/25
      
      '更新收款資料傳票號
      'Modified by Morgan 2016/3/4 收款可能會有兩家公司別不能只抓1公司
      strSql = "update ACCTMP08 set T07=(select max(a1p22||'('||a1p01||')')||decode(max(a1p01||a1p22),min(a1p01||a1p22),'',','||min(a1p22||'('||a1p01||')')) from acc1p0 where a1p02 = 'A' and a1p04=T06)" & _
         " where T05='" & Me.Name & "' and T14='" & strUserNum & "' and substr(T06,1,1)='F'"
      adoTaie.Execute strSql, intI
      
      '更新銷退資料傳票號
      strSql = "update ACCTMP08 set T07=(select max(a1p22) from acc1p0 where a1p01 = '1' and a1p02 = 'Z' and a1p04=T06)" & _
         " where T05='" & Me.Name & "' and T14='" & strUserNum & "' and substr(T06,1,1)='I'"
      adoTaie.Execute strSql, intI
      
      '更新金額欄位
      strSql = "update ACCTMP08 set T08=(select nvl(a0j09, 0)+nvl(a0j10, 0) from acc0j0 where a0j13=T01 and a0j01=T02)" & _
         ",(T09,T10,T11,T12)=(select nvl(sum(a1u04),0)+nvl(sum(a1u05),0) T09,nvl(sum(a1u06),0) T10" & _
         ",nvl(sum(a1u07),0)+nvl(sum(a1u09),0) T11,nvl(sum(a1u08),0)+nvl(sum(a1u10),0) T12 " & _
         " from acc1u0 where a1u02=T01 and a1u03=T02) where T05='" & Me.Name & "' and T14='" & strUserNum & "'"
      adoTaie.Execute strSql, intI
      
      '去除拆收據已收齊的資料
      If Text6 = "1" Then
         strSql = "delete ACCTMP08 where T05='" & Me.Name & "' and T14='" & strUserNum & "' and T14='" & strUserNum & "' and T08-T09-T11+T12=0"
         adoTaie.Execute strSql, intI
      End If
      
      'Modified by Morgan 2011/12/27 取消 a0j20,a0j21
      
      '收據
      'Modify By Sindy 2016/6/8 + decode(cp27,null,'x','o') NotSend
      'Modified by Lydia 2025/05/23 增加# 已開INVOICE => decode(a0k40,null,a0k32,'#')
      'Modify By Sindy 2021/8/2 ST02 => ST02||getlos04(a0j01,1) as ST02
      strUnion = "select decode(cp27,null,'x','o') NotSend,a0k03, a0k04, a0k11, a0k02 RDate, a0k01 RNo,decode(a0k40,null,a0k32,'#') a0k32,axc01, ST02||getlos04(a0j01,1) as ST02, a0k20, a0j02, a0j01"
      strUnion = strUnion & ", na03, getcp10desc(cp01,cp10,a0j04) cp10N, T08 RAmount, T09 EAmount, T10 DAmount, T11 CAmount"
      strUnion = strUnion & ", T12 BAmount, T08-T09-T11+T12 NAmount,'' as AccNo ,nvl(CU04,nvl(rtrim(CU05||' '||CU88||' '||CU89||' '||CU90),CU06)) CusName"
      'Modified by Lydia 2023/11/13 +a0k40
      strUnion = strUnion & " ,a0k40 from ACCTMP08, acc0j0,acc0k0, Staff,caseprogress,nation,acc431,customer"
      strUnion = strUnion & " where T05='" & Me.Name & "' and T14='" & strUserNum & "' and T06=T01 and a0j13(+)=T01 and a0j01(+)=T02"
      strUnion = strUnion & " and a0k01(+)=T01 and st01(+)=a0k20 and cp09(+)=a0j01 and na01(+)=a0j04 and axc02(+)=T01 and cu01(+)=substr(a0k03,1,8) and cu02(+)=substr(a0k03,9)"
      
      '收款
      'Modify By Sindy 2016/6/8 + decode(cp27,null,'x','o') NotSend
      'Modify By Sindy 2021/8/2 ST02 => ST02||getlos04(a0j01,1) as ST02
      strUnion = strUnion & " union select decode(cp27,null,'x','o') NotSend,a0k03, a0k04, a0k11, a0l02 RDate, a0l01 RNo,'' a0k32,'' axc01, ST02||getlos04(a0j01,1) as ST02, a0k20"
      strUnion = strUnion & ", a0j02, a0j01, na03, getcp10desc(cp01,cp10,a0j04) cp10N, 0 RAmount"
      strUnion = strUnion & ", nvl(a1u04, 0) + nvl(a1u05, 0) EAmount, nvl(a1u06,0) DAmount, 0 CAmount, 0 BAmount, 0 NAmount, T07 AccNo,nvl(CU04,nvl(rtrim(CU05||' '||CU88||' '||CU89||' '||CU90),CU06)) CusName"
      'Modified by Lydia 2023/11/13 +a0k40
      strUnion = strUnion & " ,a0k40 from ACCTMP08, acc0j0, acc0k0, Staff, acc1u0, acc0l0,caseprogress,nation,customer"
      strUnion = strUnion & " where T05='" & Me.Name & "' and T14='" & strUserNum & "' and substr(T06,1,1)='F' and a0j13(+)=T01 and a0j01(+)=T02"
      strUnion = strUnion & " and a0k01(+)=T01 and st01(+)=a0k20 and a1u01(+)=T06 and a1u02(+)=T01 and a1u03(+)=T02 and a0l01(+)=T06 and cp09(+)=a0j01 and na01(+)=a0j04 and cu01(+)=substr(a0k03,1,8) and cu02(+)=substr(a0k03,9)"
      
      If Text6 = "3" Then 'Added by Morgan 2016/4/12 查往來才抓銷退
         '消退
         'Modify By Sindy 2016/6/8 + decode(cp27,null,'x','o') NotSend
         'Modify By Sindy 2021/8/2 ST02 => ST02||getlos04(a0j01,1) as ST02
         strUnion = strUnion & " union select decode(cp27,null,'x','o') NotSend,a0k03, a0k04, a0k11, a0s03 RDate, a0s01 RNo,'' a0k32,'' axc01, ST02||getlos04(a0j01,1) as ST02, a0k20"
         strUnion = strUnion & ", a0j02, a0j01, na03, getcp10desc(cp01,cp10,a0j04) cp10N, 0 as Ramount, 0 as EAmount, 0 as Damount"
         strUnion = strUnion & ", nvl(a1u07, 0)+nvl(a1u09, 0) as CAmount, nvl(a1u08, 0)+nvl(a1u10, 0) as BAmount, 0 as NAmount, T07 AccNo,nvl(CU04,nvl(rtrim(CU05||' '||CU88||' '||CU89||' '||CU90),CU06)) CusName"
         'Modified by Lydia 2023/11/13 +a0k40
         strUnion = strUnion & " ,a0k40 from ACCTMP08, acc0j0, acc0k0, acc1u0, acc0s0, Staff,caseprogress,nation,customer"
         strUnion = strUnion & " where T05='" & Me.Name & "' and T14='" & strUserNum & "' and substr(T06,1,1)='I' and a0j13(+)=T01 and a0j01(+)=T02"
         strUnion = strUnion & " and a0k01(+)=T01 and st01(+)=a0k20 and a1u01(+)=T06 and a1u02(+)=T01 and a1u03(+)=T02 and a0S01(+)=T06 and cp09(+)=a0j01 and na01(+)=a0j04 and cu01(+)=substr(a0k03,1,8) and cu02(+)=substr(a0k03,9)"
      End If 'Added by Morgan 2016/4/12
      'end 2011/10/26
   End If
   'Modify By Sindy 2016/6/7 + 排序選項
   If Text3 = "1" Then '1.抬頭＋日期
      strUnion = strUnion & " order by a0k04 asc, RDate asc, RNo asc"
   ElseIf Text3 = "2" Then '2.收據編號
      strUnion = strUnion & " order by RNo asc,a0j01 asc"
   Else
      strUnion = strUnion & " order by Rdate asc, a0k11 asc, RNo asc"
   End If
   '2016/6/7 END
   adoadodc1.Open strUnion, adoTaie, adOpenStatic, adLockReadOnly
   Adodc1.Recordset.Requery
   Calculate 'Add by Morgan 2007/9/27
   If Adodc1.Recordset.RecordCount = 0 Then
      Adodc1.Recordset.Close
      MsgBox MsgText(28), , MsgText(5)
'      Text1.SetFocus
'      TextInverse Text1
      Exit Sub
   End If
   
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'Add by Morgan 2007/9/27
Private Sub Calculate()
   Dim dblAmt(5) As Double
   Set RsTemp = Adodc1.Recordset.Clone
   With RsTemp
   If .RecordCount > 0 Then
      Do While Not .EOF
         Select Case Left("" & .Fields("RNo"), 1)
            Case "E"
               dblAmt(0) = dblAmt(0) + Val("" & .Fields("RAmount"))
               dblAmt(5) = dblAmt(5) + Val("" & .Fields("NAmount"))
            Case "F"
               dblAmt(1) = dblAmt(1) + Val("" & .Fields("EAmount"))
               dblAmt(2) = dblAmt(2) + Val("" & .Fields("DAmount"))
            Case "I"
               dblAmt(3) = dblAmt(3) + Val("" & .Fields("CAmount"))
               dblAmt(4) = dblAmt(4) + Val("" & .Fields("BAmount"))
         End Select
         .MoveNext
      Loop
   End If
   End With
   For intI = 0 To 5
      txtAmt(intI) = Format(dblAmt(intI), "#,##0")
   Next
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Public Sub KeyDefine(KeyCode As Integer)
Dim bolShowMsg As Boolean 'Add By Sindy 2016/6/7
   
   Select Case KeyCode
      Case vbKeyF12
         'Add By Sindy 2016/6/7 主要條件三選一,必須至少要輸入一項
         If Text8 <> "" And Text9 <> "" Then
            If Len(Text9) < 6 Then Text9 = Right("00000" & Text9, 6)
            If Text10 = "" Then Text10 = "0"
            If Text11 = "" Then Text11 = "00"
         End If
         'Modify by Amy 2023/08/14 +部門
         If ((Text1 = "X" And Text2 = "X") Or (Text1 = MsgText(601) And Text2 = MsgText(601))) And _
            Text7 = MsgText(601) And _
            Text8 = MsgText(601) And Text9 = MsgText(601) And Text10 = MsgText(601) And Text11 = MsgText(601) And _
            TxtDeptS = MsgText(601) And TxtDeptE = MsgText(601) Then
            MsgBox "客戶編號 或 智權人員 或 本所案號 或 部門 四選一，至少必須要輸入一項！", , MsgText(5)
            Exit Sub
         End If
         '2016/6/7 END
         
         If FormCheck() Then
            Screen.MousePointer = vbHourglass
            StatusView MsgText(192)
            AdodcRefresh
            StatusView MsgText(601)
            Screen.MousePointer = vbDefault
            Exit Sub
         Else
            If bolShowMsg = True Then 'Modify By Sindy 2016/6/7 +if
               MsgBox MsgText(181), , MsgText(5)
            End If
         End If
   End Select
   KeyEnter KeyCode
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
   If Text1 <> MsgText(601) And Text1 <> "X" Then
      FormCheck = True
      Exit Function
   End If
   If Text2 <> MsgText(601) And Text2 <> "X" Then
      FormCheck = True
      Exit Function
   End If
   If MaskEdBox1.Text <> MsgText(29) Then
      FormCheck = True
      Exit Function
   End If
   If MaskEdBox2.Text <> MsgText(29) Then
      FormCheck = True
      Exit Function
   End If
   If Text4 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Text5 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Text6 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   'Add By Sindy 2016/6/7
   If Text7 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Text3 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   '2016/6/7 END
   
   FormCheck = False
End Function

'Add By Sindy 2016/6/7 智權人員
Private Sub Text7_Change()
   If Text7 = MsgText(601) Then
      Text13 = "" 'Add by Amy 2025/05/09 bug-輸過智權員編再拿掉,名稱仍在
      Exit Sub
   End If
   Text13 = StaffQuery(Text7)
End Sub
Private Sub Text7_GotFocus()
   TextInverse Text7
   CloseIme
End Sub
Private Sub Text7_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
'2016/6/7 END

'Add By Sindy 2016/6/7 本所案號
Private Sub Text8_GotFocus()
   TextInverse Text8
   CloseIme
End Sub
Private Sub Text8_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub Text9_GotFocus()
   TextInverse Text9
End Sub
Private Sub Text9_Validate(Cancel As Boolean)
   If Text9 = "" Then
      Exit Sub
   End If
   If Len(Text9) < 6 Then
      MsgBox MsgText(172), , MsgText(5)
      Cancel = True
      Text9.SetFocus
      Exit Sub
   End If
   Text10 = "0"
   Text11 = "00"
End Sub
Private Sub Text10_GotFocus()
   TextInverse Text10
End Sub
Private Sub Text11_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub Text11_GotFocus()
   TextInverse Text11
End Sub
'2016/6/7 END

'Add by Amy 2023/08/14
Private Sub TxtDeptE_GotFocus()
   TextInverse TxtDeptE
End Sub

Private Sub TxtDeptE_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub TxtDeptS_GotFocus()
   TextInverse TxtDeptS
End Sub

Private Sub TxtDeptS_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub TxtDeptS_LostFocus()
   If Trim(TxtDeptS) = MsgText(601) Then Exit Sub
   
   TxtDeptE = TxtDeptS
End Sub

'Added by Lydia 2023/11/13
Private Sub Text12_GotFocus(Index As Integer)
   TextInverse Text12(Index)
End Sub

Private Sub Text12_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
