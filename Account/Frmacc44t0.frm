VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc44t0 
   AutoRedraw      =   -1  'True
   Caption         =   "扣繳憑單查詢及列印"
   ClientHeight    =   5520
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   8920
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5520
   ScaleWidth      =   8920
   Begin VB.ListBox List1 
      Height          =   1120
      Left            =   4050
      MultiSelect     =   1  '簡易多重選取
      TabIndex        =   11
      Top             =   1080
      Width           =   4815
   End
   Begin VB.ComboBox cboCombNo 
      Height          =   300
      Left            =   1920
      TabIndex        =   15
      Top             =   1830
      Width           =   1845
   End
   Begin VB.CheckBox Check1 
      Caption         =   "列印備註"
      ForeColor       =   &H000000C0&
      Height          =   360
      Left            =   7890
      TabIndex        =   41
      Top             =   930
      Value           =   1  '核取
      Visible         =   0   'False
      Width           =   1010
   End
   Begin VB.CheckBox Check5 
      Caption         =   "寄測試信箱"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   6570
      TabIndex        =   40
      Top             =   2280
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.CommandButton cmdAcc21r0 
      Caption         =   "客戶/收據抬頭EMail"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5430
      Style           =   1  '圖片外觀
      TabIndex        =   21
      Top             =   2550
      Width           =   2685
   End
   Begin VB.CommandButton cmdMail 
      BackColor       =   &H00C0FFC0&
      Caption         =   "寄發Mail"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3840
      Style           =   1  '圖片外觀
      TabIndex        =   20
      Top             =   2550
      Width           =   1275
   End
   Begin VB.CommandButton cmdChkYear 
      BackColor       =   &H00C0FFC0&
      Caption         =   "扣繳確認年度"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4980
      Style           =   1  '圖片外觀
      TabIndex        =   17
      Top             =   2220
      Width           =   1545
   End
   Begin VB.TextBox txtA2802 
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
      Left            =   7200
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   8
      Top             =   720
      Width           =   612
   End
   Begin VB.ComboBox cmbPrinter 
      Height          =   300
      Left            =   1065
      Style           =   2  '單純下拉式
      TabIndex        =   16
      Top             =   2190
      Width           =   3870
   End
   Begin VB.TextBox txtLike 
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
      Left            =   7590
      MaxLength       =   1
      TabIndex        =   5
      Text            =   "N"
      Top             =   375
      Width           =   435
   End
   Begin VB.TextBox txtRecNo 
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
      Left            =   1065
      MaxLength       =   15
      TabIndex        =   6
      Top             =   720
      Width           =   1350
   End
   Begin VB.CommandButton cmdQuery 
      BackColor       =   &H00C0FFC0&
      Caption         =   "查詢(&Q)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   420
      Style           =   1  '圖片外觀
      TabIndex        =   18
      Top             =   2550
      Width           =   1395
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C0FFC0&
      Caption         =   "列印(&P)"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2130
      Style           =   1  '圖片外觀
      TabIndex        =   19
      Top             =   2550
      Width           =   1395
   End
   Begin VB.TextBox txtComp 
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
      Index           =   2
      Left            =   3480
      MaxLength       =   1
      TabIndex        =   14
      Text            =   "L"
      Top             =   1470
      Width           =   390
   End
   Begin VB.TextBox txtComp 
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
      Index           =   1
      Left            =   2850
      MaxLength       =   1
      TabIndex        =   13
      Text            =   "1"
      Top             =   1470
      Width           =   345
   End
   Begin VB.TextBox txtCustNo 
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
      Index           =   0
      Left            =   1065
      MaxLength       =   9
      TabIndex        =   2
      Top             =   375
      Width           =   1350
   End
   Begin VB.TextBox txtCustNo 
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
      Index           =   1
      Left            =   2655
      MaxLength       =   9
      TabIndex        =   3
      Top             =   375
      Width           =   1350
   End
   Begin VB.CommandButton cmdLikeSearch 
      Caption         =   "搜尋"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7830
      TabIndex        =   1
      Top             =   60
      Width           =   675
   End
   Begin VB.TextBox txtYear 
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
      Left            =   1065
      MaxLength       =   3
      TabIndex        =   12
      Top             =   1470
      Width           =   612
   End
   Begin VB.TextBox txtTaxNo 
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
      MaxLength       =   15
      TabIndex        =   7
      Top             =   720
      Width           =   1572
   End
   Begin MSMask.MaskEdBox mebRecDate 
      Height          =   300
      Index           =   1
      Left            =   1065
      TabIndex        =   9
      Top             =   1080
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mebRecDate 
      Height          =   300
      Index           =   2
      Left            =   2640
      TabIndex        =   10
      Top             =   1080
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc44t0.frx":0000
      Height          =   2370
      Left            =   30
      TabIndex        =   38
      Top             =   2910
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   4180
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   14
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
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
      ColumnCount     =   20
      BeginProperty Column00 
         DataField       =   "a0k11"
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
      BeginProperty Column01 
         DataField       =   "a0l02"
         Caption         =   "收款日期"
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
      BeginProperty Column02 
         DataField       =   "a0k01"
         Caption         =   "收據號碼"
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
      BeginProperty Column03 
         DataField       =   "a0k02"
         Caption         =   "收據日期"
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
      BeginProperty Column04 
         DataField       =   "Fee0"
         Caption         =   "收款金額"
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
      BeginProperty Column05 
         DataField       =   "Fee1"
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
      BeginProperty Column06 
         DataField       =   "Fee2"
         Caption         =   "可扣稅額"
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
      BeginProperty Column07 
         DataField       =   "Fee3"
         Caption         =   "收款扣繳額"
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
         DataField       =   "Fee4"
         Caption         =   "補扣繳額"
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
         DataField       =   "Fee5"
         Caption         =   "未扣稅額"
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
      BeginProperty Column11 
         DataField       =   "na03"
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
      BeginProperty Column12 
         DataField       =   "Fee6"
         Caption         =   "已收扣單金額"
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
      BeginProperty Column13 
         DataField       =   "A0W16"
         Caption         =   "給付總額"
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
      BeginProperty Column14 
         DataField       =   "Fee7"
         Caption         =   "調整稅額"
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
      BeginProperty Column15 
         DataField       =   "a0w04"
         Caption         =   "扣單公司"
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
         DataField       =   "a0w02"
         Caption         =   "扣單編號"
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
         DataField       =   "a0w06"
         Caption         =   "扣單備註"
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
         DataField       =   "a1p12"
         Caption         =   "票期"
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
      BeginProperty Column19 
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
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   689.953
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   929.764
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   950.173
         EndProperty
         BeginProperty Column03 
            Locked          =   -1  'True
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   929.764
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   810.142
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            Object.Visible         =   0   'False
            ColumnWidth     =   810.142
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            ColumnWidth     =   810.142
         EndProperty
         BeginProperty Column08 
            Alignment       =   1
            ColumnWidth     =   929.764
         EndProperty
         BeginProperty Column09 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   870.236
         EndProperty
         BeginProperty Column10 
            Locked          =   -1  'True
            ColumnWidth     =   1090.205
         EndProperty
         BeginProperty Column11 
            Locked          =   -1  'True
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column12 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1289.764
         EndProperty
         BeginProperty Column13 
         EndProperty
         BeginProperty Column14 
            Alignment       =   1
            ColumnWidth     =   929.764
         EndProperty
         BeginProperty Column15 
            Object.Visible         =   -1  'True
            ColumnWidth     =   870.236
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column17 
         EndProperty
         BeginProperty Column18 
         EndProperty
         BeginProperty Column19 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
      Left            =   90
      Top             =   2640
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
      Caption         =   "Adodc2"
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
   Begin MSForms.ComboBox cboTitle 
      Height          =   315
      Left            =   1065
      TabIndex        =   0
      Top             =   45
      Width           =   6690
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "11800;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtSalesNo 
      Height          =   300
      Left            =   5010
      TabIndex        =   4
      Top             =   380
      Width           =   1575
      VariousPropertyBits=   679495707
      MaxLength       =   5
      Size            =   "2778;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label17 
      Caption         =   "境外公司"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   8010
      TabIndex        =   42
      Top             =   2280
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.Label Label16 
      BackStyle       =   0  '透明
      Caption         =   "最近扣繳確認年度"
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
      Left            =   5415
      TabIndex        =   39
      Top             =   750
      Width           =   1875
   End
   Begin VB.Label Label15 
      BackStyle       =   0  '透明
      Caption         =   "印表機"
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
      Left            =   60
      TabIndex        =   37
      Top             =   2220
      Width           =   750
   End
   Begin VB.Label Label14 
      BackStyle       =   0  '透明
      Caption         =   "(Y/N)"
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
      Left            =   8130
      TabIndex        =   36
      Top             =   410
      Width           =   585
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '透明
      Caption         =   "相似查詢"
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
      Left            =   6660
      TabIndex        =   35
      Top             =   405
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '透明
      Caption         =   "扣單編號"
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
      Left            =   2625
      TabIndex        =   34
      Top             =   750
      Width           =   975
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '透明
      Caption         =   "收據編號"
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
      Left            =   60
      TabIndex        =   33
      Top             =   750
      Width           =   975
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3300
      TabIndex        =   32
      Top             =   1440
      Width           =   120
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "公司別"
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
      Left            =   2040
      TabIndex        =   31
      Top             =   1500
      Width           =   840
   End
   Begin VB.Label lblSales 
      BackStyle       =   0  '透明
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
      Left            =   7065
      TabIndex        =   30
      Top             =   420
      Width           =   1290
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2475
      TabIndex        =   29
      Top             =   390
      Width           =   255
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "扣繳年度"
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
      Left            =   60
      TabIndex        =   28
      Top             =   1500
      Width           =   975
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "合併列印客戶代號"
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
      Left            =   60
      TabIndex        =   27
      Top             =   1845
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "客戶代號"
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
      Left            =   60
      TabIndex        =   26
      Top             =   405
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "收據抬頭"
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
      Left            =   60
      TabIndex        =   25
      Top             =   75
      Width           =   975
   End
   Begin VB.Label Label3 
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
      Left            =   4065
      TabIndex        =   24
      Top             =   405
      Width           =   900
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "收款日期"
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
      Left            =   60
      TabIndex        =   23
      Top             =   1110
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2490
      TabIndex        =   22
      Top             =   1110
      Width           =   255
   End
End
Attribute VB_Name = "Frmacc44t0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2022/1/20 Form2.0已修改
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
Option Explicit

'公用
Dim rs44t0 As New ADODB.Recordset
'頁次
Dim iPage As Integer
'欄位座標
Dim PLeft(0 To 20) As Integer
'Y座標
Dim intY As Integer
'暫存金額
Dim stMoneyTemp As String
'表頭欄位
Dim stSalesName As String, stTitle As String, stCustNo As String, stCustName As String
'預設印表機
'Dim m_DefaultPrinter As String, m_Prn As Printer
'列高
Private Const RowPix As Integer = 300
'邊界調整
Private Const intDefault As Integer = 500
'一個9號字元的點數
Private Const BytePix As Integer = 90
'金錢格式
Private Const DDollar As String = "###,###,###,##0"
'群組
Dim stLstGroup1 As String, stCurGroup1 As String, stLstGroup2 As String, stCurGroup2 As String, stCompName As String
'小計
Dim arrSubTot1(0 To 7) As String, arrSubTot2(0 To 7) As String, arrGrdTot(0 To 7) As String, idx As Integer
Dim m_CU15 As String, m_StrTo As String 'Add By Sindy 2014/10/17
Dim strPrinter As String ', strTemplatePath As String 'Add By Sindy 2014/10/17
Dim m_CU16 As String, m_CU18 As String 'Add By Sindy 2014/11/5
Dim m_CU01 As String, m_CU02 As String, m_CU115 As String, m_CU20 As String, m_CU116 As String, m_CU117 As String, m_CU118 As String 'Add By Sindy 2015/1/16
Dim m_CU04 As String 'Add By Sindy 2015/2/17
Dim m_CU158 As String 'Add By Sindy 2015/12/10
'Added by Lydia 2016/12/19
Dim m_CallList As String '從會計師客戶資料查詢傳來
Dim m_CallListIdx As Integer
Dim m_CU172 As String 'Add By Sindy 2017/3/16
Dim m_bolPDF As Boolean 'Add By Sindy 2019/12/27
'Add By Sindy 2022/1/19
Dim xlsAnnuity As New Excel.Application
Dim wksAnnuity As New Worksheet
Dim m_intColumn As Integer
'2022/1/19 END
Dim m_strEmailAttch As String 'Add By Sindy 2020/4/20
Dim m_strMailKind As String 'Add By Sindy 2023/11/28
Dim stAccPerson As String, stTxtPerson As String 'Add by Amy 2024/05/17
Dim m_AccMail As String 'Add By Sindy 2024/10/8 扣繳信件信箱


'Added by Lydia 2016/12/19 從會計師客戶資料查詢傳來的字串,指定第幾筆資料
Public Function CallByA4901(ByVal pNo As String, Optional ByVal pIdx As Integer = 0) As Boolean
Dim tmpArr As Variant
Dim tmpStrA As String

   If pNo <> "" And m_CallList = "" Then m_CallList = pNo
   
   If m_CallList <> "" Then
      tmpArr = Split(m_CallList, ";")
      If Trim(tmpArr(pIdx)) <> "" Then
         Call FormClear
         If Left(Trim(tmpArr(pIdx)), 1) = "1" Then
            '直接設定客戶代號 、合併客戶代號
            txtCustNo(0) = Mid(Trim(tmpArr(pIdx)), 3, 6) & "000"
            txtCustNo(1) = Mid(Trim(tmpArr(pIdx)), 3, 6) & "ZZZ"
            'Modify By Sindy 2018/1/16
            'txtCombNo = Mid(Trim(tmpArr(pIdx)), 3, 9)
            cboCombNo = Mid(Trim(tmpArr(pIdx)), 3, 9)
            '2018/1/16 END
            cboTitle.Text = Mid(Trim(tmpArr(pIdx)), InStr(Trim(tmpArr(pIdx)), "|") + 1)
         ElseIf Left(Trim(tmpArr(pIdx)), 1) = "2" Then
            cboTitle.Text = Mid(Trim(tmpArr(pIdx)), 3)
            Call cmdLikeSearch_Click
            '若收據抬頭有搜尋到客戶代號,直接代入第1筆
            If cboTitle.ListCount > 1 Then
               cboTitle.ListIndex = 1
               '若代入客戶代號無關係企業,直接代入合併客戶代號
               If txtCustNo(0) <> "" And txtCustNo(1) = "" Then
                  txtCustNo(1) = txtCustNo(0)
                  'Modify By Sindy 2018/1/16
                  'txtCombNo = txtCustNo(0)
                  cboCombNo = txtCustNo(0)
                  '2018/1/16 END
               End If
            End If
         End If
         m_CallListIdx = pIdx
         Call cmdQuery_Click '直接呼叫查詢
         CallByA4901 = True
         Exit Function
      End If
   End If
   
   CallByA4901 = False
   
End Function

' iMode=1(查詢--預設) 2(列印)
Private Function Process(Optional ByRef iMode As Integer = 1) As Boolean
   Dim stSQL As String, stAcc0k0Con As String, stAcc0w0Con As String, stAcc0l0Con As String
   Dim stCustNoCol As String
   Dim stAcc1k0Con As String, stAcc1v0Con As String, stAcc0y0Con As String 'Add By Sindy 2015/11/16
   Dim stAcc1k0ConCP As String
   
On Error GoTo flgErr
   
   ClearQueryLog (Me.Name) 'Add By Sindy 2020/12/30 清除查詢印表記錄檔欄位
'   Label17.Caption = "" 'Add By Sindy 2013/12/6
   List1.Clear 'Add By Sindy 2014/10/17
   'Add By Sindy 2013/12/30
   If txtComp(1) = "J" Then
      MsgBox "公司別不可為 J", vbExclamation
      txtComp(1).SetFocus
      Exit Function
   End If
   If txtComp(2) = "J" Then
      MsgBox "公司別不可為 J", vbExclamation
      txtComp(2).SetFocus
      Exit Function
   End If
   '2013/12/30 END
   
   '收據抬頭
   If cboTitle = "" Then
      MsgBox "收據抬頭不可空白!!!", vbExclamation
      cboTitle.SetFocus
      Exit Function
   ElseIf cboTitle <> "" Then
      If txtLike = "N" Then
         'Modify By Sindy 2024/9/30 + ChgSQL
         stAcc0k0Con = stAcc0k0Con & " and a0k04='" & ChgSQL(cboTitle) & "'"
         stAcc1k0Con = stAcc1k0Con & " and a1k35='" & ChgSQL(cboTitle) & "'" 'Add By Sindy 2015/11/16
      Else
         '2011/10/20 MODIFY BY SONIA E10023515
         'stAcc0k0Con = stAcc0k0Con & " and instr(a0k04,'" & cboTitle & "')>0"
         'Modify By Sindy 2024/9/30 + ChgSQL
         stAcc0k0Con = stAcc0k0Con & " and instr(UPPER(a0k04),UPPER('" & ChgSQL(cboTitle) & "'))>0"
         stAcc1k0Con = stAcc1k0Con & " and instr(UPPER(a1k35),UPPER('" & ChgSQL(cboTitle) & "'))>0" 'Add By Sindy 2015/11/16
      End If
   End If
   
   '合併列印客戶代號
'   If iMode = 1 And txtCombNo = "" Then
'      MsgBox "查詢時合併列印客戶代號不可空白!!!", vbExclamation
'      txtCombNo.SetFocus
'      txtCombNo_GotFocus
'      Exit Function
'   ElseIf txtCombNo <> "" Then
'      stCustNoCol = "'" & txtCombNo & "'"
'   End If
   'Modify By Sindy 2018/1/16
   If iMode = 1 And cboCombNo = "" Then
      MsgBox "查詢時合併列印客戶代號不可空白!!!", vbExclamation
      cboCombNo.SetFocus
      Exit Function
   ElseIf cboCombNo <> "" Then
      stCustNoCol = "'" & cboCombNo & "'"
   End If
   '2018/1/16 END
   
   'Add By Sindy 2016/11/23 有會計備註及財務E-Mail按鈕變顏色
   'Add By Sindy 2017/9/27 於收據抬頭EMail按鈕上方增加
   '                       若客戶或收據抬頭為境外公司則以紅色顯示 '境外公司'
   cmdAcc21r0.BackColor = &H8000000F
'   cmdAcc11p0.BackColor = &H8000000F
   Label17.Visible = False '境外公司
   If Trim(cboTitle) <> "" Then
      pub_QL05 = pub_QL05 & ";收據抬頭:" & cboTitle 'Add By Sindy 2020/12/30
      strExc(0) = "select cu01,cu02,cu158,cu159,cu115 from customer where '" & ChgSQL(cboTitle.Text) & "'=cu04" & _
                  " or '" & ChgSQL(cboTitle.Text) & "'=cu05||' '||cu88||' '||cu89||' '||cu90" & _
                  " or '" & ChgSQL(cboTitle.Text) & "'=cu06"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If "" & RsTemp.Fields("cu159") <> "" Or "" & RsTemp.Fields("cu115") <> "" Then
            cmdAcc21r0.BackColor = &HC0FFC0
         End If
         If "" & RsTemp.Fields("cu158") = "Y" Then
            Label17.Visible = True
         End If
      Else
         '收據抬頭
         strExc(0) = "select a4201,a4216,a4217,a4218 from acc420 where a4201='" & ChgSQL(cboTitle.Text) & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If "" & RsTemp.Fields("a4217") <> "" Or "" & RsTemp.Fields("a4218") <> "" Then
               cmdAcc21r0.BackColor = &HC0FFC0
            End If
            If "" & RsTemp.Fields("a4216") = "Y" Then
               Label17.Visible = True
            End If
         End If
      End If
   End If
'   If txtCustNo(0) <> "" Then
'      strExc(0) = "select cu01 from customer" & _
'                  " where cu01='" & Left(txtCustNo(0), 8) & "' and cu02='" & Mid(txtCustNo(0), 9, 1) & "'" & _
'                  " and (cu159 is not null or cu115 is not null)"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         cmdAcc21r0.BackColor = &HC0FFC0
'      End If
'   End If
'   If Trim(cboTitle) <> "" Then
'      strExc(0) = "select a4201 from acc420" & _
'                  " where a4201='" & cboTitle & "'" & _
'                  " and (a4217 is not null or a4218 is not null)"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         cmdAcc11p0.BackColor = &HC0FFC0
'      End If
'   End If
   '2016/11/8 END
   
   '客戶編號
   If txtCustNo(0) <> "" Then
      pub_QL05 = pub_QL05 & ";客戶代號:" & txtCustNo(0) 'Add By Sindy 2020/12/30
      stAcc0k0Con = stAcc0k0Con & " and a0k03>='" & txtCustNo(0) & "'"
   End If
   If txtCustNo(1) <> "" Then
      pub_QL05 = pub_QL05 & ";客戶代號:" & txtCustNo(1) 'Add By Sindy 2020/12/30
      stAcc0k0Con = stAcc0k0Con & " and a0k03<='" & txtCustNo(1) & "'"
   End If
   'Add By Sindy 2015/11/16
   If txtCustNo(0) <> "" And txtCustNo(1) <> "" Then
      stAcc1k0Con = stAcc1k0Con & " and ((a1k03>='" & txtCustNo(0) & "' and a1k03<='" & txtCustNo(1) & "') or (a1k27>='" & txtCustNo(0) & "' and a1k27<='" & txtCustNo(1) & "') or (a1k28>='" & txtCustNo(0) & "' and a1k28<='" & txtCustNo(1) & "'))"
   ElseIf txtCustNo(0) <> "" Then
      stAcc1k0Con = stAcc1k0Con & " and (a1k03>='" & txtCustNo(0) & "' or a1k27>='" & txtCustNo(0) & "' or a1k28>='" & txtCustNo(0) & "')"
   ElseIf txtCustNo(1) <> "" Then
      stAcc1k0Con = stAcc1k0Con & " and (a1k03<='" & txtCustNo(1) & "' or a1k27<='" & txtCustNo(1) & "' or a1k28<='" & txtCustNo(1) & "')"
   End If
   '2015/11/16 END
   '智權人員
   If txtSalesNo <> "" Then
      pub_QL05 = pub_QL05 & ";智權人員:" & txtSalesNo 'Add By Sindy 2020/12/30
      stAcc0k0Con = stAcc0k0Con & " and a0k20='" & txtSalesNo & "'"
      stAcc1k0ConCP = stAcc1k0ConCP & " and cp13='" & txtSalesNo & "'" 'Add By Sindy 2015/11/16
   End If
   '收據編號
   If txtRecNo <> "" Then
      pub_QL05 = pub_QL05 & ";收據編號:" & txtRecNo 'Add By Sindy 2020/12/30
      stAcc0k0Con = stAcc0k0Con & " and a0k01='" & txtRecNo & "'"
      stAcc1k0Con = stAcc1k0Con & " and a1k01='" & txtRecNo & "'" 'Add By Sindy 2015/11/16
   End If
   '扣單編號
   If txtTaxNo <> "" Then
      pub_QL05 = pub_QL05 & ";扣單編號:" & txtTaxNo 'Add By Sindy 2020/12/30
      stAcc0w0Con = stAcc0w0Con & " and a0w02='" & txtTaxNo & "'"
      stAcc1v0Con = stAcc1v0Con & " and a0w02='" & txtTaxNo & "'" 'Add By Sindy 2015/11/16
   End If
   '收款日期
   If mebRecDate(1).Text <> MsgText(601) And mebRecDate(1).Text <> MsgText(29) Then
      pub_QL05 = pub_QL05 & ";收款日期:" & mebRecDate(1) 'Add By Sindy 2020/12/30
      stAcc0l0Con = stAcc0l0Con & " and a0l02 >= " & Val(FCDate(mebRecDate(1).Text)) & ""
      stAcc0y0Con = stAcc0y0Con & " and a0y02 >= " & Val(FCDate(mebRecDate(1).Text)) & "" 'Add By Sindy 2015/11/16
   End If
   If mebRecDate(2).Text <> MsgText(601) And mebRecDate(2).Text <> MsgText(29) Then
      pub_QL05 = pub_QL05 & ";收款日期:" & mebRecDate(2) 'Add By Sindy 2020/12/30
      stAcc0l0Con = stAcc0l0Con & " and a0l02 <= " & Val(FCDate(mebRecDate(2).Text)) & ""
      stAcc0y0Con = stAcc0y0Con & " and a0y02 <= " & Val(FCDate(mebRecDate(2).Text)) & "" 'Add By Sindy 2015/11/16
   End If
   '扣繳年度
   If txtYear <> "" Then
      pub_QL05 = pub_QL05 & ";扣繳年度:" & txtYear 'Add By Sindy 2020/12/30
      stAcc0k0Con = stAcc0k0Con & " and a0k16=" & txtYear
      stAcc1v0Con = stAcc1v0Con & " and a1v09=" & txtYear 'Add By Sindy 2015/11/16
   End If
   '公司別
   stAcc0k0Con = stAcc0k0Con & " and a0k11<>'J'" 'Add By Sindy 2013/12/30
   If txtComp(1) <> "" Then
      pub_QL05 = pub_QL05 & ";公司別:" & txtComp(1) 'Add By Sindy 2020/12/30
      stAcc0k0Con = stAcc0k0Con & " and a0k11>='" & txtComp(1) & "'"
      stAcc1v0Con = stAcc1v0Con & " and a1v03>='" & txtComp(1) & "'" 'Add By Sindy 2015/11/16
   End If
   If txtComp(2) <> "" Then
      pub_QL05 = pub_QL05 & ";公司別:" & txtComp(2) 'Add By Sindy 2020/12/30
      stAcc0k0Con = stAcc0k0Con & " and a0k11<='" & txtComp(2) & "'"
      stAcc1v0Con = stAcc1v0Con & " and a1v03<='" & txtComp(2) & "'" 'Add By Sindy 2015/11/16
   End If
   
   'Add By Sindy 2020/12/30
   If cboCombNo <> "" Then pub_QL05 = pub_QL05 & ";合併列印客戶代號:" & cboCombNo
   If txtA2802 <> "" Then pub_QL05 = pub_QL05 & ";最近扣繳確認年度:" & txtA2802
   '2020/12/30 END
   
'cancel by sonia 2016/3/28 瑞婷要求,否則同一抬頭有跨所資料時則分所無法看到完整資料
'   'Add by Morgan 2005/3/28
'   '若非北所員工, 只能查詢該所資料,暫先以收據智權人員判斷，再確認
'   If pub_strUserOffice <> "1" Then
'       stAcc0w0Con = stAcc0w0Con & " And ST01=A0K20 And ''||ST06='" & pub_strUserOffice & "'"
'       stAcc1k0ConCP = stAcc1k0ConCP & " And ST01=cp13 And ''||ST06='" & pub_strUserOffice & "'" 'Add By Sindy 2015/11/16
'   End If
'   '2005/3/28
   
   'Modify by Morgan 2005/1/21 注意1.收款金額含銷退 2.服務費=10*(已扣+未扣)
   'Fee0:收款金額,Fee1:服務費,Fee2:可扣稅額,Fee3:收款扣繳額,Fee4:補扣繳額,Fee5:未扣繳額,Fee6:已扣繳額,Fee7:調整稅額
   'Modified by Morgan 2011/11/9 考慮拆收據情形
   'stSQL = "select a0k20,st02,a0k04," & stCustNoCol & " as a0k03,a0k11,a0802,a0l02,a0k01,a0k02,substrb(a0j20,1,12) a0j20,substrb(a0j21,1,12) a0j21" & _
      ",Fee0,10*nvl(a1v04,0) Fee1,nvl(a1v04,0) Fee2,decode(a1v18,'1',nvl(a1v06,0),0) Fee3,decode(a1v18,'1',0,nvl(a1v06,0)) Fee4" & _
      ",nvl(a1v07,0) Fee5, decode(a1v15,null,null, a1v06) Fee6,nvl(a1v10,0) Fee7" & _
      ",a0w04,a0w02,a0w06,cp09,cu04,a1p12" & _
      " from ( select a0m02,max(a0l02) a0l02,min(a1p12) a1p12 from acc0k0,acc0m0, acc0l0,acc1p0" & _
      " where (a0k09 is null or a0k09 = 0)" & stAcc0k0Con & _
      " and (a0m03 is null or substr(a0m03, 1, 1) = 'E') and a0m02=a0k01 and a0l01=a0m01" & stAcc0l0Con & _
      " and a1p04(+)=a0l01 and a1p05(+)='113001' group by a0m02 ) X" & _
      ",acc0k0,acc080,staff,caseprogress" & _
      ",( select a1u03,max(a0j20) a0j20, max(a0j21) a0j21,sum(nvl(a1u04,0)+nvl(a1u05,0)-nvl(a1u08,0)-nvl(a1u10,0)) Fee0" & _
      ", sum(nvl(a1u04, 0)-nvl(a1u08, 0)) as Fee1" & _
      " From acc0k0, acc0j0, acc1u0 where (a0k09 is null or a0k09 = 0) " & stAcc0k0Con & " and a0j13=a0k01 and a1u03=a0j01 group by a1u03) Y" & _
      ",acc1v0,acc0w0,customer" & _
      " Where a0k01 = a0m02 and a0801(+)=a0k11 and st01(+)=a0k20 and cp60(+)=a0k01 and a1u03(+)=cp09" & _
      " and a1v01(+)=cp09 and a0w02(+)=a1v15 and cu01(+)=substr(a0k03,1,8) and cu02(+)=substr(a0k03,9,1)" & stAcc0w0Con & _
      " and a1v04>0 order by a0k04, a0k03, a0k11, a0l02, a0k01, cp10"
   'Modified by Morgan 2011/12/21 欄位+a0k33,a0j22,a0j25,排序加 a0j25
   'Modified by Morgan 2011/12/26 取消a0j03 改抓 cp10
   'Modified by Morgan 2011/12/27 取消 a0j20,a0j21
   'Modify By Sindy 2013/12/6 +cu115:EMail財務信箱,cu01,cu02
   'Modify By Sindy 2014/10/17 +,cu20,cu116,cu117,cu118,cu16,cu18,DECODE(CU15,'0','台端','1','貴公司','貴單位') cu15
   'Modify By Sindy 2015/12/9 cu16 ==> nvl(cu16,cu22) 沒電話則抓手機
   'modify by sonia 2016/1/25 Fee1由10*nvl(a1v04,0)改為decode(Y.a0j07,'Y',Fee0,Fee1),同時也加入Y.a0j07
   'modify by sonia 2023/5/10 +a0w16給付總額
   stSQL = "select a0k20,st02,a0k04," & IIf(stCustNoCol = "", "a0k03", stCustNoCol) & " as a0k03,a0k11,a0802,a0l02,a0k01,a0k02" & _
      ",substrb(getcp10desc(cp01,cp10,a0j04),1,12) cp10N,substrb(na03,1,12) na03" & _
      ",Fee0,decode(Y.a0j07,'Y',Fee0,Fee1) Fee1,round(nvl(a1v04,0),0) Fee2,round(decode(a1v18,'1',nvl(a1v06,0),0),0) Fee3,round(decode(a1v18,'1',0,nvl(a1v06,0)),0) Fee4" & _
      ",nvl(a1v07,0) Fee5, decode(a1v15,null,null, a1v06) Fee6,nvl(a1v10,0) Fee7" & _
      ",a0w04,a0w02,a0w06,a0j01 cp09,a1p12,a0k33,a0j22,a0j25,cp10,a0w16" & _
      " from (select a0m02,max(a0l02) a0l02,min(a1p12) a1p12 from acc0k0,acc0m0,acc0l0,acc1p0" & _
      " where (a0k09 is null or a0k09 = 0)" & stAcc0k0Con & _
      " and (a0m03 is null or substr(a0m03, 1, 1) = 'E') and a0m02(+)=a0k01 and a0l01(+)=a0m01" & stAcc0l0Con & _
      " and a1p04(+)=a0l01 and a1p05(+)='113001' group by a0m02 ) X" & _
      ",acc0k0,acc080,staff,acc0j0" & _
      ",( select a1u03,a1u02,sum(nvl(a1u04,0)+nvl(a1u05,0)-nvl(a1u08,0)-nvl(a1u10,0)) Fee0" & _
      ", sum(nvl(a1u04, 0)-nvl(a1u08, 0)) as Fee1,a0j07 From acc0k0, acc0j0, acc1u0" & _
      " where (a0k09 is null or a0k09 = 0) " & stAcc0k0Con & " and a0j13(+)=a0k01 and a1u03(+)=a0j01 and a1u02(+)=a0j13" & _
      " group by a1u03,a1u02,a0j07) Y,acc1v0,acc0w0,caseprogress,nation" & _
      " Where a0k01(+) = a0m02 and a0801(+)=a0k11 and st01(+)=a0k20 and a0j13(+)=a0k01 and a1u03(+)=a0j01 and a1u02(+)=a0j13" & _
      " and a1v01(+)=a0j01 and a1v02(+)=a0j13 and a0w02(+)=a1v15" & stAcc0w0Con & _
      " and a1v04>0 and cp09(+)=a0j01 and na01(+)=a0j04"
   'end 2011/11/9
   'Add By Sindy 2015/11/16 加入國外請款資料
   '10*nvl(a1v04,0) Fee1 ==> nvl(a1k30,0) Fee1 服務費
   'Modify By Sindy 2017/1/13 nvl(a1k30,0) Fee1 服務費 ==> nvl(a1k30,0)-nvl(a1k06,0)-nvl(a1k09,0) 請款金額要扣除折讓及規費才叫做服務費
   'Modify By Sindy 2021/1/5 nvl(a1k30,0)-nvl(a1k06,0)-nvl(a1k09,0) => 要再串回基本檔抓申請國家:
   '台灣案: 收款金額 - 規費 = 服務費
   '非台灣案: 收款金額 = 服務費
   'modify by sonia 2023/5/10 +a0w16給付總額
   stSQL = stSQL & " union " & _
      "select cp13 a0k20,st02,a1k35 a0k04," & IIf(stCustNoCol = "", "a1k28", stCustNoCol) & " as a0k03,a1v03 a0k11,a0802,a0y02 a0l02,a1k01 a0k01,a1k02 a0k02" & _
      ",substrb(GETCP10DESCCaseNO(cp10,cp01,cp02,cp03,cp04),1,12) cp10N,substrb(GETNA03DESCCaseNO(cp01,cp02,cp03,cp04),1,12) na03" & _
      ",nvl(a1k30,0) Fee0,decode(nvl(tm10,nvl(pa09,nvl(sp09,nvl(lc15,'000')))),'000',nvl(a1k30,0)-nvl(a1k06,0)-nvl(a1k09,0),nvl(a1k30,0)) Fee1" & _
      ",round(nvl(a1v04,0),0) Fee2,round(decode(a1v18,'1',nvl(a1v06,0),0),0) Fee3,round(decode(a1v18,'1',0,nvl(a1v06,0)),0) Fee4" & _
      ",nvl(a1v07,0) Fee5,decode(a1v15,null,null, a1v06) Fee6,nvl(a1v10,0) Fee7" & _
      ",a0w04,a0w02,a0w06,cp09,a1p12,'' a0k33,'' a0j22,1 a0j25,cp10,a0w16" & _
      " from (select a0z02,max(a0y02) a0y02,min(a1p12) a1p12 From acc1k0,acc0z0,acc0y0,Acc1p0" & _
      " where (a1k12 is null or a1k12 = 0)" & stAcc1k0Con & _
      " and a0z02(+)=a1k01 and a0y01(+)=a0z01" & stAcc0y0Con & _
      " and a1p04(+)=a0y01 and a1p05(+)='113001' group by a0z02) X" & _
      ",caseprogress,acc1k0,acc1v0,acc080,staff,acc0w0,patent,trademark,lawcase,servicepractice,hirecase" & _
      " Where a1k01(+) = a0z02 and (a1k12 Is Null Or a1k12 = 0)" & stAcc1k0Con & stAcc1k0ConCP & _
      " and a1k01=a1v02 and cp09(+)=a1v01" & stAcc1v0Con & _
      " and a0801(+)=a1v03" & _
      " and st01(+)=cp13" & _
      " and a0w02(+)=a1v15" & _
      " and a1v04>0"
   stSQL = stSQL & " and a1k13=pa01(+) and a1k14=pa02(+) and a1k15=pa03(+) and a1k16=pa04(+)" & _
                   " and a1k13=tm01(+) and a1k14=tm02(+) and a1k15=tm03(+) and a1k16=tm04(+)" & _
                   " and a1k13=lc01(+) and a1k14=lc02(+) and a1k15=lc03(+) and a1k16=lc04(+)" & _
                   " and a1k13=sp01(+) and a1k14=sp02(+) and a1k15=sp03(+) and a1k16=sp04(+)" & _
                   " and a1k13=hc01(+) and a1k14=hc02(+) and a1k15=hc03(+) and a1k16=hc04(+)"
'   stSQL = "select a0k20,st02,a0k04," & IIf(stCustNoCol = "", "a0k03", stCustNoCol) & " as a0k03,a0k11,a0802,a0l02,a0k01,a0k02" & _
'      ",substrb(getcp10desc(cp01,cp10,a0j04),1,12) cp10N,substrb(na03,1,12) na03" & _
'      ",Fee0,10*nvl(a1v04,0) Fee1,nvl(a1v04,0) Fee2,decode(a1v18,'1',nvl(a1v06,0),0) Fee3,decode(a1v18,'1',0,nvl(a1v06,0)) Fee4" & _
'      ",nvl(a1v07,0) Fee5, decode(a1v15,null,null, a1v06) Fee6,nvl(a1v10,0) Fee7" & _
'      ",a0w04,a0w02,a0w06,a0j01 cp09,cu04,a1p12,a0k33,a0j22,a0j25,cu115,cu01,cu02,cu20,cu116,cu117,cu118,nvl(cu16,cu22) cu16,cu18,DECODE(CU15,'0','台端','1','貴公司','貴單位') cu15,cp10" & _
'      " from (select a0m02,max(a0l02) a0l02,min(a1p12) a1p12 from acc0k0,acc0m0,acc0l0,acc1p0" & _
'      " where (a0k09 is null or a0k09 = 0)" & stAcc0k0Con & _
'      " and (a0m03 is null or substr(a0m03, 1, 1) = 'E') and a0m02(+)=a0k01 and a0l01(+)=a0m01" & stAcc0l0Con & _
'      " and a1p04(+)=a0l01 and a1p05(+)='113001' group by a0m02 ) X" & _
'      ",acc0k0,acc080,staff,acc0j0" & _
'      ",( select a1u03,a1u02,sum(nvl(a1u04,0)+nvl(a1u05,0)-nvl(a1u08,0)-nvl(a1u10,0)) Fee0" & _
'      ", sum(nvl(a1u04, 0)-nvl(a1u08, 0)) as Fee1 From acc0k0, acc0j0, acc1u0" & _
'      " where (a0k09 is null or a0k09 = 0) " & stAcc0k0Con & " and a0j13(+)=a0k01 and a1u03(+)=a0j01 and a1u02(+)=a0j13" & _
'      " group by a1u03,a1u02) Y,acc1v0,acc0w0,customer,caseprogress,nation" & _
'      " Where a0k01(+) = a0m02 and a0801(+)=a0k11 and st01(+)=a0k20 and a0j13(+)=a0k01 and a1u03(+)=a0j01 and a1u02(+)=a0j13" & _
'      " and a1v01(+)=a0j01 and a1v02(+)=a0j13 and a0w02(+)=a1v15 and cu01(+)=substr(a0k03,1,8) and cu02(+)=substr(a0k03,9,1)" & stAcc0w0Con & _
'      " and a1v04>0 and cp09(+)=a0j01 and na01(+)=a0j04"
'   'end 2011/11/9
'   'Add By Sindy 2015/11/16 加入國外請款資料
'   '客戶檔
'   stSQL = stSQL & " union " & _
'      "select cp13 a0k20,st02,a1k35 a0k04," & IIf(stCustNoCol = "", "a1k28", stCustNoCol) & " as a0k03,a1v03 a0k11,a0802,a0y02 a0l02,a1k01 a0k01,a1k02 a0k02" & _
'      ",substrb(GETCP10DESCCaseNO(cp10,cp01,cp02,cp03,cp04),1,12) cp10N,substrb(GETNA03DESCCaseNO(cp01,cp02,cp03,cp04),1,12) na03" & _
'      ",nvl(a1k30,0) Fee0,10*nvl(a1v04,0) Fee1,nvl(a1v04,0) Fee2,decode(a1v18,'1',nvl(a1v06,0),0) Fee3,decode(a1v18,'1',0,nvl(a1v06,0)) Fee4" & _
'      ",nvl(a1v07,0) Fee5,decode(a1v15,null,null, a1v06) Fee6,nvl(a1v10,0) Fee7" & _
'      ",a0w04,a0w02,a0w06,cp09,cu04,a1p12,'' a0k33,'' a0j22,1 a0j25,cu115,cu01,cu02,cu20,cu116,cu117,cu118,nvl(cu16,cu22) cu16,cu18,DECODE(CU15,'0','台端','1','貴公司','貴單位') cu15,cp10" & _
'      " from (select a0z02,max(a0y02) a0y02,min(a1p12) a1p12 From acc1k0,acc0z0,acc0y0,Acc1p0" & _
'      " where (a1k12 is null or a1k12 = 0)" & stAcc1k0Con & _
'      " and a0z02(+)=a1k01 and a0y01(+)=a0z01" & stAcc0y0Con & _
'      " and a1p04(+)=a0y01 and a1p05(+)='113001' group by a0z02) X" & _
'      ",caseprogress,acc1k0,acc1v0,acc080,staff,acc0w0,customer" & _
'      " Where a1k01(+) = a0z02 and (a1k12 Is Null Or a1k12 = 0)" & stAcc1k0Con & stAcc1k0ConCP & _
'      " and a1k01=a1v02 and cp09(+)=a1v01" & stAcc1v0Con & _
'      " and a0801(+)=a1v03" & _
'      " and st01(+)=cp13" & _
'      " and a0w02(+)=a1v15" & _
'      " and cu01(+)=substr(a1k28,1,8) and cu02(+)=substr(a1k28,9,1)" & _
'      " and a1v04>0"
'   '代理人檔
'   stSQL = stSQL & " union " & _
'      "select cp13 a0k20,st02,a1k35 a0k04," & IIf(stCustNoCol = "", "a1k28", stCustNoCol) & " as a0k03,a1v03 a0k11,a0802,a0y02 a0l02,a1k01 a0k01,a1k02 a0k02" & _
'      ",substrb(GETCP10DESCCaseNO(cp10,cp01,cp02,cp03,cp04),1,12) cp10N,substrb(GETNA03DESCCaseNO(cp01,cp02,cp03,cp04),1,12) na03" & _
'      ",nvl(a1k30,0) Fee0,10*nvl(a1v04,0) Fee1,nvl(a1v04,0) Fee2,decode(a1v18,'1',nvl(a1v06,0),0) Fee3,decode(a1v18,'1',0,nvl(a1v06,0)) Fee4" & _
'      ",nvl(a1v07,0) Fee5,decode(a1v15,null,null, a1v06) Fee6,nvl(a1v10,0) Fee7" & _
'      ",a0w04,a0w02,a0w06,cp09,fa04,a1p12,'' a0k33,'' a0j22,1 a0j25,fa79,fa01,fa02,fa16,fa80,fa81,fa82,nvl(fa12,fa13) cu16,fa14,'貴公司' cu15,cp10" & _
'      " from (select a0z02,max(a0y02) a0y02,min(a1p12) a1p12 From acc1k0,acc0z0,acc0y0,Acc1p0" & _
'      " where (a1k12 is null or a1k12 = 0)" & stAcc1k0Con & _
'      " and a0z02(+)=a1k01 and a0y01(+)=a0z01" & stAcc0y0Con & _
'      " and a1p04(+)=a0y01 and a1p05(+)='113001' group by a0z02) X" & _
'      ",caseprogress,acc1k0,acc1v0,acc080,staff,acc0w0,fagent" & _
'      " Where a1k01(+) = a0z02 and (a1k12 Is Null Or a1k12 = 0)" & stAcc1k0Con & stAcc1k0ConCP & _
'      " and a1k01=a1v02 and cp09(+)=a1v01" & stAcc1v0Con & _
'      " and a0801(+)=a1v03" & _
'      " and st01(+)=cp13" & _
'      " and a0w02(+)=a1v15" & _
'      " and fa01(+)=substr(a1k28,1,8) and fa02(+)=substr(a1k28,9,1)" & _
'      " and a1v04>0"
   '2015/11/16 END
   'modify by sonia 2024/12/2 排序A0K11改為desc
   stSQL = stSQL & " order by a0k04, a0k03, a0k11 desc, a0l02, a0k01, a0j25, cp10"
   With rs44t0
      If .State = adStateOpen Then .Close
      .CursorLocation = adUseClient
      .Open stSQL, adoTaie, adOpenForwardOnly, adLockReadOnly
      If .RecordCount > 0 Then
         InsertQueryLog .RecordCount 'Add By Sindy 2020/12/30
         Process = True
         'Add By Sindy 2017/6/19 檢查收據抬頭是否存在
         Call PUB_ChkTitleNmExist(cboTitle)
         '2017/6/19 END
      Else
         InsertQueryLog (0) 'Add By Sindy 2020/12/30
         MsgBox MsgText(28), , MsgText(5)
      End If
   End With
   
flgErr:
   
   If Err.Number <> 0 Then MsgBox Err.Description
   'Resume
End Function

'Add By Sindy 2018/1/16
Private Sub cboCombNo_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Add By Sindy 2014/9/26
Private Sub cmdAcc21r0_Click()
Dim rsA As New ADODB.Recordset
   
   If Trim(cboTitle.Text) <> "" Then
      '客戶檔
      strSql = "select cu01,cu02 from customer where '" & ChgSQL(cboTitle.Text) & "'=cu04" & _
              " or '" & ChgSQL(cboTitle.Text) & "'=cu05||' '||cu88||' '||cu89||' '||cu90" & _
              " or '" & ChgSQL(cboTitle.Text) & "'=cu06"
      If rsA.State = adStateOpen Then rsA.Close
      rsA.CursorLocation = adUseClient
      rsA.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
      If rsA.RecordCount > 0 Then
         'tool1_enabled
         Me.MousePointer = vbHourglass
         'MenuDisabled
         strUserLevel = Me.Name
         'Modify By Sindy 2020/3/10
         If Trim(cboCombNo) <> "" Then
            Frmacc21r0.txtKey = Trim(cboCombNo)
            Frmacc21r0.Command1_Click
         ElseIf Trim(txtCustNo(0)) <> "" Then
         '2020/3/10 END
            Frmacc21r0.txtKey = Trim(txtCustNo(0))
            Frmacc21r0.Command1_Click
         End If
         Frmacc21r0.Show
         Me.Hide
         Me.MousePointer = vbDefault
      Else
         '收據抬頭
         strSql = "select a4201 from acc420 where a4201='" & ChgSQL(cboTitle.Text) & "'"
         If rsA.State = adStateOpen Then rsA.Close
         rsA.CursorLocation = adUseClient
         rsA.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
         If rsA.RecordCount > 0 Then
            'tool1_enabled
            Me.MousePointer = vbHourglass
            'MenuDisabled
            strUserLevel = Me.Name
            strCompanyNo = Trim(cboTitle)
            If Trim(cboTitle) <> "" Then
               strSaveConfirm = ""
               Frmacc11p0.textA4201 = Trim(cboTitle)
               Frmacc11p0.bolCallMe = True 'Add By Sindy 2015/9/14
               Frmacc11p0.Command3_Click
            End If
            Frmacc11p0.Show
            Me.Hide
            Me.MousePointer = vbDefault
         End If
      End If
   Else
      MsgBox "無符合的收據抬頭資料！", vbCritical
   End If
   
   Set rsA = Nothing
End Sub

'Add By Sindy 2014/10/17
Private Sub cmdMail_Click()
Dim ii As Integer
Dim ArrStr() As String
   
   m_StrTo = ""
   For ii = 0 To List1.ListCount - 1
      If List1.Selected(ii) = True Then
         ArrStr = Split(List1.List(ii), "：")
         If UBound(ArrStr) > 0 Then
            m_StrTo = m_StrTo & ";" & ArrStr(1)
         End If
      End If
   Next ii
   If m_StrTo <> "" Or Check5.Value = 1 Then
      m_StrTo = Mid(m_StrTo, 2)
      If MsgBox("確定要發E-Mail嗎？" & vbCrLf & "（" & m_StrTo & "）", vbExclamation + vbYesNo + vbDefaultButton1, "重要訊息！") = vbNo Then
         Exit Sub
      End If
   Else
      MsgBox "無點選寄發的E-Mail！", vbInformation, "提醒！"
      'MsgBox "請點選欲寄發的E-Mail！", vbExclamation
      'Exit Sub
   End If
   
   'Modify By Sindy 2023/11/28 若未輸入收款日期條件則仍維持原來催扣繳憑單，不必選擇。
   If mebRecDate(1).Text <> MsgText(601) And mebRecDate(1).Text <> MsgText(29) Then
      strExc(0) = InputBox("請輸入，寄出內文選項：A.催扣繳憑單 B.催繳款書，請擇一!!", "重要訊息！")
      If Trim(strExc(0)) = "" Then
         Exit Sub
      Else
         m_strMailKind = UCase(Trim(strExc(0)))
      End If
   End If
   '2023/11/28 END
   
   Call CallcmdPrint(True)
End Sub

'Modify By Sindy 2014/10/17 +CallcmdPrint
Private Sub cmdPrint_Click()
   Call CallcmdPrint(False)
End Sub
Private Function CallcmdPrint(bolPDF As Boolean)
   Screen.MousePointer = vbHourglass
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
   'If Process(2) Then
      If rs44t0.RecordCount > 0 Then
         If FormPrint(bolPDF) = True Then
            'Add By Sindy 2014/10/17
            If bolPDF = False Then
            '2014/10/17 END
               MsgBox "列印完成！"
            End If
         End If
      End If
   'End If
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
   Screen.MousePointer = vbDefault
End Function

'Modify By Sindy 2014/10/17 +bolPDF
Private Function FormPrint(bolPDF As Boolean) As Boolean
Dim PrinterIndex As Integer
Dim i As Integer
Dim strFilePathName As String
Dim strSubject As String, strContent As String 'Add By Sindy 2014/10/20
Dim strEmp As String, strEMP_Tel As String 'Add By Sindy 2014/10/20
Dim strEmpST01 As String, strEmpST22 As String 'Add By Sindy 2023/11/28
   
On Error GoTo flgErr
   'Modify by Amy 2024/05/17 財務2個特殊設定拆成3個,設定可能多人,抓第一個人 原:Pub_GetSpecMan("財務處總帳人員")
   strExc(0) = "select st01,st02,ed01,st22" & _
               " from staff,ExtensionData" & _
               " where ST01=ED02(+)" & _
               " and st01='" & stTxtPerson & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      'Add By Sindy 2023/11/28
      strEmpST01 = RsTemp.Fields("st01")
      strEmpST22 = RsTemp.Fields("st22") 'F=女 M=男
      '2023/11/28 END
      strEmp = RsTemp.Fields("st02")
      strEMP_Tel = "" & RsTemp.Fields("ed01")
   End If
   'Modify By Sindy 2024/12/16 分所抓操作人員
'   If strUserNum = Pub_GetSpecMan("出納人員-中所") Or _
'      strUserNum = Pub_GetSpecMan("出納人員-南所") Or _
'      strUserNum = Pub_GetSpecMan("出納人員-高所") Then
   If PUB_GetST06(strUserNum) <> "1" Then
   '2024/12/16 END
      strExc(0) = "select st01,st02,ed01,st22" & _
                  " from staff,ExtensionData" & _
                  " where ST01=ED02(+)" & _
                  " and st01='" & strUserNum & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         'Add By Sindy 2023/11/28
         strEmpST01 = RsTemp.Fields("st01")
         strEmpST22 = RsTemp.Fields("st22") 'F=女 M=男
         '2023/11/28 END
         strEmp = RsTemp.Fields("st02")
         strEMP_Tel = "" & RsTemp.Fields("ed01")
      End If
   End If
   
   m_strEmailAttch = ""
   m_bolPDF = bolPDF 'Add By Sindy 2019/12/27
   'Add By Sindy 2014/10/17
   If bolPDF = True Then
      '檢查是否有安裝PDFCreator
      PrinterIndex = -1
      For i = 0 To Printers.Count - 1
         If UCase(Printers(i).DeviceName) = UCase$("PDFCreator") Then
            PrinterIndex = i
            Exit For
         End If
      Next i
      If PrinterIndex < 0 Then
         MsgBox "請通知電腦中心安裝PDFCreator !!!"
         Exit Function
      End If
      'PUB_RestorePrinter Printers(PrinterIndex).DeviceName 'Removed by Morgan 2018/2/27 不必設定否則結束時會還原錯誤的印表機
      'frmPDF.Show
      'strFilePathName = Trim(txtCustNo(0)) & "-" & txtYear & "年-客戶扣繳明細核對表"
      'frmPDF.StartProcess strExcelPath, strFilePathName
'   Else
'   '2014/10/17 END
'      For Each m_Prn In Printers
'         If m_Prn.DeviceName = cmbPrinter.Text Then
'            Set Printer = m_Prn
'            Exit For
'         End If
'      Next
   End If
   
   '設定使用者所選擇的印表機成預設印表機
   PUB_SetOsDefaultPrinter cmbPrinter 'Add By Sindy 2022/1/27
   PUB_RestorePrinter cmbPrinter
   
   'Modify By Sindy 2022/1/19 改Call Excel產生報表
'   If strSrvDate(1) >= Form20上線日 Then
      m_intColumn = 0
      With rs44t0
         .MoveFirst
         stLstGroup1 = "": stCurGroup1 = ""
         stLstGroup2 = "": stCurGroup2 = ""
         stCompName = ""
         Erase arrSubTot1()
         Erase arrSubTot2()
         Erase arrGrdTot()
         
         Call GetTitleCustData("" & .Fields("a0k04"), "" & .Fields("a0k03"), "", m_CU01, m_CU02, _
                               m_CU15, m_CU115, m_CU20, m_CU116, m_CU117, m_CU118, m_CU16, _
                               m_CU18, m_CU04, m_CU158) 'Modify By Sindy 2015/1/16 改以收據抬頭抓客戶資料
         Do While Not .EOF
            stCurGroup1 = "" & .Fields("a0k11") '公司別
            stCurGroup2 = "" & .Fields("a0k04") & .Fields("a0k03") '收據抬頭+客戶編號
   '         m_stCU16 = "" & .Fields("cu16").Value 'Add By Sindy 2014/11/5
   '         m_stCU18 = "" & .Fields("cu18").Value 'Add By Sindy 2014/11/5
            If (stCurGroup2 <> stLstGroup2) And m_intColumn > 0 Then '收據抬頭+客戶編號
               Call PrintSubTot(arrSubTot1(), 1, stCompName)
               Call PrintSubTot(arrSubTot2(), 2)
               Erase arrSubTot1()
               Erase arrSubTot2()
'               Printer.NewPage
               stSalesName = "" & .Fields("st02").Value
               stTitle = "" & .Fields("a0k04").Value
               stCustNo = "" & .Fields("a0k03").Value
               'stCustName = "" & .Fields("cu04").Value
               stCustName = m_CU04
               stCompName = .Fields("a0802").Value
               Call PrintHead(stSalesName, stTitle, stCustNo, stCustName)
            ElseIf (stCurGroup1 <> stLstGroup1) And m_intColumn > 0 Then '公司別
               Call PrintSubTot(arrSubTot1(), 1, stCompName)
               Erase arrSubTot1()
               stCompName = .Fields("a0802").Value
               'Add By Sindy 2020/4/20
               Call PrintFoot(stLstGroup1)
'               Printer.EndDoc
               m_intColumn = 0
               '2020/4/20 END
'            ElseIf m_intColumn > 0 Then
'               xlsAnnuity.Range("A" & m_intColumn & ":" & "P" & m_intColumn).Select
'               With xlsAnnuity.Selection.Borders(xlEdgeTop)
'                   .LineStyle = xlContinuous
'                   .ColorIndex = xlAutomatic
'                   .tintandshade = 0
'                   .Weight = xlThin
'               End With
'               With xlsAnnuity.Selection.Borders(xlEdgeBottom)
'                   .LineStyle = xlContinuous
'                   .ColorIndex = xlAutomatic
'                   .tintandshade = 0
'                   .Weight = xlThin
'               End With
            End If
            
            If m_intColumn = 0 Then
'               'Add By Sindy 2020/4/20
'               If bolPDF = True Then
'                  frmPDF.Show
'                  strFilePathName = Trim(txtCustNo(0)) & "-" & txtYear & "年-客戶扣繳明細核對表(" & IIf(stCurGroup1 = "1", "商標", IIf(stCurGroup1 = "2", "智慧所", "法律所")) & ")"
'                  frmPDF.StartProcess strExcelPath, strFilePathName
'                  If m_strEmailAttch <> "" Then m_strEmailAttch = m_strEmailAttch & ";"
'                  m_strEmailAttch = m_strEmailAttch & strExcelPath & strFilePathName & ".pdf"
'               End If
'               '2020/4/20 END
               stSalesName = "" & .Fields("st02").Value
               stTitle = "" & .Fields("a0k04").Value
               stCustNo = "" & .Fields("a0k03").Value
               'stCustName = "" & .Fields("cu04").Value
               stCustName = m_CU04
               stCompName = .Fields("a0802").Value
               'Call PrintHead("" & .Fields("st02").Value, "" & .Fields("a0k04").Value, "" & .Fields("a0k03").Value, "" & .Fields("cu04").Value)
               Call PrintHead("" & .Fields("st02").Value, "" & .Fields("a0k04").Value, "" & .Fields("a0k03").Value, m_CU04)
            End If
            
            PrintData
            
            stLstGroup1 = stCurGroup1
            stLstGroup2 = stCurGroup2
            For idx = LBound(arrGrdTot) To UBound(arrGrdTot)
               If Not (idx = 6 And ("" & .Fields("Fee6")) = "") Then
                  arrSubTot1(idx) = Format(Val(arrSubTot1(idx)) + Val("" & .Fields("Fee" & idx)))
                  arrSubTot2(idx) = Format(Val(arrSubTot2(idx)) + Val("" & .Fields("Fee" & idx)))
                  'arrGrdTot(idx) = Format(Val(arrGrdTot(idx)) + Val("" & .Fields("Fee" & idx)))
               End If
            Next
            .MoveNext
         Loop
      End With
      Call PrintSubTot(arrSubTot1(), 1, stCompName)
      'Call PrintSubTot(arrSubTot2(), 2)
      
      'Modify By Sindy 2020/4/20 改成Call函數
      Call PrintFoot(stLstGroup1)
      
'      If bolPDF = False Then 'Added by Morgan 2018/2/27
         PUB_SetOsDefaultPrinter strPrinter 'Add By Sindy 2022/1/27
         PUB_RestorePrinter strPrinter '復原系統預設印表機
'      End If 'Added by Morgan 2018/2/27
      
'   Else
'   '2022/1/19 END
'
'      GetPleft
'      iPage = 0
'      Printer.Orientation = 2
'      Printer.PaperSize = vbPRPSA4
'      Printer.Font = "新細明體" 'Add By Sindy 2020/3/20
'      With rs44t0
'         .MoveFirst
'         stLstGroup1 = "": stCurGroup1 = ""
'         stLstGroup2 = "": stCurGroup2 = ""
'         stCompName = ""
'         Erase arrSubTot1()
'         Erase arrSubTot2()
'         Erase arrGrdTot()
'
'         Call GetTitleCustData("" & .Fields("a0k04"), "" & .Fields("a0k03"), "", m_CU01, m_CU02, _
'                               m_CU15, m_CU115, m_CU20, m_CU116, m_CU117, m_CU118, m_CU16, _
'                               m_CU18, m_CU04, m_CU158) 'Modify By Sindy 2015/1/16 改以收據抬頭抓客戶資料
'         Do While Not .EOF
'            stCurGroup1 = "" & .Fields("a0k11")
'            stCurGroup2 = "" & .Fields("a0k04") & .Fields("a0k03")
'   '         m_stCU16 = "" & .Fields("cu16").Value 'Add By Sindy 2014/11/5
'   '         m_stCU18 = "" & .Fields("cu18").Value 'Add By Sindy 2014/11/5
'            If (stCurGroup2 <> stLstGroup2) And iPage > 0 Then
'               Call PrintSubTot(arrSubTot1(), 1, stCompName)
'               Call PrintSubTot(arrSubTot2(), 2)
'               Erase arrSubTot1()
'               Erase arrSubTot2()
'               Printer.NewPage
'               stSalesName = "" & .Fields("st02").Value
'               stTitle = "" & .Fields("a0k04").Value
'               stCustNo = "" & .Fields("a0k03").Value
'               'stCustName = "" & .Fields("cu04").Value
'               stCustName = m_CU04
'               stCompName = .Fields("a0802").Value
'               Call PrintHead(stSalesName, stTitle, stCustNo, stCustName)
'            ElseIf (stCurGroup1 <> stLstGroup1) And iPage > 0 Then
'               Call PrintSubTot(arrSubTot1(), 1, stCompName)
'               Call NewLine
'               Erase arrSubTot1()
'               stCompName = .Fields("a0802").Value
'               'Add By Sindy 2020/4/20
'               Call PrintFoot(stLstGroup1)
'               Printer.EndDoc
'               iPage = 0
'               If bolPDF = True Then
'                  frmPDF.EndtProcess
'                  Unload frmPDF
'                  Sleep 500
'               End If
'               '2020/4/20 END
'            ElseIf iPage > 0 Then
'               Call NewLine
'            End If
'
'            If iPage = 0 Then
'               'Add By Sindy 2020/4/20
'               If bolPDF = True Then
'                  frmPDF.Show
'                  strFilePathName = Trim(txtCustNo(0)) & "-" & txtYear & "年-客戶扣繳明細核對表(" & IIf(stCurGroup1 = "1", "商標", IIf(stCurGroup1 = "2", "智慧所", "法律所")) & ")"
'                  frmPDF.StartProcess strExcelPath, strFilePathName
'                  If m_strEmailAttch <> "" Then m_strEmailAttch = m_strEmailAttch & ";"
'                  m_strEmailAttch = m_strEmailAttch & strExcelPath & strFilePathName & ".pdf"
'               End If
'               '2020/4/20 END
'               GetPleft
'               Printer.Orientation = 2
'               Printer.PaperSize = vbPRPSA4
'               Printer.Font = "新細明體" 'Add By Sindy 2020/3/20
'               stSalesName = "" & .Fields("st02").Value
'               stTitle = "" & .Fields("a0k04").Value
'               stCustNo = "" & .Fields("a0k03").Value
'               'stCustName = "" & .Fields("cu04").Value
'               stCustName = m_CU04
'               stCompName = .Fields("a0802").Value
'               'Call PrintHead("" & .Fields("st02").Value, "" & .Fields("a0k04").Value, "" & .Fields("a0k03").Value, "" & .Fields("cu04").Value)
'               Call PrintHead("" & .Fields("st02").Value, "" & .Fields("a0k04").Value, "" & .Fields("a0k03").Value, m_CU04)
'            End If
'
'            PrintData
'
'            stLstGroup1 = stCurGroup1
'            stLstGroup2 = stCurGroup2
'            For idx = LBound(arrGrdTot) To UBound(arrGrdTot)
'               If Not (idx = 6 And ("" & .Fields("Fee6")) = "") Then
'                  arrSubTot1(idx) = Format(Val(arrSubTot1(idx)) + Val("" & .Fields("Fee" & idx)))
'                  arrSubTot2(idx) = Format(Val(arrSubTot2(idx)) + Val("" & .Fields("Fee" & idx)))
'                  'arrGrdTot(idx) = Format(Val(arrGrdTot(idx)) + Val("" & .Fields("Fee" & idx)))
'               End If
'            Next
'            .MoveNext
'         Loop
'      End With
'      Call PrintSubTot(arrSubTot1(), 1, stCompName)
'      'Call PrintSubTot(arrSubTot2(), 2)
'
'      'Modify By Sindy 2020/4/20 改成Call函數
'      Call PrintFoot(stLstGroup1)
'   '   'Add By Sindy 2014/10/17
'   '   Call NewLine
'   '   Call NewLine
'   '   Printer.CurrentX = PLeft(1)
'   '   Printer.CurrentY = intY
'   '   Printer.FontSize = 14 'Add By Sindy 2019/11/28
'   '   'Modify By Sindy 2015/1/21
'   '   'Printer.Print "* 以上為" & txtYear & "年" & Left(Right(strSrvDate(2), 4), 2) & "月" & Right(strSrvDate(2), 2) & "日止　" & m_CU15 & "之扣繳資料，請核對；若12月31日前有增加屬於當年度之應扣繳款項，請自行加入並合計。"
'   '   If Val(Left(strSrvDate(2), Len(strSrvDate(2)) - 4)) > Val(txtYear) Then
'   '      Printer.Print "* 以上為" & txtYear & "年12月31日止　" & m_CU15 & "之扣繳資料，請核對；若12月31日後有增加屬於當年度之應扣繳款項，請自行加入並合計。"
'   '   Else
'   '      Printer.Print "* 以上為" & Left(strSrvDate(2), Len(strSrvDate(2)) - 4) & "年" & Left(Right(strSrvDate(2), 4), 2) & "月" & Right(strSrvDate(2), 2) & "日止　" & m_CU15 & "之扣繳資料，請核對；若12月31日後有增加屬於當年度之應扣繳款項，請自行加入並合計。"
'   '   End If
'   '   '2015/1/21 END
'   '   Printer.FontSize = 9 'Add By Sindy 2019/11/28
'   '   Call NewLine
'   '   Printer.CurrentX = PLeft(1)
'   '   Printer.CurrentY = intY
'   '   'Modify By Sindy 2016/10/24
'   ''   Printer.Print " （扣繳編號：台一國際專利商標事務所 04150022　台一國際專利法律事務所 04146457）" & _
'   ''   IIf(Check1.Value = 1, vbCrLf & vbCrLf & _
'   ''                   "　　　本所有分【專利商標】及【專利法律】 二個扣繳單位：" & vbCrLf & _
'   ''                   "　　　台一國際專利商標事務所(統編 04150022) 9A 代號91商標代理  台北巿長安東路2段112號10樓" & vbCrLf & _
'   ''                   "　　　台一國際專利法律事務所(統編 04146457) 9A 代號93專利代理  台北巿長安東路2段112號9樓" & vbCrLf & _
'   ''                   "　　　●扣繳憑單開立完成，煩請郵寄至本所或(傳真02-25011666)或Mail:71006@taie.com.tw，感謝您的配合！", "")
'   '   Printer.Print " （扣繳編號：台一國際專利商標事務所 04150022　台一國際專利法律事務所 04146457）" & _
'   '   IIf(Check1.Value = 1, vbCrLf & vbCrLf & _
'   '                   "　　　本所有分【專利商標】及【專利法律】 二個扣繳單位：" & vbCrLf & _
'   '                   "　　　　　　　　　　　　　　　(統編 04150022) 9A 代號91商標代理  台北巿長安東路2段112號10樓" & vbCrLf & _
'   '                   "　　　　　　　　　　　　　　　(統編 04146457) 9A 代號93專利代理  台北巿長安東路2段112號9樓" & vbCrLf & _
'   '                   "　　　●扣繳憑單開立完成，煩請郵寄至本所或(傳真02-25011666)或Mail:71006@taie.com.tw，感謝您的配合！", "")
'   '   '2016/10/24
'   '   '2014/10/17 END
'   '   'Add By Sindy 2016/12/12 把二家事務所名稱改粗體字
'   '   If Check1.Value = 1 Then
'   '      Printer.FontBold = True
'   '      Printer.CurrentX = PLeft(1)
'   '      Printer.CurrentY = intY + 550
'   '      Printer.Print "台一國際專利商標事務所"
'   '      Printer.CurrentX = PLeft(1)
'   '      Printer.CurrentY = intY + 725
'   '      Printer.Print "台一國際專利法律事務所"
'   '      Printer.FontBold = False
'   '   End If
'   '   '2016/12/12 END
'
'      'Call PrintSubTot(arrGrdTot(), 3)
'      Printer.EndDoc
'
'      If bolPDF = False Then 'Added by Morgan 2018/2/27
'         PUB_RestorePrinter strPrinter '復原系統預設印表機
'      Else
'         frmPDF.EndtProcess
'         Unload frmPDF
'         Sleep 1000
'         'PUB_RestorePrinter strPrinter '復原系統預設印表機
'      End If 'Added by Morgan 2018/2/27
'   End If
   
   'Add By Sindy 2014/10/17
   If bolPDF = True Then
      'Add By Sindy 2023/11/28
      If m_strMailKind = "B" Then 'B=催繳款書
         strSubject = Left(cboTitle, 4) & txtYear & "年度執行業務所得明細表請核對"
         strContent = "您好，" & vbCrLf & vbCrLf & _
                      "提醒您，" & m_CU15 & txtYear & "依收據全額支付款項，未扣除10%執行業務所得稅。" & vbCrLf & _
                      "應扣繳資料請參照附件。" & vbCrLf & vbCrLf & _
                      "請盡速確認，並回覆本所稅款繳交方式。" & vbCrLf & _
                      "(1) 稅款由台一代為繳交" & vbCrLf & _
                      "請上財政部網站，填寫繳款書存成電子檔（財政部網站 https://www.etax.nat.gov.tw/etwmain/etw144w/152）" & vbCrLf & _
                      "繳款書需在本月8日前給E-Mail：" & m_AccMail & " " & Left(strEmp, 1) & IIf(strEmpST22 = "F", "小姐", "生先") & vbCrLf & _
                      "台一繳稅後，繳款書正本會再寄回給　" & m_CU15 & "。" & vbCrLf & vbCrLf & _
                      "(2) 稅款由　" & m_CU15 & "先行繳交" & vbCrLf & _
                      "繳稅完成後，請提供繳款書影本及　" & m_CU15 & "帳戶資料，本所會在1-2日內退還稅款" & vbCrLf & _
                      "請注意！稅款需在本月10日前繳納完畢才不會被加計利息！" & vbCrLf & vbCrLf
      '2023/11/28 END
      Else 'A=催扣繳憑單
         strSubject = Left(cboTitle, 4) & txtYear & "年度扣繳明細表請核對"
      'Modify By Sindy 2016/11/1 瑞婷因想事後增加郵件內容,因此又提出開啟新郵件功能
      'Modify By Sindy 2020/4/20 修改EMail內容
''      strContent = "您好，" & vbCrLf & _
''                   "附件為　" & m_CU15 & txtYear & "年度扣繳資料，煩請核對資料" & vbCrLf & _
''                   "若扣繳資料有任何問題，請務必儘速聯絡，謝謝您的合作！" & vbCrLf & vbCrLf & _
''                   "本所有分【專利商標】及【專利法律】 二個扣繳單位：" & vbCrLf & _
''                   "台一國際專利商標事務所(統編 04150022) 9A 代號91商標代理  台北巿長安東路2段112號10樓" & vbCrLf & _
''                   "台一國際專利法律事務所(統編 04146457) 9A 代號93專利代理  台北巿長安東路2段112號9樓" & vbCrLf & vbCrLf & _
''                   "●扣繳憑單開立完成，煩請郵寄至本所或(傳真02-25011666)或Mail:71006@taie.com.tw，感謝您的配合！" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & _
''                   "財務處　" & strEmp & vbCrLf & _
''                   "台一國際專利法律事務所" & vbCrLf & _
''                   "台北市長安東路２段１１２號９樓" & vbCrLf & _
''                   "電話：０２－２５０６１０２３" & IIf(strEMP_Tel <> "", "（" & strEMP_Tel & "）", "") & vbCrLf & _
''                   "傳真：０２－２５０１１６６６"
      'Modify By Sindy 2020/12/30 + 扣繳核對明細正確請不用回覆，扣單開立完成直接以MAIL提供即可。
      'Modify By Sindy 2021/12/16 + 提醒您，承辦人員已更換為 辜小姐 E-MAIL : 71005@taie.com.tw
      'Modify By Sindy 2022/2/14 "提醒您，承辦人員已更換為 辜小姐 E-MAIL : 71005@taie.com.tw" & vbCrLf & vbCrLf Mark
         strContent = "您好，" & vbCrLf & vbCrLf & _
                      "附件為　" & m_CU15 & txtYear & "年度扣繳資料，請依附件內容開立扣繳憑單。" & vbCrLf & vbCrLf & _
                      "扣繳核對明細正確請不用回覆，扣單開立完成直接以MAIL提供即可。" & vbCrLf & vbCrLf
      End If
'      If Val(Left(strSrvDate(2), Len(strSrvDate(2)) - 4)) > Val(txtYear) Then
'         strContent = strContent & "(截至" & txtYear & "/12/31)"
'      Else
'         strContent = strContent & "(截至" & Left(strSrvDate(2), Len(strSrvDate(2)) - 4) & "/" & Left(Right(strSrvDate(2), 4), 2) & "/" & Right(strSrvDate(2), 2) & ")"
'      End If
'      strContent = strContent & _
'                   "，煩請核對資料" & vbCrLf & _
'                   "<B>資料正確請勿回覆</B>，若扣繳資料有任何問題，請務必儘速聯絡，謝謝您的合作！" & vbCrLf & vbCrLf & _
'                   "本所有分【專利商標】及【專利法律】 二個扣繳單位：" & vbCrLf & _
'                   "台一國際專利商標事務所(統編 04150022) 9A 代號91商標代理  台北巿長安東路2段112號10樓" & vbCrLf & _
'                   "台一國際專利法律事務所(統編 04146457) 9A 代號93專利代理  台北巿長安東路2段112號9樓" & vbCrLf & vbCrLf & _
'                   "●扣繳憑單開立完成，煩請郵寄至本所或(傳真02-25011666)或Mail:71006@taie.com.tw，感謝您的配合！" & vbCrLf & vbCrLf & vbCrLf
      
      'Call PUB_AccSettingMail(strTemplatePath, strExcelPath & strFilePathName & ".pdf", strSubject, strContent, m_strTo)
      'PUB_SendMail strUserNum, IIf(Pub_StrUserSt03 = "M51" Or Check5.Value = 1, strUserNum, m_strTo), "", strSubject, strContent, , strExcelPath & strFilePathName & ".pdf", , , , , , , , True
      
      'Add By Sindy 2016/11/3
      Screen.MousePointer = vbDefault
      '*************************************************************************************************
      '因為分所沒有Outlook不能開新郵件,改開VB的form介面操作
      '*************************************************************************************************
      '主旨
      frm880019.txtSubject = strSubject
      '本文
      strContent = strContent & "祝好," & vbCrLf & vbCrLf 'Add By Sindy 2023/11/28
      'Modify By Sindy 2024/12/16 分所抓操作人員
      'If strUserNum = Pub_GetSpecMan("出納人員-中所") Then
      If PUB_GetST06(strUserNum) = "2" Then
      '2024/12/16 END
         strContent = strContent & "台中所(會計)　" & strEmp & vbCrLf
         strContent = strContent & "電話：04-23270288" & IIf(strEMP_Tel <> "", "（分機：" & strEMP_Tel & "）", "") & vbCrLf
         strContent = strContent & "傳真：04-23227483"
'         strContent = strContent & "台中所(會計)　" & strEmp & vbCrLf & _
'                             "台一國際專利法律事務所" & vbCrLf & _
'                             "電話：０４－２３２７０２８８" & IIf(strEMP_Tel <> "", "（" & strEMP_Tel & "）", "") & vbCrLf & _
'                             "傳真：０４－２３２２７４８３"
      'Modify By Sindy 2024/12/16 分所抓操作人員
      'ElseIf strUserNum = Pub_GetSpecMan("出納人員-南所") Then
      ElseIf PUB_GetST06(strUserNum) = "3" Then
      '2024/12/16 END
         strContent = strContent & "台南所(會計)　" & strEmp & vbCrLf
         strContent = strContent & "電話：06-2743866" & IIf(strEMP_Tel <> "", "（分機：" & strEMP_Tel & "）", "") & vbCrLf
         strContent = strContent & "傳真：06-2744030"
'         strContent = strContent & "台南所(會計)　" & strEmp & vbCrLf & _
'                             "台一國際專利法律事務所" & vbCrLf & _
'                             "電話：０６－２７４３８６６" & IIf(strEMP_Tel <> "", "（" & strEMP_Tel & "）", "") & vbCrLf & _
'                             "傳真：０６－２７４４０３０"
      'Modify By Sindy 2024/12/16 分所抓操作人員
      'ElseIf strUserNum = Pub_GetSpecMan("出納人員-高所") Then
      ElseIf PUB_GetST06(strUserNum) = "4" Then
      '2024/12/16 END
         strContent = strContent & "高雄所(會計)　" & strEmp & vbCrLf
         strContent = strContent & "電話：07-2363602" & IIf(strEMP_Tel <> "", "（分機：" & strEMP_Tel & "）", "") & vbCrLf
         strContent = strContent & "傳真：07-2364360"
'         strContent = strContent & "高雄所(會計)　" & strEmp & vbCrLf & _
'                             "台一國際專利法律事務所" & vbCrLf & _
'                             "電話：０７－２３６３６０２" & IIf(strEMP_Tel <> "", "（" & strEMP_Tel & "）", "") & vbCrLf & _
'                             "傳真：０７－２３６４３６０"
      Else
         strContent = strContent & "財務處　" & strEmp & IIf(strEMP_Tel <> "", "（分機：" & strEMP_Tel & "）", "")
'         strContent = strContent & "財務處　" & strEmp & vbCrLf & _
'                             "台一國際專利法律事務所" & vbCrLf & _
'                             "台北市長安東路2段112號9樓" & vbCrLf & _
'                             "電話：０２－２５０６１０２３" & IIf(strEMP_Tel <> "", "（" & strEMP_Tel & "）", "") & vbCrLf & _
'                             "傳真：０２－２５０１１６６６"
      End If
'      frm880019.txtContent = strContent & vbCrLf & vbCrLf & vbCrLf & _
'                             "*************保密警語********************" & vbCrLf & _
'                             "本信件僅授權於指定之收信人取閱之用，信件中可能含有機密性資訊。" & vbCrLf & _
'                             "如果您並非被指定之收信人，任何未經授權而擅自使用此信件所含之機密資訊的行為是被嚴格禁止的。" & vbCrLf & _
'                             "如果您在任何未經授權的情形之下收到本信件，煩請您立即告知原發信人並將此信件回傳至以上地址。" & vbCrLf & _
'                             "謝謝您的合作。"
      frm880019.txtContent = strContent & vbCrLf
      frm880019.m_bolPLetter = True 'Add By Sindy 2020/4/23
      '附件
      frm880019.SetAttach m_strEmailAttch 'strExcelPath & strFilePathName & ".pdf"
      frm880019.txtReceiver = m_StrTo 'IIf(Pub_StrUserSt03 = "M51" Or Check5.Value = 1, strUserNum, m_strTo)
      frm880019.cmdAttach.Visible = True 'False
       'Add By Sindy 2024/12/24
      'If PUB_GetST06(strUserNum) = "1" Then
         frm880019.lblSender = m_AccMail
      'End If
      '2024/12/24 END
      frm880019.SetParent Me
      frm880019.Show vbModal
'      pbolDone = frm880019.m_bolDone
      Unload frm880019
      '*************************************************************************************************
      Screen.MousePointer = vbHourglass
      '2016/11/3 END
   End If
   '2014/10/17 END
   
   FormPrint = True
   
flgErr:
   
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   
End Function

'Add By Sindy 2020/4/20 傳入公司別
Private Sub PrintFoot(strCmp As String)
   Dim strNoteDate As String
   Dim strTempFile As String
   Dim strNo As String, strType As String, strAddr As String
   Dim intSelRow As Integer
   
   'Modify By Sindy 2021/12/2
   If strCmp = "L" Then '台一國際法律事務所
      strNo = "77211833": strType = "9A-10": strAddr = "台北巿中山區朱園里7鄰長安東路2段110號4樓"
   Else '台一國際智慧財產事務所
      strNo = "04146457": strType = "9A-93 或 9A-91": strAddr = "台北巿中山區朱園里7鄰長安東路2段112號9樓"
   End If
   
'   'Modify By Sindy 2022/1/19 改Call Excel產生報表
'   If strSrvDate(1) >= Form20上線日 Then
      xlsAnnuity.Range("A" & 2 & ":" & "P" & m_intColumn).Select
      With xlsAnnuity.Selection.Font
         '.Bold = True '粗體
         '.Name = "新細明體"
         .Size = 10
      End With
      
      m_intColumn = m_intColumn + 2: intSelRow = m_intColumn
      'Modify By Sindy 2020/12/30 + 扣繳核對明細正確請不用回覆，扣單開立完成直接以MAIL提供即可。
      xlsAnnuity.Range("B" & m_intColumn).Value = "扣繳核對明細正確請不用回覆，扣單開立完成直接以MAIL提供即可。"
      m_intColumn = m_intColumn + 1
      'Modify By Sindy 2020/12/28 備註日期調整,前後日期是要相同
      If Val(Left(strSrvDate(2), Len(strSrvDate(2)) - 4)) > Val(txtYear) Then
         strNoteDate = txtYear & "年12月31日"
      Else
         strNoteDate = Left(strSrvDate(2), Len(strSrvDate(2)) - 4) & "年" & Left(Right(strSrvDate(2), 4), 2) & "月" & Right(strSrvDate(2), 2) & "日"
      End If
      xlsAnnuity.Range("B" & m_intColumn).Value = "* 以上為" & strNoteDate & "止　" & m_CU15 & "之扣繳資料，請核對；若" & Right(strNoteDate, 6) & "後有增加屬於當年度之應扣繳款項，請自行加入並合計。"
      '2020/12/28 END
      m_intColumn = m_intColumn + 2
      xlsAnnuity.Range("B" & m_intColumn).Value = "本所扣繳資訊如下，請開立扣繳憑單"
      m_intColumn = m_intColumn + 1
      xlsAnnuity.Range("B" & m_intColumn).Value = "1.所 得 人：" & CompNameQuery(strCmp)
      m_intColumn = m_intColumn + 1
      xlsAnnuity.Range("B" & m_intColumn).Value = "2.統一編號：" & strNo
      m_intColumn = m_intColumn + 1
      xlsAnnuity.Range("B" & m_intColumn).Value = "3.所得類別：" & strType
      m_intColumn = m_intColumn + 1
      xlsAnnuity.Range("B" & m_intColumn).Value = "4.地　　址：" & strAddr
      m_intColumn = m_intColumn + 1
      'Modify by Amy 2024/05/17 原:71005
      xlsAnnuity.Range("B" & m_intColumn).Value = "5.扣繳憑單請  E-MAIL：" & m_AccMail
      m_intColumn = m_intColumn + 2
      xlsAnnuity.Range("B" & m_intColumn).Value = "感謝您的配合與支持！"
      
      xlsAnnuity.Range("A" & intSelRow & ":" & "P" & m_intColumn).Select
      With xlsAnnuity.Selection.Font
         '.Bold = True '粗體
         '.Name = "新細明體"
         .Size = 14
      End With
      xlsAnnuity.Range(6 & ":" & 6).Select
      xlsAnnuity.Selection.WrapText = True '自動換列
      
      xlsAnnuity.ActiveSheet.PageSetup.CenterFooter = "第 &P 頁，共 &N 頁"
      'Modify By Sindy 2022/1/19 列印標題
      xlsAnnuity.ActiveSheet.PageSetup.PrintTitleRows = "$1:$6"
      With xlsAnnuity.ActiveSheet.PageSetup
         .Zoom = False
         '.FitToPagesTall = 1 '縮放成一頁高
         .FitToPagesWide = 1 '縮放成一頁寬
         .FitToPagesTall = 1000 'Added by Morgan 2022/4/8 預設為1,筆數多時會縮小
      End With
      
      'Add By Sindy 2020/4/20
      If m_bolPDF = True Then
         strTempFile = Trim(txtCustNo(0)) & "-" & txtYear & "年-客戶扣繳明細核對表(" & IIf(stLstGroup1 = "1", "商標", IIf(stLstGroup1 = "2", "智慧所", "法律所")) & ")"
         strTempFile = strExcelPath & strTempFile
         If m_strEmailAttch <> "" Then m_strEmailAttch = m_strEmailAttch & ";"
         m_strEmailAttch = m_strEmailAttch & strTempFile & ".pdf"
      Else
         strTempFile = App.path & "\" & strUserNum & "\$$demo"
      End If
      '2020/4/20 END
      
'      'Modify By Sindy 2022/3/28
'      If strUserNum = "68008" Then
'         strExc(10) = strTempFile & ".xls"
'         If Dir(strExc(10)) <> "" Then
'            '改檔案性質為一般(因為原始檔區開啟檔案,都設唯讀)
'            SetAttr strExc(10), vbNormal
'            Kill strExc(10)
'         End If
'         If Val(xlsAnnuity.Version) < 12 Then
'            xlsAnnuity.Workbooks(1).SaveAs FileName:=strExc(10), FileFormat:=-4143
'         Else
'            xlsAnnuity.Workbooks(1).SaveAs FileName:=strExc(10), FileFormat:=56
'         End If
'         PUB_SendMail strUserNum, "97038", "", "玉瑛的報表會多一頁(要測試用)", "同主旨", , strExc(10), , , , , , , , True, , , , False, , , False
'      End If
'      '2022/3/28 END
      
      If m_bolPDF = False Then
         xlsAnnuity.Workbooks(1).PrintOut
      Else
      '   'xlTypePDF   0  PDF：可攜式文件格式檔案 (.pdf)
      '   'Quality:=xlQualityStandard 0  標準品質
      '   '各參數解釋:
      '   'Type 必要  XlFixedFormatTyp  匯出目標的檔案格式類型。
      '   'FileName   選用  Variant  要儲存之檔案的檔案名稱。 可以包含完整路徑，否則 Microsoft Excel 會將檔案儲存在目前的資料夾中。
      '   'Quality 選用  Variant  選用 XlFixedFormatQuality。 這會指定已發佈檔案的品質。
      '   'IncludeDocProperties 選用  Variant  若要包含檔案屬性，則為 True 。否則 為 False。
      '   'IgnorePrintAreas  選用  Variant  True 是表示忽略所有發佈時設定的列印範圍;否則 為 False。
      '   'From 選用  Variant  要發佈的起始頁碼。 如果省略此引數，將從頭開始列印。
      '   'To   選用  Variant  要發佈的最後一頁頁碼。 如果省略此引數，將發佈至最後一頁。
      '   'OpenAfterPublish 選用  Variant  True 是表示在發佈後在檢視器中顯示檔案;否則 為 False。
      '   'FixedFormatExtClassPtr 選用  Variant  FixedFormatExt 類別的指標。
         xlsAnnuity.ActiveSheet.ExportAsFixedFormat Type:=0, FileName:=strTempFile & ".pdf", Quality:=0, _
         IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
         'ShellExecute 0, "open", strTempFile, vbNullString, vbNullString, 1
      End If
      
      xlsAnnuity.Workbooks.Close 'SaveChanges:=False
      xlsAnnuity.Quit
      Set xlsAnnuity = Nothing
'   Else
'   '2022/1/19 END
'
'      'Add By Sindy 2014/10/17
'      Call NewLine
'      Call NewLine
'      Printer.CurrentX = PLeft(1)
'      Printer.CurrentY = intY
'      Printer.FontSize = 14 'Add By Sindy 2019/11/28
'      'Modify By Sindy 2020/12/30 + 扣繳核對明細正確請不用回覆，扣單開立完成直接以MAIL提供即可。
'      Printer.Print "扣繳核對明細正確請不用回覆，扣單開立完成直接以MAIL提供即可。"
'      Call NewLine
'      Printer.CurrentX = PLeft(1)
'      Printer.CurrentY = intY
'      'Modify By Sindy 2020/12/28 備註日期調整,前後日期是要相同
'      If Val(Left(strSrvDate(2), Len(strSrvDate(2)) - 4)) > Val(txtYear) Then
'         strNoteDate = txtYear & "年12月31日"
'      Else
'         strNoteDate = Left(strSrvDate(2), Len(strSrvDate(2)) - 4) & "年" & Left(Right(strSrvDate(2), 4), 2) & "月" & Right(strSrvDate(2), 2) & "日"
'      End If
'   '   'Modify By Sindy 2015/1/21
'   '   'Printer.Print "* 以上為" & txtYear & "年" & Left(Right(strSrvDate(2), 4), 2) & "月" & Right(strSrvDate(2), 2) & "日止　" & m_CU15 & "之扣繳資料，請核對；若12月31日前有增加屬於當年度之應扣繳款項，請自行加入並合計。"
'   '   If Val(Left(strSrvDate(2), Len(strSrvDate(2)) - 4)) > Val(txtYear) Then
'   '      Printer.Print "* 以上為" & txtYear & "年12月31日止　" & m_CU15 & "之扣繳資料，請核對；若12月31日後有增加屬於當年度之應扣繳款項，請自行加入並合計。"
'   '   Else
'   '      Printer.Print "* 以上為" & Left(strSrvDate(2), Len(strSrvDate(2)) - 4) & "年" & Left(Right(strSrvDate(2), 4), 2) & "月" & Right(strSrvDate(2), 2) & "日止　" & m_CU15 & "之扣繳資料，請核對；若12月31日後有增加屬於當年度之應扣繳款項，請自行加入並合計。"
'   '   End If
'   '   '2015/1/21 END
'      'Printer.FontSize = 9 'Add By Sindy 2019/11/28
'      Printer.Print "* 以上為" & strNoteDate & "止　" & m_CU15 & "之扣繳資料，請核對；若" & Right(strNoteDate, 6) & "後有增加屬於當年度之應扣繳款項，請自行加入並合計。"
'      '2020/12/28 END
'      Call NewLine
'      Call NewLine
'
'      'Modify By Sindy 2021/12/2
'      Printer.CurrentX = PLeft(1)
'      Printer.CurrentY = intY
'      Printer.Print "本所扣繳資訊如下，請開立扣繳憑單"
'      Call NewLine
'      Printer.CurrentX = PLeft(1)
'      Printer.CurrentY = intY
'      Printer.Print "1.所 得 人：" & CompNameQuery(strCmp)
'      Call NewLine
'      Printer.CurrentX = PLeft(1)
'      Printer.CurrentY = intY
'      Printer.Print "2.統一編號：" & strNo
'      Call NewLine
'      Printer.CurrentX = PLeft(1)
'      Printer.CurrentY = intY
'      Printer.Print "3.所得類別：" & strType
'      Call NewLine
'      Printer.CurrentX = PLeft(1)
'      Printer.CurrentY = intY
'      Printer.Print "4.地　　址：" & strAddr
'      'Add By Sindy 2021/12/16
'      Call NewLine
'      Printer.CurrentX = PLeft(1)
'      Printer.CurrentY = intY
'      Printer.Print "5.扣繳憑單請  E-MAIL：71005@taie.com.tw"
'      '2021/12/16 END
'      Call NewLine
'      Call NewLine
'      Printer.CurrentX = PLeft(1)
'      Printer.CurrentY = intY
'      Printer.Print "感謝您的配合與支持！"
'      '2021/12/2 END
'
'   '   'Modify By Sindy 2021/12/2 Mark
'   '   Printer.CurrentX = PLeft(1)
'   '   Printer.CurrentY = intY
'   '   'Modify By Sindy 2020/4/20
'   '   If strCmp = "1" Then
'   '      Printer.Print " （扣繳編號：台一國際專利商標事務所 04150022）" & _
'   '      IIf(Check1.Value = 1, vbCrLf & vbCrLf & _
'   '                      "　　　　　　　　　　　　　　　(統編 04150022) 9A 代號91商標代理  台北巿長安東路2段112號10樓" & vbCrLf & _
'   '                      "　　　●扣繳憑單開立完成，煩請郵寄至本所或(傳真02-25011666)或Mail:" & Pub_GetSpecMan("財務處總帳人員") & "@taie.com.tw，感謝您的配合！", "")
'   '   ElseIf strCmp = "2" Then
'   '      Printer.Print " （扣繳編號：台一國際智慧財產事務所 04146457）" & _
'   '      IIf(Check1.Value = 1, vbCrLf & vbCrLf & _
'   '                      "　　　　　　　　　　　　　　　(統編 04146457) 9A 代號93專利代理  台北巿長安東路2段112號9樓" & vbCrLf & _
'   '                      "　　　●扣繳憑單開立完成，煩請郵寄至本所或(傳真02-25011666)或Mail:" & Pub_GetSpecMan("財務處總帳人員") & "@taie.com.tw，感謝您的配合！", "")
'   '   Else 'L
'   '      Printer.Print " （扣繳編號：台一國際法律事務所 77211833）" & _
'   '      IIf(Check1.Value = 1, vbCrLf & vbCrLf & _
'   '                      "　　　　　　　　　　　　　　　(統編 77211833) 9A 代號10律師  台北巿長安東路2段110號4樓" & vbCrLf & _
'   '                      "　　　●扣繳憑單開立完成，煩請郵寄至本所或(傳真02-25011666)或Mail:" & Pub_GetSpecMan("財務處總帳人員") & "@taie.com.tw，感謝您的配合！", "")
'   '   End If
'   ''   'Modify By Sindy 2016/10/24
'   '''   Printer.Print " （扣繳編號：台一國際專利商標事務所 04150022　台一國際專利法律事務所 04146457）" & _
'   '''   IIf(Check1.Value = 1, vbCrLf & vbCrLf & _
'   '''                   "　　　本所有分【專利商標】及【專利法律】 二個扣繳單位：" & vbCrLf & _
'   '''                   "　　　台一國際專利商標事務所(統編 04150022) 9A 代號91商標代理  台北巿長安東路2段112號10樓" & vbCrLf & _
'   '''                   "　　　台一國際專利法律事務所(統編 04146457) 9A 代號93專利代理  台北巿長安東路2段112號9樓" & vbCrLf & _
'   '''                   "　　　●扣繳憑單開立完成，煩請郵寄至本所或(傳真02-25011666)或Mail:71006@taie.com.tw，感謝您的配合！", "")
'   ''   Printer.Print " （扣繳編號：台一國際專利商標事務所 04150022　台一國際專利法律事務所 04146457）" & _
'   ''   IIf(Check1.Value = 1, vbCrLf & vbCrLf & _
'   ''                   "　　　本所有分【專利商標】及【專利法律】 二個扣繳單位：" & vbCrLf & _
'   ''                   "　　　　　　　　　　　　　　　(統編 04150022) 9A 代號91商標代理  台北巿長安東路2段112號10樓" & vbCrLf & _
'   ''                   "　　　　　　　　　　　　　　　(統編 04146457) 9A 代號93專利代理  台北巿長安東路2段112號9樓" & vbCrLf & _
'   ''                   "　　　●扣繳憑單開立完成，煩請郵寄至本所或(傳真02-25011666)或Mail:71006@taie.com.tw，感謝您的配合！", "")
'   ''   '2016/10/24
'   ''   '2014/10/17 END
'   '
'   '   'Add By Sindy 2016/12/12 把二家事務所名稱改粗體字
'   '   If Check1.Value = 1 Then
'   '      Printer.FontBold = True
'   '      Printer.CurrentX = PLeft(1)
'   '      Printer.CurrentY = intY + 500 '350 '550
'   '      'Modify By Sindy 2020/4/20
'   '      Printer.Print CompNameQuery(strCmp)
'   ''      Printer.Print "台一國際專利商標事務所"
'   ''      Printer.CurrentX = PLeft(1)
'   ''      Printer.CurrentY = intY + 725
'   ''      Printer.Print "台一國際專利法律事務所"
'   '      '2020/4/20 END
'   '      Printer.FontBold = False
'   '   End If
'   '   '2016/12/12 END
'   End If
   
End Sub

''Add By Sindy 2015/1/16 改以收據抬頭抓客戶資料
'Private Sub GetTitleCustData(strTitleNm As String, strCustNo As String)
'Dim adoquery As New ADODB.Recordset
'
'   m_cu01 = "": m_cu02 = ""
'   m_cu15 = "": m_cu115 = "": m_cu20 = "": m_cu116 = "": m_cu117 = "": m_cu118 = "": m_stCU16 = "": m_stCU18 = ""
'   m_CU04 = "" 'Add By Sindy 2015/2/17
'   m_CU158 = "" 'Add By Sindy 2015/12/10
'
'   '先以收據抬頭抓客戶檔
'   'Modify By Sindy 2015/6/30 取消 CU158.境外公司 的控制
'   'Modify By Sindy 2015/12/10 + cu158
'   strExc(0) = "select cu115,cu01,cu02,cu20,cu116,cu117,cu118,nvl(cu16,cu22) cu16,cu18,DECODE(CU15,'0','台端','1','貴公司','貴單位') cu15,nvl(cu04,nvl(cu05||' '||cu88||' '||cu89||' '||cu90,cu06)) cu04,cu158" & _
'               " from (select min(cu01||cu02) CUID from customer where ('" & strTitleNm & "'=cu04 or '" & strTitleNm & "'=cu05||' '||cu88||' '||cu89||' '||cu90 or '" & strTitleNm & "'=cu06) and cu02='0') Y,customer,staff,acc090" & _
'               " where substr(Y.CUID,1,8)=cu01(+) and substr(Y.CUID,9,1)=cu02(+) and Y.CUID is not null" & _
'               " and cu13=st01(+) and cu12=a0901(+)"
'   intI = 1
'   Set adoquery = ClsLawReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      m_cu01 = "" & adoquery.Fields("cu01")
'      m_cu02 = "" & adoquery.Fields("cu02")
'      m_cu15 = "" & adoquery.Fields("cu15")
'      m_cu115 = "" & adoquery.Fields("cu115")
'      m_cu20 = "" & adoquery.Fields("cu20")
'      m_cu116 = "" & adoquery.Fields("cu116")
'      m_cu117 = "" & adoquery.Fields("cu117")
'      m_cu118 = "" & adoquery.Fields("cu118")
'      m_stCU16 = "" & adoquery.Fields("cu16")
'      m_stCU18 = "" & adoquery.Fields("cu18")
'      m_CU04 = "" & adoquery.Fields("cu04") 'Add By Sindy 2015/2/17
'      m_CU158 = "" & adoquery.Fields("cu158") 'Add By Sindy 2015/12/10
'   Else
'      adoquery.Close
'      'Add By Sindy 2015/6/9
'      '再讀收據抬頭基本資料檔
'      'Modify By Sindy 2015/6/30 取消 a4216.境外公司 的控制(and a4216 is null)
'      'Modify By Sindy 2015/12/10 + a4216
'      strExc(0) = "select a4201,a4204,a4205,a4203,a4217,a4206,st02,a4218,DECODE(a4219,'0','台端','1','貴公司','貴單位') a4219,a4216" & _
'                  " from acc420,staff" & _
'                  " where '" & strTitleNm & "'=a4201 and a4206=st01(+)"
'      intI = 1
'      Set adoquery = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         m_cu01 = ""
'         m_cu02 = ""
'         m_cu15 = "" & adoquery.Fields("a4219")
'         m_cu115 = "" & adoquery.Fields("a4218")
'         m_cu20 = ""
'         m_cu116 = ""
'         m_cu117 = ""
'         m_cu118 = ""
'         m_stCU16 = "" & adoquery.Fields("a4204")
'         m_stCU18 = "" & adoquery.Fields("a4205")
'         m_CU04 = "" & adoquery.Fields("a4201")
'         m_CU158 = "" & adoquery.Fields("a4216") 'Add By Sindy 2015/12/10
'      Else
'      '2015/6/9 END
'         adoquery.Close
'         'Modify By Sindy 2015/1/19
'         '再以客戶編號抓客戶檔
'         'Modify By Sindy 2015/6/30 取消 CU158.境外公司 的控制
'         'Modify By Sindy 2015/12/10 + cu158
'         strExc(0) = "select cu115,cu01,cu02,cu20,cu116,cu117,cu118,nvl(cu16,cu22) cu16,cu18,DECODE(CU15,'0','台端','1','貴公司','貴單位') cu15,nvl(cu04,nvl(cu05||' '||cu88||' '||cu89||' '||cu90,cu06)) cu04,cu158" & _
'                     " from customer,staff,acc090" & _
'                     " where substr('" & strCustNo & "',1,8)=cu01(+) and substr('" & strCustNo & "',9,1)=cu02(+)" & _
'                     " and cu13=st01(+) and cu12=a0901(+)"
'         intI = 1
'         Set adoquery = ClsLawReadRstMsg(intI, strExc(0))
'         If intI = 1 Then
'            m_cu01 = "" & adoquery.Fields("cu01")
'            m_cu02 = "" & adoquery.Fields("cu02")
'            m_cu15 = "" & adoquery.Fields("cu15")
'            m_cu115 = "" & adoquery.Fields("cu115")
'            m_cu20 = "" & adoquery.Fields("cu20")
'            m_cu116 = "" & adoquery.Fields("cu116")
'            m_cu117 = "" & adoquery.Fields("cu117")
'            m_cu118 = "" & adoquery.Fields("cu118")
'            m_stCU16 = "" & adoquery.Fields("cu16")
'            m_stCU18 = "" & adoquery.Fields("cu18")
'            m_CU04 = "" & adoquery.Fields("cu04") 'Add By Sindy 2015/2/17
'            m_CU158 = "" & adoquery.Fields("cu158") 'Add By Sindy 2015/12/10
'         'Add By Sindy 2015/11/18 再以收據抬頭抓國外代理人檔
'         Else
'            adoquery.Close
'            strExc(0) = "select fa79 cu115,fa01 cu01,fa02 cu02,fa16 cu20,fa80 cu116,fa81 cu117," & _
'                        "fa82 cu118,fa12 cu16,fa14 cu18,'貴公司' cu15," & _
'                        "nvl(fa04,nvl(fa05||' '||fa63||' '||fa64||' '||fa65,fa06)) cu04" & _
'                        " from (select min(fa01||fa02) CUID from fagent where ('" & strTitleNm & "'=fa04 or '" & strTitleNm & "'=fa05||' '||fa63||' '||fa64||' '||fa65 or '" & strTitleNm & "'=fa06) and fa02='0') Y,fagent" & _
'                        " where substr(Y.CUID,1,8)=fa01(+) and substr(Y.CUID,9,1)=fa02(+) and Y.CUID is not null"
'            intI = 1
'            Set adoquery = ClsLawReadRstMsg(intI, strExc(0))
'            If intI = 1 Then
'               m_cu01 = "" & adoquery.Fields("cu01")
'               m_cu02 = "" & adoquery.Fields("cu02")
'               m_cu15 = "" & adoquery.Fields("cu15")
'               m_cu115 = "" & adoquery.Fields("cu115")
'               m_cu20 = "" & adoquery.Fields("cu20")
'               m_cu116 = "" & adoquery.Fields("cu116")
'               m_cu117 = "" & adoquery.Fields("cu117")
'               m_cu118 = "" & adoquery.Fields("cu118")
'               m_stCU16 = "" & adoquery.Fields("cu16")
'               m_stCU18 = "" & adoquery.Fields("cu18")
'               m_CU04 = "" & adoquery.Fields("cu04")
'            End If
'         '2015/11/18 END
'         End If
'      End If
'   End If
'   adoquery.Close
'
'   Set adoquery = Nothing
'End Sub

Private Function SetGrid() As Boolean
   'Dim ADF
   Dim rsNew As New ADODB.Recordset
   Dim iField As Integer
   'Dim ColInfo()
   Dim jj As Integer, ii As Integer
   Dim strText As String
   Dim bolComp As Boolean
   Dim ArrStr As Variant
   Dim strQ As String
   Dim strManyRowCUid As String 'Modify By Sindy 2018/1/16
   
On Error GoTo flgErr
   'Modify by Amy 2016/10/17 改寫入暫存檔
    strQ = "Delete  Accrpt44t0 Where ID='" & strUserNum & "'"
    cnnConnection.Execute strQ
     
   'modify by sonia 2023/5/10 +a0w16給付總額R029
   strQ = "Select R001 as A0K20,R002 as ST02,R003 as A0K04,R004 as A0K03,R005 as A0K11,R006 as A0802,R007 as A0L02, " & _
                "R008 as A0K01,R009 as A0K02,R010 as CP10N,R011 as NA03,R012 as FEE0,R013 as FEE1,R014 as FEE2,R015 as FEE3," & _
                "R016 as FEE4,R017 as FEE5,R018 as FEE6,R019 as FEE7,R020 as A0W04,R021 as A0W02,R022 as A0W06,R023 as CP09," & _
                "R024 as A1P12,R025 as A0K33,R026 as A0J22,R027 as A0J25,R028 as CP10,R029 as A0W16,ID " & _
                "From Accrpt44t0 Where ID='" & strUserNum & "'"
   rsNew.CursorLocation = adUseClient
   rsNew.Open strQ, adoTaie, adOpenDynamic, adLockBatchOptimistic
   'end 2016/10/17
   cmdMail.Enabled = False 'Modify By Sindy 2017/2/2
   With rs44t0
      .MoveFirst
      stLstGroup1 = "": stCurGroup1 = ""
      stLstGroup2 = "": stCurGroup2 = ""
      stCompName = ""
      Erase arrSubTot1()
      Erase arrSubTot2()
      Erase arrGrdTot()
      
      'Mark by Amy 2016/10/17 改寫入暫存檔
'      ReDim ColInfo(.Fields.Count - 1)
'      For iField = 0 To UBound(ColInfo)
'         ColInfo(iField) = Array(.Fields(iField).Name, CInt(129), CInt(2000), True)
'      Next
'      Set rsNew = ADF.CreateRecordset(ColInfo)
      'end 2016/10/17
      'cmdMail.Enabled = False 'Add By Sindy 2014/10/17
      cmdMail.Enabled = True 'Modify By Sindy 2017/2/2
      'Modify By Sindy 2018/1/16 + strManyRowCUid
      Call GetTitleCustData("" & .Fields("a0k04"), "" & .Fields("a0k03"), "", m_CU01, m_CU02, _
                            m_CU15, m_CU115, m_CU20, m_CU116, m_CU117, m_CU118, m_CU16, _
                            m_CU18, m_CU04, m_CU158, , , , , , , , , , , , , , , m_CU172, , _
                            strManyRowCUid) 'Modify By Sindy 2015/1/16 改以收據抬頭抓客戶資料
      'Modify By Sindy 2018/1/16
      cboCombNo.Tag = cboCombNo.Text
      If strManyRowCUid <> "" Then
         cboCombNo.Clear
         ArrStr = Split(strManyRowCUid, ",")
         For ii = 0 To UBound(ArrStr)
            cboCombNo.AddItem ArrStr(ii)
         Next ii
         cboCombNo.Text = cboCombNo.Tag
      End If
      '2018/1/16 END
      
      'Modify By Sindy 2014/10/17
      For jj = 1 To 5
         strText = ""
         'Modify By Sindy 2015/1/16 改以收據抬頭抓客戶資料
'            If jj = 1 And Trim("" & .Fields("cu115")) <> "" Then strText = .Fields("cu01") & .Fields("cu02") & " 財：" & Trim(.Fields("cu115"))
'            If jj = 2 And Trim("" & .Fields("cu20")) <> "" Then strText = .Fields("cu01") & .Fields("cu02") & " 代：" & Trim(.Fields("cu20"))
'            If jj = 3 And Trim("" & .Fields("cu116")) <> "" Then strText = .Fields("cu01") & .Fields("cu02") & " 其：" & Trim(.Fields("cu116"))
'            If jj = 4 And Trim("" & .Fields("cu117")) <> "" Then strText = .Fields("cu01") & .Fields("cu02") & " 其：" & Trim(.Fields("cu117"))
'            If jj = 5 And Trim("" & .Fields("cu118")) <> "" Then strText = .Fields("cu01") & .Fields("cu02") & " 其：" & Trim(.Fields("cu118"))
         If jj = 1 And m_CU115 <> "" Then strText = m_CU01 & m_CU02 & Mid(m_CU04, 1, 2) & " 財：" & m_CU115
         If jj = 2 And m_CU20 <> "" Then strText = m_CU01 & m_CU02 & Mid(m_CU04, 1, 2) & " 代：" & m_CU20
         If jj = 3 And m_CU116 <> "" Then strText = m_CU01 & m_CU02 & Mid(m_CU04, 1, 2) & " 其：" & m_CU116
         If jj = 4 And m_CU117 <> "" Then strText = m_CU01 & m_CU02 & Mid(m_CU04, 1, 2) & " 其：" & m_CU117
         If jj = 5 And m_CU118 <> "" Then strText = m_CU01 & m_CU02 & Mid(m_CU04, 1, 2) & " 其：" & m_CU118
         If strText <> "" Then
            bolComp = False
            For ii = 0 To List1.ListCount - 1
               If List1.List(ii) = strText Then bolComp = True: Exit For
            Next ii
            If bolComp = False Then
               'Modify By Sindy 2015/12/10
               ArrStr = Split(strText, ";")
               If UBound(ArrStr) > 0 Then
                  For ii = 0 To UBound(ArrStr)
                     If InStr(ArrStr(ii), "：") > 0 Then
                        List1.AddItem ArrStr(ii)
                     Else
                        List1.AddItem Left(strText, InStr(strText, "：")) & ArrStr(ii)
                     End If
                  Next ii
               Else
               '2015/12/10 END
                  List1.AddItem strText
               End If
'               cmdMail.Enabled = True
            End If
         End If
      Next jj
      '2014/10/17 END
      
      Do While Not .EOF
         'Add By Sindy 2015/12/10 踢除境外公司或個人收據且未扣繳的資料
         strExc(0) = "select a1v01,a1v02,a1v06 from acc1v0,acc0k0" & _
                     " where a0k01='" & .Fields("A0K01") & "' and a0k05='1' and a0k01=a1v02(+) and a1v06=0 and a1v02 is not null"
         If m_CU158 = "Y" Then
            strExc(0) = strExc(0) & " Union" & _
                     " select a1v01,a1v02,a1v06 from acc1v0" & _
                     " where a1v02='" & .Fields("A0K01") & "' and a1v06=0 and a1v02 is not null"
         End If
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            GoTo ReadNext
         End If
         '2015/12/10 END
         If m_CU172 = "N" Then GoTo ReadNext 'Add By Sindy 2017/3/16 剔除不寄發扣繳核對資料
         
         '公司別
         stCurGroup1 = "" & .Fields("a0k11")
         '抬頭+編號
         stCurGroup2 = "" & .Fields("a0k04") & .Fields("a0k03")
         '紀錄第一筆群組資料
         If .AbsolutePosition = 1 Then
            stCompName = Replace(.Fields("a0802").Value, "台一國際", "")
            stCompName = Replace(stCompName, "事務所", "")
         '群組小計 1,2
         'Modify By Sindy 2016/1/28 晴天美食 105年
         'ElseIf (stCurGroup2 <> stLstGroup2) Then
         ElseIf (stCurGroup2 <> stLstGroup2 And stLstGroup2 <> "") Then
         '2016/1/28 END
            rsNew.AddNew
            rsNew.Fields("ID") = strUserNum 'Add by Amy 2016/10/17
            For idx = LBound(arrGrdTot) To UBound(arrGrdTot)
               rsNew.Fields("a0k11") = stLstGroup1
               rsNew.Fields("a0l02") = stCompName
               rsNew.Fields("a0k02") = "小計: "
               If Not (idx = 6 And ("" & .Fields("Fee6")) = "") Then
                  rsNew.Fields("Fee" & idx) = arrSubTot1(idx)
               End If
               arrSubTot1(idx) = 0
            Next
            rsNew.AddNew
            rsNew.Fields("ID") = strUserNum 'Add by Amy 2016/10/17
            For idx = LBound(arrGrdTot) To UBound(arrGrdTot)
               rsNew.Fields("a0k02") = "合計: "
               If Not (idx = 6 And ("" & .Fields("Fee6")) = "") Then
                  rsNew.Fields("Fee" & idx) = arrSubTot2(idx)
               End If
               arrSubTot2(idx) = 0
            Next
            Erase arrSubTot1()
            Erase arrSubTot2()
            stCompName = Replace(.Fields("a0802").Value, "台一國際", "")
            stCompName = Replace(stCompName, "事務所", "")
         '群組小計 1
         'Modify By Sindy 2016/1/28 晴天美食 105年
         'ElseIf (stCurGroup1 <> stLstGroup1) Then
         ElseIf (stCurGroup1 <> stLstGroup1 And stLstGroup1 <> "") Then
         '2016/1/28 END
            rsNew.AddNew
            rsNew.Fields("ID") = strUserNum 'Add by Amy 2016/10/17
            rsNew.Fields("a0k11") = stLstGroup1
            rsNew.Fields("a0l02") = stCompName
            rsNew.Fields("a0k02") = "小計: "
            For idx = LBound(arrGrdTot) To UBound(arrGrdTot)
               If Not (idx = 6 And ("" & .Fields("Fee6")) = "") Then
                  rsNew.Fields("Fee" & idx) = arrSubTot1(idx)
               End If
               arrSubTot1(idx) = 0
            Next
            Erase arrSubTot1()
            stCompName = Replace(.Fields("a0802").Value, "台一國際", "")
            stCompName = Replace(stCompName, "事務所", "")
         End If
         stLstGroup1 = stCurGroup1
         stLstGroup2 = stCurGroup2
         For idx = LBound(arrSubTot1) To UBound(arrSubTot1)
            If Not (idx = 6 And ("" & .Fields("Fee6")) = "") Then
               arrSubTot1(idx) = Format(Val(arrSubTot1(idx)) + Val("" & .Fields("Fee" & idx)))
               arrSubTot2(idx) = Format(Val(arrSubTot2(idx)) + Val("" & .Fields("Fee" & idx)))
               'arrGrdTot(idx) = Format(Val(arrGrdTot(idx)) + Val("" & .Fields("Fee" & idx)))
            End If
         Next
         '複製原來資料
         rsNew.AddNew
         rsNew.Fields("ID") = strUserNum 'Add by Amy 2016/10/17
         For iField = 0 To .Fields.Count - 1
            'Modify by Amy 2016/10/17
            rsNew.Fields(iField) = .Fields(iField)
         Next
         
ReadNext: 'Add By Sindy 2015/12/10
         .MoveNext
      Loop
      'Add By Sindy 2015/12/10
      If rsNew.RecordCount = 0 Then
         MsgBox MsgText(28), , MsgText(5)
      Else
      '2015/12/10 END
         rsNew.AddNew
         rsNew.Fields("ID") = strUserNum 'Add by Amy 2016/10/17
         rsNew.Fields("a0k11") = stCurGroup1
         rsNew.Fields("a0l02") = stCompName
         rsNew.Fields("a0k02") = "小計: "
         For idx = LBound(arrSubTot1) To UBound(arrSubTot1)
            rsNew.Fields("Fee" & idx) = arrSubTot1(idx)
         Next
         rsNew.AddNew
         rsNew.Fields("ID") = strUserNum 'Add by Amy 2016/10/17
         rsNew.Fields("a0k02") = "合計: "
         For idx = LBound(arrSubTot2) To UBound(arrSubTot2)
            rsNew.Fields("Fee" & idx) = arrSubTot2(idx)
         Next
         rsNew.UPDATE
      End If
   End With
   Set Adodc1.Recordset = rsNew.Clone
   Set DataGrid1.DataSource = Adodc1
   
   'Add By Sindy 2016/11/1 + 會計師E-Mail
   'Modify By Sindy 2024/9/30 + ChgSQL
   strExc(0) = "select * from acc490" & _
               " where a4901='" & m_CU01 & m_CU02 & "' and a4905 is not null" & _
               " union " & _
               "select * from acc490" & _
               " where a4901='" & ChgSQL(cboTitle) & "' and a4905 is not null"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If RsTemp.Fields("a4901") = m_CU01 & m_CU02 Then
         strText = m_CU01 & m_CU02 & Mid(m_CU04, 1, 2) & " 會：" & RsTemp.Fields("a4905")
      Else
         strText = Left(Trim(RsTemp.Fields("a4901")), 10) & " 會：" & RsTemp.Fields("a4905")
      End If
      ArrStr = Split(strText, ";")
      If UBound(ArrStr) > 0 Then
         For ii = 0 To UBound(ArrStr)
            If InStr(ArrStr(ii), "：") > 0 Then
               List1.AddItem ArrStr(ii)
            Else
               List1.AddItem Left(strText, InStr(strText, "：")) & ArrStr(ii)
            End If
         Next ii
      Else
         List1.AddItem strText
      End If
'      cmdMail.Enabled = True
   End If
   '2016/11/1 END
   
   SetListScroll List1 'Add By Sindy 2014/10/17
   
   SetGrid = True
   
flgErr:
   Set rsNew = Nothing
   'Set ADF = Nothing 'Mark by Amy 2016/10/17
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   'Resume
End Function

'Modify By Sindy 2014/12/16
'Private Sub cmdQuery_Click()
Public Sub cmdQuery_Click()
'2014/12/16 END
   Screen.MousePointer = vbHourglass
   Set DataGrid1.DataSource = Nothing
   DataGrid1.Refresh
   cmdPrint.Enabled = False
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
   If Process(1) = True Then
      If SetGrid = True Then cmdPrint.Enabled = True
   End If
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
   Screen.MousePointer = vbDefault
End Sub

'Add By Sindy 2013/11/27
Private Sub cmdChkYear_Click()
   If Trim(cboTitle.Text) = "" Then
      MsgBox "請輸入收據抬頭！", vbCritical
      Exit Sub
   End If
   If Trim(txtYear.Text) = "" Then
      MsgBox "請輸入扣繳年度！", vbCritical
      txtYear.SetFocus
      Exit Sub
   End If
   If Val(txtA2802) >= Val(txtYear) Then
      MsgBox "扣繳年度必須大於最近扣繳確認年度！", vbCritical
      txtYear.SetFocus
      Exit Sub
   End If
   'Modify By Sindy 2024/9/30 + ChgSQL
   strSql = "insert into ACC280(a2801,a2802) values('" & ChgSQL(Trim(cboTitle.Text)) & "'," & txtYear & ")"
   cnnConnection.Execute strSql
   txtA2802 = PUB_GetA2802LastYear(Trim(cboTitle.Text))
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
   If KeyCode <> vbKeyEscape Then
      Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
   End If
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 9048
   Me.Height = 5985 'Modify by Amy 2024/08/21 原:5700
   MoveFormToCenter Me
   Image1 = LoadPicture(strBackPicPath4)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   mebRecDate(1).Mask = DFormat
   mebRecDate(2).Mask = DFormat
   'Modify by Morgan 2004/4/8
   '預設年度改判斷4月
   'txtYear = IIf(Val(Mid(CFDate(ACDate(ServerDate)), 5, 2)) < 7, Val(Mid(CFDate(ACDate(ServerDate)), 1, 3)) - 1, Val(Mid(CFDate(ACDate(ServerDate)), 1, 3)))
   If Val(Right(strSrvDate(2), 4)) >= 401 Then
      txtYear = strSrvDate(2) \ 10000
   Else
      txtYear = strSrvDate(2) \ 10000 - 1
   End If
   
'   m_DefaultPrinter = Printer.DeviceName
'   For Each m_Prn In Printers
'      '2008/11/3 CANCEL BY SONIA
'      'If m_Prn.DeviceName <> m_DefaultPrinter Then
'         cmbPrinter.AddItem m_Prn.DeviceName
'      'End If
'   Next
'   If cmbPrinter.ListCount > 0 Then
'      cmbPrinter.ListIndex = 0
'   End If
   PUB_SetPrinter Me.Name, cmbPrinter, strPrinter 'Add By Sindy 2014/10/17
   
   Me.DataGrid1.Columns(0).Width = 450 '公司別
   Me.DataGrid1.Columns(1).Width = 825 '收款日期
   Me.DataGrid1.Columns(3).Width = Me.DataGrid1.Columns(1).Width '收據日期
   Me.DataGrid1.Columns(8).Width = 860 '補扣繳額
   
   'Add By Sindy 2014/10/20
   '取得郵件範本檔名
'   'Modify By Sindy 2016/11/1 又改要開新郵件
'   strTemplatePath = "$$TOT-000M31-0-01.oft"
'   Call PUB_GetSampleFile(strTemplatePath, Replace(Left(strTemplatePath, Len(strTemplatePath) - 4), "$$", ""))
'   strTemplatePath = App.path & "\" & strTemplatePath
'   '2016/11/1 END
   Check5.Caption = "寄測試信箱（" & strUserNum & "）"
   '2014/10/20 END
   
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
'   Label17.Caption = "" 'Add By Sindy 2013/12/6
   List1.Clear 'Add By Sindy 2014/10/17
   
   'Add By Sindy 2022/3/30
   If Dir(strExcelPath, vbDirectory) = "" Then
      MkDir strExcelPath
   End If
   'Add by Amy 2024/05/17 財務2個特殊設定拆成3個
   If Val(strSrvDate(1)) >= Val(財務拆總帳出納國內應收啟用日) Then
       stAccPerson = Pub_GetSpecMan("財務處應收處理人員")
   Else
      stAccPerson = Pub_GetSpecMan("財務處總帳人員")
   End If
   stTxtPerson = stAccPerson '取第一個人
   If InStr(stTxtPerson, ";") > 0 Then stTxtPerson = Mid(stTxtPerson, 1, Val(InStr(stTxtPerson, ";")) - 1)
   'end 2024/05/17
   
   'Modify By Sindy 2024/12/16 寄信者若為分所(中南高)修改為操作人員的信箱;
   '北所人員操作時維持用共用信箱 (taieacc)
   If PUB_GetST06(strUserNum) = "1" Then
      m_AccMail = "taieacc@taie.com.tw" 'Add By Sindy 2024/10/8 扣繳信件信箱
   Else
      m_AccMail = strUserNum & "@taie.com.tw" 'Add By Sindy 2024/10/8 扣繳信件信箱
   End If
   '2024/12/16 END
End Sub

'Added by Lydia 2016/12/19
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   '移到下一筆勾選客戶
   If CallByA4901(m_CallList, m_CallListIdx + 1) Then
      Cancel = True
      Exit Sub
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   'Add By Sindy 2014/11/12
   '若印表機變動, 則更新列印設定
   If Me.cmbPrinter.Text <> Me.cmbPrinter.Tag Then
      PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.cmbPrinter.Name, "0", "0", Me.cmbPrinter.Text
   End If
   '2014/11/12 END
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   
'   For Each m_Prn In Printers
'      If m_Prn.DeviceName = m_DefaultPrinter Then
'         Set Printer = m_Prn
'         Exit For
'      End If
'   Next

   Set rs44t0 = Nothing
   Set Frmacc44t0 = Nothing
   
   'Added by Lydia 2016/12/19 跳回會計師客戶資料查詢
   If m_CallList <> "" Then
      Frmacc44z0.Show
      tool3_enabled
   End If
End Sub

'Add By Sindy 2014/10/17
Private Sub SetListScroll(oList As ListBox)
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

Private Sub txtLike_GotFocus()
   'edit by nickc 2007/06/11  切換輸入法改用API
   'txtSalesNo.IMEMode = 2
   CloseIme
   TextInverse txtLike
End Sub

Private Sub txtLike_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> Asc("Y") Then
      KeyAscii = Asc("N")
   End If
End Sub

Private Sub txtRecNo_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtSalesNo_GotFocus()
   'edit by nickc 2007/06/11  切換輸入法改用API
   'txtSalesNo.IMEMode = 2
   CloseIme
   TextInverse txtSalesNo
End Sub

Private Sub txtSalesNo_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtSalesNo_Validate(Cancel As Boolean)
   If txtSalesNo = "" Then
      lblSales = ""
   Else
      lblSales = GetStaffName(txtSalesNo)
      If lblSales = "" Then
         MsgBox "智權人員不存在，請重新輸入！"
         txtSalesNo_GotFocus
         Cancel = True
      End If
   End If
End Sub

Private Sub txtTaxNo_GotFocus()
   TextInverse txtTaxNo
End Sub

Private Sub txtTaxNo_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtYear_GotFocus()
   'edit by nickc 2007/06/11  切換輸入法改用API
   'txtYear.IMEMode = 2
   CloseIme
   TextInverse txtYear
End Sub

Private Sub txtYear_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And (KeyAscii < 48 Or KeyAscii > 57) Then
      KeyAscii = 0
   End If
End Sub

'Private Sub txtCombNo_GotFocus()
'   TextInverse txtCombNo
'End Sub
'
'Private Sub txtCombNo_KeyPress(KeyAscii As Integer)
'   KeyAscii = UpperCase(KeyAscii)
'End Sub

'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
   
   cboTitle.Clear
   txtCustNo(0) = "": txtCustNo(1) = ""
   txtSalesNo = ""
   txtRecNo = ""
   txtTaxNo = ""
   mebRecDate(1).Mask = ""
   mebRecDate(1).Text = ""
   mebRecDate(1).Mask = DFormat
   mebRecDate(2).Mask = ""
   mebRecDate(2).Text = ""
   mebRecDate(2).Mask = DFormat
   'Modify by Lydia 2016/12/19 預設年度改判斷4月
   'txtYear = IIf(Val(Mid(CFDate(ACDate(ServerDate)), 5, 2)) < 7, Val(Mid(CFDate(ACDate(ServerDate)), 1, 3)) - 1, Val(Mid(CFDate(ACDate(ServerDate)), 1, 3)))
   If Val(Right(strSrvDate(2), 4)) >= 401 Then
      txtYear = strSrvDate(2) \ 10000
   Else
      txtYear = strSrvDate(2) \ 10000 - 1
   End If
   
   txtComp(1) = "1": txtComp(2) = "L" '"8"
   'Modify By Sindy 2018/1/16
   'txtCombNo = ""
   cboCombNo = ""
   '2018/1/16 END
End Sub

Private Sub cboTitle_Click()
   If cboTitle.ListIndex > 0 Then
      If txtCustNo(0).Text = "" Then
         txtCustNo(0).Text = Right(cboTitle.Text, 9)
      ElseIf txtCustNo(1).Text = "" Then
         txtCustNo(1).Text = Right(cboTitle.Text, 9)
      End If
      Dim strTmp As String
      strTmp = cboTitle.List(cboTitle.ListIndex)
      cboTitle.List(0) = RTrim(Left(strTmp, Len(strTmp) - 9))
      'Modify By Sindy 2018/1/16
      'If txtCombNo = "" Then
      If cboCombNo = "" Then
      '2018/1/16 END
         Call GetCoCustNo(cboTitle.List(0), Left(Right(strTmp, 9), 6))
      End If
   End If
   cboTitle.ListIndex = 0
   txtA2802 = PUB_GetA2802LastYear(Trim(cboTitle.Text)) 'Add By Sindy 2013/11/27
End Sub

Private Sub GetCoCustNo(strCustName As String, strCustNo As String)
   Dim strSql As String
   Dim adoquery As New ADODB.Recordset
   
   'Modify By Sindy 2024/9/30 + ChgSQL
   strSql = "Select CU01||CU02 From Customer Where CU04 like '" & ChgSQL(strCustName) & "%' And CU01 like '" & strCustNo & "%' Order by CU04 ASC,CU01 ASC,CU02 ASC"
   
On Error GoTo ErrHand

   adoquery.MaxRecords = 2
   adoquery.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
   If Not (adoquery.EOF And adoquery.BOF) Then
      'Modify By Sindy 2018/1/16
      'txtCombNo = "" & adoquery.Fields(0).Value
      cboCombNo = "" & adoquery.Fields(0).Value
      '2018/1/16 END
   'Add By Sindy 2015/2/17
   Else
      adoquery.Close
      'Modify By Sindy 2024/9/30 + ChgSQL
      strSql = "Select CU01||CU02 From Customer Where CU04 like '" & ChgSQL(strCustName) & "%' Order by CU04 ASC,CU01 ASC,CU02 ASC"
      adoquery.MaxRecords = 2
      adoquery.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
      If Not (adoquery.EOF And adoquery.BOF) Then
         'Modify By Sindy 2018/1/16
         'txtCombNo = "" & adoquery.Fields(0).Value
         cboCombNo = "" & adoquery.Fields(0).Value
         '2018/1/16 END
      'add by sonia 2025/8/11 無客戶檔資料時則合併列印客戶代號預設客戶代號起號
      Else
         cboCombNo = txtCustNo(0)
      'end 2025/8/11
      End If
   '2015/2/17 END
   End If
   adoquery.Close
   Set adoquery = Nothing
   Exit Sub
   
ErrHand:
   MsgBox Err.Description
   
End Sub

Private Sub cboTitle_GotFocus()
   'edit by nickc 2007/06/11  切換輸入法改用API
   'cboTitle.IMEMode = 1
   OpenIme
   TextInverse cboTitle    'add by sonia 2025/8/11
End Sub

Private Sub cboTitle_Validate(Cancel As Boolean)
   If CheckLen(Label1, cboTitle, 100) = MsgText(603) Then
      Cancel = True
      Exit Sub
   End If
   'edit by nickc 2007/06/11  切換輸入法改用API
   If Cancel = False Then CloseIme
End Sub

Private Sub cmdLikeSearch_Click()
Dim strTitle As String 'Add By Sindy 2021/8/19

   strTitle = cboTitle.Text 'Add By Sindy 2021/8/19
   If cboTitle.Text = "" Then
      MsgBox "請輸入收據抬頭！", vbCritical
   Else
      'Modify By Sindy 2021/8/19 從cboTitle_KeyPress搬過來
      If txtCustNo(0) <> "" Or txtCustNo(1) <> "" Or cboTitle.ListCount > 0 Then
         txtCustNo(0) = "": txtCustNo(1) = ""
         txtSalesNo = "": lblSales = ""
         cboTitle.Clear
         cboTitle.Text = strTitle 'Add By Sindy 2021/8/19
      End If
      'Modify By Sindy 2018/1/16
   '   If txtCombNo <> "" Then
   '      txtCombNo = ""
   '   End If
      If cboCombNo <> "" Then
         cboCombNo = ""
      End If
      '2018/1/16 END
      '2021/8/19 END
      
      'Modify by Morgan 2011/3/11 改呼叫共用函數
      'AddItem2CboTitle
      'Modify By Sindy 2013/12/30
      PUB_AddItem2CboTitle cboTitle, txtCustNo(0), txtCustNo(1), txtYear, True
   End If
   txtA2802 = PUB_GetA2802LastYear(Trim(cboTitle.Text)) 'Add By Sindy 2013/11/27
End Sub

'Private Function AddItem2CboTitle() As Boolean
'
'   Dim strSql As String, strCon1 As String, strCon2 As String
'   Dim adoQuery As New ADODB.Recordset
'   Dim strItem As String
'
'On Error GoTo ErrHand
'
'   strCon1 = ""
'   If txtYear <> "" Then
'      strCon1 = " and a0k16=" & txtYear
'   End If
'
'   strCon2 = ""
'   If txtCustNo(0) <> "" Then
'      strCon2 = strCon2 & " and a0k03>='" & txtCustNo(0).Text & "'"
'   End If
'   If txtCustNo(1) <> "" Then
'      strCon2 = strCon2 & " and a0k03<='" & txtCustNo(1).Text & "'"
'   End If
'
'   '2011/10/20 MODIFY BY SONIA E10023515
'   'strSql = "Select distinct rpad(a0k04, 60,' ') C01, a0k03 C02" & _
'      " from Acc0k0 where a0k04 like '" & cboTitle.Text & "%'" & strCon1 & strCon2 & _
'      " order by 2,1"
'   strSql = "Select distinct rpad(a0k04, 60,' ') C01, a0k03 C02" & _
'      " from Acc0k0 where instr(upper(a0k04),upper('" & cboTitle.Text & "'))>0" & strCon1 & strCon2 & _
'      " order by 2,1"
'
'   adoQuery.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
'   If Not (adoQuery.EOF And adoQuery.BOF) Then
'      strItem = cboTitle.Text
'      cboTitle.Clear
'      cboTitle.AddItem strItem
'      Do While Not adoQuery.EOF
'         strItem = "" & adoQuery.Fields(0) & " " & adoQuery.Fields(1)
'         cboTitle.AddItem strItem
'         adoQuery.MoveNext
'      Loop
'      cboTitle.ListIndex = 0
'   End If
'   adoQuery.Close
'   Set adoQuery = Nothing
'   AddItem2CboTitle = True
'   Exit Function
'
'ErrHand:
'   MsgBox Err.Description
'
'End Function

Private Sub txtComp_GotFocus(Index As Integer)
    TextInverse txtComp(Index)
End Sub

Private Sub txtComp_KeyPress(Index As Integer, KeyAscii As Integer)
   'Modify By Sindy 2020/4/20 + And Chr(KeyAscii) <> "L"
   If KeyAscii <> 8 And (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And Chr(KeyAscii) <> "L" Then
      KeyAscii = 0
   End If
End Sub

Private Sub txtComp_Validate(Index As Integer, Cancel As Boolean)
    If txtComp(Index) <> "" Then
        If txtComp(1) > txtComp(2) Then
            MsgBox "公司別範圍錯誤！", vbCritical
            Cancel = True
            Call txtComp_GotFocus(Index)
        End If
    End If
End Sub

Private Sub txtCustNo_GotFocus(Index As Integer)
   'edit by nickc 2007/06/11  切換輸入法改用API
   'txtCustNo(Index).IMEMode = 2
   CloseIme
   If Index = 1 Then
      If txtCustNo(0) <> "" And txtCustNo(1) = "" Then
         txtCustNo(1) = txtCustNo(0)
         txtCustNo(1).SelStart = 6
         txtCustNo(1).SelLength = 3
      Else
         TextInverse txtCustNo(Index)
      End If
   Else
      TextInverse txtCustNo(Index)
   End If
   
End Sub

Private Sub txtCustNo_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCustNo_Validate(Index As Integer, Cancel As Boolean)
   If Index = 1 Then
      If txtCustNo(Index) <> "" And Left(txtCustNo(0), 6) <> Left(txtCustNo(1), 6) Then
         MsgBox "前六碼必需相同！", vbCritical
         Cancel = True
      End If
   End If
   If txtCustNo(Index) <> "" Then
      txtCustNo(Index) = Left(txtCustNo(Index) + "000000000", 9)
   End If

End Sub

Private Sub cboTitle_KeyPress(KeyAscii As MSForms.ReturnInteger)
'   If txtCustNo(0) <> "" Or txtCustNo(1) <> "" Or cboTitle.ListCount > 0 Then
'      txtCustNo(0) = "": txtCustNo(1) = ""
'      txtSalesNo = "": lblSales = ""
'      cboTitle.Clear
'   End If
'   'Modify By Sindy 2018/1/16
''   If txtCombNo <> "" Then
''      txtCombNo = ""
''   End If
'   If cboCombNo <> "" Then
'      cboCombNo = ""
'   End If
'   '2018/1/16 END
End Sub


Sub GetPleft()
   Dim ii As Integer
   Erase PLeft
   ' 公司別
   PLeft(0) = 300
   '1 收款日期
   PLeft(1) = PLeft(0) + 3 * BytePix
   '收據號碼
   PLeft(2) = PLeft(1) + 9 * BytePix
   '收據日期
   PLeft(3) = PLeft(2) + 10 * BytePix
   '服務費
   PLeft(4) = PLeft(3) + 9 * BytePix
   '可扣稅額
   'pLeft(5) = pLeft(4) + 9 * BytePix
   PLeft(5) = PLeft(4) + 11 * BytePix
   '收款扣繳額
   'pLeft(6) = pLeft(5) + 11 * BytePix
   PLeft(6) = PLeft(5) + 2 * BytePix
   '補扣繳額
   PLeft(7) = PLeft(6) + 11 * BytePix
   '未扣稅額
   PLeft(8) = PLeft(7) + 11 * BytePix
   '案件性質
   PLeft(9) = PLeft(8) + 9 * BytePix
   '申請國家
   PLeft(10) = PLeft(9) + 13 * BytePix
   '已收扣單金額
   PLeft(11) = PLeft(10) + 13 * BytePix
   '調整稅額
   PLeft(12) = PLeft(11) + 13 * BytePix
   '扣單公司
   PLeft(13) = PLeft(12) + 11 * BytePix
   '扣單編號
   PLeft(14) = PLeft(13) + 5 * BytePix
   '扣單備註
   PLeft(15) = PLeft(14) + 10 * BytePix
   '票期
   PLeft(16) = PLeft(15) + 13 * BytePix
End Sub

'*************************************************
'  抬頭列印
'
'*************************************************
Private Sub PrintHead(Optional stSalesName As String = "", Optional stTitle As String, Optional stCustNo As String, _
                      Optional stCustName As String = "")
   'Modify By Sindy 2022/1/19 改Call Excel產生報表
'   If strSrvDate(1) >= Form20上線日 Then
      '預設A4紙張/橫式/比例 80%/水平置中/邊界左右都改0
      Set xlsAnnuity = New Excel.Application
      'xlsAnnuity.Visible = True
      xlsAnnuity.SheetsInNewWorkbook = 1 '預設工作表數量
      xlsAnnuity.Workbooks.add
      Set wksAnnuity = xlsAnnuity.Worksheets(1)
      xlsAnnuity.ActiveWindow.Zoom = 75 '畫面比例100%太大了,調整為75%
      '把Excel的警告訊息關掉
      xlsAnnuity.DisplayAlerts = False
      
      wksAnnuity.PageSetup.PaperSize = 9 'A4
      wksAnnuity.PageSetup.Orientation = xlLandscape '橫印
      'wksAnnuity.PageSetup.Orientation = wdOrientLandscape '直印
      wksAnnuity.PageSetup.LeftMargin = xlsAnnuity.InchesToPoints(0.1) '邊界
      wksAnnuity.PageSetup.RightMargin = xlsAnnuity.InchesToPoints(0.1)
      wksAnnuity.PageSetup.TopMargin = xlsAnnuity.InchesToPoints(0.4)
      wksAnnuity.PageSetup.BottomMargin = xlsAnnuity.InchesToPoints(0.4)
      wksAnnuity.PageSetup.CenterHorizontally = True '版面設定->邊界->水平置中
      
   '   xlsAnnuity.SheetsInNewWorkbook = 1 '預設工作表數量
   '   xlsAnnuity.Workbooks.add
   '   Set wksAnnuity = xlsAnnuity.Worksheets(1)
      wksAnnuity.Activate
      
      '設定各欄位長度
      wksAnnuity.Columns("A").ColumnWidth = 2
      wksAnnuity.Columns("B").ColumnWidth = 8
      wksAnnuity.Columns("C").ColumnWidth = 8
      wksAnnuity.Columns("D").ColumnWidth = 8
      wksAnnuity.Columns("E").ColumnWidth = 8
      wksAnnuity.Columns("F").ColumnWidth = 9
      wksAnnuity.Columns("G").ColumnWidth = 8
      'Modify By Sindy 2025/9/30
      If m_bolPDF = False Then
         wksAnnuity.Columns("H").ColumnWidth = 8 '未扣稅額
      '2025/9/30 END
      Else
         wksAnnuity.Columns("H").ColumnWidth = 2
      End If
      wksAnnuity.Columns("I").ColumnWidth = 10
      wksAnnuity.Columns("J").ColumnWidth = 10
      wksAnnuity.Columns("K").ColumnWidth = 8
      wksAnnuity.Columns("L").ColumnWidth = 8
      wksAnnuity.Columns("M").ColumnWidth = 4
      wksAnnuity.Columns("N").ColumnWidth = 8
      wksAnnuity.Columns("O").ColumnWidth = 10
      wksAnnuity.Columns("P").ColumnWidth = 8
      
      '標題
      m_intColumn = 1
      xlsAnnuity.Range("D" & m_intColumn).Value = ReportTitle(422) '***  年度扣繳明細核對表  ***"
      xlsAnnuity.Range("A" & m_intColumn & ":" & "P" & m_intColumn).Select
      With xlsAnnuity.Selection
         .HorizontalAlignment = xlCenter '置中
         .VerticalAlignment = xlCenter
         .WrapText = False
         .Orientation = 0
         .AddIndent = False
         .IndentLevel = 0
         .ShrinkToFit = False
         .ReadingOrder = xlContext
         .MergeCells = True
      End With
      With xlsAnnuity.Selection.Font
         .Bold = True '粗體
         .Name = "新細明體"
         .Size = 16
      End With
      m_intColumn = m_intColumn + 1
      xlsAnnuity.Range("I" & m_intColumn).Value = "扣繳年度：" & txtYear
      xlsAnnuity.Range("L" & m_intColumn).Value = "TEL1: " & m_CU16
      m_intColumn = m_intColumn + 1
      xlsAnnuity.Range("A" & m_intColumn).Value = "列印人員：" & StaffQuery(strUserNum)
      xlsAnnuity.Range("L" & m_intColumn).Value = "FAX1: " & m_CU18
      xlsAnnuity.Range("O" & m_intColumn).Value = "列印日期：" & IIf(Mid(CFDate(ACDate(ServerDate)), 1, 1) = "0", Mid(CFDate(ACDate(ServerDate)), 2, 8), CFDate(ACDate(ServerDate)))
      m_intColumn = m_intColumn + 1
      xlsAnnuity.Range("A" & m_intColumn).Value = "智權人員：" & stSalesName
'      '換頁
'      .Range("A" & intCounter).Select
'      .HPageBreaks.add Before:=.Application.ActiveCell
'      Call SetExcel(2, , , stCmp) 'Modify by Amy 2020/0
      'xlsAnnuity.Range("O" & m_intColumn).Value = "頁　次：" & wksAnnuity.HPageBreaks.Count + 1
      
      m_intColumn = m_intColumn + 1
      xlsAnnuity.Range("A" & m_intColumn).Value = "收據抬頭：" & stTitle
      If cboCombNo <> "" Then
         xlsAnnuity.Range("F" & m_intColumn).Value = "合併客戶代號：" & cboCombNo.Text
      Else
         xlsAnnuity.Range("F" & m_intColumn).Value = "客戶代號：" & stCustNo
         xlsAnnuity.Range("J" & m_intColumn).Value = "客戶名稱：" & stCustName
      End If
      
      m_intColumn = m_intColumn + 1
      xlsAnnuity.Range("A" & m_intColumn).Value = "公司"
      'xlsAnnuity.Range("A" & m_intColumn).HorizontalAlignment = xlCenter
      xlsAnnuity.Range("B" & m_intColumn).Value = "收款日期"
      xlsAnnuity.Range("C" & m_intColumn).Value = "收據號碼"
      xlsAnnuity.Range("D" & m_intColumn).Value = "收據日期"
      xlsAnnuity.Range("E" & m_intColumn).Value = "服務費"
      xlsAnnuity.Range("F" & m_intColumn).Value = "收款扣繳額"
      xlsAnnuity.Range("G" & m_intColumn).Value = "補扣繳額"
      'Modify By Sindy 2025/9/30
      If m_bolPDF = False Then
      '2025/9/30 END
         xlsAnnuity.Range("H" & m_intColumn).Value = "未扣稅額"
      End If
      xlsAnnuity.Range("I" & m_intColumn).Value = "案件性質"
      xlsAnnuity.Range("J" & m_intColumn).Value = "申請國家"
      xlsAnnuity.Range("K" & m_intColumn).Value = "已收扣單金額"
      xlsAnnuity.Range("L" & m_intColumn).Value = "調整稅額"
      xlsAnnuity.Range("M" & m_intColumn).Value = "扣單公司"
      xlsAnnuity.Range("N" & m_intColumn).Value = "扣單編號"
      xlsAnnuity.Range("O" & m_intColumn).Value = "扣單備註"
      xlsAnnuity.Range("P" & m_intColumn).Value = "票期"
      xlsAnnuity.Range("A" & m_intColumn & ":" & "P" & m_intColumn).Select
      '上框線
      With xlsAnnuity.Selection.Borders(xlEdgeTop)
          .LineStyle = xlContinuous
          .ColorIndex = xlAutomatic
          .tintandshade = 0
          .Weight = xlThin
      End With
      '下框線
      With xlsAnnuity.Selection.Borders(xlEdgeBottom)
          .LineStyle = xlContinuous
          .ColorIndex = xlAutomatic
          .tintandshade = 0
          .Weight = xlThin
      End With
      xlsAnnuity.Range("E" & m_intColumn & ":H" & m_intColumn).Select
      xlsAnnuity.Selection.HorizontalAlignment = xlRight '靠右
      xlsAnnuity.Range("K" & m_intColumn & ":L" & m_intColumn).Select
      xlsAnnuity.Selection.HorizontalAlignment = xlRight '靠右
'   Else
'   '2022/1/19 END
'
'      '起始列印列
'      intY = 1000 - intDefault
'      iPage = iPage + 1
'      With Printer
'         '表頭
'         .FontSize = 16
'         .CurrentX = (.ScaleWidth - .TextWidth(ReportTitle(422))) / 2
'         .CurrentY = intY
'         Printer.Print ReportTitle(422)
'         '跳列
'         intY = intY + 500
'
'         '條件
'         .FontSize = 9
'
'         .CurrentX = (.ScaleWidth - .TextWidth("扣繳年度: " & txtYear)) / 2
'         .CurrentY = intY
'         Printer.Print "扣繳年度: " & txtYear
'
'         'Add By Sindy 2014/10/17
'         .CurrentX = PLeft(13) - 1 * BytePix
'         .CurrentY = intY
'         Printer.Print "TEL1: " & m_CU16 'rs44t0("cu16").Value 'Modify By Sindy 2014/11/5
'         '2014/10/17
'
'         '跳列
'         intY = intY + RowPix
'
'         .CurrentX = 300
'         .CurrentY = intY
'         Printer.Print "列印人員: "
'
'         .CurrentX = 1200
'         .CurrentY = intY
'         Printer.Print StaffQuery(strUserNum)
'
'         'Add By Sindy 2014/10/17
'         .CurrentX = PLeft(13) - 1 * BytePix
'         .CurrentY = intY
'         Printer.Print "FAX1: " & m_CU18 'rs44t0("cu18").Value 'Modify By Sindy 2014/11/5
'         '2014/10/17
'
'         .CurrentX = .ScaleWidth - 20 * BytePix
'         .CurrentY = intY
'         Printer.Print "列印日期: " & IIf(Mid(CFDate(ACDate(ServerDate)), 1, 1) = "0", Mid(CFDate(ACDate(ServerDate)), 2, 8), CFDate(ACDate(ServerDate)))
'
'         '跳列
'         intY = intY + RowPix
'
'         .CurrentX = 300
'         .CurrentY = intY
'         Printer.Print "智權人員: "
'
'         .CurrentX = 1200
'         .CurrentY = intY
'         Printer.Print stSalesName
'
'         .CurrentX = .ScaleWidth - 20 * BytePix
'         .CurrentY = intY
'         Printer.Print "頁次: " & Right(Space(4) + Format(iPage), 4)
'
'         '跳列
'         intY = intY + RowPix
'
'         .CurrentY = intY
'         .CurrentX = 300
'         Printer.Print "收據抬頭: "
'
'         .CurrentX = 1200
'         .CurrentY = intY
'         Printer.Print stTitle
'
'         'Modify By Sindy 2018/1/16
'         'If txtCombNo <> "" Then
'         If cboCombNo <> "" Then
'         '2018/1/16 END
'            .CurrentX = 5000
'            .CurrentY = intY
'            Printer.Print "合併客戶代號: "
'            .CurrentX = 5000 + 14 * BytePix
'            .CurrentY = intY
'            'Modify By Sindy 2018/1/16
'            'Printer.Print txtCombNo.Text
'            Printer.Print cboCombNo.Text
'            '2018/1/16 END
'         Else
'            .CurrentX = 5000
'            .CurrentY = intY
'            Printer.Print "客戶代號: "
'            .CurrentX = 5000 + 10 * BytePix
'            .CurrentY = intY
'            Printer.Print stCustNo
'            .CurrentX = 5000 + 35 * BytePix
'            .CurrentY = intY
'            Printer.Print "客戶名稱: "
'            .CurrentX = 5000 + 45 * BytePix
'            .CurrentY = intY
'            Printer.Print stCustName
'         End If
'
'         '跳列
'         intY = intY + RowPix
'
'         .CurrentX = PLeft(0)
'         .CurrentY = intY
'         Printer.Print "公"
'
'         .CurrentX = PLeft(0)
'         .CurrentY = intY + 2 * BytePix
'         Printer.Print "司"
'
'         .CurrentX = PLeft(1)
'         .CurrentY = intY
'         Printer.Print "收款日期"
'         .CurrentX = PLeft(2)
'         .CurrentY = intY
'         Printer.Print "收據號碼"
'         .CurrentX = PLeft(3)
'         .CurrentY = intY
'         Printer.Print "收據日期"
'
'         .CurrentX = PLeft(5) - 1 * BytePix - .TextWidth("服務費")
'         .CurrentY = intY
'         Printer.Print "服務費"
'   '      .CurrentX = pLeft(6) - 1 * BytePix - .TextWidth("可扣稅額")
'   '      .CurrentY = intY
'   '      Printer.Print "可扣稅額"
'         .CurrentX = PLeft(7) - 1 * BytePix - .TextWidth("收款扣繳額")
'         .CurrentY = intY
'         Printer.Print "收款扣繳額"
'         .CurrentX = PLeft(8) - 1 * BytePix - .TextWidth("補扣繳額")
'         .CurrentY = intY
'         Printer.Print "補扣繳額"
'
'         'Modify By Sindy 2019/12/27
'         If m_bolPDF = False Then
'         '2019/12/27 END
'            .CurrentX = PLeft(9) - 1 * BytePix - .TextWidth("未收稅額")
'            .CurrentY = intY
'            Printer.Print "未扣稅額"
'         End If
'         '2019/12/27 END
'
'         .CurrentX = PLeft(9)
'         .CurrentY = intY
'         Printer.Print "案件性質"
'         .CurrentX = PLeft(10)
'         .CurrentY = intY
'         Printer.Print "申請國家"
'
'         .CurrentX = PLeft(12) - 1 * BytePix - .TextWidth("已收扣單金額")
'         .CurrentY = intY
'         Printer.Print "已收扣單金額"
'         .CurrentX = PLeft(13) - 1 * BytePix - .TextWidth("調整稅額")
'         .CurrentY = intY
'         Printer.Print "調整稅額"
'         .CurrentX = PLeft(13)
'         .CurrentY = intY
'         Printer.Print "扣單"
'         .CurrentX = PLeft(13)
'         .CurrentY = intY + 2 * BytePix
'         Printer.Print "公司"
'         .CurrentX = PLeft(14)
'         .CurrentY = intY
'         Printer.Print "扣單編號"
'         .CurrentX = PLeft(15)
'         .CurrentY = intY
'         Printer.Print "扣單備註"
'         .CurrentX = PLeft(16)
'         .CurrentY = intY
'         Printer.Print "票期"
'
'         If NewLine(5 * BytePix) = False Then
'            Printer.DrawStyle = vbSolid
'            Printer.Line (PLeft(0), intY)-(.ScaleWidth - 4 * BytePix, intY)
'            Call NewLine(1 * BytePix)
'         End If
'      End With
'   End If
End Sub

Private Sub PrintData()
   'Added by Morgan 2011/12/21
   Dim stLstItem As String, stLstRecNo As String
   Dim dblAddAmt(7) As Double
   Dim idx As Integer
   
   'Modify By Sindy 2022/1/19 改Call Excel產生報表
'   If strSrvDate(1) >= Form20上線日 Then
      '列印明細
      m_intColumn = m_intColumn + 1
      xlsAnnuity.Range("A" & m_intColumn).Value = "" & rs44t0("a0k11").Value '公司
      xlsAnnuity.Range("B" & m_intColumn).Value = ChangeTStringToTDateString("" & rs44t0("a0l02").Value) '收款日期
      xlsAnnuity.Range("C" & m_intColumn).Value = "" & rs44t0("a0k01").Value '收據號碼
      xlsAnnuity.Range("D" & m_intColumn).Value = ChangeTStringToTDateString("" & rs44t0("a0k02").Value) '收據日期
      '案件性質
      If rs44t0("a0k33") = "Y" Then
         xlsAnnuity.Range("I" & m_intColumn).Value = "" & rs44t0("a0j22").Value
      Else
         xlsAnnuity.Range("I" & m_intColumn).Value = "" & rs44t0("cp10N").Value
      End If
      xlsAnnuity.Range("J" & m_intColumn).Value = "" & rs44t0("na03").Value '申請國家
      xlsAnnuity.Range("M" & m_intColumn).Value = "" & rs44t0("a0w04").Value '扣單公司
      xlsAnnuity.Range("N" & m_intColumn).Value = "" & rs44t0("a0w02").Value '扣單編號
      xlsAnnuity.Range("O" & m_intColumn).Value = "" & Left(rs44t0("a0w06").Value, 14) '扣單備註
      xlsAnnuity.Range("P" & m_intColumn).Value = Format("" & rs44t0("a1p12").Value, "###/##/##") '票期
      
      'Added by Morgan 2011/12/21 若收據有變更帳款類別則相同的依照列印順序合併
      Erase dblAddAmt
      
      For idx = 1 To 7
         dblAddAmt(idx) = Val("" & rs44t0("Fee" & idx))
      Next
      
      stLstRecNo = "" & rs44t0("a0k01")
      stLstItem = "" & rs44t0("a0j22")
      If rs44t0("a0k33") = "Y" Then
         rs44t0.MoveNext
         Do While Not rs44t0.EOF
            If stLstRecNo = rs44t0("a0k01") And rs44t0("a0k33") = "Y" And stLstItem = rs44t0("a0j22") Then
               
               rs44t0.MovePrevious
               For idx = LBound(arrGrdTot) To UBound(arrGrdTot)
                  If Not (idx = 6 And ("" & rs44t0("Fee6")) = "") Then
                     arrSubTot1(idx) = Format(Val(arrSubTot1(idx)) + Val("" & rs44t0("Fee" & idx)))
                     arrSubTot2(idx) = Format(Val(arrSubTot2(idx)) + Val("" & rs44t0("Fee" & idx)))
                  End If
               Next
               rs44t0.MoveNext
               
               For idx = 1 To 7
                  dblAddAmt(idx) = dblAddAmt(idx) + Val("" & rs44t0("Fee" & idx))
               Next
            Else
               Exit Do
            End If
            rs44t0.MoveNext
         Loop
         rs44t0.MovePrevious
      End If
      
      '"已收扣單金額"
      'Modified by Morgan 2011/12/21 若收據有變更帳款類別則相同的依照列印順序合併
      'stMoneyTemp = IIf(IsNull(rs44t0("Fee6").Value), "", Format(rs44t0("Fee6").Value, DDollar))
      stMoneyTemp = IIf(IsNull(rs44t0("Fee6").Value), "", Format(dblAddAmt(6), DDollar))
      xlsAnnuity.Range("K" & m_intColumn).Value = stMoneyTemp '已收扣單金額
      
      '"調整稅額"
      'Modified by Morgan 2011/12/21 若收據有變更帳款類別則相同的依照列印順序合併
      'stMoneyTemp = Format("" & rs44t0("Fee7").Value, DDollar)
      stMoneyTemp = Format(dblAddAmt(7), DDollar)
      xlsAnnuity.Range("L" & m_intColumn).Value = stMoneyTemp '調整稅額
      
      '數字靠右
      '"服務費"
      'Modified by Morgan 2011/12/21 若收據有變更帳款類別則相同的依照列印順序合併
      'stMoneyTemp = Format("" & rs44t0("Fee1").Value, DDollar)
      stMoneyTemp = Format(dblAddAmt(1), DDollar)
      xlsAnnuity.Range("E" & m_intColumn).Value = stMoneyTemp '服務費
      
      '"收款扣繳額"
      'Modified by Morgan 2011/12/21 若收據有變更帳款類別則相同的依照列印順序合併
      'stMoneyTemp = Format("" & rs44t0("Fee3").Value, DDollar)
      stMoneyTemp = Format(dblAddAmt(3), DDollar)
      xlsAnnuity.Range("F" & m_intColumn).Value = stMoneyTemp '收款扣繳額
      
      '"補扣繳額"
      'Modified by Morgan 2011/12/21 若收據有變更帳款類別則相同的依照列印順序合併
      'stMoneyTemp = Format("" & rs44t0("Fee4").Value, DDollar)
      stMoneyTemp = Format(dblAddAmt(4), DDollar)
      xlsAnnuity.Range("G" & m_intColumn).Value = stMoneyTemp '補扣繳額
      
      '"未扣稅額"
      'Modify By Sindy 2019/12/27
      If m_bolPDF = False Then
      '2019/12/27 END
         'Modified by Morgan 2011/12/21 若收據有變更帳款類別則相同的依照列印順序合併
         'stMoneyTemp = Format("" & rs44t0("Fee5").Value, DDollar)
         stMoneyTemp = Format(dblAddAmt(5), DDollar)
         xlsAnnuity.Range("H" & m_intColumn).Value = stMoneyTemp '未扣稅額
      End If
      '2019/12/27 END
'      m_intColumn = m_intColumn + 2
'      xlsAnnuity.Range("A" & m_intColumn).Value = "共 " & .Rows - 1 & " 筆"
'   Else
'   '2022/1/19 END
'
'      With Printer
'         '"公司"
'         .CurrentX = PLeft(0)
'         .CurrentY = intY
'         Printer.Print "" & rs44t0("a0k11").Value
'         '"收款日期"
'         .CurrentX = PLeft(1)
'         .CurrentY = intY
'         Printer.Print ChangeTStringToTDateString("" & rs44t0("a0l02").Value)
'         '"收據號碼"
'         .CurrentX = PLeft(2)
'         .CurrentY = intY
'         Printer.Print "" & rs44t0("a0k01").Value
'         '"收據日期"
'         .CurrentX = PLeft(3)
'         .CurrentY = intY
'         Printer.Print ChangeTStringToTDateString("" & rs44t0("a0k02").Value)
'
'         '"案件性質"
'         .CurrentX = PLeft(9)
'         .CurrentY = intY
'         'Added by Morgan 2011/12/21
'         If rs44t0("a0k33") = "Y" Then
'            Printer.Print "" & rs44t0("a0j22").Value
'         Else
'         'end 2011/12/21
'            'Modified by Morgan 2011/12/27 取消 a0j20
'            Printer.Print "" & rs44t0("cp10N").Value
'         End If
'
'         '"申請國家"
'         .CurrentX = PLeft(10)
'         .CurrentY = intY
'         'Modified by Morgan 2011/12/30 取消 a0j21
'         Printer.Print "" & rs44t0("na03").Value
'
'         ' "扣單公司"
'         .CurrentX = PLeft(13)
'         .CurrentY = intY
'         Printer.Print "" & rs44t0("a0w04").Value
'
'         '"扣單編號"
'         .CurrentX = PLeft(14)
'         .CurrentY = intY
'         Printer.Print "" & rs44t0("a0w02").Value
'
'         '扣單備註
'         .CurrentX = PLeft(15)
'         .CurrentY = intY
'         Printer.Print "" & Left(rs44t0("a0w06").Value, 14)
'
'         'Add by Morgan 2005/5/23
'         '票期
'         .CurrentX = PLeft(16)
'         .CurrentY = intY
'         Printer.Print Format("" & rs44t0("a1p12").Value, "###/##/##")
'
'         'Added by Morgan 2011/12/21 若收據有變更帳款類別則相同的依照列印順序合併
'         Erase dblAddAmt
'
'         For idx = 1 To 7
'            dblAddAmt(idx) = Val("" & rs44t0("Fee" & idx))
'         Next
'
'         stLstRecNo = "" & rs44t0("a0k01")
'         stLstItem = "" & rs44t0("a0j22")
'         If rs44t0("a0k33") = "Y" Then
'            rs44t0.MoveNext
'            Do While Not rs44t0.EOF
'               If stLstRecNo = rs44t0("a0k01") And rs44t0("a0k33") = "Y" And stLstItem = rs44t0("a0j22") Then
'
'                  rs44t0.MovePrevious
'                  For idx = LBound(arrGrdTot) To UBound(arrGrdTot)
'                     If Not (idx = 6 And ("" & rs44t0("Fee6")) = "") Then
'                        arrSubTot1(idx) = Format(Val(arrSubTot1(idx)) + Val("" & rs44t0("Fee" & idx)))
'                        arrSubTot2(idx) = Format(Val(arrSubTot2(idx)) + Val("" & rs44t0("Fee" & idx)))
'                     End If
'                  Next
'                  rs44t0.MoveNext
'
'                  For idx = 1 To 7
'                     dblAddAmt(idx) = dblAddAmt(idx) + Val("" & rs44t0("Fee" & idx))
'                  Next
'               Else
'                  Exit Do
'               End If
'               rs44t0.MoveNext
'            Loop
'            rs44t0.MovePrevious
'         End If
'
'         '"已收扣單金額"
'         'Modified by Morgan 2011/12/21 若收據有變更帳款類別則相同的依照列印順序合併
'         'stMoneyTemp = IIf(IsNull(rs44t0("Fee6").Value), "", Format(rs44t0("Fee6").Value, DDollar))
'         stMoneyTemp = IIf(IsNull(rs44t0("Fee6").Value), "", Format(dblAddAmt(6), DDollar))
'
'         .CurrentX = PLeft(12) - 1 * BytePix - .TextWidth(stMoneyTemp)
'         .CurrentY = intY
'         Printer.Print stMoneyTemp
'
'         '"調整稅額"
'         'Modified by Morgan 2011/12/21 若收據有變更帳款類別則相同的依照列印順序合併
'         'stMoneyTemp = Format("" & rs44t0("Fee7").Value, DDollar)
'         stMoneyTemp = Format(dblAddAmt(7), DDollar)
'
'         .CurrentX = PLeft(13) - 1 * BytePix - .TextWidth(stMoneyTemp)
'         .CurrentY = intY
'         Printer.Print stMoneyTemp
'
'         '數字靠右
'         '"服務費"
'         'Modified by Morgan 2011/12/21 若收據有變更帳款類別則相同的依照列印順序合併
'         'stMoneyTemp = Format("" & rs44t0("Fee1").Value, DDollar)
'         stMoneyTemp = Format(dblAddAmt(1), DDollar)
'
'         .CurrentX = PLeft(5) - 1 * BytePix - .TextWidth(stMoneyTemp)
'         .CurrentY = intY
'         Printer.Print stMoneyTemp
'
'   '      '"可扣稅額"
'   '      'Modified by Morgan 2011/12/21 若收據有變更帳款類別則相同的依照列印順序合併
'   '      'stMoneyTemp = Format("" & rs44t0("Fee2").Value, DDollar)
'   '      stMoneyTemp = Format(dblAddAmt(2), DDollar)
'   '
'   '      .CurrentX = pLeft(6) - 1 * BytePix - .TextWidth(stMoneyTemp)
'   '      .CurrentY = intY
'   '      Printer.Print stMoneyTemp
'
'         '"收款扣繳額"
'         'Modified by Morgan 2011/12/21 若收據有變更帳款類別則相同的依照列印順序合併
'         'stMoneyTemp = Format("" & rs44t0("Fee3").Value, DDollar)
'         stMoneyTemp = Format(dblAddAmt(3), DDollar)
'
'         .CurrentX = PLeft(7) - 1 * BytePix - .TextWidth(stMoneyTemp)
'         .CurrentY = intY
'         Printer.Print stMoneyTemp
'         '"補扣繳額"
'         'Modified by Morgan 2011/12/21 若收據有變更帳款類別則相同的依照列印順序合併
'         'stMoneyTemp = Format("" & rs44t0("Fee4").Value, DDollar)
'         stMoneyTemp = Format(dblAddAmt(4), DDollar)
'
'         .CurrentX = PLeft(8) - 1 * BytePix - .TextWidth(stMoneyTemp)
'         .CurrentY = intY
'         Printer.Print stMoneyTemp
'
'         '"未扣稅額"
'         'Modify By Sindy 2019/12/27
'         If m_bolPDF = False Then
'         '2019/12/27 END
'            'Modified by Morgan 2011/12/21 若收據有變更帳款類別則相同的依照列印順序合併
'            'stMoneyTemp = Format("" & rs44t0("Fee5").Value, DDollar)
'            stMoneyTemp = Format(dblAddAmt(5), DDollar)
'            .CurrentX = PLeft(9) - 1 * BytePix - .TextWidth(stMoneyTemp)
'            .CurrentY = intY
'            Printer.Print stMoneyTemp
'         End If
'         '2019/12/27 END
'      End With
'   End If
End Sub

Private Function NewLine(Optional ByVal iLine As Integer = RowPix) As Boolean
   intY = intY + iLine
   If intY > 10000 Then
      Printer.NewPage
      Call PrintHead(stSalesName, stTitle, stCustNo, stCustName)
      NewLine = True
   End If
End Function

Private Sub PrintSubTot(ByRef arrTot() As String, Optional ByVal iTag As Integer = 1, Optional ByVal stSubDesc As String = "")
   'Modify By Sindy 2022/1/19 改Call Excel產生報表
'   If strSrvDate(1) >= Form20上線日 Then
      xlsAnnuity.Range("E" & m_intColumn & ":" & "L" & m_intColumn).Select
'      With xlsAnnuity.Selection.Borders(xlEdgeTop)
'          .LineStyle = xlContinuous
'          .ColorIndex = xlAutomatic
'          .tintandshade = 0
'          .Weight = xlThin
'      End With
      With xlsAnnuity.Selection.Borders(xlEdgeBottom)
          .LineStyle = xlContinuous
          .ColorIndex = xlAutomatic
          .tintandshade = 0
          .Weight = xlThin
      End With
      
      m_intColumn = m_intColumn + 1
      If iTag = 1 Then
         xlsAnnuity.Range("B" & m_intColumn).Value = stSubDesc & "  小計:"
      ElseIf iTag = 2 Then
         xlsAnnuity.Range("B" & m_intColumn).Value = " 合計:"
      ElseIf iTag = 3 Then
         xlsAnnuity.Range("B" & m_intColumn).Value = " 總計:"
      End If
      stMoneyTemp = Format(arrTot(1), DDollar)
      xlsAnnuity.Range("E" & m_intColumn).Value = stMoneyTemp '服務費
      stMoneyTemp = Format(arrTot(3), DDollar)
      xlsAnnuity.Range("F" & m_intColumn).Value = stMoneyTemp '收款扣繳額
      stMoneyTemp = Format(arrTot(4), DDollar)
      xlsAnnuity.Range("G" & m_intColumn).Value = stMoneyTemp '補扣繳額
      If m_bolPDF = False Then
         stMoneyTemp = Format(arrTot(5), DDollar)
         xlsAnnuity.Range("H" & m_intColumn).Value = stMoneyTemp '未扣稅額
      End If
      stMoneyTemp = IIf(arrTot(6) = "", "", Format(arrTot(6), DDollar))
      xlsAnnuity.Range("K" & m_intColumn).Value = stMoneyTemp '已收扣單金額
      stMoneyTemp = Format(arrTot(7), DDollar)
      xlsAnnuity.Range("L" & m_intColumn).Value = stMoneyTemp '調整稅額
      
      If iTag = 3 Then
         xlsAnnuity.Range("E" & m_intColumn & ":" & "L" & m_intColumn).Select
         With xlsAnnuity.Selection.Borders(xlEdgeTop)
             .LineStyle = xlContinuous
             .ColorIndex = xlAutomatic
             .tintandshade = 0
             .Weight = xlThin
         End With
         With xlsAnnuity.Selection.Borders(xlEdgeBottom)
             .LineStyle = xlContinuous
             .ColorIndex = xlAutomatic
             .tintandshade = 0
             .Weight = xlThin
         End With
         m_intColumn = m_intColumn + 1
         xlsAnnuity.Range("I" & m_intColumn).Value = "*** 結束 ***"
      End If
'   Else
'   '2022/1/19 END
'
'      If NewLine(4 * BytePix) = False Then
'         Printer.DrawStyle = vbDot
'         Printer.Line (PLeft(4), intY)-(PLeft(13), intY)
'         Call NewLine(1 * BytePix)
'      End If
'      With Printer
'         If iTag = 1 Then
'            stMoneyTemp = stSubDesc & "  小計:"
'         ElseIf iTag = 2 Then
'            stMoneyTemp = " 合計:"
'         ElseIf iTag = 3 Then
'            stMoneyTemp = " 總計:"
'         End If
'         .CurrentX = PLeft(4) - .TextWidth(stMoneyTemp)
'         .CurrentY = intY
'         Printer.Print stMoneyTemp
'
'         '數字靠右
'         '"服務費"
'         stMoneyTemp = Format(arrTot(1), DDollar)
'         .CurrentX = PLeft(5) - 1 * BytePix - .TextWidth(stMoneyTemp)
'         .CurrentY = intY
'         Printer.Print stMoneyTemp
'
'   '      '"可扣稅額"
'   '      stMoneyTemp = Format(arrTot(2), DDollar)
'   '      .CurrentX = pLeft(6) - 1 * BytePix - .TextWidth(stMoneyTemp)
'   '      .CurrentY = intY
'   '      Printer.Print stMoneyTemp
'
'         '"收款扣繳額"
'         stMoneyTemp = Format(arrTot(3), DDollar)
'         .CurrentX = PLeft(7) - 1 * BytePix - .TextWidth(stMoneyTemp)
'         .CurrentY = intY
'         Printer.Print stMoneyTemp
'
'         '"補扣繳額"
'         stMoneyTemp = Format(arrTot(4), DDollar)
'         .CurrentX = PLeft(8) - 1 * BytePix - .TextWidth(stMoneyTemp)
'         .CurrentY = intY
'         Printer.Print stMoneyTemp
'
'         '"未扣稅額"
'         'Modify By Sindy 2019/12/27
'         If m_bolPDF = False Then
'         '2019/12/27 END
'            stMoneyTemp = Format(arrTot(5), DDollar)
'            .CurrentX = PLeft(9) - 1 * BytePix - .TextWidth(stMoneyTemp)
'            .CurrentY = intY
'            Printer.Print stMoneyTemp
'         End If
'         '2019/12/27 END
'
'         '"已收扣單金額"
'         stMoneyTemp = IIf(arrTot(6) = "", "", Format(arrTot(6), DDollar))
'         .CurrentX = PLeft(12) - 1 * BytePix - .TextWidth(stMoneyTemp)
'         .CurrentY = intY
'         Printer.Print stMoneyTemp
'
'         '"調整稅額"
'         stMoneyTemp = Format(arrTot(7), DDollar)
'         .CurrentX = PLeft(13) - 1 * BytePix - .TextWidth(stMoneyTemp)
'         .CurrentY = intY
'         Printer.Print stMoneyTemp
'
'         If iTag = 3 Then
'            If NewLine(4 * BytePix) = False Then
'               Printer.DrawStyle = vbInsideSolid
'               Printer.Line (PLeft(4), intY)-(PLeft(13), intY)
'               Call NewLine(1 * BytePix)
'               .CurrentX = PLeft(9)
'               .CurrentY = intY
'               Printer.Print "*** 結束 ***"
'            End If
'         End If
'      End With
'   End If
End Sub
