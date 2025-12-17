VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc11c0 
   AutoRedraw      =   -1  'True
   Caption         =   "扣繳憑單維護"
   ClientHeight    =   5736
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   9408
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5736
   ScaleWidth      =   9408
   Begin VB.TextBox Text3 
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   720
      MaxLength       =   1
      TabIndex        =   7
      Top             =   2670
      Width           =   345
   End
   Begin VB.TextBox Text17 
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4920
      MaxLength       =   14
      TabIndex        =   9
      Top             =   2670
      Width           =   1155
   End
   Begin VB.CommandButton Command3 
      Caption         =   "整張收款單號改扣繳年度"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   5910
      Style           =   1  '圖片外觀
      TabIndex        =   50
      Top             =   2070
      Width           =   1425
   End
   Begin VB.CommandButton Command4 
      Caption         =   "收據抬頭修改"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4440
      Style           =   1  '圖片外觀
      TabIndex        =   49
      Top             =   2070
      Width           =   1425
   End
   Begin VB.TextBox txtA2802 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00C0FFFF&
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
      Left            =   4980
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   6
      Top             =   690
      Width           =   612
   End
   Begin VB.ComboBox cboSelComp 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      ItemData        =   "Frmacc11c0.frx":0000
      Left            =   8760
      List            =   "Frmacc11c0.frx":0002
      TabIndex        =   47
      Top             =   2100
      Width           =   570
   End
   Begin VB.CommandButton Command2 
      Caption         =   "全選收據公司　　 "
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7365
      TabIndex        =   46
      Top             =   2070
      Width           =   2010
   End
   Begin VB.TextBox Text16 
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
      Height          =   288
      Left            =   2340
      TabIndex        =   45
      Top             =   5100
      Width           =   960
   End
   Begin VB.TextBox Text15 
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
      Left            =   4260
      TabIndex        =   44
      Top             =   5100
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.TextBox Text1 
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
      Height          =   288
      Left            =   5940
      TabIndex        =   43
      Top             =   5100
      Width           =   804
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
      Height          =   285
      Index           =   5
      Left            =   5940
      TabIndex        =   42
      Top             =   5415
      Width           =   804
   End
   Begin VB.ComboBox CboComp 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      ItemData        =   "Frmacc11c0.frx":0004
      Left            =   1230
      List            =   "Frmacc11c0.frx":0006
      TabIndex        =   11
      Top             =   2070
      Width           =   750
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
      Height          =   288
      Index           =   4
      Left            =   8160
      TabIndex        =   40
      Top             =   5415
      Width           =   852
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
      Height          =   288
      Index           =   3
      Left            =   6765
      TabIndex        =   39
      Top             =   5415
      Width           =   804
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
      Height          =   288
      Index           =   2
      Left            =   5010
      TabIndex        =   38
      Top             =   5415
      Width           =   900
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
      Height          =   288
      Index           =   1
      Left            =   3315
      TabIndex        =   37
      Top             =   5415
      Width           =   828
   End
   Begin VB.ComboBox cboSubTotal 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      ItemData        =   "Frmacc11c0.frx":0008
      Left            =   1620
      List            =   "Frmacc11c0.frx":000A
      TabIndex        =   35
      Top             =   5415
      Width           =   750
   End
   Begin VB.CommandButton cmdSearch1 
      Caption         =   "相似尋找"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   6450
      TabIndex        =   33
      Top             =   720
      Width           =   1100
   End
   Begin VB.TextBox txtSales 
      BackColor       =   &H00C0FFFF&
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
      Left            =   5700
      MaxLength       =   6
      TabIndex        =   4
      Top             =   360
      Width           =   1065
   End
   Begin VB.CommandButton cmdLikeSearch 
      Caption         =   "相似搜尋"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   8085
      TabIndex        =   1
      Top             =   60
      Width           =   1100
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "尋找"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   8085
      TabIndex        =   30
      Top             =   390
      Width           =   1100
   End
   Begin VB.TextBox txtCustNo 
      BackColor       =   &H00C0FFFF&
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
      Index           =   1
      Left            =   3090
      MaxLength       =   9
      TabIndex        =   3
      Top             =   360
      Width           =   1572
   End
   Begin VB.TextBox txtCustNo 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   1170
      MaxLength       =   9
      TabIndex        =   2
      Top             =   360
      Width           =   1572
   End
   Begin VB.TextBox Text14 
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
      Height          =   288
      Left            =   3315
      TabIndex        =   29
      Top             =   5100
      Width           =   828
   End
   Begin VB.TextBox Text13 
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
      Height          =   288
      Left            =   5010
      TabIndex        =   28
      Top             =   5100
      Width           =   900
   End
   Begin VB.TextBox Text12 
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
      Left            =   7680
      TabIndex        =   27
      Top             =   5100
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.TextBox Text5 
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
      Height          =   288
      Left            =   8160
      TabIndex        =   26
      Top             =   5100
      Width           =   852
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "Frmacc11c0.frx":000C
      Height          =   2010
      Left            =   45
      TabIndex        =   12
      Top             =   3060
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   3535
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
      ColumnCount     =   18
      BeginProperty Column00 
         DataField       =   "a1v14"
         Caption         =   "選擇"
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
         DataField       =   "a1v03"
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
         DataField       =   "R11C04"
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
      BeginProperty Column03 
         DataField       =   "a1v02"
         Caption         =   "收據編號"
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
         DataField       =   "a1v04"
         Caption         =   "應扣繳額"
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
      BeginProperty Column05 
         DataField       =   "a1v05"
         Caption         =   "部份收款"
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
         DataField       =   "a1v06"
         Caption         =   "已扣繳額"
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
      BeginProperty Column07 
         DataField       =   "a1v07"
         Caption         =   "未扣繳額"
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
      BeginProperty Column08 
         DataField       =   "a1v08"
         Caption         =   "退費否"
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
         DataField       =   "a1v09"
         Caption         =   "扣繳年度"
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
         DataField       =   "R11C02"
         Caption         =   "收款日期"
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
         DataField       =   "a1v10"
         Caption         =   "調整稅款"
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
         DataField       =   "a1v17"
         Caption         =   "發票號碼"
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
      BeginProperty Column13 
         DataField       =   "a1v12"
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
      BeginProperty Column14 
         DataField       =   "a1v13"
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
      BeginProperty Column15 
         DataField       =   "R11C05"
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
      BeginProperty Column16 
         DataField       =   "a1v01"
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
      BeginProperty Column17 
         DataField       =   "R11C06"
         Caption         =   "收據別"
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
            ColumnWidth     =   492.095
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   492.095
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   852.095
         EndProperty
         BeginProperty Column03 
            Locked          =   -1  'True
            ColumnWidth     =   912.189
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   852.095
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   875.906
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   875.906
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   840.189
         EndProperty
         BeginProperty Column08 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   671.811
         EndProperty
         BeginProperty Column09 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   828.284
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   852.095
         EndProperty
         BeginProperty Column11 
            Alignment       =   1
            ColumnWidth     =   815.811
         EndProperty
         BeginProperty Column12 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column13 
            Locked          =   -1  'True
            ColumnWidth     =   1091.906
         EndProperty
         BeginProperty Column14 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column15 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column16 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column17 
            Object.Visible         =   0   'False
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   315
      Left            =   45
      Top             =   3000
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc11c0.frx":0021
      Height          =   990
      Left            =   30
      TabIndex        =   25
      Top             =   1080
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   1736
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   16
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
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
         DataField       =   "a0w04"
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
         DataField       =   "a0w02"
         Caption         =   "扣繳憑單編號"
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
      BeginProperty Column02 
         DataField       =   "a0w05"
         Caption         =   "扣單稅額"
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
      BeginProperty Column03 
         DataField       =   "a0w16"
         Caption         =   "給付總額"
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
      BeginProperty Column04 
         DataField       =   "a0w14"
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
      BeginProperty Column05 
         DataField       =   "a0w06"
         Caption         =   "備註"
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
            ColumnWidth     =   815.811
         EndProperty
         BeginProperty Column01 
            Alignment       =   1
            ColumnWidth     =   1488.189
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1031.811
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1307.906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   3767.811
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   3780.284
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
      Left            =   30
      Top             =   1515
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
   Begin VB.TextBox Text11 
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3060
      MaxLength       =   14
      TabIndex        =   8
      Top             =   2670
      Width           =   915
   End
   Begin VB.TextBox Text10 
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
      Height          =   288
      Left            =   6765
      TabIndex        =   23
      Top             =   5100
      Width           =   804
   End
   Begin VB.TextBox Text9 
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
      Left            =   2370
      TabIndex        =   22
      Top             =   5400
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.TextBox Text8 
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
      Height          =   288
      Left            =   1635
      TabIndex        =   21
      Top             =   5100
      Visible         =   0   'False
      Width           =   735
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
      Height          =   276
      Left            =   4050
      Picture         =   "Frmacc11c0.frx":0036
      Style           =   1  '圖片外觀
      TabIndex        =   13
      ToolTipText     =   "取消"
      Top             =   2100
      Width           =   350
   End
   Begin VB.TextBox Text7 
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
      Left            =   2490
      TabIndex        =   19
      Top             =   2070
      Width           =   1524
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
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
      Left            =   1170
      MaxLength       =   3
      TabIndex        =   5
      Top             =   720
      Width           =   612
   End
   Begin MSForms.TextBox Text6 
      Height          =   315
      Left            =   6600
      TabIndex        =   10
      Top             =   2670
      Width           =   2655
      VariousPropertyBits=   -1466939365
      ScrollBars      =   2
      Size            =   "4683;556"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblSales 
      Height          =   195
      Left            =   6870
      TabIndex        =   53
      Top             =   420
      Width           =   690
      VariousPropertyBits=   268435475
      Caption         =   "lblSales"
      Size            =   "1217;344"
      FontName        =   "新細明體"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin MSForms.ComboBox cboTitle 
      Height          =   315
      Left            =   1170
      TabIndex        =   0
      Top             =   30
      Width           =   6780
      VariousPropertyBits=   679495707
      BackColor       =   12648447
      DisplayStyle    =   3
      Size            =   "11959;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label14 
      BackStyle       =   0  '透明
      Caption         =   "公司名稱"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   1110
      TabIndex        =   52
      Top             =   2700
      Width           =   945
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '透明
      Caption         =   "給付總額"
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
      Left            =   4050
      TabIndex        =   51
      Top             =   2700
      Width           =   945
   End
   Begin VB.Label Label16 
      BackStyle       =   0  '透明
      Caption         =   "最近扣繳確認年度"
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
      Left            =   3090
      TabIndex        =   48
      Top             =   705
      Width           =   1875
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '透明
      Caption         =   "收據公司別"
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
      Left            =   60
      TabIndex        =   41
      Top             =   2100
      Width           =   1185
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '透明
      Caption         =   "公司別小計"
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
      Left            =   360
      TabIndex        =   36
      Top             =   5415
      Width           =   1230
   End
   Begin VB.Label Label10 
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
      Left            =   4750
      TabIndex        =   34
      Top             =   420
      Width           =   900
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "客戶編號"
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
      Left            =   135
      TabIndex        =   32
      Top             =   420
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2850
      TabIndex        =   31
      Top             =   420
      Width           =   255
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF0000&
      Height          =   1050
      Left            =   45
      Top             =   0
      Width           =   9300
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "扣單稅額"
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
      Left            =   2190
      TabIndex        =   24
      Top             =   2700
      Width           =   945
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   -120
      Top             =   4872
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "合計"
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
      Left            =   1035
      TabIndex        =   20
      Top             =   5100
      Width           =   510
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   420
      Left            =   60
      Top             =   2625
      Width           =   9300
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "合計"
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
      Left            =   2010
      TabIndex        =   18
      Top             =   2100
      Width           =   645
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "備註"
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
      Left            =   6150
      TabIndex        =   17
      Top             =   2700
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "公司別"
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
      TabIndex        =   16
      Top             =   2700
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "扣繳年度"
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
      Left            =   135
      TabIndex        =   15
      Top             =   705
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "收據抬頭"
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
      Left            =   135
      TabIndex        =   14
      Top             =   90
      Width           =   975
   End
End
Attribute VB_Name = "Frmacc11c0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/10/8 Form2.0已修改
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/11/26 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit

Public adoaccsum As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Public adoadodc2 As New ADODB.Recordset
Public adoacc0w0 As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
Public adocheck As New ADODB.Recordset
Dim strSql As String
Dim strSQL1 As String
Dim strSQL2 As String
Dim strRNo As String 'A1V01.收文號
Dim intY As Integer
Dim strTaxNo As String
Dim intCconfirm As Integer
Dim intRefresh As Integer
Dim arrSum(0 To 9) As String
'Add by Morgan 2006/5/2
'避免單一資料同時被兩程式處理
Dim m_NewKey As String
Dim m_OldKey As String
'Add by Morgan 2006/5/17
Dim m_Title As String '扣單抬頭
Dim m_bolExact As Boolean '前次搜尋方式
Dim m_bClick As Boolean '是否有點選
Dim m_yynotsub As Integer    'add by sonia 2023/11/24 轉年度服務費超過2萬未扣繳收據張數
Public bolUpdData As Boolean 'Add By Sindy 2017/10/31

Private Sub cboSelComp_Click()
   If cboSelComp.Text <> "" Then
      If cboSelComp <> Text3 Then
         Text3 = cboSelComp
         Text3_Change
      End If
   End If
End Sub

'Add By Sindy 2017/10/31
'整張收款單號改扣繳年度
Private Sub Command3_Click()
Dim rsRead As New ADODB.Recordset
Dim sqlB As String
Dim intB As Integer, i As Integer
Dim rstClone As New ADODB.Recordset
   
   Me.Tag = "": bolUpdData = False
   If Adodc2.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   Me.Enabled = False
   Set rstClone = Adodc2.Recordset.Clone
   With rstClone
      .MoveFirst
      Do While Not .EOF
         If ("" & .Fields("A1V14").Value) = "Y" Then
            Me.Tag = .Fields("a1v02").Value
            Exit Do
         End If
         .MoveNext
      Loop
      .MoveFirst
   End With
   Set rstClone = Nothing
   Adodc2.Recordset.Resync
   Me.Enabled = True
   If Me.Tag = "" Then Exit Sub
   
   sqlB = "select '',a1v03,a0k02,a0k01,a0k04,a1v09,a1v06,sum(a1v07)" & _
          " From acc0m0 m2, acc0l0, acc0k0, acc1v0,(select a0m01,a0m02 from acc0m0 where a0m02='" & Me.Tag & "') m1" & _
          " where m2.a0m01=m1.a0m01 and m2.a0m01=a0l01(+) and m2.a0m02=a0k01(+) and m2.a0m02=a1v02" & _
          " group by a1v03,a0k02,a0k01,a0k04,a1v09,a1v06" & _
          " Union" & _
          " select '',a1v03,a1k02,a1k01,a1k35,a1v09,a1v06,sum(a1v07)" & _
          " From acc0z0 z2, acc0y0, acc1k0, acc1v0,(select a0z01,a0z02 from acc0z0 where a0z02='" & Me.Tag & "') z1" & _
          " where z2.a0z01=z1.a0z01 and z2.a0z01=a0y01(+) and z2.a0z02=a1k01(+) and z2.a0z02=a1v02" & _
          " group by a1v03,a1k02,a1k01,a1k35,a1v09,a1v06"
   intB = 0
   Set rsRead = ClsLawReadRstMsg(intB, sqlB)
   If intB = 1 Then
      Set Frmacc11c0_1.grdDataList.Recordset = rsRead
      '最大的收款日期
      strExc(0) = "select max(a0l02)" & _
                  " From acc0m0 m2, acc0l0, acc0k0, acc1v0,(select a0m01,a0m02 from acc0m0 where a0m02='" & Me.Tag & "') m1" & _
                  " where m2.a0m01=m1.a0m01 and m2.a0m01=a0l01(+) and m2.a0m02=a0k01(+) and m2.a0m02=a1v02" & _
                  " group by a1v03,a0k02,a0k01,a0k04,a1v09,a1v06" & _
                  " Union" & _
                  " select max(a0y02)" & _
                  " From acc0z0 z2, acc0y0, acc1k0, acc1v0,(select a0z01,a0z02 from acc0z0 where a0z02='" & Me.Tag & "') z1" & _
                  " where z2.a0z01=z1.a0z01 and z2.a0z01=a0y01(+) and z2.a0z02=a1k01(+) and z2.a0z02=a1v02" & _
                  " group by a1v03,a1k02,a1k01,a1k35,a1v09,a1v06"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         Frmacc11c0_1.Label1 = "" & RsTemp.Fields(0)
         If Frmacc11c0_1.Label1 <> "" Then
            Frmacc11c0_1.Label1 = ChangeTStringToTDateString(Frmacc11c0_1.Label1)
         End If
      Else
         Frmacc11c0_1.Label1 = ""
      End If
      '***
      Frmacc11c0_1.Show vbModal
      '有異動資料,重新查詢
      If bolUpdData = True Then
         Call cmdSearch_Click
      End If
   End If
End Sub

'Added by Sindy 2017/10/31
'點選呼叫"收據抬頭修改"
Private Sub Command4_Click()
Dim rstClone As New ADODB.Recordset
   
   strItemNo = ""
   If Adodc2.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   Me.Enabled = False
   Set rstClone = Adodc2.Recordset.Clone
   With rstClone
      .MoveFirst
      Do While Not .EOF
         If ("" & .Fields("A1V14").Value) = "Y" Then
            strItemNo = .Fields("a1v02").Value
            Exit Do
         End If
         .MoveNext
      Loop
      .MoveFirst
   End With
   Set rstClone = Nothing
   Adodc2.Recordset.Resync
   Me.Enabled = True
   If strItemNo = "" Then Exit Sub
   
   strTitle = Me.Name
   If Mid(strItemNo, 1, 1) = "E" Then
      tool14_enabled
      MenuDisabled
      Frmacc1140.Show
      Me.Enabled = False
   Else
      MsgBox "請點選收據/請款單資料..."
      strItemNo = ""
      strTitle = ""
   End If
End Sub

'Add by Morgan 2008/2/18
Private Sub Command2_Click()
Dim bolAlv05isY As Boolean 'Add By Sindy 2023/3/28
   
   bolAlv05isY = False 'Add By Sindy 2023/3/28
   '設定值 Y/NULL
   Dim strValue As String, rstClone As New ADODB.Recordset
   Dim bolOK As Boolean
   
   If Adodc2.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   
   Me.Enabled = False
   Screen.MousePointer = vbHourglass
   Frmacc0000.StatusBar1.Panels(1).Text = "資料更新中...."
   Set rstClone = Adodc2.Recordset.Clone
   strValue = ""
   With rstClone
   .MoveFirst
   Do While Not .EOF
     If cboSelComp = "" Then
        bolOK = True
     ElseIf "" & .Fields("A1V03") = cboSelComp Then
        bolOK = True
     Else
        bolOK = False
     End If
     If bolOK And ("" & .Fields("A1V14").Value) <> "Y" Then
        strValue = "Y"
        Exit Do
     End If
     'Add By Sindy 2023/3/28
     If "" & .Fields("A1V05").Value = "Y" Then
         bolAlv05isY = True
     End If
     '2023/3/28 END
     .MoveNext
   Loop
   
   .MoveFirst
   Do While Not .EOF
     If cboSelComp = "" Then
        bolOK = True
     ElseIf "" & .Fields("A1V03") = cboSelComp Then
        bolOK = True
     Else
        bolOK = False
     End If
     
     If bolOK = True And ("" & .Fields("A1V14").Value) <> strValue Then
        .Fields("A1V14").Value = strValue
     End If
     .MoveNext
   Loop
   .UpdateBatch
   End With
   Set rstClone = Nothing
   Adodc2.Recordset.Resync
   SumShow3
   Frmacc0000.StatusBar1.Panels(1).Text = "按 尋找 調出補扣繳資料 / " & MsgText(98)
   Screen.MousePointer = vbDefault
   
   'Add By Sindy 2023/3/28 部分收款=Y者,彈提醒"尾款未收齊"
   If bolAlv05isY = True Then
      MsgBox "尾款未收齊", vbInformation
   End If
   '2023/3/28 END
   
   Me.Enabled = True
End Sub

Private Sub DataGrid2_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyInsert, vbKeyF12, vbKeyEscape
         KeyDefine KeyCode
         
      Case Else
         'Modified by Morgan 2019/4/9 +可以輸入負值--瑞婷
         If Not (KeyCode > 47 And KeyCode < 58) And Not (KeyCode > 95 And KeyCode < 106) And KeyCode <> vbKeyBack And KeyCode <> vbKeyDelete And KeyCode <> vbKeyReturn And KeyCode <> 189 And KeyCode <> 109 Then
            KeyCode = 0
         End If
   End Select
End Sub

Private Sub Form_Activate()
Dim tmpArr As Variant
   
   strFormLink = ""
   strCon1 = ""
   strCon2 = ""
   lngTotal = 0
   If intRefresh = 1 Then
      AdodcRefresh
      AdodcClear
      intRefresh = 0
   End If
   
   strFormName = Name
   'Add By Sindy 2015/8/11
   If strItemNo = MsgText(601) Then
      Exit Sub
   Else
      strExc(0) = "SELECT * FROM acc0k0 WHERE a0k01='" & Left(strItemNo, 9) & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         cboTitle = "" & RsTemp.Fields("A0K04")
         txtCustNo(0) = "" & RsTemp.Fields("A0K03")
         txtCustNo(1) = "" & RsTemp.Fields("A0K03")
         Text2 = "" & RsTemp.Fields("A0K16")
         Call cmdSearch_Click
      Else
         '國外請款資料時,只要帶出收據抬頭,扣繳年度,扣繳憑單資料
         strExc(0) = "SELECT * FROM acc1k0 WHERE a1k01='" & Left(strItemNo, 9) & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            ClearAll
            tmpArr = Split(strItemNo, ",")
            For intI = 1 To UBound(tmpArr)
               If intI = 1 Then cboTitle = tmpArr(intI)
               If intI = 2 Then Text2 = tmpArr(intI)
            Next intI
            If adoadodc1.State = adStateOpen Then adoadodc1.Close
            adoadodc1.CursorLocation = adUseClient
            adoadodc1.Open "select * from acc0w0 where a0w15 is null and a0w01 = " & Val(Text2) & " and instr(a0w03, '" & cboTitle & "') > 0 order by a0w02 desc", adoTaie, adOpenDynamic, adLockBatchOptimistic
            Adodc1.Recordset.Requery
         End If
      End If
   End If
   '2015/8/11 END
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Call PUB_SaveTrackMode(0, KeyCode) 'Added by Morgan 2023/2/7
End Sub

Private Sub Form_Load()
Dim i As Integer

   PUB_InitForm Me, 9500, 6350
   lblSales.Caption = ""
   Me.Label14 = ""
   
   strItemNo = MsgText(601)
   strFormName = Name 'Add By Sindy 2015/8/11
   '預設年度改判斷4月
   If Val(Right(strSrvDate(2), 4)) >= 401 Then
      Text2 = strSrvDate(2) \ 10000
   Else
      Text2 = strSrvDate(2) \ 10000 - 1
   End If
   OpenTable
   
   Frmacc0000.StatusBar1.Panels(1).Text = "按 尋找 調出補扣繳資料 / " & MsgText(98)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add by Morgan 2006/5/2
   '請除鎖定資料
   If PUB_GetLock("", m_OldKey) = False Then
      Cancel = 1
      Exit Sub
   End If
   '2006/5/2 end
   'ADD BY SONIA 2014/4/21 若無LOCKREC資料則刪除工作檔ACCRPT11C
   If adoquery.State = adStateOpen Then adoquery.Close
   adoquery.CursorLocation = adUseClient
   adoquery.Open "select * from LOCKREC", adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount = 0 Then
      If adoacc0w0.State = adStateOpen Then adoacc0w0.Close
      adoacc0w0.CursorLocation = adUseClient
      adoacc0w0.Open "delete accrpt11c", adoTaie, adOpenStatic, adLockReadOnly
      If adoacc0w0.State = adStateOpen Then adoacc0w0.Close
   End If
   If adoquery.State = adStateOpen Then adoquery.Close
   '2014/4/21 END
   Call fnUnCheck
   StatusClear
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc11c0 = Nothing
End Sub

Private Sub cboTitle_GotFocus()
   'edit by nickc 2007/06/11  切換輸入法改用API
   'cboTitle.IMEMode = 1
   OpenIme
End Sub

Private Sub cboTitle_KeyPress(KeyAscii As MSForms.ReturnInteger)
'   If txtCustNo(0) <> "" Or txtCustNo(1) <> "" Then
'      txtCustNo(0) = "": txtCustNo(1) = ""
'      txtSales = "": lblSales = ""
'   End If
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
      'Add by Morgan 2008/2/18
      If txtCustNo(1).Text = "" Then
         strExc(0) = txtCustNo(0).Text
         For intI = cboTitle.ListIndex + 1 To cboTitle.ListCount - 1
            If InStr(cboTitle.List(intI), cboTitle.List(0)) = 1 Then
               'Modify By Sindy 2016/2/18 讀取最大客戶編號
               If strExc(0) < Right(cboTitle.List(intI), 9) Then
               '2016/2/18 END
                  strExc(0) = Right(cboTitle.List(intI), 9)
               End If
            Else
               Exit For
            End If
         Next
         txtCustNo(1).Text = strExc(0)
      End If
      cboTitle.ListIndex = 0
      cmdSearch_Click
   End If
   txtA2802 = PUB_GetA2802LastYear(Trim(cboTitle.Text)) 'Add By Sindy 2013/11/27
End Sub

Private Sub cboTitle_Validate(Cancel As Boolean)
   If CheckLen(Label1, cboTitle, 100) = MsgText(603) Then
      Cancel = True
      Exit Sub
   End If
   'edit by nickc 2007/06/11  切換輸入法改用API
   If Cancel = False Then CloseIme
End Sub

Public Sub cmdSearch_Click()
   If ChkConOK = True Then
      ProcessData
   End If
   Frmacc0000.StatusBar1.Panels(1).Text = "按 尋找 調出補扣繳資料 / " & MsgText(98)
End Sub

Private Sub cmdSearch1_Click()
   If ChkConOK = True Then
      ProcessData (False)
   End If
   Frmacc0000.StatusBar1.Panels(1).Text = "按 尋找 調出補扣繳資料 / " & MsgText(98)
End Sub

Private Sub cboComp_Click()
   Screen.MousePointer = vbHourglass
   AdodcRefresh m_bolExact
   Screen.MousePointer = vbDefault
End Sub

Private Sub cboSubTotal_Click()
    Dim ArrStr() As String, ii As Integer
    If cboSubTotal.ListIndex <> -1 Then
        ArrStr = Split(arrSum(cboSubTotal.ListIndex), ",")
        For ii = 1 To 5
            TxtSum(ii) = ArrStr(ii)
        Next
    End If
End Sub

Private Sub Text11_GotFocus()
   TextInverse Text11
End Sub

'Add By Sindy 2020/5/4
Private Sub Text17_GotFocus()
   TextInverse Text17
End Sub

'Add by Morgan 2006/2/17 備註預設勾選扣單合計--辜
Private Sub Text16_Change()
   'Modify By Sindy 2020/5/4 新增給付總額欄位
   'Text6.Text = Text16.Text
   Text17.Text = Text16.Text
   '2020/5/4 END
End Sub

Private Sub Text2_GotFocus()
   'edit by nickc 2007/06/11  切換輸入法改用API
   'Text2.IMEMode = 2
   CloseIme
   TextInverse Text2
End Sub

Private Sub Text3_Change()
   Label14 = CompNameQuery(Text3, "4")
End Sub

Private Sub Text3_GotFocus()
   TextInverse Text3
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text6_GotFocus()
   TextInverse Text6
End Sub

Private Sub Command1_Click()
   AdodcDelete
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

Private Sub cmdLikeSearch_Click()
   If cboTitle.Text = "" Then
      MsgBox "請輸入收據抬頭！", vbCritical
   Else
      'Modify By Sindy 2021/12/29 從cboTitle_KeyPress搬過來
      If txtCustNo(0) <> "" Or txtCustNo(1) <> "" Or cboTitle.ListCount > 0 Then
         txtCustNo(0) = "": txtCustNo(1) = ""
         txtSales = "": lblSales = ""
      End If
      '2021/12/29 END
      
      'Modify by Morgan 2007/10/2 改呼叫共用函數
      'AddItem2CboTitle
      'Modify by Sindy 2013/12/30
      PUB_AddItem2CboTitle cboTitle, txtCustNo(0), txtCustNo(1), Text2, True
      'end 2007/10/2
   End If
   txtA2802 = PUB_GetA2802LastYear(Trim(cboTitle.Text)) 'Add By Sindy 2013/11/27
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

Private Sub DataGrid1_SelChange(Cancel As Integer)
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   AdodcShow
   strTaxNo = Adodc1.Recordset.Fields("a0w02").Value
End Sub

Private Sub DataGrid2_AfterColUpdate(ByVal ColIndex As Integer)
   If Adodc2.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   
   'Add By Sindy 2015/12/10
   With DataGrid2
      Select Case ColIndex
      Case 0, 8 '個人不可扣繳
         If Format(.Columns(ColIndex)) <> "" Then
            If PUB_ChkIsPerson(.Columns(3)) = True Then
              .Columns(ColIndex).Value = ""
            End If
         End If
      End Select
   End With
   '2015/12/10 END
   
   Adodc2.Recordset.UpdateBatch 'Modified by Morgan 2024/1/12 恢復要執行
   Adodc2.Recordset.Resync 'Added by Morgan 2024/1/12
   SumShow3
End Sub

Private Sub DataGrid2_Click()
   m_bClick = True
End Sub

Private Sub DataGrid2_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim strENo As String, ii As Integer, intCurCol As Integer 'Add By Sindy 2023/6/16
   
   If Adodc2.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   
   If DataGrid2.row < 0 Then Exit Sub 'Added by Morgan 2024/1/12
   
   '註:
   'Adodc2 從 1 開始
   'DataGrid2 從 0 開始
   If m_bClick = True Then
      With DataGrid2
      If DataGrid2.Columns(3).Text <> "" Then
         'Add By Sindy 2023/6/16
         strENo = DataGrid2.Columns(3).Text '收據號碼
         'intCurCol = DataGrid2.row
         intCurCol = Adodc2.Recordset.AbsolutePosition '目前列數
         '2023/6/16 END
         
         If .col = 0 Or .col = 8 Then
            'Add By Sindy 2015/12/10
            '個人不可扣繳
            If PUB_ChkIsPerson(DataGrid2.Columns(3).Text) = True Then
               .Text = ""
               Exit Sub
            End If
            '2015/12/10 END
            
            If .Text = "Y" Then
               .Text = ""
               'Add By Sindy 2023/6/16
               Adodc2.Recordset.MoveFirst
               Do While Adodc2.Recordset.EOF = False
                  If Adodc2.Recordset.Fields("a1v02").Value = strENo Then
                     Adodc2.Recordset.Fields("a1v14").Value = ""
                  End If
                  Adodc2.Recordset.MoveNext
               Loop
'                  For ii = 0 To Adodc2.Recordset.RecordCount - 1
'                     DataGrid2.row = ii
'                     If DataGrid2.Columns(3).Text = strENo Then
'                        DataGrid2.Columns(8).Text = ""
'                     End If
'                  Next ii
               '2023/6/16 END
            Else
               .Text = "Y"
               
               'Add By Sindy 2023/6/16
               Adodc2.Recordset.MoveFirst
               Do While Adodc2.Recordset.EOF = False
                  If Adodc2.Recordset.Fields("a1v02").Value = strENo Then
                     Adodc2.Recordset.Fields("a1v14").Value = "Y"
                  End If
                  Adodc2.Recordset.MoveNext
               Loop
'                  For ii = 0 To Adodc2.Recordset.RecordCount - 1
'                     DataGrid2.row = ii
'                     If DataGrid2.Columns(3).Text = strENo Then
'                        DataGrid2.Columns(8).Text = "Y"
'                     End If
'                  Next ii
               '2023/6/16 END
            End If
            'Add By Sindy 2023/6/16
            'DataGrid2.row = intCurCol
            Adodc2.Recordset.AbsolutePosition = intCurCol
            '2023/6/16 END
            Adodc2.Recordset.UpdateBatch 'Modified by Morgan 2024/1/12 恢復要執行
            Adodc2.Recordset.Resync 'Added by Morgan 2024/1/12
            SumShow3
            
            'Add By Sindy 2023/3/28 部分收款=Y者,彈提醒"尾款未收齊"
            If DataGrid2.Columns(0).Text = "Y" Then
               If DataGrid2.Columns(5).Text = "Y" Then
                  MsgBox "尾款未收齊", vbInformation
               End If
            End If
            '2023/3/28 END
            .col = 10
         End If
      End If
      End With
   End If
   m_bClick = False
End Sub

Private Sub DataGrid2_HeadClick(ByVal ColIndex As Integer)
   '設定值 Y/NULL
   Dim strValue As String, rstClone As New ADODB.Recordset
   If Adodc2.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   If ColIndex = 0 Then
       Me.Enabled = False
       Screen.MousePointer = vbHourglass
       Frmacc0000.StatusBar1.Panels(1).Text = "資料更新中...."
       Set rstClone = Adodc2.Recordset.Clone
       strValue = ""
       With rstClone
       .MoveFirst
       Do While Not .EOF
           If ("" & .Fields("A1V14").Value) <> "Y" Then
               strValue = "Y"
               Exit Do
           End If
           .MoveNext
       Loop
       .MoveFirst
       Do While Not .EOF
           If ("" & .Fields("A1V14").Value) <> strValue Then
              'Add By Sindy 2017/2/14
              If PUB_ChkIsPerson(.Fields("A1V02").Value, False) = True Then '收據編號
               .Fields("A1V14").Value = ""
              Else
              '2017/2/14 END
               .Fields("A1V14").Value = strValue
              End If
           End If
           .MoveNext
       Loop
       .UpdateBatch
       End With
       Set rstClone = Nothing
       Adodc2.Recordset.Resync
       SumShow3
       Frmacc0000.StatusBar1.Panels(1).Text = "按 尋找 調出補扣繳資料 / " & MsgText(98)
       Screen.MousePointer = vbDefault
       Me.Enabled = True
   End If
End Sub

'========================================================================================================================

'*************************************************
'  重新整理 Adodc 之資料
'
'*************************************************
Private Sub AdodcRefresh(Optional ByVal bolExact As Boolean = True)
Dim strName As String
Dim strA1V02 As String   'add by sonia 2014/3/20

   'Add by Morgan 2005/2/25
   If fnUnCheck = False Then Exit Sub
   
   m_bolExact = bolExact
   
On Error GoTo Checking
   strName = ""
   strA1V02 = ""   'add by sonia 2014/3/20
   Screen.MousePointer = vbHourglass

   strSql = MsgText(601)
   strSQL1 = MsgText(601)
'   strSQL2 = MsgText(601)

   If m_Title <> MsgText(601) Then
      If bolExact = True Then
         strSql = " and a0k04= '" & m_Title & "'"
      'Modify By Sindy 2020/2/13 工業技術研究院 or 財團法人工業技術研究院 都要查的出來
      Else
         'strSql = " and instrb(a0k04, '" & m_Title & "') = 1"
         strSql = " and a0k04 like '%" & m_Title & "%'"
      End If
      'strSQL1 = " and a1k35 = '" & m_Title & "'"
      If bolExact = True Then
         strSQL1 = " and a1k35 = '" & m_Title & "'"
      Else
         'strSQL1 = " and instrb(a1k35, '" & m_Title & "') = 1"
         strSQL1 = " and a1k35 like '%" & m_Title & "%'"
      End If
      '2020/2/13 END
'      strSQL2 = " and fa04 = '" & m_Title & "'"
   End If

   If txtCustNo(0) <> "" Then
      strSql = strSql & " and a0k03>='" & txtCustNo(0).Text & "'"
   End If
   If txtCustNo(1) <> "" Then
      strSql = strSql & " and a0k03<='" & txtCustNo(1).Text & "'"
   End If
         
   If txtSales <> "" Then
      strSql = strSql & " and a0k20||''='" & txtSales & "'"
   End If

   If Text2 <> MsgText(601) Then
      strSql = strSql & " and a0k16 = " & Val(Text2) & ""
'      strSQL1 = strSQL1 & " and a1v09 = " & Val(Text2) & ""
'      strSQL2 = strSQL2 & " and a1v09 = " & Val(Text2) & ""
      strSQL1 = strSQL1 & " and a1v09 = " & Val(Text2)
   End If
   '改由cboComp控制公司別條件
   strSql = strSql & " and a0k11<>'J'" 'Add By Sindy 2013/12/30
    If cboComp <> MsgText(601) Then
        strSql = strSql & " and a0k11 = '" & cboComp & "'"
    End If
    
   If adoadodc1.State = adStateOpen Then adoadodc1.Close
   adoadodc1.CursorLocation = adUseClient
'Modify by Morgan 2006/5/17
'   adoadodc1.Open "select * from acc0w0 where a0w15 is null and a0w01 = " & Val(Text2) & " and instr(a0w03, '" & cboTitle & "') > 0 order by a0w02 desc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoadodc1.Open "select * from acc0w0 where a0w15 is null and a0w01 = " & Val(Text2) & " and instr(a0w03, '" & m_Title & "') > 0 order by a0w02 desc", adoTaie, adOpenDynamic, adLockBatchOptimistic
'end 2006/5/17

   Adodc1.Recordset.Requery
   SumShow1 'Modify By Sindy 2020/5/6
   
   adoquery.CursorLocation = adUseClient
'Modify by Morgan 2005/3/1 不判斷是否已被選取改判斷未開扣單
'Modified by Morgan 2013/2/1 +a1v02
'Modify By Sindy 2015/5/4 +acc1k0
'"select a1v01,a1v02 from acc1k0, acc1v0, customer where a1k01 = a1v02 and substr(a1k28, 1, 8) = cu01 (+) and substr(a1k28, 9, 1) = cu02 (+) and a1v15 is null and (a1v17 is null or substr(a1v17, 1, 1) = 'E') and (a1k12 is null or a1k12 = 0) and a1k25 is null" & strSQL1 & " union " & _
'"select a1v01,a1v02 from acc1k0, acc1v0, fagent where a1k01 = a1v02 and substr(a1k28, 1, 8) = fa01 (+) and substr(a1k28, 9, 1) = fa02 (+) and a1v15 is null and (a1v17 is null or substr(a1v17, 1, 1) = 'E') and (a1k12 is null or a1k12 = 0) and a1k25 is null" & strSQL2
   adoquery.Open "select a1v01,a1v02 from acc0k0, acc1v0 where a0k01 = a1v02 and a1v15 is null and (a1v17 is null or substr(a1v17, 1, 1) = 'E') and (a0k09 is null or a0k09 = 0) and (a0k10 is null or a1v04 <> 0)" & strSql & " union " & _
                 "select a1v01,a1v02 from acc1k0, acc1v0 where a1k01 = a1v02 and a1v15 is null and (a1v17 is null or substr(a1v17, 1, 1) = 'E') and (a1k12 is null or a1k12 = 0) and a1k25 is null" & strSQL1 & _
                 " order by a1v01 asc", adoTaie, adOpenStatic, adLockReadOnly
'2005/3/1 end
   
   Do While adoquery.EOF = False
      'Modified by Morgan 2013/2/1
      'strName = strName & "'" & adoquery.Fields("a1v01").Value & "', "
      strName = strName & "'" & adoquery.Fields("a1v01").Value & adoquery.Fields("a1v02").Value & "', "
      'Add By Sindy 2020/2/13
      'Modify By Sindy 2021/3/17 "'" & +  & "'"
      If InStr(strA1V02, "'" & adoquery.Fields("a1v02").Value & "'") = 0 Then
      '2020/2/13 END
         strA1V02 = strA1V02 & "'" & adoquery.Fields("a1v02").Value & "', "   'add by sonia 2014/3/20
      End If
      adoquery.MoveNext
   Loop
   If adoquery.State = adStateOpen Then adoquery.Close
   
   If strName <> "" Then
      strName = Mid(strName, 1, Len(strName) - 2)
      strA1V02 = Mid(strA1V02, 1, Len(strA1V02) - 2)  'add by sonia 2014/3/20
   Else
      strName = "'z'"
      strA1V02 = "'z'"  'add by sonia 2014/3/20
   End If
   If adoadodc2.State = adStateOpen Then adoadodc2.Close
   adoadodc2.CursorLocation = adUseClient
   'Modify by Morgan 2007/1/29 改先收據號後收文號
   'adoadodc2.Open "select a1v01, a1v02, a1v03, a1v04, a1v05, a1v06, a1v07, a1v08, a1v09, a1v10, a1v11, a1v12, a1v13, a1v14, a1v15, a1v16, a1v17, a0k02,st02 from acc1v0, acc0k0,staff where a1v02=a0k01(+) and a0k20=st01(+) and a1v01 in (" & strName & ") order by a1v03 asc, a1v01 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   'Modified by Morgan 2013/2/1
   'adoadodc2.Open "select a1v01, a1v02, a1v03, a1v04, a1v05, a1v06, a1v07, a1v08, a1v09, a1v10, a1v11, a1v12, a1v13, a1v14, a1v15, a1v16, a1v17, a0k02,st02 from acc1v0, acc0k0,staff where a1v02=a0k01(+) and a0k20=st01(+) and a1v01 in (" & strName & ") order by a1v03 asc,a1v02 asc,a1v01 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   'modify by sonia 2014/3/24 加入最後收款日期欄,並加入a1v09條件,又因要更新datagrid必須為實際存在之欄位及table,故寫入workfile再讀出來
   'adoadodc2.Open "select a1v01, a1v02, a1v03, a1v04, a1v05, a1v06, a1v07, a1v08, a1v09, a1v10, a1v11, a1v12, a1v13, a1v14, a1v15, a1v16, a1v17, a0k02,st02 from acc1v0, acc0k0,staff, " & _
                  " where a1v02=a0k01(+) and a0k20=st01(+) and a1v09 = " & Val(Text2) & " and a1v01||a1v02 in (" & strName & ") order by a1v03 asc,a1v02 asc,a1v01 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   If adoacc0w0.State = adStateOpen Then adoacc0w0.Close
   adoacc0w0.CursorLocation = adUseClient
   
   'Add By Sindy 2020/2/13
   'adoacc0w0.Open "delete accrpt11c where R11C03='" & cboTitle & "'", adoTaie, adOpenStatic, adLockReadOnly
   'adoacc0w0.Open "delete accrpt11c where R11C03 like '%" & cboTitle & "%'", adoTaie, adOpenStatic, adLockReadOnly
   adoacc0w0.Open "delete accrpt11c", adoTaie, adOpenStatic, adLockReadOnly
   '2020/2/13 END
   
   'Modify By Sindy 2015/5/4
   'adoacc0w0.Open "insert into accrpt11c (select a0m02,max(a0l02),'" & cboTitle & "' from acc0m0,acc0l0 where a0m01=a0l01(+) and a0m02 in (" & strA1V02 & ") group by a0m02)", adoTaie, adOpenStatic, adLockReadOnly
   'Modify By Sindy 2015/12/10 + R11C06=a0k05
   adoacc0w0.Open "insert into accrpt11c(R11C01,R11C02,R11C03,R11C04,R11C05,R11C06) select a0m02,max(a0l02),'" & cboTitle & "',a0k02,st02,a0k05 from acc0m0,acc0l0,acc0k0,staff where a0m01=a0l01(+) and a0m02 in (" & strA1V02 & ") and a0m02=a0k01(+) and a0k20=st01(+) group by a0m02,a0k02,st02,a0k05 union select a0z02,max(a0y02),'" & cboTitle & "',a1k02,st02,'' from acc0z0,acc0y0,acc1k0,caseprogress,staff where a0z01=a0y01(+) and a0z02 in (" & strA1V02 & ") and a0z02=a1k01(+) and a0z01=cp09(+) and cp13=st01(+) group by a0z02,a1k02,st02", adoTaie, adOpenStatic, adLockReadOnly
   '2015/5/4 END
   'Modify By Sindy 2015/5/4
'   adoadodc2.Open "select a1v01, a1v02, a1v03, a1v04, a1v05, a1v06, a1v07, a1v08, a1v09, a1v10, a1v11, a1v12, a1v13, a1v14, a1v15, a1v16, a1v17, a0k02,st02,R11C02 from acc1v0, acc0k0,staff,accrpt11c " & _
'                  " where a1v02=a0k01(+) and a0k20=st01(+) and a1v02=R11C01(+) and a1v09 = " & Val(Text2) & " and a1v01||a1v02 in (" & strName & ") order by a1v03 asc,a1v02 asc,a1v01 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   'Modify By Sindy 2015/12/10 + R11C06
   adoadodc2.Open "select a1v01, a1v02, a1v03, a1v04, a1v05, a1v06, a1v07, a1v08, a1v09, a1v10, a1v11, a1v12, a1v13, a1v14, a1v15, a1v16, a1v17, R11C04,R11C05,R11C02,R11C06 from acc1v0,accrpt11c" & _
                  " where a1v02=R11C01(+) and a1v09 = " & Val(Text2) & " and a1v01||a1v02 in (" & strName & ")" & _
                  " order by a1v03 asc,a1v02 asc,a1v01 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   '2015/5/4 END
   '2014/3/24 end
   If intRefresh = 1 Then
      intRefresh = 0
   Else
      If adoadodc2.State = adStateOpen Then
         If adoadodc2.RecordCount = 0 Then
            If adoacc0w0.State = adStateOpen Then adoacc0w0.Close
            MsgBox MsgText(28), , MsgText(5)
            Exit Sub
         End If
      Else
         If adoacc0w0.State = adStateOpen Then adoacc0w0.Close
         MsgBox MsgText(28), , MsgText(5)
         Exit Sub
      End If
   End If
   Adodc2.Recordset.Requery
   
   'SumShow1 'Modify By Sindy 2020/5/6 Mark:程式往上移
   SumShow3
   Screen.MousePointer = vbDefault
   Frmacc0000.StatusBar1.Panels(1).Text = "按 尋找 調出補扣繳資料 / " & MsgText(98)
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'產生公司別小計資料
Private Sub SumShow3()
    'Dim intComp As Integer
    Dim strCompNo As String 'Add By Sindy 2020/5/7
    Dim arrlngSum(1 To 5) As Long, iArrIdx As Integer, ii As Integer
    Dim rstClone As New ADODB.Recordset
    
On Error GoTo ErrHnd

    cboSubTotal.Clear
    
    '有勾選擇的調整稅額
    Text8 = ""
    '有勾選擇的已扣繳額
    Text9 = ""
    '有勾選擇且有勾退費的調整稅額
    Text12 = ""
    '應扣繳額
    Text14 = ""
    '已扣繳額
    Text13 = ""
    '未扣繳額
    Text1 = ""
    '有勾退費的未扣繳款+調整稅款
    Text10 = ""
    '有勾退費的調整稅款
    Text5 = ""
    '扣繳憑單金額
    Text11 = ""
    '有勾選擇且有勾退費的未扣繳款
    Text15 = ""
    '有勾選的已收服務費
    Text16 = ""
    For ii = 1 To 5
        TxtSum(ii) = ""
    Next

    If Adodc2.Recordset.State <> adStateOpen Or Adodc2.Recordset.EOF Then
        Exit Sub
    End If

    cboSubTotal.Clear
    Set rstClone = Adodc2.Recordset.Clone()
    With rstClone
    .MoveFirst
    iArrIdx = 0
    'Modify By Sindy 2020/5/7
    'intComp = .Fields("A1V03").Value
    strCompNo = .Fields("A1V03").Value
    '2020/5/7 END
    strRNo = "" 'Add by Morgan 2005/2/21
    strCon6 = "" 'Add By Sindy 2015/11/18
    strCon7 = "" 'Add By Sindy 2015/11/18
    m_yynotsub = 0 'add by sonia 2023/11/24
    Do While Not .EOF
      'Add by Morgan 2005/2/21
      If "" & .Fields("A1V14") = "Y" Then
         strRNo = strRNo & "'" & .Fields("A1V01").Value & "',"
      End If
      '2005/2/21
      
      'Add By Sindy 2015/11/18
      If "" & .Fields("A1V08").Value = "Y" Then
         'Modify By Sindy 2017/3/17 只做請款單Acc1k0=X編號的
         If Left(.Fields(1), 1) = "X" Then
         '2017/3/17 END
            strCon6 = cboTitle.Text '收據抬頭
            strCon7 = Text2 '扣繳年度
         End If
      End If
      '2015/11/18 END
      'Modify By Sindy 2020/5/7
'      If intComp <> ("" & .Fields("A1V03").Value) Then
'         cboSubTotal.AddItem intComp
'         arrSum(iArrIdx) = intComp & "," & arrlngSum(1) & "," & arrlngSum(2) & "," & arrlngSum(3) & "," & arrlngSum(4) & "," & arrlngSum(5)
      If strCompNo <> ("" & .Fields("A1V03").Value) Then
         cboSubTotal.AddItem strCompNo
         arrSum(iArrIdx) = strCompNo & "," & arrlngSum(1) & "," & arrlngSum(2) & "," & arrlngSum(3) & "," & arrlngSum(4) & "," & arrlngSum(5)
      '2020/5/7 END
         For ii = 1 To 5
             arrlngSum(ii) = 0
         Next
         iArrIdx = iArrIdx + 1
      End If
      arrlngSum(1) = arrlngSum(1) + Val(Format("" & .Fields("A1V04").Value, "0"))
      arrlngSum(2) = arrlngSum(2) + Val(Format("" & .Fields("A1V06").Value, "0"))
      arrlngSum(3) = arrlngSum(3) + IIf("" & .Fields("A1V08").Value = "Y", Val(Format("" & .Fields("A1V07").Value, "0")) + Val(Format("" & .Fields("A1V10").Value, "0")), 0)
      arrlngSum(4) = arrlngSum(4) + IIf("" & .Fields("A1V08").Value = "Y", Val(Format("" & .Fields("A1V10").Value, "0")), 0)
      arrlngSum(5) = arrlngSum(5) + Val(Format("" & .Fields("A1V07").Value, "0"))
      
      Text14 = Val(Text14) + Val(Format("" & .Fields("A1V04").Value, "0"))
      Text13 = Val(Text13) + Val(Format("" & .Fields("A1V06").Value, "0"))
      Text10 = Val(Text10) + IIf("" & .Fields("A1V08").Value = "Y", Val(Format("" & .Fields("A1V07").Value, "0")) + Val(Format("" & .Fields("A1V10").Value, "0")), 0)
      Text5 = Val(Text5) + IIf("" & .Fields("A1V08").Value = "Y", Val(Format("" & .Fields("A1V10").Value, "0")), 0)
      '未扣繳額
      Text1 = Val(Text1) + Val(Format("" & .Fields("A1V07").Value, "0"))
      Text8 = Val(Text8) + IIf(("" & .Fields("A1V14") = "Y"), Val(Format("" & .Fields("A1V10").Value, "0")), 0)
      Text9 = Val(Text9) + IIf(("" & .Fields("A1V14") = "Y"), Val(Format("" & .Fields("A1V06").Value, "0")), 0)
      Text12 = Val(Text12) + IIf(("" & .Fields("A1V14") = "Y" And "" & .Fields("A1V08").Value = "Y"), Val(Format("" & .Fields("A1V10").Value, "0")), 0)
      Text15 = Val(Text15) + IIf(("" & .Fields("A1V14") = "Y" And "" & .Fields("A1V08").Value = "Y"), Val(Format("" & .Fields("A1V07").Value, "0")), 0)
      Text16 = Val(Text16) + IIf(("" & .Fields("A1V14") = "Y"), 10 * Val(Format("" & .Fields("A1V04").Value, "0")), 0)
      'add by sonia 2023/11/24 計算扣繳轉年度(與收款年度不符)之服務費超過2萬且未扣繳收據張數
      If Val(Format("" & .Fields("A1V07").Value, "0")) >= 2000 And Val(Format("" & .Fields("A1V09").Value, "0")) <> Val(Left(Format("" & .Fields("R11C02").Value, "0"), 3)) Then
         m_yynotsub = m_yynotsub + 1
      End If
      'end 2023/11/24
      'Modify By Sindy 2020/5/7
      'intComp = .Fields("A1V03").Value
      strCompNo = .Fields("A1V03").Value
      '2020/5/7 END
      .MoveNext
    Loop
    'Modify By Sindy 2020/5/7
'    cboSubTotal.AddItem intComp
'    arrSum(iArrIdx) = intComp & "," & arrlngSum(1) & "," & arrlngSum(2) & "," & arrlngSum(3) & "," & arrlngSum(4) & "," & arrlngSum(5)
    cboSubTotal.AddItem strCompNo
    arrSum(iArrIdx) = strCompNo & "," & arrlngSum(1) & "," & arrlngSum(2) & "," & arrlngSum(3) & "," & arrlngSum(4) & "," & arrlngSum(5)
    '2020/5/7 END
    
    '扣繳憑單金額=有勾選擇的已扣繳額+有勾選擇的調整稅額+有勾選擇和退費的未扣繳額
    Text11 = Val(Text9) + Val(Text8) + Val(Text15)
    End With
    cboSubTotal.ListIndex = 0
    Set rstClone = Nothing
    
ErrHnd:
    If Err.Number <> 0 Then
        MsgBox Err.Description
    End If
End Sub

'*************************************************
'  清除顯示資料
'
'*************************************************
Private Sub AdodcClear()
   Text3 = ""
   Label14 = ""
   Text6 = ""
   Text11 = ""
   Text17 = "" 'Add By Sindy 2020/5/4 新增給付總額欄位
End Sub

'*************************************************
'  顯示資料表(扣繳憑單資料)
'
'*************************************************
Private Sub AdodcShow()
   '扣單公司別
   If IsNull(Adodc1.Recordset.Fields("a0w04").Value) Then
      Text3 = MsgText(601)
   Else
      Text3 = Adodc1.Recordset.Fields("a0w04").Value
      Text3_Change
   End If
   If IsNull(Adodc1.Recordset.Fields("a0w05").Value) Then
      Text11 = MsgText(601)
   Else
      Text11 = Adodc1.Recordset.Fields("a0w05").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a0w06").Value) Then
      Text6 = MsgText(601)
   Else
      Text6 = Adodc1.Recordset.Fields("a0w06").Value
   End If
   'Add By Sindy 2020/5/4 新增給付總額欄位
   If IsNull(Adodc1.Recordset.Fields("a0w16").Value) Then
      Text17 = MsgText(601)
   Else
      Text17 = Adodc1.Recordset.Fields("a0w16").Value
   End If
   '2020/5/4 END
End Sub

'*************************************************
'  計算並顯示合計(扣繳憑單資料)
'
'*************************************************
Private Sub SumShow1()

   adoaccsum.CursorLocation = adUseClient
'Modify by Morgan 2006/5/17
'   adoaccsum.Open "select sum(a0w05) from acc0w0 where a0w01 = " & Val(Text2) & " and a0w03 = '" & cboTitle.Text & "'", adoTaie, adOpenStatic, adLockReadOnly
   '依年度,收據抬頭取得扣繳憑單金額的合計
   adoaccsum.Open "select sum(a0w05) from acc0w0 where a0w01 = " & Val(Text2) & " and a0w03 = '" & m_Title & "'", adoTaie, adOpenStatic, adLockReadOnly
'end 2006/5/17

   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0).Value) Then
         Text7 = MsgText(601)
      Else
         Text7 = adoaccsum.Fields(0).Value
      End If
   Else
      Text7 = MsgText(601)
   End If
   If adoaccsum.State = adStateOpen Then adoaccsum.Close
End Sub

''*************************************************
''  計算並顯示合計(國內收據資料)
''
''*************************************************
'Private Sub SumShow2()
'   adoaccsum.CursorLocation = adUseClient
'   'Modify By Sindy 2015/5/4
'   '"select sum(a1v04), sum(a1v06), sum(a1v07) from acc1k0, acc1v0, customer where a1k01 = a1v02 and substr(a1k28, 1, 8) = cu01 (+) and substr(a1k28, 9, 1) = cu02 (+) and a1v15 is null and (a1k12 is null or a1k12 = 0) and a1k25 is null" & strSQL1 & " union " & _
'   '"select sum(a1v04), sum(a1v06), sum(a1v07) from acc1k0, acc1v0, fagent where a1k01 = a1v02 and substr(a1k28, 1, 8) = fa01 (+) and substr(a1k28, 9, 1) = fa02 (+) and a1v15 is null and (a1k12 is null or a1k12 = 0) and a1k25 is null" & strSQL2
'   adoaccsum.Open "select sum(a1v04), sum(a1v06), sum(a1v07) from acc0k0, acc1v0 where a0k01 = a1v02 and a1v15 is null and (a0k09 is null or a0k09 = 0) and (a0k10 is null or a1v04 <> 0)" & strSql & " union " & _
'                  "select sum(a1v04), sum(a1v06), sum(a1v07) from acc1k0, acc1v0 where a1k01 = a1v02 and a1v15 is null and (a1k12 is null or a1k12 = 0) and a1k25 is null" & strSQL1, adoTaie, adOpenStatic, adLockReadOnly
'   If adoaccsum.RecordCount <> 0 Then
'      '應扣繳額
'      If IsNull(adoaccsum.Fields(0).Value) Then
'         Text14 = "0"
'      Else
'         Text14 = adoaccsum.Fields(0).Value
'      End If
'      '已扣繳額
'      If IsNull(adoaccsum.Fields(1).Value) Then
'         Text13 = "0"
'      Else
'         Text13 = adoaccsum.Fields(1).Value
'      End If
'   Else
'      Text14 = "0"
'      Text13 = "0"
'   End If
'   If adoaccsum.State = adStateOpen Then adoaccsum.Close
'   adoaccsum.CursorLocation = adUseClient
'   'Modify By Sindy 2015/5/4
'   '"select sum(a1v04), sum(a1v06), sum(a1v07) from acc1k0, acc1v0, customer where a1k01 = a1v02 and substr(a1k28, 1, 8) = cu01 (+) and substr(a1k28, 9, 1) = cu02 (+) and (a1v14 = '" & MsgText(602) & "') and (a1v15 is null or a1v15 = '') and (a1k12 is null or a1k12 = 0) and a1k25 is null" & strSQL1 & " union " & _
'   '"select sum(a1v04), sum(a1v06), sum(a1v07) from acc1k0, acc1v0, fagent where a1k01 = a1v02 and substr(a1k28, 1, 8) = fa01 (+) and substr(a1k28, 9, 1) = fa02 (+) and (a1v14 = '" & MsgText(602) & "') and (a1v15 is null or a1v15 = '') and (a1k12 is null or a1k12 = 0) and a1k25 is null" & strSQL2
'   adoaccsum.Open "select sum(a1v04), sum(a1v06), sum(a1v07) from acc0k0, acc1v0 where a0k01 = a1v02 and (a1v14 = '" & MsgText(602) & "') and (a1v15 is null or a1v15 = '') and (a0k09 is null or a0k09 = 0) and (a0k10 is null or a1v04 <> 0)" & strSql & " union " & _
'                  "select sum(a1v04), sum(a1v06), sum(a1v07) from acc1k0, acc1v0 where a1k01 = a1v02 and (a1v14 = '" & MsgText(602) & "') and (a1v15 is null or a1v15 = '') and (a1k12 is null or a1k12 = 0) and a1k25 is null" & strSQL1, adoTaie, adOpenStatic, adLockReadOnly
'   If adoaccsum.RecordCount <> 0 Then
'      '有勾選擇的應扣繳額
'      If IsNull(adoaccsum.Fields(0).Value) Then
'         Text8 = "0"
'      Else
'         Text8 = adoaccsum.Fields(0).Value
'      End If
'      '有勾選擇的已扣繳額
'      If IsNull(adoaccsum.Fields(1).Value) Then
'         Text9 = "0"
'      Else
'         Text9 = adoaccsum.Fields(1).Value
'      End If
'      '有勾選擇的未扣繳額
'      If IsNull(adoaccsum.Fields(2).Value) Then
'         Text10 = "0"
'      Else
'         Text10 = adoaccsum.Fields(2).Value
'      End If
'      adoquery.CursorLocation = adUseClient
'      'Modify By Sindy 2015/5/4
'      '"select sum(a1v10) from acc1k0, acc1v0, customer where a1k01 = a1v02 and substr(a1k28, 1, 8) = cu01 (+) and substr(a1k28, 9, 1) = cu02 (+) and a1v14 = '" & MsgText(602) & "' and a1v15 is null and (a1k12 is null or a1k12 = 0) and a1k25 is null" & strSQL1 & " union " & _
'      '"select sum(a1v10) from acc1k0, acc1v0, fagent where a1k01 = a1v02 and substr(a1k28, 1, 8) = fa01 (+) and substr(a1k28, 9, 1) = fa02 (+) and a1v14 = '" & MsgText(602) & "' and a1v15 is null and (a1k12 is null or a1k12 = 0) and a1k25 is null" & strSQL2
'      adoquery.Open "select sum(a1v10) from acc0k0, acc1v0 where a0k01 = a1v02 and a1v14 = '" & MsgText(602) & "' and a1v15 is null and (a0k09 is null or a0k09 = 0) and (a0k10 is null or a1v04 <> 0)" & strSql & " union " & _
'                    "select sum(a1v10) from acc1k0, acc1v0 where a1k01 = a1v02 and a1v14 = '" & MsgText(602) & "' and a1v15 is null and (a1k12 is null or a1k12 = 0) and a1k25 is null" & strSQL1, adoTaie, adOpenStatic, adLockReadOnly
'      If adoquery.RecordCount <> 0 Then
'         '有勾退費的調整稅款
'         If IsNull(adoquery.Fields(0).Value) Then
'            Text5 = "0"
'         Else
'            Text5 = adoquery.Fields(0).Value
'         End If
'      Else
'         Text5 = "0"
'      End If
'      If adoquery.State = adStateOpen Then adoquery.Close
'      adoquery.CursorLocation = adUseClient
'      'Modify By Sindy 2015/5/4
'      '"select sum(a1v10) from acc1k0, acc1v0, customer where a1k01 = a1v02 and substr(a1k28, 1, 8) = cu01 (+) and substr(a1k28, 9, 1) = cu02 (+) and a1v14 = '" & MsgText(602) & "' and (a1v08 = '" & MsgText(603) & "') and a1v15 is null and (a1k12 is null or a1k12 = 0) and a1k25 is null" & strSQL1 & " union " & _
'      '"select sum(a1v10) from acc1k0, acc1v0, fagent where a1k01 = a1v02 and substr(a1k28, 1, 8) = fa01 (+) and substr(a1k28, 9, 1) = fa02 (+) and a1v14 = '" & MsgText(602) & "' and (a1v08 = '" & MsgText(603) & "') and a1v15 is null and (a1k12 is null or a1k12 = 0) and a1k25 is null" & strSQL2
'      adoquery.Open "select sum(a1v10) from acc0k0, acc1v0 where a0k01 = a1v02 and a1v14 = '" & MsgText(602) & "' and (a1v08 = '" & MsgText(603) & "') and a1v15 is null and (a0k09 is null or a0k09 = 0) and (a0k10 is null or a1v04 <> 0)" & strSql & " union " & _
'                    "select sum(a1v10) from acc1k0, acc1v0 where a1k01 = a1v02 and a1v14 = '" & MsgText(602) & "' and (a1v08 = '" & MsgText(603) & "') and a1v15 is null and (a1k12 is null or a1k12 = 0) and a1k25 is null" & strSQL1, adoTaie, adOpenStatic, adLockReadOnly
'      If adoquery.RecordCount <> 0 Then
'         '有勾選擇且有勾退費的調整稅額
'         If IsNull(adoquery.Fields(0).Value) Then
'            Text12 = "0"
'         Else
'            Text12 = adoquery.Fields(0).Value
'         End If
'      Else
'         Text12 = "0"
'      End If
'      If adoquery.State = adStateOpen Then adoquery.Close
'   Else
'      Text8 = "0"
'      Text9 = "0"
'      Text10 = "0"
'      Text5 = "0"
'      Text12 = "0"
'   End If
'   If adoaccsum.State = adStateOpen Then adoaccsum.Close
'   '扣繳憑單金額
'   Text11 = (Val(Text9) + Val(Text10) + Val(Text5) + Val(Text12))
'End Sub

'檢查公司別是否錯選 1,2公司才可合併
Private Function CheckSelect() As Boolean
    Dim stComp As String
    Dim rstClone As New ADODB.Recordset
    
On Error GoTo ErrHnd
           
    If Text3 = "1" Or Text3 = "2" Then
        stComp = "12"
    Else
        stComp = Text3
    End If
    Set rstClone = Adodc2.Recordset.Clone()
    With rstClone
    .MoveFirst
    Do While Not .EOF
        If ("" & .Fields("A1v14").Value) = "Y" And InStr(stComp, .Fields("A1V03").Value) = 0 Then
            Exit Do
        End If
        .MoveNext
    Loop
    If .EOF Then
        CheckSelect = True
    End If
    End With
    Set rstClone = Nothing
    
ErrHnd:
    If Err.Number <> 0 Then
        MsgBox Err.Description
    End If
End Function

'*************************************************
'  儲存資料表(扣繳憑單資料)
'
'*************************************************
Private Function Acc0w0Save() As Boolean
   Dim strYes As String
   Dim strCaseNo As String

   If Text2 = MsgText(601) Then
      MsgBox MsgText(10), , MsgText(5)
      strControlButton = MsgText(602)
      Text2.SetFocus
      Exit Function
   Else
      If Text3.Text = "" Then
         MsgBox MsgText(170), , MsgText(5)
         strControlButton = MsgText(602)
         Text3.SetFocus
         Exit Function
      End If
      If Text3 <> MsgText(601) Then
         If ExistCheck("acc080", "a0801", Text3, Label3) = False Then
            MsgBox MsgText(28) & Label3, , MsgText(5)
            strControlButton = MsgText(602)
            Text3.SetFocus
            Exit Function
         End If
      End If
      
'Remove by Morgan 2006/5/17 固定用尋找資料時的抬頭
'      If CheckLen(Label1, cboTitle, 100) = MsgText(603) Then
'         strControlButton = MsgText(602)
'         cboTitle.SetFocus
'         Exit Sub
'      End If
'End 2006/5/17

      '檢查公司別是否錯選 1,2公司才可合併
      If CheckSelect = False Then
        MsgBox "選擇公司別有錯！！", vbCritical
        strControlButton = MsgText(602)
        Text3.SetFocus
        Text3_GotFocus
        Exit Function
      End If
      If Adodc2.Recordset.RecordCount <> 0 Then
         Adodc2.Recordset.MoveFirst
         Do While Adodc2.Recordset.EOF = False
            If Adodc2.Recordset.Fields("a1v14").Value = MsgText(602) And Adodc2.Recordset.Fields("a1v08").Value = MsgText(602) And IsNull(Adodc2.Recordset.Fields("a1v17").Value) = False And (Adodc2.Recordset.Fields("a1v17").Value < "E" Or Adodc2.Recordset.Fields("a1v17").Value > "F") Then
               MsgBox MsgText(182), , MsgText(5)
               Exit Function
            End If
            Adodc2.Recordset.MoveNext
         Loop
      End If
    '扣繳憑單金額=有勾選擇的已扣繳額+有勾選擇的調整稅額+有勾選擇和退費的未扣繳額
      'If Val(Text11) <> (Val(Text9) + Val(Text10) + Val(Text5) + Val(Text12)) Then
      If Text11 <> Val(Text9) + Val(Text8) + Val(Text15) Then
         MsgBox MsgText(120), , MsgText(5)
         strControlButton = MsgText(602)
         If Adodc2.Recordset.RecordCount <> 0 Then
            DataGrid2.SetFocus
         End If
         Exit Function
      End If
   End If
   
   'add by sonia 2023/11/23 轉年度未扣繳提醒
   If m_yynotsub > 0 Then
      MsgBox "含有轉年度單筆服務費超過2萬且未扣繳的收據共" & m_yynotsub & "張 !"
   End If
   'end 2023/11/23
   
   FeeShow
   If intCconfirm = vbCancel Then
      Exit Function
   End If
   
   Adodc1.Recordset.AddNew
   
'Modify by Morgan 2006/5/17
'   If cboTitle.Text <> MsgText(601) Then
'      Adodc1.Recordset.Fields("a0w03").Value = cboTitle.Text
'   Else
'      Adodc1.Recordset.Fields("a0w03").Value = Null
'   End If
   Adodc1.Recordset.Fields("a0w03").Value = LeftB(m_Title, 100)
'end 2006/5/17

   Adodc1.Recordset.Fields("a0w01").Value = Val(Text2)
   If Text3 <> MsgText(601) Then
      Adodc1.Recordset.Fields("a0w04").Value = Text3
   Else
      Adodc1.Recordset.Fields("a0w04").Value = Null
   End If
   'Modify by Morgan 2011/1/5 都改抓資料最大單號
   'If Val(Text2) = Mid(CFDate(ACDate(ServerDate)), 1, 3) Then
   '   strTaxNo = AutoNo(MsgText(807), 5, 1)
   'Else
      If adoquery.State = adStateOpen Then adoquery.Close
      adoquery.CursorLocation = adUseClient
      'Modify By Sindy 2018/4/18
      'adoquery.Open "select max(a0w02) from acc0w0 where a0w01 = " & Val(Text2), adoTaie, adOpenStatic, adLockReadOnly
      adoquery.Open "select nvl(max(a0w02),0) from acc0w0 where a0w01 = " & Val(Text2), adoTaie, adOpenStatic, adLockReadOnly
      '2018/4/18 END
      If Not (adoquery.EOF And adoquery.BOF) Then
         If IsNull(adoquery.Fields(0).Value) Then
            strTaxNo = MsgText(807) & IIf(Len(Text2) = 2, "0" & Text2, Text2) & ZeroBeforeNo("0", 5)
         Else
            strTaxNo = MsgText(807) & IIf(Len(Text2) = 2, "0" & Text2, Text2) & ZeroBeforeNo(Mid(adoquery.Fields(0).Value, 5, 5), 5)
         End If
      Else
         strTaxNo = MsgText(807) & IIf(Len(Text2) = 2, "0" & Text2, Text2) & ZeroBeforeNo("0", 5)
      End If
      If adoquery.State = adStateOpen Then adoquery.Close
   'End If
   Adodc1.Recordset.Fields("a0w02").Value = strTaxNo
   If Text11 <> MsgText(601) Then
      Adodc1.Recordset.Fields("a0w05").Value = Val(Text11)
   Else
      Adodc1.Recordset.Fields("a0w05").Value = 0
   End If
   If Text6 <> MsgText(601) Then
      Adodc1.Recordset.Fields("a0w06").Value = Text6
   Else
      Adodc1.Recordset.Fields("a0w06").Value = Null
   End If
   'Add By Sindy 2020/5/4 新增給付總額欄位
   If Text17 <> MsgText(601) Then
      Adodc1.Recordset.Fields("a0w16").Value = Text17
   Else
      Adodc1.Recordset.Fields("a0w16").Value = Null
   End If
   '2020/5/4 END
   Adodc1.Recordset.Fields("a0w07").Value = Val(strSrvDate(2))
   Adodc1.Recordset.Fields("a0w08").Value = ServerTime
   Adodc1.Recordset.Fields("a0w09").Value = strUserNum
   Adodc1.Recordset.Fields("a0w13").Value = Val(strSrvDate(2))
   If Adodc2.Recordset.RecordCount <> 0 Then
      Adodc2.Recordset.MoveFirst
   End If
   
   'Modify by Morgan 2007/1/29 因欄位不夠長,改不紀錄全部收據編號,連號時用"起號-迄號"表示
   strCaseNo = ""
   strExc(0) = "" '前一筆收據號
   strExc(1) = "" '這一筆收據號
   strExc(2) = "" '最後連號收據號
   Adodc2.Recordset.MoveFirst
   Do While Adodc2.Recordset.EOF = False
      If IsNull(Adodc2.Recordset.Fields("a1v15").Value) And Adodc2.Recordset.Fields("a1v14").Value = MsgText(602) Then
         strExc(1) = "" & Adodc2.Recordset.Fields("a1v02").Value
         If strExc(0) = "" Then
            strCaseNo = strExc(1)
         ElseIf strExc(0) <> strExc(1) Then
            If Val(Mid(strExc(1), 2)) = Val(Mid(strExc(0), 2)) + 1 Then
               strExc(2) = strExc(1)
            Else
               If strExc(2) <> "" Then
                  strCaseNo = strCaseNo & "-" & strExc(2)
                  strExc(2) = ""
               End If
               strCaseNo = strCaseNo & "," & strExc(1)
            End If
         End If
         strExc(0) = strExc(1)
      End If
      Adodc2.Recordset.MoveNext
   Loop
   If strExc(2) <> "" Then
      strCaseNo = strCaseNo & "-" & strExc(2)
      strExc(2) = ""
   End If
   If strCaseNo <> "" Then
      'Adodc1.Recordset.Fields("a0w14").Value = Mid(strCaseNo, 1, Len(strCaseNo) - 1)
      Adodc1.Recordset.Fields("a0w14").Value = strCaseNo
   End If
   'end 2007/1/29
   
   Adodc1.Recordset.UpdateBatch
   Acc0w0Save = True
End Function

Private Function FormSave() As Boolean

   adoTaie.BeginTrans
   
On Error GoTo ErrHnd

   If Acc0w0Save = False Then
      strControlButton = MsgText(601)
      GoTo ErrHnd
   End If
         
   strItemNo = strTaxNo
   Text3.SetFocus
   If Val(Text10) > 0 Then
      strCon1 = Val(Text10)
      If Adodc2.Recordset.RecordCount <> 0 Then
         Adodc2.Recordset.MoveFirst
      End If
      Do While Adodc2.Recordset.EOF = False
         If IsNull(Adodc2.Recordset.Fields("a1v15").Value) And Adodc2.Recordset.Fields("a1v14").Value = MsgText(602) Then
            adocheck.CursorLocation = adUseClient
            adocheck.Open "select a0k03 from acc0k0 where a0k01 = '" & Adodc2.Recordset.Fields("a1v02").Value & "' union " & _
                          "select a1k03 as a0k03 from acc1k0 where a1k01 = '" & Adodc2.Recordset.Fields("a1v02").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
            If adocheck.RecordCount <> 0 Then
               If IsNull(adocheck.Fields(0).Value) Then
                  strCon2 = MsgText(601)
               Else
                  strCon2 = "" & adocheck.Fields(0).Value
               End If
            Else
               strCon2 = MsgText(601)
            End If
            If adocheck.State = adStateOpen Then adocheck.Close
            Exit Do
         End If
         Adodc2.Recordset.MoveNext
      Loop
'Modify by Morgan 2006/5/17
'            strCon3 = cboTitle.Text & Text2
      strCon3 = m_Title & Text2
'end 2006/5/17
      strFormLink = Name
      tool3_enabled
      '傳票摘要
'Modify by Morgan 2006/5/17
'            Frmacc11b1.strMemo = GetSaleShortName & "/" & cboTitle & Text2
      Frmacc11b1.strMemo = GetSaleShortName & "/" & m_Title & Text2
'end 2006/5/17
      Frmacc11b1.intForm = 1
      Frmacc11b1.Show
      Me.Enabled = False
   End If
   
   If Adodc2.Recordset.RecordCount <> 0 Then
      Adodc2.Recordset.MoveFirst
   End If
   Do While Adodc2.Recordset.EOF = False
      If IsNull(Adodc2.Recordset.Fields("a1v15").Value) And Adodc2.Recordset.Fields("a1v14").Value = MsgText(602) Then
         '只更新有退費的
         'modify by sonia 2018/3/15 K10601986之收據會因為K10601987而錯誤
         'adoTaie.Execute "update acc1v0 set a1v15 = '" & strTaxNo & "' where a1v14 = 'Y' and a1v01 = '" & Adodc2.Recordset.Fields("a1v01").Value & "' and a1v01 in " & strCon4
         adoTaie.Execute "update acc1v0 set a1v15 = '" & strTaxNo & "' where a1v14 = 'Y' and a1v01 = '" & Adodc2.Recordset.Fields("a1v01").Value & "' and a1v02 = '" & Adodc2.Recordset.Fields("a1v02").Value & "' and a1v01 in " & strCon4, intI
         'cancel by sonia 2023/3/13 抓資料條件就是A0K16,看不出為何又要更新A0K16,為免A0K15不一致故取消
         ''更新扣繳年度
         'adoTaie.Execute "update acc0k0 set a0k16 = " & Val(Adodc2.Recordset.Fields("a1v09").Value) & " where a0k01 = '" & Adodc2.Recordset.Fields("a1v02").Value & "'"
         'end 2023/3/13
      End If
      Adodc2.Recordset.MoveNext
   Loop
   
   If Me.Enabled = True Then
      adoTaie.CommitTrans
   End If
   FormSave = True
   Exit Function
   
ErrHnd:
   adoTaie.RollbackTrans
   If Err.Number <> 0 Then
      MsgBox Err.Description
   End If
End Function

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)
   'Added by Morgan 2023/2/7
   Call PUB_SaveTrackMode(1, KeyCode)
   If PUB_ChkTrackMode = False Then
       Exit Sub
   End If
   'end 2023/2/7

   Select Case KeyCode
      Case vbKeyInsert
         If strRNo <> "" Then
            If Right(strRNo, 1) = "," Then
               strCon4 = "(" & Mid(strRNo, 1, Len(strRNo) - 1) & ")"
            Else
               strCon4 = "(" & strRNo & ")"
            End If
         Else
            Exit Sub
         End If
         If adoquery.State = adStateOpen Then adoquery.Close
         adoquery.CursorLocation = adUseClient
         adoquery.Open "select a1v17 from acc0k0, acc1v0 where a0k01 = a1v02 and a1v14 = '" & MsgText(602) & "' and (a1v17 is not null and substr(a1v17, 1, 1) <> 'E') and a1v01 in " & strCon4 & strSql, adoTaie, adOpenStatic, adLockReadOnly
         If adoquery.RecordCount <> 0 Then
            MsgBox MsgText(182), , MsgText(5)
            If adoquery.State = adStateOpen Then adoquery.Close
            Exit Sub
         End If
         If adoquery.State = adStateOpen Then adoquery.Close
         
         If FormSave = True Then
            intRefresh = 1
            If Frmacc11b1.intForm <> 1 Then
               AdodcRefresh
               AdodcClear
            End If
            If Me.Enabled = True Then cboTitle.SetFocus
         End If
         
      Case vbKeyF12
         cmdSearch_Click
   
      Case Else
         KeyEnter KeyCode
         
   End Select
   Frmacc0000.StatusBar1.Panels(1).Text = "按 尋找 調出補扣繳資料 / " & MsgText(98)
End Sub

'Add by Morgan 2006/2/27 檢查是否已過帳
Private Function CheckPosted() As Boolean
On Error GoTo Checking
   strSql = "select a1p22 from acc1p0 where a1p01 = '1' and a1p02 = 'E' and substr(a1p04, 1, 9) = '" & strTaxNo & "' and a1p22 is not null"
   CheckOC
   With adoRecordset
      .CursorLocation = adUseClient
      .Open strSql, adoTaie, adOpenForwardOnly, adLockReadOnly
      If .RecordCount > 0 Then
         If PUB_CheckPosted(.Fields(0), True) = True Then
            CheckPosted = True
         End If
      End If
   End With
Checking:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
End Function

'*************************************************
'  刪除 Adodc1 之資料
'
'*************************************************
Private Sub AdodcDelete()
On Error GoTo Checking
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   'Add by Morgan 2006/2/27 檢查是否已過帳
   If CheckPosted = True Then
      Exit Sub
   End If
   '2006/2/27 end
   
   'Remove by Morgan 2006/2/27 取消--會刪到補扣繳的資料(這裡只作調整,不會補扣繳,應該不會新增收款資料)
   'adoTaie.Execute "delete from acc1u0 where a1u01 in (select a1v01 from acc1v0 where a1v15 = '" & strTaxNo & "')"
   '2006/2/27 end
   adoTaie.Execute "delete from acc0w0 where a0w02 = '" & strTaxNo & "'"
   'Modify by Morgan 2006/2/27 a1v10不必清掉--辜
   'adoTaie.Execute "update acc1v0 set a1v14 = null, a1v08 = null, a1v10 = 0 where a1v15 = '" & strTaxNo & "'"
   adoTaie.Execute "update acc1v0 set a1v14 = null, a1v08 = null where a1v15 = '" & strTaxNo & "'"
   '2006/2/27 end
   adoTaie.Execute "update acc1v0 set a1v15 = null where a1v15 = '" & strTaxNo & "'"
   adoTaie.Execute "delete from acc1p0 where a1p01 = '1' and a1p02 = 'E' and substr(a1p04, 1, 9) = '" & strTaxNo & "'"
   AdodcRefresh
   AdodcClear
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'Add By Sindy 2015/3/27
Private Sub ClearAll()
Dim ii As Integer
   
   Adodc1.Recordset.Requery
   Adodc2.Recordset.Requery
   cboSubTotal.Clear
   '有勾選擇的調整稅額
   Text8 = ""
   '有勾選擇的已扣繳額
   Text9 = ""
   '有勾選擇且有勾退費的調整稅額
   Text12 = ""
   '應扣繳額
   Text14 = ""
   '已扣繳額
   Text13 = ""
   '未扣繳額
   Text1 = ""
   '有勾退費的未扣繳款+調整稅款
   Text10 = ""
   '有勾退費的調整稅款
   Text5 = ""
   '扣繳憑單金額
   Text11 = ""
   '有勾選擇且有勾退費的未扣繳款
   Text15 = ""
   '有勾選的已收服務費
   Text16 = ""
   For ii = 1 To 5
      TxtSum(ii) = ""
   Next
End Sub

'*************************************************
'  產生扣繳明細資料
'
'*************************************************
Public Sub ProcessData(Optional ByVal bolExact As Boolean = True)

Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)

ClearAll 'Add By Sindy 2015/3/27
'Add by Morgan 2004/3/5-----------
'國內客戶不用再產生1v0資料，跳過中間程式碼
If txtCustNo(0) <> "" Then GoTo flgSkip
'Add end 2004/3/3-----------------

Dim StrSQLa As String, ii As Integer
'清除收據公司選單
   strSql = MsgText(601)
   strSQL1 = MsgText(601)
   strSQL2 = MsgText(601)
   
   If cboTitle.Text <> MsgText(601) Then
      If bolExact = True Then
         strSql = " and a0k04= '" & cboTitle.Text & "'"
      Else
         strSql = " and instrb(a0k04, '" & cboTitle.Text & "') = 1"
      End If
      strSQL1 = " and instrb(cu04, '" & cboTitle.Text & "') = 1"
      strSQL2 = " and instrb(fa04, '" & cboTitle.Text & "') = 1"
   End If
   
   If txtCustNo(0) <> "" Then
      strSql = strSql & " and a0k03>='" & txtCustNo(0).Text & "'"
   End If
   If txtCustNo(1) <> "" Then
      strSql = strSql & " and a0k03<='" & txtCustNo(1).Text & "'"
   End If
   
   If txtSales <> "" Then
      strSql = strSql & " and a0k20||''='" & txtSales & "'"
   End If

   If Text2 <> MsgText(601) Then
      strSql = strSql & " and a0k16 = " & Val(Text2) & ""
      strSQL1 = strSQL1 & " and a0y02>=" & Val(Text2) & "0101 And a0y02<=" & Val(Text2) & "1231 "
      strSQL2 = strSQL2 & " and a0y02>=" & Val(Text2) & "0101 And a0y02<=" & Val(Text2) & "1231 "
   End If

   Screen.MousePointer = vbHourglass
   adoquery.CursorLocation = adUseClient
   'Memo by Morgan 2011/12/27 取消 a0j20,a0j21(這裡不用改)
   'Modify By Sindy 2020/6/4 nvl(a0z13, '1') => nvl(a0z13, '2')
   StrSQLa = " select cp09, cp60, nvl(a0z13, '2') as a0k11, DECODE((nvl(a0z04, 0) - nvl(A1K09, 0)) * 0.1  - nvl(a0z12, 0),0,0,((nvl(cp16, 0) - nvl(cp17, 0)) * 0.1)) as TAmount, nvl(cp16, 0) as cp16, a0z04 as cp75, to_number(substr(a0y02, 1, length(a0y02) - 4)) as a0k16, '' as a0j20, '' as a0j21, DECODE(a0z12,0,0,((nvl(cp16, 0) - nvl(cp17, 0)) * 0.1)) as cp76, ((nvl(cp16, 0) - nvl(cp17, 0)) * 0.1) as RAmount, fa04 as a0k04, '' as a0k13 from acc1k0, caseprogress, acc0z0, acc0y0, fagent where a1k01 = cp60 and a1k01 = a0z02 and a0z01 = a0y01 and substr(a1k28, 1, 8) = fa01 (+) and substr(a1k28, 9, 1) = fa02 (+) and cp09 is not null" & strSQL2 & _
            " union select cp09, cp60, nvl(a0z13, '2') as a0k11, DECODE((nvl(a0z04, 0) - nvl(A1K09, 0)) * 0.1  - nvl(a0z12, 0),0,0,((nvl(cp16, 0) - nvl(cp17, 0)) * 0.1)) as TAmount, nvl(cp16, 0) as cp16, a0z04 as cp75, to_number(substr(a0y02, 1, length(a0y02) - 4)) as a0k16, '' as a0j20, '' as a0j21, DECODE(a0z12,0,0,((nvl(cp16, 0) - nvl(cp17, 0)) * 0.1)) as cp76, ((nvl(cp16, 0) - nvl(cp17, 0)) * 0.1) as RAmount, cu04 as a0k04, '' as a0k13 from acc1k0, caseprogress, acc0z0, acc0y0, customer where a1k01 = cp60 and a1k01 = a0z02 and a0z01 = a0y01 and substr(a1k28, 1, 8) = cu01 (+) and substr(a1k28, 9, 1) = cu02 (+) and cp09 is not null" & strSQL1 & _
            " order by a0k11 asc, cp09 asc"

   adoquery.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly

   Do While adoquery.EOF = False
      With adoquery
         adocheck.CursorLocation = adUseClient
         adocheck.Open "select a1v01 from acc1v0 where a1v01 = '" & .Fields("cp09").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
         If adocheck.RecordCount = 0 Then
            If adoaccsum.State = adStateOpen Then adoaccsum.Close
            adoaccsum.CursorLocation = adUseClient
            adoaccsum.Open "select a0m03 from acc0m0 where a0m02 = '" & adoquery.Fields("cp60").Value & "' union " & _
                           "select null as a0m03 from acc0z0 where a0z02 = '" & adoquery.Fields("cp60").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
            'Memo by Morgan 2011/12/27 取消 a0j20,a0j21 (這裡不用改,因為上面都是放Null)
            If adoaccsum.RecordCount <> 0 Then
               If IsNull(adoaccsum.Fields("a0m03").Value) = False Then
                  adoTaie.Execute "insert into acc1v0 values ('" & .Fields("cp09").Value & "', '" & .Fields("cp60").Value & "', '" & .Fields("a0k11").Value & "', " & .Fields("RAmount").Value & ", '" & IIf(IsNull(.Fields("a0k13").Value) = True Or .Fields("a0k13").Value = "", "N", "Y") & "', " & .Fields("cp76").Value & ", " & .Fields("TAmount").Value & ", null, " & .Fields("a0k16").Value & ", 0, 0, '" & IIf(IsNull(.Fields("a0j20").Value), "", .Fields("a0j20").Value) & "', '" & IIf(IsNull(.Fields("a0j21").Value), "", .Fields("a0j21").Value) & "', null, null, null, '" & adoaccsum.Fields("a0m03").Value & "', " & IIf(IsNull(.Fields("cp76").Value) = True Or .Fields("cp76").Value = 0, "null", "'1'") & ")"
               Else
                  adoTaie.Execute "insert into acc1v0 values ('" & .Fields("cp09").Value & "', '" & .Fields("cp60").Value & "', '" & .Fields("a0k11").Value & "', " & .Fields("RAmount").Value & ", '" & IIf(IsNull(.Fields("a0k13").Value) = True Or .Fields("a0k13").Value = "", "N", "Y") & "', " & .Fields("cp76").Value & ", " & .Fields("TAmount").Value & ", null, " & .Fields("a0k16").Value & ", 0, 0, '" & IIf(IsNull(.Fields("a0j20").Value), "", .Fields("a0j20").Value) & "', '" & IIf(IsNull(.Fields("a0j21").Value), "", .Fields("a0j21").Value) & "', null, null, null, null, " & IIf(IsNull(.Fields("cp76").Value) = True Or .Fields("cp76").Value = 0, "null", "'1'") & ")"
               End If
            Else
               adoTaie.Execute "insert into acc1v0 values ('" & .Fields("cp09").Value & "', '" & .Fields("cp60").Value & "', '" & .Fields("a0k11").Value & "', " & .Fields("RAmount").Value & ", '" & IIf(IsNull(.Fields("a0k13").Value) = True Or .Fields("a0k13").Value = "", "N", "Y") & "', " & .Fields("cp76").Value & ", " & .Fields("TAmount").Value & ", null, " & .Fields("a0k16").Value & ", 0, 0, '" & IIf(IsNull(.Fields("a0j20").Value), "", .Fields("a0j20").Value) & "', '" & IIf(IsNull(.Fields("a0j21").Value), "", .Fields("a0j21").Value) & "', null, null, null, null, " & IIf(IsNull(.Fields("cp76").Value) = True Or .Fields("cp76").Value = 0, "null", "'1'") & ")"
            End If
            If adoaccsum.State = adStateOpen Then adoaccsum.Close
         End If
         If adocheck.State = adStateOpen Then adocheck.Close
      End With
      adoquery.MoveNext
   Loop
   If adoquery.State = adStateOpen Then adoquery.Close
   
flgSkip:
   'adoTaie.Execute "delete from acc1v0 where a1v06 = 0 and a1v07 = 0"
   cboComp.Clear
   cboSelComp.Clear
   AdodcRefresh bolExact
   '將收據公司別加進下拉選單
   cboComp.AddItem ""
   cboSelComp.AddItem ""
   For ii = 0 To cboSubTotal.ListCount - 1
      cboComp.AddItem cboSubTotal.List(ii)
      cboSelComp.AddItem cboSubTotal.List(ii)
   Next ii
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(601)
   
   Screen.MousePointer = vbDefault
End Sub

'*************************************************
'  計算並顯示退費資料
'
'*************************************************
Public Sub FeeShow()
    If Val(Text10) > 0 Then
      intCconfirm = MsgBox(MsgText(121) & Format(Val(Text10), DDollar), vbOKCancel + vbDefaultButton1, MsgText(5))
    Else
      intCconfirm = vbOK
    End If
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
'   If Text2 <> "" Then
'      strCon1 = " and a0k16=" & Text2
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
'   If adoQuery.State = adStateOpen Then adoQuery.Close
'   Set adoQuery = Nothing
'
'   AddItem2CboTitle = True
'   Exit Function
'
'ErrHand:
'   MsgBox Err.Description
'
'End Function

Private Function ChkConOK() As Boolean
   If cboTitle = "" Then
      MsgBox "收據抬頭不可空白！", vbCritical
      Exit Function
   End If
   If CheckLen(Label1, cboTitle, 100) = MsgText(603) Then
      Exit Function
   End If
   'Add by Morgan 2006/5/2
   If Len(txtCustNo(0)) <> 9 Or Len(txtCustNo(1)) <> 9 Then
      MsgBox "客戶編號資料不完整！", vbExclamation
      txtCustNo(0).SetFocus
      Exit Function
   End If
   
   'Add By Sindy 2013/12/30
   If cboComp.Text = "J" Then
      MsgBox "收據公司別不可為 J 公司！", vbCritical
      cboComp.SetFocus
      Exit Function
   End If
   '2013/12/30 END
   
'Remove by Morgan 2007/5/9
'   If Left(txtCustNo(0), 6) <> Left(txtCustNo(1), 6) Then
'      MsgBox "客戶編號前六碼必須相同！", vbExclamation
'      txtCustNo(0).SetFocus
'      Exit Function
'   End If
'end 2007/5/9

   'Modify by Morgan 2007/5/9 改放收據抬頭
   'm_NewKey = Left(txtCustNo(0), 6)
   m_NewKey = cboTitle
   'end 2007/5/9
   If PUB_GetLock(m_NewKey, m_OldKey) = True Then
      ChkConOK = True
      'Add by Morgan 2006/5/17 設定扣單抬頭
      m_Title = cboTitle
   End If
   '2006/5/2 end
End Function

Private Sub txtSales_GotFocus()
   'edit by nickc 2007/06/11  切換輸入法改用API
   'txtSales.IMEMode = 2
   CloseIme
   TextInverse txtSales
End Sub

Private Sub txtSales_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtSales_Validate(Cancel As Boolean)
   If txtSales = "" Then
      lblSales = ""
   Else
      lblSales = GetStaffName(txtSales)
      If lblSales = "" Then
         MsgBox "智權人員不存在，請重新輸入！"
         Cancel = True
      End If
   End If
End Sub

'========================================================================================================================

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking

   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "select * from acc0w0 where rownum<1", adoTaie, adOpenForwardOnly, adLockReadOnly
   Set Adodc1.Recordset = adoadodc1
   
   adoadodc2.CursorLocation = adUseClient
   'Modify By Sindy 2015/5/4
   'adoadodc2.Open "select * from acc0k0, acc1v0 where a0k01 = a1v02 and rownum<1", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoadodc2.Open "select * from acc1v0 where rownum<1", adoTaie, adOpenDynamic, adLockBatchOptimistic
   '2015/5/4 END
   Set Adodc2.Recordset = adoadodc2
   
Checking:
   If Err.Number <> 0 Then
      MsgBox Err.Description, , MsgText(5)
   End If
End Sub

'用最大收據的最大收文號抓智權人員簡稱
Private Function GetSaleShortName() As String
   Dim rstClone As New ADODB.Recordset, stSQL As String
   
   Set rstClone = Adodc2.Recordset.Clone
   rstClone.Sort = "a1v02 desc,a1v01 desc"
   rstClone.MoveFirst
   rstClone.Find "a1v08='Y'", 0, adSearchForward, 1
   If Not rstClone.EOF Then
      '2012/8/23 modif by sonia 改抓最大收據之收據智權人員
      'stSQL = "select sn01,CP13 from caseprogress,salesno where cp13=sn02(+) and cp09='" & rstClone.Fields("a1v01") & "'"
      'Modify By Sindy 2015/5/4
      If Left(Trim(rstClone.Fields("a1v02")), 1) = "X" Then '國外請款編號
         stSQL = "select sn01,CP13 from caseprogress,salesno where cp13=sn02(+) and cp09='" & rstClone.Fields("a1v01") & "'"
         If adoquery.State = adStateOpen Then adoquery.Close
         adoquery.CursorLocation = adUseClient
         adoquery.Open stSQL, adoTaie, adOpenStatic, adLockReadOnly
         If Not adoquery.EOF And Not IsNull(adoquery.Fields(0)) Then
            GetSaleShortName = adoquery.Fields(0)
         Else
            GetSaleShortName = adoquery.Fields(1)
         End If
      Else
      '2015/5/4 END
         '國內請款編號
         stSQL = "select sn01,a0k20 from acc0k0,salesno where a0k20=sn02(+) and a0k01='" & rstClone.Fields("a1v02") & "'"
         If adoquery.State = adStateOpen Then adoquery.Close
         adoquery.CursorLocation = adUseClient
         adoquery.Open stSQL, adoTaie, adOpenStatic, adLockReadOnly
         If Not adoquery.EOF And Not IsNull(adoquery.Fields(0)) Then
            GetSaleShortName = adoquery.Fields(0)
         Else
            GetSaleShortName = adoquery.Fields(1)
         End If
      End If
   End If
   If adoquery.State = adStateOpen Then adoquery.Close
   Set rstClone = Nothing
End Function

'Add by Morgan 2005/2/25 取消選擇
Private Function fnUnCheck() As Boolean
On Error GoTo ErrHnd

   If strRNo <> "" Then
      strRNo = "(" & Mid(strRNo, 1, Len(strRNo) - 1) & ")"
      '將已選取無扣單編號的收文號,更新為未選取,無退費,調整稅款=0
      adoTaie.Execute "update acc1v0 set a1v14 = null, a1v08 = null, a1v10 = 0 where a1v15 is null and a1v14 = '" & MsgText(602) & "' and a1v01 in " & strRNo
      strRNo = ""
   End If
   fnUnCheck = True
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Function
