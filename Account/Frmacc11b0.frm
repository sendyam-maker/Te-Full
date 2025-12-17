VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc11b0 
   AutoRedraw      =   -1  'True
   Caption         =   "補扣繳作業"
   ClientHeight    =   5340
   ClientLeft      =   50
   ClientTop       =   280
   ClientWidth     =   9410
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5340
   ScaleWidth      =   9410
   Begin VB.CommandButton cmdA49 
      Caption         =   "繳款書寄件維護"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6540
      Style           =   1  '圖片外觀
      TabIndex        =   33
      Top             =   1440
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.TextBox txtSum 
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
      Height          =   288
      Index           =   5
      Left            =   5400
      TabIndex        =   32
      Top             =   5016
      Width           =   828
   End
   Begin VB.TextBox Text1 
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
      Height          =   288
      Left            =   5400
      TabIndex        =   31
      Top             =   4704
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
      Height          =   276
      ItemData        =   "Frmacc11b0.frx":0000
      Left            =   1530
      List            =   "Frmacc11b0.frx":0002
      TabIndex        =   29
      Top             =   5016
      Width           =   750
   End
   Begin VB.TextBox txtSum 
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
      Height          =   288
      Index           =   1
      Left            =   2748
      TabIndex        =   28
      Top             =   5016
      Width           =   900
   End
   Begin VB.TextBox txtSum 
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
      Height          =   288
      Index           =   2
      Left            =   4464
      TabIndex        =   27
      Top             =   5016
      Width           =   948
   End
   Begin VB.TextBox txtSum 
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
      Height          =   288
      Index           =   3
      Left            =   6240
      TabIndex        =   26
      Top             =   5016
      Width           =   828
   End
   Begin VB.TextBox txtSum 
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
      Height          =   288
      Index           =   4
      Left            =   7692
      TabIndex        =   25
      Top             =   5016
      Width           =   924
   End
   Begin VB.TextBox txtSales 
      BackColor       =   &H00C0FFFF&
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
      Left            =   1200
      MaxLength       =   6
      TabIndex        =   3
      Top             =   735
      Width           =   1065
   End
   Begin VB.CommandButton cmdSearch1 
      Caption         =   "相似尋找"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   3592
      TabIndex        =   23
      Top             =   765
      Width           =   1100
   End
   Begin VB.TextBox txtCustNo 
      BackColor       =   &H00C0FFFF&
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
      Left            =   1200
      MaxLength       =   9
      TabIndex        =   1
      Top             =   420
      Width           =   1572
   End
   Begin VB.TextBox txtCustNo 
      BackColor       =   &H00C0FFFF&
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
      Left            =   3030
      MaxLength       =   9
      TabIndex        =   2
      Top             =   420
      Width           =   1572
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "尋找"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   8115
      TabIndex        =   20
      Top             =   495
      Width           =   1100
   End
   Begin VB.CommandButton cmdLikeSearch 
      Caption         =   "相似搜尋"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   8115
      TabIndex        =   19
      Top             =   150
      Width           =   1100
   End
   Begin VB.TextBox Text8 
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
      Height          =   288
      Left            =   2748
      TabIndex        =   18
      Top             =   4704
      Width           =   900
   End
   Begin VB.TextBox Text9 
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
      Height          =   288
      Left            =   4464
      TabIndex        =   17
      Top             =   4704
      Width           =   948
   End
   Begin VB.TextBox Text10 
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
      Height          =   288
      Left            =   6240
      TabIndex        =   16
      Top             =   4704
      Width           =   828
   End
   Begin VB.TextBox Text5 
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
      Height          =   288
      Left            =   7692
      TabIndex        =   15
      Top             =   4704
      Width           =   924
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
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
      Left            =   1200
      MaxLength       =   3
      TabIndex        =   4
      Top             =   1065
      Width           =   612
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1320
      TabIndex        =   7
      Top             =   1515
      Width           =   852
   End
   Begin VB.CommandButton Command2 
      Caption         =   "退費輸入"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   8115
      TabIndex        =   10
      Top             =   840
      Width           =   1100
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "Frmacc11b0.frx":0004
      Height          =   2796
      Left            =   0
      TabIndex        =   9
      Top             =   1872
      Width           =   9336
      _ExtentX        =   16439
      _ExtentY        =   4904
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
      ColumnCount     =   15
      BeginProperty Column00 
         DataField       =   "a1v03"
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
      BeginProperty Column02 
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
      BeginProperty Column03 
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
      BeginProperty Column04 
         DataField       =   "ex01"
         Caption         =   "尚欠金額"
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
      BeginProperty Column06 
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
      BeginProperty Column07 
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
      BeginProperty Column08 
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
      BeginProperty Column09 
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
      BeginProperty Column10 
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
      BeginProperty Column11 
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
      BeginProperty Column12 
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
      BeginProperty Column13 
         DataField       =   "ST02"
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
      BeginProperty Column14 
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
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   679.748
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   909.921
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   909.921
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   849.827
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   879.874
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   879.874
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            ColumnWidth     =   840.189
         EndProperty
         BeginProperty Column07 
            Alignment       =   2
            ColumnWidth     =   670.11
         EndProperty
         BeginProperty Column08 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   829.984
         EndProperty
         BeginProperty Column09 
            Alignment       =   1
            ColumnWidth     =   840.189
         EndProperty
         BeginProperty Column10 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column11 
            Locked          =   -1  'True
            ColumnWidth     =   1549.984
         EndProperty
         BeginProperty Column12 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column13 
         EndProperty
         BeginProperty Column14 
            Object.Visible         =   0   'False
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
      Left            =   30
      Top             =   1785
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
   Begin MSMask.MaskEdBox mebRecDate 
      Height          =   300
      Index           =   1
      Left            =   3045
      TabIndex        =   5
      Top             =   1065
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   529
      _Version        =   393216
      BackColor       =   12648447
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
   Begin MSMask.MaskEdBox mebRecDate 
      Height          =   300
      Index           =   2
      Left            =   4620
      TabIndex        =   6
      Top             =   1065
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   529
      _Version        =   393216
      BackColor       =   12648447
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
   Begin MSForms.TextBox Text3 
      Height          =   285
      Left            =   2220
      TabIndex        =   8
      Top             =   1530
      Width           =   3945
      VariousPropertyBits=   679495709
      BackColor       =   14737632
      ScrollBars      =   3
      Size            =   "6959;503"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblSales 
      Height          =   255
      Left            =   2370
      TabIndex        =   36
      Top             =   810
      Width           =   1095
      VariousPropertyBits=   19
      Caption         =   "lblSales"
      Size            =   "1931;450"
      FontName        =   "新細明體"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin MSForms.ComboBox cboTitle 
      Height          =   315
      Left            =   1200
      TabIndex        =   0
      Top             =   90
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
   Begin VB.Label Label9 
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
      Left            =   4440
      TabIndex        =   35
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "收款日期"
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
      Left            =   2040
      TabIndex        =   34
      Top             =   1095
      Width           =   975
   End
   Begin VB.Label Label11 
      Alignment       =   1  '靠右對齊
      BackStyle       =   0  '透明
      Caption         =   "公司別小計"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   276
      TabIndex        =   30
      Top             =   5040
      Width           =   1236
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "智權人員"
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
      Left            =   165
      TabIndex        =   24
      Top             =   765
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
      Height          =   255
      Left            =   2850
      TabIndex        =   22
      Top             =   450
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "客戶編號"
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
      Left            =   165
      TabIndex        =   21
      Top             =   450
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "扣繳年度"
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
      Left            =   180
      TabIndex        =   14
      Top             =   1095
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   1470
      Left            =   60
      Top             =   30
      Width           =   9270
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   -240
      Top             =   4680
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label5 
      Alignment       =   1  '靠右對齊
      BackStyle       =   0  '透明
      Caption         =   "合計"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   888
      TabIndex        =   13
      Top             =   4716
      Width           =   612
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "公司別"
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
      Left            =   360
      TabIndex        =   12
      Top             =   1515
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "收據抬頭"
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
      Left            =   180
      TabIndex        =   11
      Top             =   150
      Width           =   975
   End
End
Attribute VB_Name = "Frmacc11b0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/10/7 Form2.0已修改
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/11/26 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit

Public adoacc0k0 As New ADODB.Recordset
Public adoacc0m0 As New ADODB.Recordset
Public adoacctmp03 As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Public adocheck As New ADODB.Recordset
Dim intCconfirm As Integer
'Dim strSQL1 As String
'Dim strSQL2 As String
Dim strSql0k As String, strSql1k As String
Dim str1vCon As String 'Added by Morgan 2014/1/16 考慮拆收據會有不同扣繳年度
Dim strRNo As String
Dim intY As Integer
'Dim strSql As String
'公司別小計
Dim arrSum(0 To 9) As String
'Add by Morgan 2006/5/2
'避免單一資料同時被兩程式處理
Dim m_NewKey As String
Dim m_OldKey As String


'Add by Morgan 2004/1/29
Private Sub cboSubTotal_Click()
    Dim ArrStr() As String, ii As Integer
    If cboSubTotal.ListIndex <> -1 Then
        ArrStr = Split(arrSum(cboSubTotal.ListIndex), ",")
        For ii = 1 To 5
            txtSum(ii) = ArrStr(ii)
        Next
    End If
End Sub

Private Sub cboTitle_GotFocus()
   'edit by nickc 2007/06/11  切換輸入法改用API
   'cboTitle.IMEMode = 1
   OpenIme
End Sub

Private Sub cboTitle_KeyPress(KeyAscii As MSForms.ReturnInteger)
   If txtCustNo(0) <> "" Or txtCustNo(1) <> "" Or cboTitle.ListCount > 0 Then
      txtCustNo(0) = "": txtCustNo(1) = ""
      txtSales = "": lblSales = ""
      cboTitle.Clear
   End If
End Sub

Private Sub cmdA49_Click()
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
         Frmacc21r0.SetParent Me 'Add By Sindy 2016/11/29
         If Trim(txtCustNo(0)) <> "" Then
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
            Frmacc11p0.SetParent Me 'Add By Sindy 2016/11/29
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

Private Sub Combo1_Change()
   If Combo1 = "" Then
      Text3 = ""
      Exit Sub
   End If
   Text3 = A0802Query(Combo1)
End Sub

Private Sub Combo1_Click()
   Screen.MousePointer = vbHourglass
   Text3 = A0802Query(Combo1)
'   ProcessData
   AdodcRefresh
'Remove by Morgan 2004/1/29
'併入 AdodcRefresh 內的 SumShow3
'   SumShow
'Remove End-----
   Screen.MousePointer = vbDefault
End Sub

Private Sub Combo1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

'Add by Morgan 2004/2/11
'用最大收據的最大收文號抓智權人員簡稱
Private Function GetSaleShortName() As String
    Dim rstClone As New ADODB.Recordset, stSQL As String
    
    Set rstClone = Adodc1.Recordset.Clone
    'Modify By Sindy 2016/1/6 最大收據日期的最大收據的最大收文號
    'rstClone.Sort = "a1v02 desc,a1v01 desc"
    rstClone.Sort = "a0k02 desc,a1v02 desc,a1v01 desc"
    '2016/1/6 END
    rstClone.MoveFirst
    rstClone.Find "a1v08='Y'", 0, adSearchForward, 1
    If Not rstClone.EOF Then
        '2012/8/23 modif by sonia 改抓最大收據之收據智權人員
        'stSQL = "select sn01,CP13 from caseprogress, salesno where cp13 = sn02 (+) and cp09='" & rstClone.Fields("a1v01") & "'"
        'Modify By Sindy 2016/1/6 + acc1k0
        stSQL = "select sn01,a0k20 from acc0k0, salesno where a0k20 = sn02 (+) and a0k01 ='" & rstClone.Fields("a1v02") & "'" & _
                " union " & _
                "select sn01,CP13 from caseprogress, salesno, acc1k0 where cp13 = sn02 (+) and a1k01 ='" & rstClone.Fields("a1v02") & "' and cp09='" & rstClone.Fields("a1v01") & "'"
        If adoquery.State = adStateOpen Then adoquery.Close
        adoquery.CursorLocation = adUseClient
        adoquery.Open stSQL, adoTaie, adOpenStatic, adLockReadOnly
        If Not adoquery.EOF And Not IsNull(adoquery.Fields(0)) Then
            GetSaleShortName = adoquery.Fields(0)
        Else
            GetSaleShortName = adoquery.Fields(1)
        End If
    End If
    If adoquery.State = adStateOpen Then adoquery.Close
    Set rstClone = Nothing
End Function

'退費輸入
Private Sub Command2_Click()
Dim strCompNo As String 'Add By Sindy 2020/5/20
Dim ii As Integer 'Add By Sindy 2023/7/20
   
   'Add By Sindy 2020/5/20
   With Adodc1.Recordset
      .MoveFirst
      Do While Not .EOF
         If "" & .Fields("A1V08") <> "" Then
            If strCompNo <> "" Then
               If strCompNo = "L" Or "" & .Fields("A1V03") = "L" Then
                  If strCompNo <> "" & .Fields("A1V03") Then
                     MsgBox "法律所不可以跟其他公司合併扣繳!!", vbExclamation, "扣繳檢查"
                     Exit Sub
                  End If
               End If
            End If
            strCompNo = "" & .Fields("A1V03") '公司別
            'Add By Sindy 2023/7/20 檢查是否有收據編號多筆的漏勾選
            strSql = "select * from acc1v0 where a1v02='" & "" & .Fields("A1V02") & "' and a1v08 is null"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               MsgBox "收據編號( " & .Fields("A1V02") & " )漏勾選!!", vbExclamation, "扣繳檢查"
               DataGrid2.row = ii
               Exit Sub
            End If
            '2023/7/20 END
         End If
         ii = ii + 1 'Add By Sindy 2023/7/20
         .MoveNext
      Loop
      '游標指到Y第一筆
      .MoveFirst
      Do While Not .EOF
         If "" & .Fields("A1V08") <> "" Then
            Exit Do
         End If
         .MoveNext
      Loop
   End With
   '2020/5/20 END
   
   If strRNo <> "" Then
      If Right(strRNo, 1) = "," Then
         strCon4 = "(" & Mid(strRNo, 1, Len(strRNo) - 1) & ") "
      Else
         strCon4 = "(" & strRNo & ") "
      End If
   Else
      Exit Sub
   End If
   If adoquery.State = adStateOpen Then
      adoquery.Close
   End If
   adoquery.CursorLocation = adUseClient
   'Modify By Sindy 2016/1/6 + acc1k0
   'adoquery.Open "select a1v17 from acc0k0, acc1v0 where a0k01 = a1v02 and a1v04 <> a1v06 and a1v08 = '" & MsgText(602) & "' and (a1v17 is not null and substr(a1v17, 1, 1) <> 'E') and a1v01 in " & strCon4 & strSql, adoTaie, adOpenStatic, adLockReadOnly
   adoquery.Open "select a1v17 from acc0k0, acc1v0, acc0M0, acc0L0 where a0k01 = a1v02 and a1v04 <> a1v06 and a1v08 = '" & MsgText(602) & "' and (a1v17 is not null and substr(a1v17, 1, 1) <> 'E') and a0k01 = a0m02 and a0m01 = a0L01 and a1v01 in " & strCon4 & str1vCon & strSql0k & _
          " union select a1v17 from acc1k0, acc1v0, acc0z0, acc0y0 where a1k01 = a1v02 and a1v04 <> a1v06 and a1v08 = '" & MsgText(602) & "' and (a1v17 is not null and substr(a1v17, 1, 1) <> 'E') and a1k01 = a0z02 and a0z01 = a0y01 and a1v01 in " & strCon4 & str1vCon & strSql1k _
                 , adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      MsgBox MsgText(182), , MsgText(5)
      adoquery.Close
      Exit Sub
   End If
   adoquery.Close
   FeeShow
   If intCconfirm = vbCancel Then
      'Remove by Morgan 2007/5/8 離開時會清
      'If strRNo <> "" Then
      '   strRNo = "(" & Mid(strRNo, 1, Len(strRNo) - 1) & ")"
      '   adoTaie.Execute "update acc1v0 set a1v08 = null, a1v10 = 0 where a1v04 <> a1v06 and a1v08 = '" & MsgText(602) & "' and a1v01 in " & strRNo
      '   strRNo = ""
      'End If
      'end 2007/5/8
      
      AdodcRefresh
      'Remove by Morgan 2004/1/29
      '併入 AdodcRefresh 內的 SumShow3
      'SumShow
      'Remove End ------
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   
'Modify by Morgan 2003/12/15

'   If Text1 <> MsgText(601) Then
'      strItemNo = Trim(Text1) & Text2
'   Else
'      MsgBox MsgText(10), , MsgText(5)
'      Text1.SetFocus
'      Exit Sub
'   End If
   If cboTitle.Text <> MsgText(601) Then
       strItemNo = Trim(cboTitle.Text) & Text2
   Else
      MsgBox MsgText(10), , MsgText(5)
      cboTitle.SetFocus
      Exit Sub
   End If
   
'End 2003/12/15

   strCon1 = ""
   'Modify by Morgan 2004/2/7
   'Text10 已改為退費金額
'   If Text10 <> MsgText(601) And Val(Text10) <> 0 Then
'      strCon1 = Val(Text10) + Val(Text5)
   If Text10 <> MsgText(601) <> 0 Then
      strCon1 = Val(Text10)
   Else
      MsgBox MsgText(115), , MsgText(5)
      Exit Sub
   End If
   adocheck.CursorLocation = adUseClient
   adocheck.Open "select a0k03 from acc0k0 where a0k01 = '" & Adodc1.Recordset.Fields("a1v02").Value & "' union " & _
                 "select a1k03 as a0k03 from acc1k0 where a1k01 = '" & Adodc1.Recordset.Fields("a1v02").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adocheck.RecordCount <> 0 Then
      If IsNull(adocheck.Fields(0).Value) Then
         strCon2 = MsgText(601)
      Else
         strCon2 = adocheck.Fields(0).Value
      End If
   Else
      strCon2 = MsgText(601)
   End If
   adocheck.Close
   strCon3 = ""
   If Adodc1.Recordset.RecordCount <> 0 Then
      Adodc1.Recordset.MoveFirst
      Do While Adodc1.Recordset.EOF = False
         If Adodc1.Recordset.Fields("a1v08").Value = MsgText(602) Then
            strCon3 = strCon3 & "'" & Adodc1.Recordset.Fields("a1v01").Value & "', "
         End If
         Adodc1.Recordset.MoveNext
      Loop
      Adodc1.Recordset.MoveFirst
      If strCon3 <> "" Then
         strCon3 = " and a1v01 in (" & Mid(strCon3, 1, Len(strCon3) - 2) & ")"
      End If
   End If
   strFormLink = Name
   tool3_enabled
   'Add by Morgan 2005/2/25
   Frmacc11b1.strMemo = GetSaleShortName & "/" & Trim(cboTitle.Text) & Text2
   Frmacc11b1.intForm = 2
   Frmacc11b1.Show
   'Add by Morgan
   '傳票摘要
   'Modify by Morgan 2005/2/25 移到上面這樣才讀得到
   'Frmacc11b1.strMemo = GetSaleShortName & "/" & Trim(cboTitle.Text) & Text2
   
   
   Me.Enabled = False
   Screen.MousePointer = vbDefault
End Sub

Private Sub DataGrid2_AfterColUpdate(ByVal ColIndex As Integer)
'Dim strUptENo As String, intCurCol As Integer 'Add By Sindy 2023/7/20
'Dim strUpdSQL As String
   
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If

'   strUptENo = DataGrid2.Columns(2) '收據號碼
'   intCurCol = DataGrid2.row '目前列數
   'Modify by Morgan 2007/5/8 改在SumShow3做
   'Select Case DataGrid2.Col
   '   Case 7
   '      If DataGrid2.Columns(7).Text = MsgText(602) Then
   '         strRNo = strRNo & "'" & DataGrid2.Columns(14).Text & "',"
   '      End If
   '   Case 6
   'End Select
   'end 2007/5/8

   'Add By Sindy 2015/8/6
   With DataGrid2
      Select Case ColIndex
      Case 7 '退費否,個人不可扣繳
         If Format(.Columns(7)) <> "" Then
            'Add By Sindy 2015/10/26
            If Val(.Columns(4)) > 0 Then
               '未全額收齊,不能扣繳
               MsgBox "未全額收齊,不能扣繳!!", vbExclamation, "扣繳檢查"
               .Columns(7).Value = ""
            '2015/10/26 END
            ElseIf PUB_ChkIsPerson(.Columns(2)) = True Then
               .Columns(7).Value = ""
'            'Add By Sindy 2023/7/20
'            Else
'               If strUptENo <> "" Then
'                  strUpdSQL = "update acc1v0 set a1v08 = 'Y'" & _
'                              " where a1v02 = '" & strUptENo & "'"
'               End If
'               '2023/7/20 END
'            End If
'         'Add By Sindy 2023/7/20
'         Else
'            If strUptENo <> "" Then
'               strUpdSQL = "update acc1v0 set a1v08 = null" & _
'                           " where a1v02 = '" & strUptENo & "'"
'            '2023/7/20 END
            End If
         End If
      End Select
   End With
   '2015/8/6 END

   If Val(Format(DataGrid2.Columns(5).Value, "0")) + Val(Format(DataGrid2.Columns(6).Value, "0")) > Val(Format(DataGrid2.Columns(3).Value, "0")) Then
      MsgBox MsgText(122), , MsgText(5)
      DataGrid2.Columns(6).Value = 0
   End If

   Adodc1.Recordset.UpdateBatch

'   'Add By Sindy 2023/7/20
'   If strUpdSQL <> "" Then
'      adoTaie.Execute strUpdSQL
'      Adodc1.Recordset.Requery '*****
'      AdodcRefresh
'      DataGrid2.row = intCurCol
''               DataGrid2.row = intCurCol
''               SumShow3 'Add By Sindy 2023/7/17
''               Call cmdSearch_Click
'   End If
'   '2023/7/20 END

   'Modify by Morgan 2004/1/29
   'SumShow
   SumShow3
   'Modify End-----
End Sub

'Add by Morgan 2006/8/14
Private Sub DataGrid2_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
   Select Case ColIndex
      Case 6
         If Not IsNumeric(DataGrid2.Columns(ColIndex).Value) Then
            Cancel = True
         End If
   End Select
   If Val(Format(DataGrid2.Columns(5).Value, "0")) + Val(Format(DataGrid2.Columns(6).Value, "0")) > Val(Format(DataGrid2.Columns(3).Value, "0")) Then
      MsgBox MsgText(122), , MsgText(5)
      Cancel = True
   End If
End Sub

Private Sub DataGrid2_Click()
'Private Sub DataGrid2_DblClick()
'Dim strUptENo As String, intCurCol As Integer 'Add By Sindy 2023/7/17
   
'   'Add By Sindy 2023/7/17
'   strUptENo = DataGrid2.Columns(2).Text '收據號碼
'   intCurCol = DataGrid2.row '目前列數
'   '2023/7/17 END
'   Debug.Print intCurCol & " " & Now
   Select Case DataGrid2.col
      Case 7 '退費否
         'Modify by Morgan 2004/1/29
         '扣繳金額<>0才能上Y
         'If DataGrid2.Columns(7).Text = MsgText(601) Then
         If DataGrid2.Columns(7).Text = MsgText(601) And DataGrid2.Columns(6).Text <> "0" Then
            SendKeys "{Y}"
            
'            'Add By Sindy 2023/7/17
'            'Add By Sindy 2015/10/26
'            If Val(DataGrid2.Columns(4).Text) > 0 Then
'               '未全額收齊,不能扣繳
'               MsgBox "未全額收齊,不能扣繳!!", vbExclamation, "扣繳檢查"
'            '2015/10/26 END
'            'Add By Sindy 2015/8/6 個人不可扣繳
'            ElseIf PUB_ChkIsPerson(DataGrid2.Columns(2).Text) = True Then
'               MsgBox "個人不可扣繳!!", vbExclamation, "扣繳檢查"
'            Else
'            '2015/8/6 END
'               If strUptENo <> "" Then
'                  adoTaie.Execute "update acc1v0 set a1v08 = 'Y'" & _
'                                  " where a1v02 = '" & strUptENo & "'"
'                  Adodc1.Recordset.Requery '*****
'                  DataGrid2.row = intCurCol
'                  AdodcRefresh
''                  DataGrid2.row = intCurCol
''                  SumShow3 'Add By Sindy 2023/7/17
''                  Call cmdSearch_Click
'               End If
'            End If
'            '2023/7/17 END
         Else
            SendKeys "{BACKSPACE}"
            'SendKeys "{DEL}"
            
'            'Add By Sindy 2023/7/17
'            If strUptENo <> "" Then
'               adoTaie.Execute "update acc1v0 set a1v08 = null" & _
'                               " where a1v02 = '" & strUptENo & "'"
'               Adodc1.Recordset.Requery '*****
'               DataGrid2.row = intCurCol
'               AdodcRefresh
''               DataGrid2.row = intCurCol
''               SumShow3 'Add By Sindy 2023/7/17
''               Call cmdSearch_Click
'            End If
'            '2023/7/17 END
         End If
   End Select
   
'   If Val(Format(DataGrid2.Columns(5).Text, "0")) + Val(Format(DataGrid2.Columns(6).Text, "0")) > Val(Format(DataGrid2.Columns(3).Text, "0")) Then
'      MsgBox MsgText(122), , MsgText(5)
'      'DataGrid2.Columns(6).Value = 0
'      'Add By Sindy 2023/7/17
'      If strUptENo <> "" Then
'         adoTaie.Execute "update acc1v0 set a1v07 = 0" & _
'                         " where a1v02 = '" & strUptENo & "'"
'         Adodc1.Recordset.Requery '*****
'         DataGrid2.row = intCurCol
'         AdodcRefresh
''         DataGrid2.row = intCurCol
''         SumShow3
''         Call cmdSearch_Click
'      End If
'      '2023/7/17 END
'   End If
End Sub

Private Sub DataGrid2_GotFocus()
   DataGrid2.col = 6
End Sub

Private Sub DataGrid2_HeadClick(ByVal ColIndex As Integer)
'設定值 Y/NULL
Dim strValue As String, rstClone As New ADODB.Recordset

    If Adodc1.Recordset.State <> adStateOpen Or Adodc1.Recordset.EOF Then
        Exit Sub
    End If
    If ColIndex = 7 Then
        Me.Enabled = False
        Screen.MousePointer = vbHourglass
        Frmacc0000.StatusBar1.Panels(1).Text = "資料更新中...."
        Set rstClone = Adodc1.Recordset.Clone
        With rstClone
        .MoveFirst
        Do While Not .EOF
            If Val("" & .Fields("A1V07")) <> 0 And ("" & .Fields("A1V08").Value) <> "Y" Then
                strValue = "Y"
                Exit Do
            End If
            .MoveNext
        Loop
        .MoveFirst
        Do While Not .EOF
            If Val("" & .Fields("A1V07")) <> 0 And ("" & .Fields("A1V08").Value) <> strValue Then
                .Fields("A1V08").Value = strValue
            End If
            .MoveNext
        Loop
        .UpdateBatch
        End With
        Set rstClone = Nothing
        'Modify By Sindy 2016/1/21
        'Adodc1.Recordset.Resync
        Adodc1.Recordset.Requery
        '2016/1/21 END
        SumShow3
        Frmacc0000.StatusBar1.Panels(1).Text = ""
        Screen.MousePointer = vbDefault
        Me.Enabled = True
    End If
End Sub

Private Sub DataGrid2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   Select Case DataGrid2.col
      Case 7
      If KeyAscii = 89 Then
         intY = KeyAscii
      End If
   End Select
End Sub

Private Sub DataGrid2_KeyUp(KeyCode As Integer, Shift As Integer)
Dim intCounter As Integer

   Select Case KeyCode
      Case vbKeyReturn
         Select Case DataGrid2.col
            Case 6
               SendKeys "{RIGHT}"
            Case 7
               SendKeys "{RIGHT}"
               'Add by Morgan 2006/5/8 繳費年度鎖住,要向右多跳一格
               SendKeys "{RIGHT}"
            Case 8
               SendKeys "{RIGHT}"
            Case 9
               SendKeys "{DOWN}"
               For intCounter = 1 To 3
                  SendKeys "{LEFT}"
               Next intCounter
         End Select
   End Select
   KeyDefine KeyCode
End Sub

Private Sub Form_Activate()
   If strItemNo <> MsgText(601) Then
      AdodcRefresh
      'Remove by Morgan 2004/1/29
      '併入 AdodcRefresh 內的 SumShow3
      'SumShow
      'Remove end-------
   End If
   strFormName = Name
   strFormLink = MsgText(601)
   strItemNo = ""
   strCon1 = ""
   strCon2 = ""
   strCon3 = ""
   strCon6 = "" 'Add By Sindy 2016/1/6
   strCon7 = "" 'Add By Sindy 2016/1/6
   Frmacc0000.StatusBar1.Panels(1).Text = "按 尋找 調出補扣繳資料 / " & MsgText(98)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
   Frmacc0000.StatusBar1.Panels(1).Text = "按 尋找 調出補扣繳資料 / " & MsgText(98)
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 9645
   Me.Height = 5910
'Modify by Morgan 2003/12/17
'   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   MoveFormToCenter Me
   lblSales.Caption = ""
'Modify 2003/12/17
   Image1 = LoadPicture(strBackPicPath1)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   'Modify by Morgan 2004/4/8
   '預設年度改判斷4月
   'If Val(Right(ServerDate, 4)) >= 501 Then
   If Val(Right(strSrvDate(2), 4)) >= 401 Then
      Text2 = strSrvDate(2) \ 10000
   Else
      Text2 = strSrvDate(2) \ 10000 - 1
   End If
   strItemNo = MsgText(601)
   OpenTable
   Combo1.AddItem ""
   adoquery.CursorLocation = adUseClient
   adoquery.Open "select a0801 from acc080 order by a0801 asc", adoTaie, adOpenStatic, adLockReadOnly
   Do While adoquery.EOF = False
      Combo1.AddItem adoquery.Fields(0).Value
      adoquery.MoveNext
   Loop
   adoquery.Close
   Frmacc0000.StatusBar1.Panels(1).Text = "按 尋找 調出補扣繳資料 / " & MsgText(98)
   'Add By Sindy 2017/9/15
   mebRecDate(1).Mask = ""
   mebRecDate(1).Text = ""
   mebRecDate(1).Mask = DFormat
   mebRecDate(2).Mask = ""
   mebRecDate(2).Text = ""
   mebRecDate(2).Mask = DFormat
   '2017/9/15 END
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add by Morgan 2006/5/2
   '請除鎖定資料
   If PUB_GetLock("", m_OldKey) = False Then
      Cancel = 1
      Exit Sub
   End If
   '2006/5/2 end
   
   'Modify by Morgan 2007/5/8
   'If strRNo <> "" Then
   '   strRNo = "(" & Mid(strRNo, 1, Len(strRNo) - 1) & ")"
   '   adoTaie.Execute "update acc1v0 set a1v08 = '" & MsgText(601) & "' where a1v08 = '" & MsgText(602) & "' and a1v01 in " & strRNo
   '   strRNo = ""
   'End If
   fnUnCheck
   'end 2007/5/8
   
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   Set Frmacc11b0 = Nothing
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
'   strSql = MsgText(601)
'
''Modify by Morgan 2003/12/15
''   If Text1 <> MsgText(601) Then
''      strSQL = " and instr(a0k04, '" & Text1 & "') > 0"
''   End If
'
'   If cboTitle.Text <> MsgText(601) Then
'      '2011/10/20 MODIFY BY SONIA E10023515
'      'strSql = " and instr(a0k04, '" & cboTitle.Text & "') > 0"
'      strSql = " and instr(UPPER(a0k04), UPPER('" & cboTitle.Text & "')) > 0"
'   End If
'
'   If txtCustNo(0) <> "" Then
'      strSql = strSql & " and a0k03>='" & txtCustNo(0).Text & "'"
'   End If
'   If txtCustNo(1) <> "" Then
'      strSql = strSql & " and a0k03<='" & txtCustNo(1).Text & "'"
'   End If
'
''End 2003/12/15
'
'   If Text2 <> MsgText(601) Then
'      strSql = strSql & " and a0k16 = " & Val(Text2) & ""
'   End If
'   Select Case Combo1
'      Case "1", "2"
'         strSql = strSql & " and a0k11 in ('1', '2')"
'      Case Else
'         strSql = strSql & " and a0k11 = '" & Combo1 & "'"
'   End Select
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "select * from acc1v0 where a1v01 = 'z'", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  重新整理 Adodc 之資料
'
'*************************************************
Public Sub AdodcRefresh(Optional ByVal bolExact As Boolean = True)
Dim strName As String
Dim StrSQLa As String 'Add By Sindy 2014/10/15
   
On Error GoTo Checking
   
   Screen.MousePointer = vbHourglass
   strSql0k = MsgText(601)
   strSql1k = MsgText(601)
   str1vCon = MsgText(601)
   
   '收據抬頭
   If cboTitle.Text <> MsgText(601) Then
      If bolExact = True Then
         strSql0k = " and a0k04= '" & cboTitle.Text & "'"
         strSql1k = " and a1k35= '" & cboTitle.Text & "'"
      Else
         strSql0k = " and instrb(a0k04, '" & cboTitle.Text & "') = 1"
         strSql1k = " and instrb(a1k35, '" & cboTitle.Text & "') = 1"
      End If
'      strSQL1 = " and cu04 = '" & cboTitle.Text & "'"
'      strSQL2 = " and fa04 = '" & cboTitle.Text & "'"
   End If
   
   '客戶編號
   If txtCustNo(0) <> "" Then
      strSql0k = strSql0k & " and a0k03>='" & txtCustNo(0).Text & "'"
   End If
   If txtCustNo(1) <> "" Then
      strSql0k = strSql0k & " and a0k03<='" & txtCustNo(1).Text & "'"
   End If
   If txtCustNo(0) <> "" And txtCustNo(1) <> "" Then
      strSql1k = strSql1k & " and ((a1k03>='" & txtCustNo(0).Text & "' and a1k03<='" & txtCustNo(1).Text & "') or" & _
                                  "(a1k27>='" & txtCustNo(0).Text & "' and a1k27<='" & txtCustNo(1).Text & "') or" & _
                                  "(a1k28>='" & txtCustNo(0).Text & "' and a1k28<='" & txtCustNo(1).Text & "'))"
   ElseIf txtCustNo(0) <> "" Then
      strSql1k = strSql1k & " and (a1k03>='" & txtCustNo(0).Text & "' or a1k27>='" & txtCustNo(0).Text & "' or a1k28>='" & txtCustNo(0).Text & "')"
   ElseIf txtCustNo(1) <> "" Then
      strSql1k = strSql1k & " and (a1k03<='" & txtCustNo(1).Text & "' or a1k27<='" & txtCustNo(1).Text & "' or a1k28<='" & txtCustNo(1).Text & "')"
   End If
   
   '智權人員
   If txtSales <> "" Then
      strSql0k = strSql0k & " and a0k20||''='" & txtSales & "'"
      strSql1k = strSql1k & " and cp13||''='" & txtSales & "'"
   End If
   
   '扣繳年度
   If Text2 <> MsgText(601) Then
      'strSql0k = strSql0k & " and a0k16 = " & Val(Text2) & ""
      str1vCon = str1vCon & " and a1v09 = " & Val(Text2) & ""
   End If
   strSql0k = strSql0k & " and a0k11<>'J'" 'Add By Sindy 2013/12/30
   
   '公司別
   If Combo1 <> MsgText(601) Then
      'strSql0k = strSql0k & " and a0k11 = '" & Combo1 & "'"
      str1vCon = str1vCon & " and a1v03 = '" & Combo1 & "'"
   End If
   
   'Add By Sindy 2017/9/13
   '收款日期
   If mebRecDate(1).Text <> MsgText(601) And mebRecDate(1).Text <> MsgText(29) Then
      strSql0k = strSql0k & " and a0l02 >= " & Val(FCDate(mebRecDate(1).Text))
      strSql1k = strSql1k & " and a0y02 >= " & Val(FCDate(mebRecDate(1).Text))
   End If
   If mebRecDate(2).Text <> MsgText(601) And mebRecDate(2).Text <> MsgText(29) Then
      strSql0k = strSql0k & " and a0l02 <= " & Val(FCDate(mebRecDate(2).Text))
      strSql1k = strSql1k & " and a0y02 <= " & Val(FCDate(mebRecDate(2).Text))
   End If
   '2017/9/13 END
   
   adoquery.CursorLocation = adUseClient
'Modify by Morgan 2005/3/1 不判斷是否已被選取改判斷未開扣單
'   adoquery.Open "select a1v01 from acc0k0, acc1v0 where a0k01 = a1v02 and (a1v14 is null or a1v14 = '') and (a1v17 is null or substr(a1v17, 1, 1) = 'E') and (a0k09 is null or a0k09 = 0) and (a0k10 is null or a1v04 <> 0)" & strSQL & " union " & _
'                 "select a1v01 from acc1k0, acc1v0, customer where a1k01 = a1v02 and substr(a1k28, 1, 8) = cu01 (+) and substr(a1k28, 9, 1) = cu02 (+) and (a1v14 is null or a1v14 = '') and (a1v17 is null or substr(a1v17, 1, 1) = 'E') and (a1k12 is null or a1k12 = 0) and a1k25 is null" & strSQL1 & " union " & _
'                 "select a1v01 from acc1k0, acc1v0, fagent where a1k01 = a1v02 and substr(a1k28, 1, 8) = fa01 (+) and substr(a1k28, 9, 1) = fa02 (+) and (a1v14 is null or a1v14 = '') and (a1v17 is null or substr(a1v17, 1, 1) = 'E') and (a1k12 is null or a1k12 = 0) and a1k25 is null" & strSQL2 & _
'                 " order by a1v01 asc", adoTaie, adOpenStatic, adLockReadOnly
   'Modify By Sindy 2015/11/11
'   adoquery.Open "select a1v01 from acc0k0, acc1v0 where a0k01 = a1v02 and a1v15 is null and (a1v17 is null or substr(a1v17, 1, 1) = 'E') and (a0k09 is null or a0k09 = 0) and (a0k10 is null or a1v04 <> 0)" & strSql & " union " & _
'                 "select a1v01 from acc1k0, acc1v0, customer where a1k01 = a1v02 and substr(a1k28, 1, 8) = cu01 (+) and substr(a1k28, 9, 1) = cu02 (+) and a1v15 is null and (a1v17 is null or substr(a1v17, 1, 1) = 'E') and (a1k12 is null or a1k12 = 0) and a1k25 is null" & strSQL1 & " union " & _
'                 "select a1v01 from acc1k0, acc1v0, fagent where a1k01 = a1v02 and substr(a1k28, 1, 8) = fa01 (+) and substr(a1k28, 9, 1) = fa02 (+) and a1v15 is null and (a1v17 is null or substr(a1v17, 1, 1) = 'E') and (a1k12 is null or a1k12 = 0) and a1k25 is null" & strSQL2 & _
'                 " order by a1v01 asc", adoTaie, adOpenStatic, adLockReadOnly
   
   adoquery.Open "select a1v01 from acc0k0, acc1v0, acc0M0, acc0L0 where a0k01 = a1v02 and a1v15 is null and (a1v17 is null or substr(a1v17, 1, 1) = 'E') and (a0k09 is null or a0k09 = 0) and (a0k10 is null or a1v04 <> 0) and a0k01 = a0m02 and a0m01 = a0L01" & str1vCon & strSql0k & " union " & _
                 "select a1v01 from acc1k0, acc1v0, caseprogress, acc0z0, acc0y0 where a1k01 = a1v02 and a1v01=cp09(+) and cp09 is not null and a1v15 is null and (a1v17 is null or substr(a1v17, 1, 1) = 'E') and (a1k12 is null or a1k12 = 0) and a1k25 is null and a1k01 = a0z02 and a0z01 = a0y01" & str1vCon & strSql1k & _
                 " order by a1v01 asc", adoTaie, adOpenStatic, adLockReadOnly
   '2015/11/11 END
'2005/3/1 end
   Do While adoquery.EOF = False
      strName = strName & "'" & adoquery.Fields("a1v01").Value & "', "
      adoquery.MoveNext
   Loop
   adoquery.Close
   If strName <> "" Then
      strName = Mid(strName, 1, Len(strName) - 2)
   Else
      strName = "'z'"
   End If
   adoadodc1.Close
   adoadodc1.CursorLocation = adUseClient
   'Modify by Morgan 2003/12/12
   'adoadodc1.Open "select a1v01, a1v02, a1v03, a1v04, a1v05, a1v06, a1v07, a1v08, a1v09, a1v10, a1v11, a1v12, a1v13, a1v14, a1v15, a1v16, a1v17 from acc1v0 where a1v01 in (" & strName & ") order by a1v03 asc, a1v01 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   'Modify by Morgan 2006/8/14 加顯示欄位ex01
'   StrSQLa = "select a1v01, a1v02, a1v03, a1v04, a1v05, a1v06, a1v07, a1v08, a1v09, a1v10, a1v11, a1v12, a1v13, a1v14, a1v15" & _
'             ", a1v16, a1v17, a0k02,ST02, decode(a1v05,'Y',nvl(cp79,0)) ex01" & _
'             " from acc1v0, acc0k0, STAFF, caseprogress where cp09(+)=a1v01 and a0k01=a1v02 AND ST01(+)=A0K20" & _
'             " and a1v01 in (" & strName & ")" & str1vCon
   'Modify By Sindy 2014/10/13
   'Modify By Sindy 2016/2/25 + and nvl(a0k04,a1k35)='" & cboTitle.Text & "' Ex: 輸入鑫港企業 104年 E10419109 收據不應該顯示
   StrSQLa = "select a1v01, a1v02, a1v03, a1v04, a1v05, a1v06, a1v07, a1v08, a1v09, a1v10, a1v11, a1v12, a1v13, a1v14, a1v15" & _
             ", a1v16, a1v17, nvl(a0k02,a1k02) a0k02,nvl(s1.ST02,s2.ST02) ST02, decode(a1v05,'Y',nvl(cp79,0)) ex01" & _
             " from acc1v0, acc0k0, STAFF s1, caseprogress, STAFF s2, acc1k0" & _
             " where cp09(+)=a1v01 and a0k01(+)=a1v02 AND s1.ST01(+)=A0K20" & _
             " and a1v01 in (" & strName & ")" & str1vCon & _
             " and a1k01(+)=a1v02 AND s2.ST01(+)=cp13" & _
             " and nvl(a0k04,a1k35)='" & cboTitle.Text & "'"
'   'Add By Sindy 2014/10/13
'   If cboTitle.Text <> MsgText(601) Then
'      If bolExact = True Then
'         StrSQLa = StrSQLa & " and ((a0k04 is not null and a0k04= '" & cboTitle.Text & "') or a0k04 is null)"
'      Else
'         StrSQLa = StrSQLa & " and ((a0k04 is not null and instrb(a0k04, '" & cboTitle.Text & "') = 1) or a0k04 is null)"
'      End If
'   End If
'   '2014/10/13 END
'   'Add By Sindy 2015/11/11
'   StrSQLa = StrSQLa & " union select a1v01, a1v02, a1v03, a1v04, a1v05, a1v06, a1v07, a1v08, a1v09, a1v10, a1v11, a1v12, a1v13, a1v14, a1v15" & _
'             ", a1v16, a1v17, a1k02 a0k02,ST02, decode(a1v05,'Y',nvl(cp79,0)) ex01" & _
'             " from acc1v0, acc1k0, STAFF, caseprogress where cp09(+)=a1v01 and a1k01=a1v02 AND ST01(+)=cp13" & _
'             " and a1v01 in (" & strName & ")" & str1vCon
'   If cboTitle.Text <> MsgText(601) Then
'      If bolExact = True Then
'         StrSQLa = StrSQLa & " and ((a1k35 is not null and a1k35= '" & cboTitle.Text & "') or a1k35 is null)"
'      Else
'         StrSQLa = StrSQLa & " and ((a1k35 is not null and instrb(a1k35, '" & cboTitle.Text & "') = 1) or a1k35 is null)"
'      End If
'   End If
'   '2015/11/11 END
   StrSQLa = StrSQLa & " order by a1v03 asc, a1v01 asc"
   adoadodc1.Open StrSQLa, adoTaie, adOpenDynamic, adLockBatchOptimistic
   'End 2003/12/12
   cmdA49.Visible = False 'Add By Sindy 2016/11/15
   If adoadodc1.State = adStateOpen Then
      If adoadodc1.RecordCount = 0 Then
         MsgBox MsgText(28), , MsgText(5)
      'Add By Sindy 2016/11/15
      Else
         cmdA49.Visible = True
      End If
      '2016/11/15 END
   Else
      MsgBox MsgText(28), , MsgText(5)
   End If
   Adodc1.Recordset.Requery
   'Add by Morgan 2004/1/29
   '公司別小計
   SumShow3
   'Add end -------
   
   Screen.MousePointer = vbDefault
   Frmacc0000.StatusBar1.Panels(1).Text = "按 尋找 調出補扣繳資料 / " & MsgText(98)
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)
   Select Case KeyCode
      Case vbKeyF12
         If strControlButton = MsgText(602) Then
            strControlButton = MsgText(601)
            Exit Sub
         End If
         AdodcRefresh
   End Select
   KeyEnter KeyCode
End Sub

'*************************************************
'  計算並顯示合計
'
'*************************************************
Public Sub SumShow()
   adoaccsum.CursorLocation = adUseClient
   'Modify By Sindy 2015/11/11
'   adoaccsum.Open "select sum(a1v04), sum(a1v06), sum(a1v07) from acc0k0, acc1v0 where a0k01 = a1v02 and (a1v14 is null or a1v14 = '') and (a1v17 is null or substr(a1v17, 1, 1) = 'E') and (a0k09 is null or a0k09 = 0) and (a0k10 is null or a1v04 <> 0)" & strSql & " union " & _
'                  "select sum(a1v04), sum(a1v06), sum(a1v07) from acc1k0, acc1v0, customer where a1k01 = a1v02 and substr(a1k28, 1, 8) = cu01 (+) and substr(a1k28, 9, 1) = cu02 (+) and (a1v14 is null or a1v14 = '') and (a1v17 is null or substr(a1v17, 1, 1) = 'E') and (a1k12 is null or a1k12 = 0) and a1k25 is null" & strSQL1 & " union " & _
'                  "select sum(a1v04), sum(a1v06), sum(a1v07) from acc1k0, acc1v0, fagent where a1k01 = a1v02 and substr(a1k28, 1, 8) = fa01 (+) and substr(a1k28, 9, 1) = fa02 (+) and (a1v14 is null or a1v14 = '') and (a1v17 is null or substr(a1v17, 1, 1) = 'E') and (a1k12 is null or a1k12 = 0) and a1k25 is null" & strSQL2, adoTaie, adOpenStatic, adLockReadOnly
   adoaccsum.Open "select sum(a1v04), sum(a1v06), sum(a1v07) from (" & _
                  "select a0k01,a1v01,a1v02,a1v04,a1v06,a1v07 from acc0k0, acc1v0, acc0M0, acc0L0 where a0k01 = a1v02 and (a1v14 is null or a1v14 = '') and (a1v17 is null or substr(a1v17, 1, 1) = 'E') and (a0k09 is null or a0k09 = 0) and (a0k10 is null or a1v04 <> 0) and a0k01 = a0m02 and a0m01 = a0L01" & str1vCon & strSql0k & " union " & _
                  "select a1k01,a1v01,a1v02,a1v04,a1v06,a1v07 from acc1k0, acc1v0, caseprogress, acc0z0, acc0y0 where a1k01 = a1v02 and a1v01=cp09(+) and cp09 is not null and (a1v14 is null or a1v14 = '') and (a1v17 is null or substr(a1v17, 1, 1) = 'E') and (a1k12 is null or a1k12 = 0) and a1k25 is null and a1k01 = a0z02 and a0z01 = a0y01" & str1vCon & strSql1k & _
                  ")", adoTaie, adOpenStatic, adLockReadOnly
   '2015/11/11 END
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0).Value) Then
         Text8 = "0"
      Else
         Text8 = adoaccsum.Fields(0).Value
      End If
      If IsNull(adoaccsum.Fields(1).Value) Then
         Text9 = "0"
      Else
         Text9 = adoaccsum.Fields(1).Value
      End If
      adoquery.CursorLocation = adUseClient
      'Modify By Sindy 2015/11/11
'      adoquery.Open "select sum(a1v10), sum(a1v07) from acc0k0, acc1v0 where a0k01 = a1v02 and a1v04 <> a1v06 and a1v08 = '" & MsgText(602) & "' and (a1v14 is null or a1v14 = '') and (a1v17 is null or substr(a1v17, 1, 1) = 'E') and (a0k09 is null or a0k09 = 0) and (a0k10 is null or a1v04 <> 0)" & strSql & " union " & _
'                    "select sum(a1v10), sum(a1v07) from acc1k0, acc1v0, customer where a1k01 = a1v02 and substr(a1k28, 1, 8) = cu01 (+) and substr(a1k28, 9, 1) = cu02 (+) and a1v04 <> a1v06 and a1v08 = '" & MsgText(602) & "' and (a1v14 is null or a1v14 = '') and (a1v17 is null or substr(a1v17, 1, 1) = 'E') and (a1k12 is null or a1k12 = 0) and a1k25 is null" & strSQL1 & " union " & _
'                    "select sum(a1v10), sum(a1v07) from acc1k0, acc1v0, fagent where a1k01 = a1v02 and substr(a1k28, 1, 8) = fa01 (+) and substr(a1k28, 9, 1) = fa02 (+) and a1v04 <> a1v06 and a1v08 = '" & MsgText(602) & "' and (a1v14 is null or a1v14 = '') and (a1v17 is null or substr(a1v17, 1, 1) = 'E') and (a1k12 is null or a1k12 = 0) and a1k25 is null" & strSQL2, adoTaie, adOpenStatic, adLockReadOnly
      adoquery.Open "select sum(a1v10), sum(a1v07) from (" & _
                    "select a0k01,a1v01,a1v02,a1v10,a1v07 from acc0k0, acc1v0, acc0M0, acc0L0 where a0k01 = a1v02 and a1v04 <> a1v06 and a1v08 = '" & MsgText(602) & "' and (a1v14 is null or a1v14 = '') and (a1v17 is null or substr(a1v17, 1, 1) = 'E') and (a0k09 is null or a0k09 = 0) and (a0k10 is null or a1v04 <> 0) and a0k01 = a0m02 and a0m01 = a0L01" & str1vCon & strSql0k & " union " & _
                    "select a1k01,a1v01,a1v02,a1v10,a1v07 from acc1k0, acc1v0, caseprogress, acc0z0, acc0y0 where a1k01 = a1v02 and a1v01=cp09(+) and cp09 is not null and a1v04 <> a1v06 and a1v08 = '" & MsgText(602) & "' and (a1v14 is null or a1v14 = '') and (a1v17 is null or substr(a1v17, 1, 1) = 'E') and (a1k12 is null or a1k12 = 0) and a1k25 is null and a1k01 = a0z02 and a0z01 = a0y01" & str1vCon & strSql1k & _
                    ")", adoTaie, adOpenStatic, adLockReadOnly
      '2015/11/11 END
      If adoquery.RecordCount <> 0 Then
         If IsNull(adoquery.Fields(1).Value) Then
            Text10 = "0"
         Else
            Text10 = adoquery.Fields(1).Value
         End If
         If IsNull(adoquery.Fields(0).Value) Then
            Text5 = "0"
         Else
            Text5 = adoquery.Fields(0).Value
         End If
      Else
         Text10 = "0"
         Text5 = "0"
      End If
      adoquery.Close
   Else
      Text8 = "0"
      Text9 = "0"
      Text10 = "0"
      Text5 = "0"
   End If
   adoaccsum.Close
End Sub

''*************************************************
''  儲存資料表(國內收據資料)
''
''*************************************************
'Private Sub Acc0k0Save()
'On Error GoTo Checking
'   adoacc0k0.Close
'   adoacc0k0.CursorLocation = adUseClient
'   adoacc0k0.Open "select * from acc0k0 where a0k01 = '" & Adodc1.Recordset.Fields("t0301").Value & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
'   If adoacc0k0.RecordCount <> 0 Then
'      If IsNull(Adodc1.Recordset.Fields("t0304").Value) Then
'         adoacc0k0.Fields("a0k13").Value = Null
'      Else
'         adoacc0k0.Fields("a0k13").Value = Adodc1.Recordset.Fields("t0304").Value
'      End If
'
'      If Text2 = MsgText(601) Then
'         adoacc0k0.Fields("a0k16").Value = 0
'      Else
'         adoacc0k0.Fields("a0k16").Value = Val(Text2)
'      End If
'      adoacc0k0.UpdateBatch
'   End If
'Checking:
'   If Err.Number = 0 Then
'      Exit Sub
'   End If
'   MsgBox Err.Description, , MsgText(5)
'End Sub

'*************************************************
'  產生扣繳明細資料
'
'*************************************************
Public Sub ProcessData(Optional ByVal bolExact As Boolean = True)

Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
   
'Add by Morgan 2004/3/5-----------
'國內客戶不用再產生1v0資料，跳過中間程式碼
If txtCustNo(0) <> "" Then GoTo flgSkip
'Add end 2004/3/3-----------------
   
   'Modify by Morgan 2007/5/8
   'If strRNo <> "" Then
   '   strRNo = "(" & Mid(strRNo, 1, Len(strRNo) - 1) & ")"
   '   adoTaie.Execute "update acc1v0 set a1v08 = '" & MsgText(601) & "' where a1v08 = '" & MsgText(602) & "' and a1v01 in " & strRNo
   '   strRNo = ""
   'End If
   fnUnCheck
   'end 2007/5/8
   
   StatusView MsgText(192)
   strSql0k = MsgText(601)
   strSql1k = MsgText(601)
   str1vCon = MsgText(601)
   
'Modify by Morgan 2003/12/17
'   If Text1 <> MsgText(601) Then
'      strSQL = " and instrb(a0k04, '" & Text1 & "') = 1"
'      strSQL1 = " and instrb(cu04, '" & Text1 & "') = 1"
'      strSQL2 = " and instrb(fa04, '" & Text1 & "') = 1"
'   End If

   If cboTitle.Text <> MsgText(601) Then
      If bolExact = True Then
         strSql0k = " and a0k04= '" & cboTitle.Text & "'"
         strSql1k = " and a1k35= '" & cboTitle.Text & "'"
      Else
         strSql0k = " and instrb(a0k04, '" & cboTitle.Text & "') = 1"
         strSql1k = " and instrb(a1k35, '" & cboTitle.Text & "') = 1"
      End If
'      strSQL1 = " and instrb(cu04, '" & cboTitle.Text & "') = 1"
'      strSQL2 = " and instrb(fa04, '" & cboTitle.Text & "') = 1"
   End If
   
   If txtCustNo(0) <> "" Then
      strSql0k = strSql0k & " and a0k03>='" & txtCustNo(0).Text & "'"
   End If
   If txtCustNo(1) <> "" Then
      strSql0k = strSql0k & " and a0k03<='" & txtCustNo(1).Text & "'"
   End If
   If txtCustNo(0) <> "" And txtCustNo(1) <> "" Then
      strSql1k = strSql1k & " and ((a1k03>='" & txtCustNo(0).Text & "' and a1k03<='" & txtCustNo(1).Text & "') or" & _
                                  "(a1k27>='" & txtCustNo(0).Text & "' and a1k27<='" & txtCustNo(1).Text & "') or" & _
                                  "(a1k28>='" & txtCustNo(0).Text & "' and a1k28<='" & txtCustNo(1).Text & "'))"
   ElseIf txtCustNo(0) <> "" Then
      strSql1k = strSql1k & " and (a1k03>='" & txtCustNo(0).Text & "' or a1k27>='" & txtCustNo(0).Text & "' or a1k28>='" & txtCustNo(0).Text & "')"
   ElseIf txtCustNo(1) <> "" Then
      strSql1k = strSql1k & " and (a1k03<='" & txtCustNo(1).Text & "' or a1k27<='" & txtCustNo(1).Text & "' or a1k28<='" & txtCustNo(1).Text & "')"
   End If
   
   If txtSales <> "" Then
      strSql0k = strSql0k & " and a0k20||''='" & txtSales & "'"
      strSql1k = strSql1k & " and cp13||''='" & txtSales & "'"
   End If
'End 2003/12/17
   
   '扣繳年度
   If Text2 <> MsgText(601) Then
     strSql0k = strSql0k & " and a0k16 = " & Val(Text2) & ""
     '92.4.17 MODIFY BY SONIA
     'strSQL1 = strSQL1 & " and to_number(substr(a0y02, 1, length(a0y02) - 4)) = " & Val(Text2) & ""
     'strSQL2 = strSQL2 & " and to_number(substr(a0y02, 1, length(a0y02) - 4)) = " & Val(Text2) & ""
'     strSQL1 = strSQL1 & " and a0y02 >= " & (Val(Text2) * 10000 + 101) & " and a0y02 <= " & (Val(Text2) * 10000 + 1231)
'     strSQL2 = strSQL2 & " and a0y02 >= " & (Val(Text2) * 10000 + 101) & " and a0y02 <= " & (Val(Text2) * 10000 + 1231)
     '92.4.17 END
     strSql1k = strSql1k & " and a0y02 >= " & (Val(Text2) * 10000 + 101) & " and a0y02 <= " & (Val(Text2) * 10000 + 1231)
   End If
   
   '公司別
   If Combo1 <> MsgText(601) Then
      Select Case Combo1
         Case "1", "2"
            strSql0k = strSql0k & " and a0k11 in ('1', '2')"
            'Modify By Sindy 2020/6/4 nvl(a0z13, '1') => nvl(a0z13, '2')
            strSql1k = strSql1k & " and nvl(a0z13, '2') in ('1', '2')"
         Case Else
            strSql0k = strSql0k & " and a0k11 = '" & Combo1 & "'"
            'Modify By Sindy 2020/6/4 nvl(a0z13, '1') => nvl(a0z13, '2')
            strSql1k = strSql1k & " and nvl(a0z13, '2') = '" & Combo1 & "'"
      End Select
   'Add By Sindy 2013/12/30
   Else
      strSql = strSql & " and a0k11<>'J'"
   '2013/12/30 END
   End If
   
   'Add By Sindy 2017/9/13
   '收款日期
   If mebRecDate(1).Text <> MsgText(601) And mebRecDate(1).Text <> MsgText(29) Then
      strSql1k = strSql1k & " and a0y02 >= " & Val(FCDate(mebRecDate(1).Text))
   End If
   If mebRecDate(2).Text <> MsgText(601) And mebRecDate(2).Text <> MsgText(29) Then
      strSql1k = strSql1k & " and a0y02 <= " & Val(FCDate(mebRecDate(2).Text))
   End If
   '2017/9/13 END
   
   Screen.MousePointer = vbHourglass
   adoquery.CursorLocation = adUseClient
'   adoquery.Open "select cp09, cp60, a0k11, decode(a0k30, 'Y', (nvl(cp16, 0)) * 0.1 - nvl(cp76, 0), (nvl(cp16, 0) - nvl(cp17, 0)) * 0.1  - nvl(cp76, 0)) as TAmount, cp16, cp75, a0k16, a0j20, a0j21, cp76, decode(a0k30, 'Y', (nvl(cp16, 0)) * 0.1, (nvl(cp16, 0) - nvl(cp17, 0)) * 0.1) as RAmount, a0k04 from acc0k0, caseprogress, acc0j0 where a0k01 = cp60 (+) and a0k01 = a0j13 (+) and cp09 is not null" & strSQL & " order by a0k11 asc, cp09 asc", adoTaie, adOpenStatic, adLockReadOnly
   '92.12.18 ADD BY SONIA
   'adoquery.Open "select cp09, cp60, a0k11, decode(a0k30, 'Y', (nvl(cp16, 0) - nvl(cp77, 0)) * 0.1 - nvl(cp76, 0), (nvl(cp16, 0) - nvl(cp17, 0) - nvl(cp77, 0)) * 0.1  - nvl(cp76, 0)) as TAmount, cp16, cp75, a0k16, a0j20, a0j21, cp76, decode(a0k30, 'Y', (nvl(cp16, 0)- nvl(cp77, 0)) * 0.1, (nvl(cp16, 0) - nvl(cp17, 0) - nvl(cp77, 0)) * 0.1) as RAmount, a0k04 from acc0k0, caseprogress, acc0j0 where a0k01 = cp60 (+) and a0k01 = a0j13 (+) and cp09 is not null" & strSQL & " union " & _
   '              "select cp09, cp60, nvl(a0z13, '1') as a0k11, (nvl(cp16, 0) - nvl(cp17, 0)) * 0.1  - nvl(a0z12, 0) as TAmount, nvl(cp16, 0) as cp16, a0z04 as cp75, to_number(substr(a0y02, 1, length(a0y02) - 4)) as a0k16, '' as a0j20, '' as a0j21, a0z12 as cp76, ((nvl(cp16, 0) - nvl(cp17, 0)) * 0.1) as RAmount, fa04 as a0k04 from acc1k0, caseprogress, acc0z0, acc0y0, fagent where a1k01 = cp60 and a1k01 = a0z02 and a0z01 = a0y01 and substr(a1k28, 1, 8) = fa01 (+) and substr(a1k28, 9, 1) = fa02 (+) and cp09 is not null" & strSQL2 & " union " & _
   '              "select cp09, cp60, nvl(a0z13, '1') as a0k11, (nvl(cp16, 0) - nvl(cp17, 0)) * 0.1  - nvl(a0z12, 0) as TAmount, nvl(cp16, 0) as cp16, a0z04 as cp75, to_number(substr(a0y02, 1, length(a0y02) - 4)) as a0k16, '' as a0j20, '' as a0j21, a0z12 as cp76, ((nvl(cp16, 0) - nvl(cp17, 0)) * 0.1) as RAmount, cu04 as a0k04 from acc1k0, caseprogress, acc0z0, acc0y0, customer where a1k01 = cp60 and a1k01 = a0z02 and a0z01 = a0y01 and substr(a1k28, 1, 8) = cu01 (+) and substr(a1k28, 9, 1) = cu02 (+) and cp09 is not null" & strSQL1 & " order by a0k11 asc, cp09 asc", adoTaie, adOpenStatic, adLockReadOnly
   
   'Modify by Morgan 2004/3/16
   '只抓國外請款資料
   'adoquery.Open "select cp09, cp60, a0k11, decode(a0k30, 'Y', (nvl(cp16, 0) - nvl(a1u08, 0) - nvl(a1u10, 0)) * 0.1 - nvl(cp76, 0), (nvl(cp16, 0) - nvl(cp17, 0) - nvl(a1u08, 0)) * 0.1  - nvl(cp76, 0)) as TAmount, cp16, cp75, a0k16, a0j20, a0j21, cp76, decode(a0k30, 'Y', (nvl(cp16, 0) - nvl(a1u08, 0) - nvl(a1u10, 0)) * 0.1, (nvl(cp16, 0) - nvl(cp17, 0) - nvl(a1u08, 0)) * 0.1) as RAmount, a0k04, a0k13 from acc0k0, caseprogress, acc0j0, (select a1u01, sum(a1u08) as a1u08, sum(a1u10) as a1u10 from acc1u0 group by a1u01) where a0k01 = cp60 (+) and cp09 = a0j01 (+) and a0j01 = a1u01 (+) and cp09 is not null" & strSQL & " union " & _
                 "select cp09, cp60, nvl(a0z13, '1') as a0k11, (nvl(a0z04, 0) - nvl(A1K09, 0)) * 0.1  - nvl(a0z12, 0) as TAmount, nvl(A1K11, 0) as cp16, a0z04 as cp75, to_number(substr(a0y02, 1, length(a0y02) - 4)) as a0k16, '' as a0j20, '' as a0j21, a0z12 as cp76, ((nvl(cp16, 0) - nvl(cp17, 0)) * 0.1) as RAmount, fa04 as a0k04, '' as a0k13 from acc1k0, caseprogress, acc0z0, acc0y0, fagent where a1k01 = cp60 and a1k01 = a0z02 and a0z01 = a0y01 and substr(a1k28, 1, 8) = fa01 (+) and substr(a1k28, 9, 1) = fa02 (+) and cp09 is not null" & strSQL2 & " union " & _
                 "select cp09, cp60, nvl(a0z13, '1') as a0k11, (nvl(a0z04, 0) - nvl(A1K09, 0)) * 0.1  - nvl(a0z12, 0) as TAmount, nvl(A1K11, 0) as cp16, a0z04 as cp75, to_number(substr(a0y02, 1, length(a0y02) - 4)) as a0k16, '' as a0j20, '' as a0j21, a0z12 as cp76, ((nvl(cp16, 0) - nvl(cp17, 0)) * 0.1) as RAmount, cu04 as a0k04, '' as a0k13 from acc1k0, caseprogress, acc0z0, acc0y0, customer where a1k01 = cp60 and a1k01 = a0z02 and a0z01 = a0y01 and substr(a1k28, 1, 8) = cu01 (+) and substr(a1k28, 9, 1) = cu02 (+) and cp09 is not null" & strSQL1 & " order by a0k11 asc, cp09 asc", adoTaie, adOpenStatic, adLockReadOnly
   '92.12.18 END
   '93.12.7 MODIFY BY SONIA 一請款單多個收文號時扣繳金額應個別抓
   'adoquery.Open "select cp09, cp60, nvl(a0z13, '1') as a0k11, (nvl(a0z04, 0) - nvl(A1K09, 0)) * 0.1  - nvl(a0z12, 0) as TAmount, nvl(A1K11, 0) as cp16, a0z04 as cp75, to_number(substr(a0y02, 1, length(a0y02) - 4)) as a0k16, '' as a0j20, '' as a0j21, a0z12 as cp76, ((nvl(cp16, 0) - nvl(cp17, 0)) * 0.1) as RAmount, fa04 as a0k04, '' as a0k13 from acc1k0, caseprogress, acc0z0, acc0y0, fagent where a1k01 = cp60 and a1k01 = a0z02 and a0z01 = a0y01 and substr(a1k28, 1, 8) = fa01 (+) and substr(a1k28, 9, 1) = fa02 (+) and cp09 is not null" & strSQL2 & " union " & _
   '              "select cp09, cp60, nvl(a0z13, '1') as a0k11, (nvl(a0z04, 0) - nvl(A1K09, 0)) * 0.1  - nvl(a0z12, 0) as TAmount, nvl(A1K11, 0) as cp16, a0z04 as cp75, to_number(substr(a0y02, 1, length(a0y02) - 4)) as a0k16, '' as a0j20, '' as a0j21, a0z12 as cp76, ((nvl(cp16, 0) - nvl(cp17, 0)) * 0.1) as RAmount, cu04 as a0k04, '' as a0k13 from acc1k0, caseprogress, acc0z0, acc0y0, customer where a1k01 = cp60 and a1k01 = a0z02 and a0z01 = a0y01 and substr(a1k28, 1, 8) = cu01 (+) and substr(a1k28, 9, 1) = cu02 (+) and cp09 is not null" & strSQL1 & " order by a0k11 asc, cp09 asc", adoTaie, adOpenStatic, adLockReadOnly
   'Memo by Morgan 2011/12/27 取消 a0j20,a0j21 (這裡不用改)
   'Modify Sindy 2015/11/11
'   adoquery.Open "select cp09, cp60, nvl(a0z13, '1') as a0k11, DECODE((nvl(a0z04, 0) - nvl(A1K09, 0)) * 0.1  - nvl(a0z12, 0),0,0,((nvl(cp16, 0) - nvl(cp17, 0)) * 0.1)) as TAmount, nvl(A1K11, 0) as cp16, a0z04 as cp75, to_number(substr(a0y02, 1, length(a0y02) - 4)) as a0k16, '' as a0j20, '' as a0j21, DECODE(a0z12,0,0,((nvl(cp16, 0) - nvl(cp17, 0)) * 0.1)) as cp76, ((nvl(cp16, 0) - nvl(cp17, 0)) * 0.1) as RAmount, fa04 as a0k04, '' as a0k13 from acc1k0, caseprogress, acc0z0, acc0y0, fagent where a1k01 = cp60 and a1k01 = a0z02 and a0z01 = a0y01 and substr(a1k28, 1, 8) = fa01 (+) and substr(a1k28, 9, 1) = fa02 (+) and cp09 is not null" & strSQL2 & " union " & _
'                 "select cp09, cp60, nvl(a0z13, '1') as a0k11, DECODE((nvl(a0z04, 0) - nvl(A1K09, 0)) * 0.1  - nvl(a0z12, 0),0,0,((nvl(cp16, 0) - nvl(cp17, 0)) * 0.1)) as TAmount, nvl(A1K11, 0) as cp16, a0z04 as cp75, to_number(substr(a0y02, 1, length(a0y02) - 4)) as a0k16, '' as a0j20, '' as a0j21, DECODE(a0z12,0,0,((nvl(cp16, 0) - nvl(cp17, 0)) * 0.1)) as cp76, ((nvl(cp16, 0) - nvl(cp17, 0)) * 0.1) as RAmount, cu04 as a0k04, '' as a0k13 from acc1k0, caseprogress, acc0z0, acc0y0, customer where a1k01 = cp60 and a1k01 = a0z02 and a0z01 = a0y01 and substr(a1k28, 1, 8) = cu01 (+) and substr(a1k28, 9, 1) = cu02 (+) and cp09 is not null" & strSQL1 & _
'                 " order by a0k11 asc, cp09 asc", adoTaie, adOpenStatic, adLockReadOnly
   'Modify By Sindy 2020/6/4 nvl(a0z13, '1') => nvl(a0z13, '2')
   adoquery.Open "select cp09, cp60, nvl(a0z13, '2') as a0k11, DECODE((nvl(a0z04, 0) - nvl(A1K09, 0)) * 0.1  - nvl(a0z12, 0),0,0,((nvl(cp16, 0) - nvl(cp17, 0)) * 0.1)) as TAmount, nvl(A1K11, 0) as cp16, a0z04 as cp75, to_number(substr(a0y02, 1, length(a0y02) - 4)) as a0k16, '' as a0j20, '' as a0j21, DECODE(a0z12,0,0,((nvl(cp16, 0) - nvl(cp17, 0)) * 0.1)) as cp76, ((nvl(cp16, 0) - nvl(cp17, 0)) * 0.1) as RAmount, a1k35 as a0k04, '' as a0k13 from acc1k0, caseprogress, acc0z0, acc0y0 where a1k01 = cp60 and a1k01 = a0z02 and a0z01 = a0y01 and cp09 is not null" & strSql1k & _
                 " order by a0k11 asc, cp09 asc", adoTaie, adOpenStatic, adLockReadOnly
   '2015/11/11 END
   '93.12.7 END
    '2004/3/16
   If adoquery.RecordCount <> 0 Then
      If IsNull(adoquery.Fields("a0k04").Value) = False Then
      
'Remove by Morgan 2003/12/17
''Modify by Morgan 2003/12/15
''         Text1 = Trim(adoQuery.Fields("a0k04").Value)
'         cboTitle.Clear
'         cboTitle.AddItem Trim(adoquery.Fields("a0k04").Value)
'         cboTitle.ListIndex = 0
''Modify 2003/12/15
'Remove 2003/12/17

      End If
   End If
   Do While adoquery.EOF = False
      With adoquery
         adocheck.CursorLocation = adUseClient
         adocheck.Open "select a1v01 from acc1v0 where a1v01 = '" & .Fields("cp09").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
         If adocheck.RecordCount = 0 Then
            If adoaccsum.State = adStateOpen Then
               adoaccsum.Close
            End If
            adoaccsum.CursorLocation = adUseClient
            adoaccsum.Open "select a0m03 from acc0m0 where a0m02 = '" & adoquery.Fields("cp60").Value & "' union " & _
                           "select null as a0m03 from acc0z0 where a0z02 = '" & adoquery.Fields("cp60").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
            'Memo by Morgan 2011/12/27 取消 a0j20,a0j21 (這裡不用改,因為上面都是放Null)
            If adoaccsum.RecordCount <> 0 Then
               If IsNull(adoaccsum.Fields("a0m03").Value) = False Then
                  adoTaie.Execute "insert into acc1v0 values ('" & .Fields("cp09").Value & "', '" & .Fields("cp60").Value & "', '" & .Fields("a0k11").Value & "', " & .Fields("RAmount").Value & ", '" & IIf(IsNull(.Fields("a0k13").Value) = True Or .Fields("a0k13").Value = "", "N", "Y") & "', " & IIf(IsNull(.Fields("cp76").Value), 0, .Fields("cp76").Value) & ", " & .Fields("TAmount").Value & ", null, " & .Fields("a0k16").Value & ", 0, 0, '" & IIf(IsNull(.Fields("a0j20").Value), "", .Fields("a0j20").Value) & "', '" & IIf(IsNull(.Fields("a0j21").Value), "", .Fields("a0j21").Value) & "', null, null, null, '" & adoaccsum.Fields("a0m03").Value & "', " & IIf(IsNull(.Fields("cp76").Value) = True Or .Fields("cp76").Value = 0, "null", "'1'") & ")"
               Else
                  adoTaie.Execute "insert into acc1v0 values ('" & .Fields("cp09").Value & "', '" & .Fields("cp60").Value & "', '" & .Fields("a0k11").Value & "', " & .Fields("RAmount").Value & ", '" & IIf(IsNull(.Fields("a0k13").Value) = True Or .Fields("a0k13").Value = "", "N", "Y") & "', " & IIf(IsNull(.Fields("cp76").Value), 0, .Fields("cp76").Value) & ", " & .Fields("TAmount").Value & ", null, " & .Fields("a0k16").Value & ", 0, 0, '" & IIf(IsNull(.Fields("a0j20").Value), "", .Fields("a0j20").Value) & "', '" & IIf(IsNull(.Fields("a0j21").Value), "", .Fields("a0j21").Value) & "', null, null, null, null, " & IIf(IsNull(.Fields("cp76").Value) = True Or .Fields("cp76").Value = 0, "null", "'1'") & ")"
               End If
            Else
               adoTaie.Execute "insert into acc1v0 values ('" & .Fields("cp09").Value & "', '" & .Fields("cp60").Value & "', '" & .Fields("a0k11").Value & "', " & .Fields("RAmount").Value & ", '" & IIf(IsNull(.Fields("a0k13").Value) = True Or .Fields("a0k13").Value = "", "N", "Y") & "', " & IIf(IsNull(.Fields("cp76").Value), 0, .Fields("cp76").Value) & ", " & .Fields("TAmount").Value & ", null, " & .Fields("a0k16").Value & ", 0, 0, '" & IIf(IsNull(.Fields("a0j20").Value), "", .Fields("a0j20").Value) & "', '" & IIf(IsNull(.Fields("a0j21").Value), "", .Fields("a0j21").Value) & "', null, null, null, null, " & IIf(IsNull(.Fields("cp76").Value) = True Or .Fields("cp76").Value = 0, "null", "'1'") & ")"
            End If
            adoaccsum.Close
         End If
         adocheck.Close
      End With
      adoquery.MoveNext
   Loop
   adoquery.Close
   
flgSkip:
   adoTaie.Execute "delete from acc1v0 where a1v06 = 0 and a1v07 = 0"
''Modify by Morgan 2003/12/17
'   AdodcRefresh
   AdodcRefresh bolExact
'Modify by 2003/12/17
'Remove by Morgan 2004/1/29
'併入 AdodcRefresh 內的 SumShow3
'   SumShow
'Remove End-------
   Screen.MousePointer = vbDefault
   StatusView MsgText(601)
   Frmacc0000.StatusBar1.Panels(1).Text = "按 尋找 調出補扣繳資料 / " & MsgText(98)
End Sub

Private Sub Text2_GotFocus()
   'edit by nickc 2007/06/11  切換輸入法改用API
   'Text2.IMEMode = 2
   CloseIme
   TextInverse Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And (KeyAscii < 48 Or KeyAscii > 57) Then
      KeyAscii = 0
   End If
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
'   AdodcRefresh
'   SumShow
End Sub

'*************************************************
'  計算並顯示退費資料
'
'*************************************************
Public Sub FeeShow()
    'Modify by Morgan 2004/2/7
    'text10 已改為退費金額
    'intCconfirm = MsgBox(MsgText(121) & Format(Val(Text10) + Val(Text5), DDollar), vbOKCancel + vbDefaultButton1, MsgText(5))
    intCconfirm = MsgBox(MsgText(121) & Format(Val(Text10), DDollar), vbOKCancel + vbDefaultButton1, MsgText(5))
End Sub

'Add by Morgan 2003/12/15
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

'Add by Morgan 2003/12/15
Private Sub txtCustNo_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Add by Morgan 2003/12/15
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

'Add by Morgan 2003/12/15
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
   End If
   cboTitle.ListIndex = 0
End Sub

'Add by Morgan 2003/12/15
Private Sub cboTitle_Validate(Cancel As Boolean)
   If CheckLen(Label1, cboTitle, 100) = MsgText(603) Then
      Cancel = True
      Exit Sub
   End If
   If Cancel = False Then CloseIme
End Sub

'Add by Morgan 2003/12/15
Private Sub cmdLikeSearch_Click()
   
   If cboTitle.Text = "" Then
      MsgBox "請輸入收據抬頭！", vbCritical
   Else
      txtCustNo(0) = "": txtCustNo(1) = ""  'add by sonia 2023/11/7 輸第二筆時要清前一筆的客戶編號
      'Modify by Morgan 2011/3/10 改呼叫共用函數
      'AddItem2CboTitle
      'Modify By Sindy 2013/12/30 +true(不含J公司)
      PUB_AddItem2CboTitle cboTitle, txtCustNo(0), txtCustNo(1), Text2, True
   End If
End Sub

''Add by Morgan 2003/12/15
'Private Function AddItem2CboTitle() As Boolean
'
'   Dim strSql As String, strConA As String, strConB As String
'   Dim adoQuery As New ADODB.Recordset
'   Dim strItem As String
'
'On Error GoTo ErrHand
'
'   strConA = ""
'   If Text2 <> "" Then
'      strConA = " and a0k16=" & Text2
'   End If
'
'   strConB = ""
'   If txtCustNo(0) <> "" Then
'      strConB = strConB & " and a0k03>='" & txtCustNo(0).Text & "'"
'   End If
'   If txtCustNo(1) <> "" Then
'      strConB = strConB & " and a0k03<='" & txtCustNo(1).Text & "'"
'   End If
'
'   '2011/10/20 MODIFY BY SONIA E10023515
'   'strSql = "Select distinct rpad(a0k04, 60,' ') C01, a0k03 C02" & _
'      " from Acc0k0 where substrb(a0k01,-4)>'2000' and a0k04 like '" & cboTitle.Text & "%'" & strConA & strConB & _
'      " order by 2,1"
'   strSql = "Select distinct rpad(a0k04, 60,' ') C01, a0k03 C02" & _
'      " from Acc0k0 where substrb(a0k01,-4)>'2000' and instr(upper(a0k04),upper('" & cboTitle.Text & "'))>0" & strConA & strConB & _
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
'
'   AddItem2CboTitle = True
'   Exit Function
'
'ErrHand:
'   MsgBox Err.Description
'
'End Function


'Add by Morgan 2003/12/17
Private Sub txtSales_GotFocus()
   'edit by nickc 2007/06/11  切換輸入法改用API
   'txtSales.IMEMode = 2
   CloseIme
   TextInverse txtSales
End Sub

Private Sub txtSales_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Add by Morgan 2003/12/17
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
'Add by Morgan 2003/12/15
Private Sub cmdSearch_Click()
   If ChkConOK = True Then
      ProcessData
   End If
   Frmacc0000.StatusBar1.Panels(1).Text = "按 尋找 調出補扣繳資料 / " & MsgText(98)
End Sub
'Add by Morgan 2003/12/17
Private Sub cmdSearch1_Click()
   If ChkConOK = True Then
      ProcessData (False)
   End If
   Frmacc0000.StatusBar1.Panels(1).Text = "按 尋找 調出補扣繳資料 / " & MsgText(98)
End Sub

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
   If Combo1.Text = "J" Then
      MsgBox "公司別不可為 J 公司！", vbCritical
      Combo1.SetFocus
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
   End If
   '2006/5/2 end
End Function

'Remove by Morgan 2003/12/15
'Private Sub Text1_GotFocus()
'   TextInverse Text1
'   StatusView MsgText(65) & "100" & " / " & MsgText(157)
'End Sub

'Private Sub Text1_LostFocus()
'   StatusView MsgText(601)
'End Sub
'
'Private Sub Text1_Validate(Cancel As Boolean)
'   If CheckLen(Label1, Text1, 100) = MsgText(603) Then
'      Cancel = True
'      Exit Sub
'   End If
'   If Text1 = "" Then
'      Exit Sub
'   End If
'   ProcessData
'   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(157) & " / " & MsgText(98)
'End Sub

'Add by Morgan 2004/1/29
'產生公司別小計資料
Private Sub SumShow3()
    'Modify By Sindy 2020/5/7
    'Dim intComp As Integer
    Dim strCompNo As String
    '2020/5/7 END
    Dim arrlngSum(1 To 5) As Long, iArrIdx As Integer, ii As Integer
    Dim rstClone As New ADODB.Recordset
    
On Error GoTo ErrHnd
    
    cboSubTotal.Clear
    '應扣繳款
    Text8 = ""
    '已扣繳款
    Text9 = ""
    '有勾退費的未扣繳額+調整稅款
    Text10 = ""
    '有勾退費的調整稅款
    Text5 = "":
    '未扣繳額
    Text1 = ""
    For ii = 1 To 5
        txtSum(ii) = ""
    Next
    
    If Adodc1.Recordset.State <> adStateOpen Or Adodc1.Recordset.EOF Then
        Exit Sub
    End If
    
    Set rstClone = Adodc1.Recordset.Clone()
    With rstClone
    .MoveFirst
    iArrIdx = 0
    
    'Modify By Sindy 2020/5/7
    'intComp = "" & .Fields("A1V03").Value
    strCompNo = "" & .Fields("A1V03").Value
    '2020/5/7 END
    
    strRNo = "" 'Add by Morgan 2007/5/8
    strCon6 = "" 'Add By Sindy 2017/3/17
    strCon7 = "" 'Add By Sindy 2017/3/17
    Do While Not .EOF
         'Add by Morgan 2007/5/8
         If "" & .Fields("A1V08") = "Y" Then
            strRNo = strRNo & "'" & .Fields("A1V01").Value & "',"
            'Add By Sindy 2015/11/11
            'Modify By Sindy 2017/3/17 只做請款單Acc1k0=X編號的
            If Left(.Fields("A1V02").Value, 1) = "X" Then
            '2017/3/17 END
               strCon6 = cboTitle.Text '收據抬頭
               strCon7 = .Fields("A1V09") '扣繳年度
            End If
            '2015/11/11 END
         End If
         '2007/5/8
         
        'Modify By Sindy 2020/5/7
        'If intComp <> ("" & .Fields("A1V03").Value) Then
        If strCompNo <> ("" & .Fields("A1V03").Value) Then
            'cboSubTotal.AddItem intComp
            'arrSum(iArrIdx) = intComp & "," & arrlngSum(1) & "," & arrlngSum(2) & "," & arrlngSum(3) & "," & arrlngSum(4) & "," & arrlngSum(5)
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
        
        Text8 = Val(Text8) + Val(Format("" & .Fields("A1V04").Value, "0"))
        Text9 = Val(Text9) + Val(Format("" & .Fields("A1V06").Value, "0"))
        Text10 = Val(Text10) + IIf("" & .Fields("A1V08").Value = "Y", Val(Format("" & .Fields("A1V07").Value, "0")) + Val(Format("" & .Fields("A1V10").Value, "0")), 0)
        Text5 = Val(Text5) + IIf("" & .Fields("A1V08").Value = "Y", Val(Format("" & .Fields("A1V10").Value, "0")), 0)
        Text1 = Val(Text1) + Val(Format("" & .Fields("A1V07").Value, "0"))
        
        'Modify By Sindy 2020/5/7
        'intComp = "" & .Fields("A1V03").Value
        strCompNo = "" & .Fields("A1V03").Value
        '2020/5/7 END
        .MoveNext
    Loop
    'Modify By Sindy 2020/5/7
    'cboSubTotal.AddItem intComp
    'arrSum(iArrIdx) = intComp & "," & arrlngSum(1) & "," & arrlngSum(2) & "," & arrlngSum(3) & "," & arrlngSum(4) & "," & arrlngSum(5)
    cboSubTotal.AddItem strCompNo
    arrSum(iArrIdx) = strCompNo & "," & arrlngSum(1) & "," & arrlngSum(2) & "," & arrlngSum(3) & "," & arrlngSum(4) & "," & arrlngSum(5)
    '2020/5/7 END
    .MoveFirst
    End With
    cboSubTotal.ListIndex = 0
    Set rstClone = Nothing
    
ErrHnd:

    If Err.Number <> 0 Then
        MsgBox Err.Description
    End If
End Sub

'Add by Morgan 2007/5/8 取消選擇
Private Function fnUnCheck() As Boolean
On Error GoTo ErrHnd

   If strRNo <> "" Then
      strRNo = "(" & Mid(strRNo, 1, Len(strRNo) - 1) & ")"
      adoTaie.Execute "update acc1v0 set a1v08 = null where a1v15 is null and a1v14 = '" & MsgText(602) & "' and a1v01 in " & strRNo
      strRNo = ""
   End If
   fnUnCheck = True
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   
End Function
