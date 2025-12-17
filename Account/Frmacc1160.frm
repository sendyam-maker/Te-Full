VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc1160 
   AutoRedraw      =   -1  'True
   Caption         =   "廠商基本資料"
   ClientHeight    =   5360
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5360
   ScaleWidth      =   8760
   Begin VB.TextBox Text15 
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
      Left            =   2055
      MaxLength       =   7
      TabIndex        =   12
      Top             =   1800
      Width           =   1995
   End
   Begin VB.TextBox Text14 
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
      Left            =   1320
      MaxLength       =   3
      TabIndex        =   3
      Top             =   540
      Width           =   1125
   End
   Begin VB.TextBox Text12 
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
      Left            =   1320
      MaxLength       =   15
      TabIndex        =   14
      Top             =   2130
      Width           =   2745
   End
   Begin VB.TextBox Text11 
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
      Left            =   2025
      MaxLength       =   10
      TabIndex        =   16
      Top             =   2460
      Width           =   2025
   End
   Begin VB.TextBox Text10 
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
      Left            =   5337
      TabIndex        =   9
      Top             =   1155
      Width           =   3075
   End
   Begin VB.TextBox Text9 
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
      Left            =   5337
      MaxLength       =   50
      TabIndex        =   15
      Top             =   2130
      Width           =   3075
   End
   Begin VB.TextBox Text8 
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
      Left            =   1320
      MaxLength       =   9
      TabIndex        =   10
      Top             =   1470
      Width           =   2745
   End
   Begin VB.TextBox Text7 
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
      Left            =   5337
      MaxLength       =   15
      TabIndex        =   13
      Top             =   1800
      Width           =   3075
   End
   Begin VB.TextBox Text5 
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
      Left            =   1320
      MaxLength       =   1
      TabIndex        =   8
      Top             =   1155
      Width           =   405
   End
   Begin VB.CommandButton Command3 
      Height          =   300
      Left            =   2550
      Picture         =   "Frmacc1160.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   210
      Width           =   350
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc1160.frx":0102
      Height          =   2115
      Left            =   240
      TabIndex        =   18
      Top             =   2940
      Width           =   8295
      _ExtentX        =   14623
      _ExtentY        =   3739
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   18
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.5
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
      Caption         =   "廠商資料"
      ColumnCount     =   17
      BeginProperty Column00 
         DataField       =   "a0i01"
         Caption         =   "廠商編號"
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
         DataField       =   "a0i02"
         Caption         =   "廠商名稱"
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
         DataField       =   "a0i05"
         Caption         =   "電話"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "####-####"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "a0i18"
         Caption         =   "身分證/統編"
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
         DataField       =   "a0i07"
         Caption         =   "傳真"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "####-####"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "a0i08"
         Caption         =   "聯絡人"
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
         DataField       =   "a0i03"
         Caption         =   "地址"
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
         DataField       =   "a0i12"
         Caption         =   "付款方式"
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
         DataField       =   "A0I17"
         Caption         =   "帳戶名稱"
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
         DataField       =   "a0i13"
         Caption         =   "付款銀行"
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
         DataField       =   "a0i20"
         Caption         =   "銀行及分行代碼"
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
         DataField       =   "a0i14"
         Caption         =   "付款帳號"
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
         DataField       =   "a0i15"
         Caption         =   "一信編碼"
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
         DataField       =   "a0i16"
         Caption         =   "EMAIL"
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
         DataField       =   "a0i17"
         Caption         =   "帳戶名稱"
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
         DataField       =   "a0i18"
         Caption         =   "身分證字號/統編"
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
         DataField       =   "a0i19"
         Caption         =   "手機"
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
            ColumnWidth     =   1450.205
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3559.748
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1349.858
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1349.858
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1319.811
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   3809.764
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   980.221
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1539.78
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1920.189
         EndProperty
         BeginProperty Column10 
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1830.047
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   1869.732
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   2489.953
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   1440
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   1810.205
         EndProperty
         BeginProperty Column16 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text1 
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
      Left            =   1320
      MaxLength       =   5
      TabIndex        =   0
      Top             =   210
      Width           =   1215
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1320
      TabIndex        =   5
      Top             =   855
      Width           =   1575
      _ExtentX        =   2787
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   15
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
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   300
      Left            =   4080
      TabIndex        =   6
      Top             =   855
      Width           =   1575
      _ExtentX        =   2787
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   15
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
      Left            =   0
      Top             =   3000
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
   Begin MSForms.TextBox Text13 
      Height          =   315
      Left            =   5337
      TabIndex        =   17
      Top             =   2460
      Width           =   3075
      VariousPropertyBits=   679493659
      MaxLength       =   500
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text6 
      Height          =   315
      Left            =   5337
      TabIndex        =   11
      Top             =   1470
      Width           =   3075
      VariousPropertyBits=   679493659
      MaxLength       =   15
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text4 
      Height          =   315
      Left            =   6840
      TabIndex        =   7
      Top             =   855
      Width           =   1572
      VariousPropertyBits=   679493659
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text3 
      Height          =   315
      Left            =   2445
      TabIndex        =   4
      Top             =   540
      Width           =   5970
      VariousPropertyBits=   679493659
      MaxLength       =   35
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text2 
      Height          =   315
      Left            =   4080
      TabIndex        =   2
      Top             =   210
      Width           =   4332
      VariousPropertyBits=   679493659
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "銀行及分行代碼"
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
      Left            =   360
      TabIndex        =   34
      Top             =   1830
      Width           =   1575
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "備註"
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
      Left            =   4365
      TabIndex        =   33
      Top             =   2520
      Width           =   450
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "手機"
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
      Left            =   360
      TabIndex        =   32
      Top             =   2182
      Width           =   450
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '透明
      Caption         =   "身分證字號/統編"
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
      Left            =   360
      TabIndex        =   31
      Top             =   2520
      Width           =   1665
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "帳戶名稱"
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
      Left            =   4365
      TabIndex        =   30
      Top             =   1200
      Width           =   915
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "EMAIL"
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
      Left            =   4365
      TabIndex        =   29
      Top             =   2182
      Width           =   690
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "瑞興編碼"
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
      Left            =   360
      TabIndex        =   28
      Top             =   1515
      Width           =   900
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "付款帳號"
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
      Left            =   4365
      TabIndex        =   27
      Top             =   1845
      Width           =   900
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "付款銀行"
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
      Left            =   4365
      TabIndex        =   26
      Top             =   1522
      Width           =   915
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "付款方式　　（1.電匯 2.瑞興直存）"
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
      Left            =   360
      TabIndex        =   25
      Top             =   1200
      Width           =   3480
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   4800
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   2745
      Left            =   225
      Top             =   120
      Width           =   8295
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "聯絡人"
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
      Left            =   5880
      TabIndex        =   24
      Top             =   900
      Width           =   690
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "傳真"
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
      Left            =   3120
      TabIndex        =   23
      Top             =   900
      Width           =   465
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "電話"
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
      Left            =   360
      TabIndex        =   22
      Top             =   900
      Width           =   465
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "地址"
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
      Left            =   360
      TabIndex        =   21
      Top             =   592
      Width           =   465
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "廠商名稱"
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
      Left            =   3120
      TabIndex        =   20
      Top             =   255
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "廠商編號"
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
      Left            =   360
      TabIndex        =   19
      Top             =   255
      Width           =   915
   End
End
Attribute VB_Name = "Frmacc1160"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/08/09 Form2.0已修改 Text2/Text3/Text4/Text6/Text13/DataGrid1
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/11/26 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/28 日期欄已修改
Option Explicit
Public adoadodc1 As New ADODB.Recordset

Private Sub Command3_Click()
   If Adodc1.Recordset.RecordCount = 0 Or Text1 = MsgText(601) Then
      Exit Sub
   End If
   Adodc1.Recordset.Find "a0i01 = '" & Text1 & "'", 0, adSearchForward, 1
   If Adodc1.Recordset.EOF = False Then
      FormShow
      RecordShow
   Else
      MsgBox MsgText(33), , MsgText(5)
      Adodc1.Recordset.MoveFirst
   End If
End Sub

Private Sub Command3_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Command3_Click
         Exit Sub
   End Select
   KeyEnter KeyCode
End Sub

Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub DataGrid1_SelChange(Cancel As Integer)
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   FormShow
   RecordShow
End Sub

Private Sub Form_Activate()
   strFormName = Name
   If strCompanyNo = MsgText(601) Then
      Exit Sub
   End If
   If Adodc1.Recordset.RecordCount <> 0 Then
      Adodc1.Recordset.MoveFirst
   End If
   Adodc1.Recordset.Find "a0i01 = '" & strCompanyNo & "'", 0, adSearchForward, 1
   If Adodc1.Recordset.EOF = False Then
      FormShow
      RecordShow
   End If
   strCompanyNo = MsgText(601)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
Dim intCounter As Integer
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 8980
   Me.Height = 5920
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath1)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   strCompanyNo = MsgText(601)
   OpenTable
   If Adodc1.Recordset.RecordCount <> 0 Then
      RecordShow
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
      Cancel = 1
      Exit Sub
   End If
   StatusClear
   strConTitle = MsgText(601)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc1160 = Nothing
End Sub

Private Sub MaskEdBox1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub MaskEdBox2_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "select * from acc0i0 order by a0i01 desc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  顯示資料表(廠商資料)
'
'*************************************************
Public Sub FormShow()
   Text1 = Adodc1.Recordset.Fields("a0i01").Value
   If IsNull(Adodc1.Recordset.Fields("a0i02").Value) Then
      Text2 = MsgText(601)
   Else
      Text2 = Adodc1.Recordset.Fields("a0i02").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a0i03").Value) Then
      Text3 = MsgText(601)
   Else
      Text3 = Adodc1.Recordset.Fields("a0i03").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a0i05").Value) Then
      MaskEdBox1.Text = MsgText(601)
   Else
      MaskEdBox1.Text = Adodc1.Recordset.Fields("a0i05").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a0i07").Value) Then
      MaskEdBox2.Text = MsgText(601)
   Else
      MaskEdBox2.Text = Adodc1.Recordset.Fields("a0i07").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a0i08").Value) Then
      Text4 = MsgText(601)
   Else
      Text4 = Adodc1.Recordset.Fields("a0i08").Value
   End If
   'Add by Morgan 2006/6/12
   Text5 = "" & Adodc1.Recordset.Fields("a0i12").Value
   Text6 = "" & Adodc1.Recordset.Fields("a0i13").Value
   Text7 = "" & Adodc1.Recordset.Fields("a0i14").Value
   Text8 = "" & Adodc1.Recordset.Fields("a0i15").Value
   Text9 = "" & Adodc1.Recordset.Fields("a0i16").Value
   'end 2006/6/12
   Text10 = "" & Adodc1.Recordset.Fields("a0i17").Value 'Add by Morgan 2007/2/5
   Text11 = "" & Adodc1.Recordset.Fields("a0i18").Value 'Add by Morgan 2007/6/6
   Text12 = "" & Adodc1.Recordset.Fields("a0i19").Value 'Add by Morgan 2007/6/6
   Text14 = "" & Adodc1.Recordset.Fields("a0i04").Value 'Add by Morgan 2009/1/21
   Text13 = "" & Adodc1.Recordset.Fields("a0i06").Value 'Add by Morgan 2009/1/21
   Text15 = "" & Adodc1.Recordset.Fields("a0i20").Value 'Add by Morgan 2010/6/20
   
   'Added by Morgan 2012/3/26
   '外譯編號員工檔欄位不可修改
   'Modified by Morgan 2012/11/19 改用Locked以方便複製(原用Enabled)
   If Left(Text1, 1) = "F" Then
      Text2.Locked = True
      Text14.Locked = True
      Text3.Locked = True
      MaskEdBox1.Enabled = False
      MaskEdBox2.Enabled = False
      Text12.Locked = True
      Text9.Locked = True
   Else
      Text2.Locked = False
      Text14.Locked = False
      Text3.Locked = False
      MaskEdBox1.Enabled = True
      MaskEdBox2.Enabled = True
      Text12.Locked = False
      Text9.Locked = False
   End If
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   If strSaveConfirm = MsgText(3) Then
      strExc(0) = "select st02,st08 from staff where st01='" & Text1 & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         Text2 = "" & RsTemp.Fields("st02")
         Text3 = "" & RsTemp.Fields("st08")
      End If
   End If
End Sub

Private Sub Text10_GotFocus()
'edit by nickc 2007/06/11  切換輸入法改用API
OpenIme
End Sub

Private Sub Text10_LostFocus()
'edit by nickc 2007/06/11  切換輸入法改用API
CloseIme
End Sub

'Add by Amy 2023/05/15
Private Sub Text10_Validate(Cancel As Boolean)
    If Text10 = MsgText(601) Then Exit Sub
    'Memo 此處放寬,廠商名稱也要修改
    If CheckLen(Label12, Text10, 80) = MsgText(603) Then
        Cancel = True
        Exit Sub
    End If
End Sub

Private Sub Text11_GotFocus()
   TextInverse Text11
End Sub

Private Sub Text11_Validate(Cancel As Boolean)
   Dim strTmp As String, i As Integer
   If Text11.Text <> "" Then
      If Len(Text11) = 10 Then
         i = 0
      ElseIf Len(Text11) = 8 Then
         i = 1
      Else
         strTmp = "身分證字號/統編碼數錯誤是否要繼續?"
         If MsgBox(strTmp, vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
            Cancel = True
         End If
         Exit Sub
      End If
      If CheckID(i, Text11.Text) = False Then
         If i = 0 Then
            strTmp = "身分證字號錯誤，是否確定 ?"
         Else
            strTmp = "統一編號錯誤，是否確定 ?"
         End If
         If MsgBox(strTmp, vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
            Cancel = True
         End If
      End If
   End If
End Sub

Private Sub Text12_GotFocus()
   TextInverse Text12
End Sub
'Add by Morgan 2009/1/21 郵遞區號
Private Sub Text14_GotFocus()
   TextInverse Text14
   CloseIme
End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
   KeyAscii = ChangeZIP(KeyAscii)
End Sub

Private Sub Text15_GotFocus()
   CloseIme
   TextInverse Text15
End Sub
'Add by Morgan 2011/6/20
Private Sub Text15_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text2_GotFocus()
   StatusView MsgText(65) & "40"
   TextInverse Text2
   'edit by nickc 2007/06/11  切換輸入法改用API
   OpenIme
End Sub

'Modify by Amy 2021/08/09 改Form2.0 原:Integer
Private Sub Text2_KeyUp(KeyCode As MSForms.ReturnInteger, Shift As Integer)
   KeyEnter Val(KeyCode)
End Sub

Private Sub Text2_LostFocus()
   StatusView MsgText(601)
   'edit by nickc 2007/06/11  切換輸入法改用API
   CloseIme
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
   'Memo 此處放寬,帳戶名稱也要修改
   'Modify by Amy 2023/05/15 原:40
   If CheckLen(Label2, Text2, 80) = MsgText(603) Then
      Cancel = True
      Exit Sub
   End If
End Sub

Private Sub Text3_GotFocus()
   StatusView MsgText(65) & "70"
   TextInverse Text3
   'edit by nickc 2007/06/11  切換輸入法改用API
   OpenIme
End Sub

'Add by Morgan 2009/1/21 轉全形
'Modify by Amy 2021/08/09 改Form2.0 原:Integer
Private Sub Text3_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = ChangeZIP(KeyAscii)
End Sub

'Modify by Amy 2021/08/09 改Form2.0 原:Integer
Private Sub Text3_KeyUp(KeyCode As MSForms.ReturnInteger, Shift As Integer)
   KeyEnter Val(KeyCode)
End Sub

Private Sub Text3_LostFocus()
   StatusView MsgText(601)
   'edit by nickc 2007/06/11  切換輸入法改用API
   CloseIme
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
   If CheckLen(Label3, Text3, 70) = MsgText(603) Then
      Cancel = True
      Exit Sub
   End If
End Sub

Private Sub Text4_GotFocus()
   StatusView MsgText(65) & "10"
   TextInverse Text4
   'edit by nickc 2007/06/11  切換輸入法改用API
   OpenIme
End Sub

'Modify by Amy 2021/08/09 改Form2.0 原:Integer
Private Sub Text4_KeyUp(KeyCode As MSForms.ReturnInteger, Shift As Integer)
   KeyEnter Val(KeyCode)
End Sub

'*************************************************
'  重新整理 Adodc 之資料
'
'*************************************************
Public Sub AdodcRefresh()
On Error GoTo Checking
   adoadodc1.Close
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "select * from acc0i0 order by a0i01 desc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Adodc1.Recordset.Requery
   If Adodc1.Recordset.RecordCount <> 0 Then
      If Text1 <> MsgText(601) Then
         Adodc1.Recordset.Find "a0i01 = '" & Text1 & "'", 0, adSearchForward, 1
         If Adodc1.Recordset.EOF = False Then
            FormShow
            RecordShow
         End If
      End If
   End If
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  顯示筆數
'
'*************************************************
Public Sub RecordShow()
   Frmacc0000.StatusBar1.Panels(2).Text = Adodc1.Recordset.Bookmark & MsgText(35) & Adodc1.Recordset.RecordCount
End Sub

Private Sub Text4_LostFocus()
   StatusView MsgText(601)
   'edit by nickc 2007/06/11  切換輸入法改用API
   CloseIme
End Sub

Private Sub Text4_Validate(Cancel As Boolean)
   If CheckLen(Label6, Text4, 10) = MsgText(603) Then
      Cancel = True
      Exit Sub
   End If
End Sub
'Add by Morgan 2006/6/12
Private Sub Text5_GotFocus()
   TextInverse Text5
   'edit by nickc 2007/06/11  切換輸入法改用API
   'If pub_OS = 1 Then
   '   Text5.IMEMode = 2
   'End If
   CloseIme
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text6_GotFocus()
   TextInverse Text6
   'edit by nickc 2007/06/11  切換輸入法改用API
   'If pub_OS = 1 Then
   '   Text6.IMEMode = 1
   'End If
   OpenIme
End Sub

Private Sub Text6_Validate(Cancel As Boolean)
   If CheckLen(Label8, Text6, 30) = MsgText(603) Then
      Cancel = True
      Exit Sub
   End If
End Sub

Private Sub Text7_GotFocus()
   TextInverse Text7
   'edit by nickc 2007/06/11  切換輸入法改用API
   'If pub_OS = 1 Then
   '   Text7.IMEMode = 2
   'End If
   CloseIme
End Sub

Private Sub Text8_GotFocus()
   TextInverse Text8
   'edit by nickc 2007/06/11  切換輸入法改用API
   'If pub_OS = 1 Then
   '   Text8.IMEMode = 2
   'End If
   CloseIme
End Sub

Private Sub Text9_GotFocus()
   TextInverse Text9
   'edit by nickc 2007/06/11  切換輸入法改用API
   'If pub_OS = 1 Then
   '   Text9.IMEMode = 2
   'End If
   CloseIme
End Sub
'end 2006/6/12

'Add by Amy 2014/01/07
Public Function TxtValidate() As Boolean
    Dim bCancel As Boolean
    TxtValidate = False
    
    bCancel = False
    'Modify by Amy 2023/05/15 從aacc_sav 搬過來
    If Text1 = MsgText(601) Then
        MsgBox MsgText(10), , MsgText(5)
        strControlButton = MsgText(602)
        Text1.SetFocus
        Exit Function
    Else
        Call Text2_Validate(bCancel)
        If bCancel = True Then
            strControlButton = MsgText(602)
            Text2.SetFocus
            Exit Function
         End If
        If Trim(Text10) <> MsgText(601) Then
            Call Text10_Validate(bCancel)
            If bCancel = True Then
                strControlButton = MsgText(602)
                Text10.SetFocus
                Exit Function
             End If
         End If
         'Add by Amy 2015/09/10 +有值才檢查
         If Trim(Text3) <> MsgText(601) Then
            If CheckLen(Label3, Text3, 70) = MsgText(603) Then
               strControlButton = MsgText(602)
               Text3.SetFocus
               Exit Function
            '2011/10/18 add by sonia 檢查地址
            ElseIf CheckTaiwanAddr(Text3, "000", "地址") = False Then
               strControlButton = MsgText(602)
               Text3.SetFocus
               Exit Function
            '2011/10/18 end
            End If
         End If
         'end 2015/09/10
         If CheckLen(Label6, Text4, 10) = MsgText(603) Then
            strControlButton = MsgText(602)
            Text4.SetFocus
            Exit Function
         End If
         'Add by Morgan 2011/6/20
         If Text5 = "1" Then
            If Text15 = "" Then
               MsgBox "電匯廠商的【" & Label16 & "】欄位不可空白！"
               strControlButton = MsgText(602)
               Text15.SetFocus
               Exit Function
               
            ElseIf Len(Text15) <> 7 Then
               MsgBox "【" & Label16 & "】欄位必須為 7 碼數字！"
               strControlButton = MsgText(602)
               Text15.SetFocus
               Exit Function
            End If
            If Text7 = "" Then
               MsgBox "電匯廠商的【" & Label9 & "】欄位不可空白！"
               strControlButton = MsgText(602)
               Text7.SetFocus
               Exit Function
               
            ElseIf Len(Text7) <> 14 Then
               MsgBox "【" & Label9 & "】欄位必須為 14 碼數字！"
               strControlButton = MsgText(602)
               Text7.SetFocus
               Exit Function
            End If
         End If
         'end 2011/6/20
    End If
    
    Call Text11_Validate(bCancel)
    If bCancel = True Then
         strControlButton = MsgText(602)
        Text11.SetFocus
        Exit Function
    End If
    'Add by Morgan 2007/8/20
    If Left(Text1, 1) = "F" Then
        If Text3 = "" Then
            If MsgBox("此編號為翻譯人員但未輸入" & Label3 & "，是否要繼續？", vbYesNo + vbDefaultButton2) = vbNo Then
                strControlButton = MsgText(602)
                Text3.SetFocus
                Exit Function
            End If
        End If
    End If
    If strSaveConfirm = MsgText(3) Then
        If Adodc1.Recordset.RecordCount <> 0 Then
            Adodc1.Recordset.Find "a0i01 = '" & Text1 & "'", 0, adSearchForward, 1
            If Adodc1.Recordset.EOF = False Then
                MsgBox "廠商資料已存在！", vbExclamation
                strControlButton = MsgText(602)
                Exit Function
            End If
        End If
    End If
    'end 2023/05/15
    'Add by Amy 2021/08/20 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
    If PUB_ChkUniText(Me) = False Then
        strControlButton = MsgText(602)
        Exit Function
    End If
    TxtValidate = True
End Function

'Add by Amy 2022/03/11 從aacc_cls搬過來
Public Sub Frmacc1160_Clear()
   With Frmacc1160
      .Text1 = ""
      .Text2 = ""
      .Text3 = ""
      .MaskEdBox1.Text = ""
      .MaskEdBox2.Text = ""
      .Text4 = ""
      'Add by Morgan 2006/7/5
      .Text5 = ""
      .Text6 = ""
      .Text7 = ""
      .Text8 = ""
      .Text9 = ""
      'end 2006/7/5
      .Text10 = "" 'Add by Morgan 2007/8/10
      'Add by Morgan 2007/12/24
      .Text11 = ""
      .Text12 = ""
      'end 2007/12/24
      'Add by Amy 2022/03/11 新增資料要清空-瑞婷
      .Text13 = ""
      .Text14 = ""
      .Text15 = ""
      'end 2022/03/11
      .Text1.SetFocus
   End With
End Sub
