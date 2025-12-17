VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc2211 
   AutoRedraw      =   -1  'True
   Caption         =   "請款資料查詢"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   350
   ClientWidth     =   9110
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5070
   ScaleWidth      =   9110
   Begin VB.ComboBox Combo4 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "Frmacc2211.frx":0000
      Left            =   6650
      List            =   "Frmacc2211.frx":0010
      Style           =   2  '單純下拉式
      TabIndex        =   31
      Top             =   1500
      Width           =   2070
   End
   Begin VB.TextBox Text10 
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
      Height          =   330
      Left            =   7550
      TabIndex        =   29
      Top             =   1095
      Width           =   945
   End
   Begin VB.TextBox Text9 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7320
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   27
      Top             =   716
      Width           =   372
   End
   Begin VB.CommandButton cmdSharePoint 
      Caption         =   "點數分配"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7260
      TabIndex        =   26
      Top             =   1890
      Width           =   1245
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1230
      MaxLength       =   15
      TabIndex        =   0
      Top             =   50
      Width           =   1572
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1230
      TabIndex        =   13
      Top             =   383
      Width           =   1572
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7950
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   12
      Top             =   50
      Width           =   372
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6270
      TabIndex        =   11
      Top             =   1095
      Width           =   615
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6675
      TabIndex        =   10
      Top             =   383
      Width           =   492
   End
   Begin VB.TextBox Text13 
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
      Left            =   4110
      TabIndex        =   9
      Top             =   4680
      Width           =   1575
   End
   Begin VB.TextBox Text14 
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
      Left            =   7470
      TabIndex        =   8
      Top             =   4680
      Width           =   1092
   End
   Begin VB.TextBox Text21 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7155
      TabIndex        =   7
      Top             =   383
      Width           =   852
   End
   Begin VB.TextBox Text22 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7995
      TabIndex        =   6
      Top             =   383
      Width           =   252
   End
   Begin VB.TextBox Text23 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8235
      TabIndex        =   5
      Top             =   383
      Width           =   372
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1230
      Locked          =   -1  'True
      MaxLength       =   9
      TabIndex        =   2
      Top             =   716
      Width           =   1572
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3990
      Locked          =   -1  'True
      MaxLength       =   9
      TabIndex        =   3
      Top             =   716
      Width           =   1572
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc2211.frx":0042
      Height          =   2420
      Left            =   30
      TabIndex        =   4
      Top             =   2220
      Width           =   9050
      _ExtentX        =   15963
      _ExtentY        =   4269
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   17
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
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "a1l02"
         Caption         =   "項次"
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
         DataField       =   "a1l04"
         Caption         =   "請款項目"
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
         DataField       =   "a1j03"
         Caption         =   "中文名稱"
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
         DataField       =   "a1l05"
         Caption         =   "請款金額(台幣)"
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
         DataField       =   "a1l07"
         Caption         =   "折扣(台幣)"
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
         DataField       =   "a1l16"
         Caption         =   "輸入幣別"
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
         DataField       =   "a1l17"
         Caption         =   "輸入幣別金額"
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
            ColumnWidth     =   500.032
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   909.921
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2229.732
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1069.795
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   920.126
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1230.236
         EndProperty
      EndProperty
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   330
      Left            =   3990
      TabIndex        =   1
      Top             =   50
      Width           =   1575
      _ExtentX        =   2787
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   150
      Top             =   1935
      Visible         =   0   'False
      Width           =   1200
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
   Begin MSForms.TextBox Text3 
      Height          =   330
      Left            =   2820
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   390
      Width           =   2715
      VariousPropertyBits=   671105055
      BackColor       =   14737632
      Size            =   "4789;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text11 
      Height          =   420
      Left            =   1230
      TabIndex        =   35
      Top             =   1050
      Width           =   4380
      VariousPropertyBits=   -1466941409
      BackColor       =   16777215
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "7726;741"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text12 
      Height          =   630
      Left            =   1230
      TabIndex        =   34
      Top             =   1500
      Width           =   4380
      VariousPropertyBits=   -1466941409
      BackColor       =   16777215
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "7726;1111"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "注意事項"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   270
      TabIndex        =   33
      Top             =   1575
      Width           =   840
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "列印幣別"
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
      Left            =   5715
      TabIndex        =   32
      Top             =   1567
      Width           =   900
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "匯率"
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
      Left            =   7020
      TabIndex        =   30
      Top             =   1155
      Width           =   450
   End
   Begin VB.Label Label20 
      BackStyle       =   0  '透明
      Caption         =   "是否特殊請款單      (Y:是 C:整批)"
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
      Left            =   5715
      TabIndex        =   28
      Top             =   754
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "請款編號"
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
      Left            =   270
      TabIndex        =   25
      Top             =   88
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "請款日期"
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
      Left            =   3030
      TabIndex        =   24
      Top             =   88
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "代理人"
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
      Left            =   270
      TabIndex        =   23
      Top             =   421
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "是否列印申請人(Y/N)"
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
      Left            =   5715
      TabIndex        =   22
      Top             =   88
      Width           =   2295
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "幣別"
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
      Left            =   5715
      TabIndex        =   21
      Top             =   1133
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
      Left            =   5715
      TabIndex        =   20
      Top             =   421
      Width           =   975
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "列印備註"
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
      Left            =   270
      TabIndex        =   19
      Top             =   1155
      Width           =   900
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "合計"
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
      Left            =   2070
      TabIndex        =   18
      Top             =   4710
      Width           =   495
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "折扣後金額 NTD"
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
      Left            =   5790
      TabIndex        =   17
      Top             =   4710
      Width           =   1665
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '透明
      Caption         =   "美金金額 USD"
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
      Left            =   2670
      TabIndex        =   16
      Top             =   4710
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Top             =   3840
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label15 
      BackStyle       =   0  '透明
      Caption         =   "列印對象"
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
      Left            =   270
      TabIndex        =   15
      Top             =   754
      Width           =   975
   End
   Begin VB.Label Label17 
      BackStyle       =   0  '透明
      Caption         =   "請款對象"
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
      Left            =   3030
      TabIndex        =   14
      Top             =   754
      Width           =   975
   End
End
Attribute VB_Name = "Frmacc2211"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/09 改成Form2.0 ; DataGrid1改字型=新細明體-ExtB、Text3、Text12、Text11
'Memo By Sonia 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit

Public adoacc1k0 As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
Dim strSql As String
Dim strNo As String
Dim douAmount As Double
Dim strAmount As String
Dim intLength As Integer
Dim intCounter As Integer
Dim douUSDollar As Double
Dim strLanguage As String
Dim strMaxNo As String
Dim strDiscount As String

'Add by Morgan 2010/4/22
Private Sub cmdSharePoint_Click()
   Frmacc21h3.m_bolQuery = True
   Frmacc21h3.Show vbModal
   strFormName = Me.Name
   tool3_enabled
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   'Modified by Lydia 2021/12/09 改成模組
'   Me.Icon = LoadPicture(strIcoPath)
'   strFormName = Name
'   'Modified by Lydia 2015/04/15
'   'Me.Width = 8850
'   Me.Width = 9225
'   Me.Height = 5100
'   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
'   Image1 = LoadPicture(strBackPicPath1)
'   sglWidth = Image1.Width
'   sglHeight = Image1.Height
'   For intX = 0 To Int(ScaleWidth / sglWidth)
'       For intY = 0 To Int(ScaleHeight / sglHeight)
'           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
'       Next
'   Next
   strFormName = Name
   PUB_InitForm Me, 9225, 5500, strBackPicPath1
   'end 2021/12/09
   OpenTable
   SumShow
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strItemNo = ""
   tool3_enabled
   Select Case strFormLink
      Case "Frmacc2210"
         Frmacc2210.Enabled = True
      Case "Frmacc2220"
         Frmacc2220.Enabled = True
      'Add By Sindy 2014/2/18
      Case "frm040205a"
         frm040205a.Enabled = True
      '2014/2/18 END
   End Select
   Set Frmacc2211 = Nothing
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
Dim strSystemKind As String
   
On Error GoTo Checking
   
   Text10 = MsgText(601) 'Add By Sindy 2011/01/06
   adoacc1k0.CursorLocation = adUseClient
   adoacc1k0.Open "select * from acc1k0 where a1k01 = '" & strItemNo & "'", adoTaie, adOpenStatic, adLockReadOnly
   adoadodc1.CursorLocation = adUseClient
   FormShow
   'modify by sonia 2025/8/27 加入left join(+)，否則改系統類別無請款項目代號時就會缺資料X11410363由S案轉入FCT，原請款項目代號0012在FCT沒有
   adoadodc1.Open "select * from acc1l0, acc1j0 where a1l03 = a1j01(+) and a1l04 = a1j02(+) and a1l01 = '" & Text1 & "' order by a1l02 asc", adoTaie, adOpenStatic, adLockReadOnly
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  顯示資料表
'
'*************************************************
Public Sub FormShow()
   Text1 = strItemNo
   MaskEdBox1.Mask = MsgText(601)
   If IsNull(adoacc1k0.Fields("a1k02").Value) Then
      MaskEdBox1.Text = CFDate(strSrvDate(2))
   Else
      MaskEdBox1.Text = CFDate(adoacc1k0.Fields("a1k02").Value)
   End If
   MaskEdBox1.Mask = DFormat
   If IsNull(adoacc1k0.Fields("a1k03").Value) Then
      Text2 = MsgText(601)
   Else
      If Len(adoacc1k0.Fields("a1k03").Value) = 6 Then
         Text2 = adoacc1k0.Fields("a1k03").Value & "000"
      Else
         Text2 = adoacc1k0.Fields("a1k03").Value
      End If
   End If
   '2005/7/8 MODIFY BY SONIA
   'Text3 = FagentQuery(Text2, 2)
   Select Case Mid(Text2, 1, 1)
      Case "Y"
         Text3 = FagentQuery(Text2, 2)
         If Text3 = "" Then
            Text3 = FagentQuery(Text2, 1)
         End If
         If Text3 = "" Then
            Text3 = FagentQuery(Text2, 3)
         End If
      Case "X"
         Text3 = CustomerQuery(Text2, 2)
         If Text3 = "" Then
            Text3 = CustomerQuery(Text2, 1)
         End If
         If Text3 = "" Then
            Text3 = CustomerQuery(Text2, 3)
         End If
   End Select
   '2005/7/8 END
   If IsNull(adoacc1k0.Fields("a1k04").Value) Then
      Text4 = MsgText(601)
   Else
      Text4 = adoacc1k0.Fields("a1k04").Value
   End If
   If IsNull(adoacc1k0.Fields("a1k18").Value) Then
      Text5 = MsgText(601)
   Else
      Text5 = adoacc1k0.Fields("a1k18").Value
   End If
   If IsNull(adoacc1k0.Fields("a1k05").Value) Then
      Text11 = MsgText(601)
   Else
      Text11 = adoacc1k0.Fields("a1k05").Value
   End If
   '2013/10/22 add by sonia
   If IsNull(adoacc1k0.Fields("a1k34").Value) Then
      Text12 = MsgText(601)
   Else
      Text12 = adoacc1k0.Fields("a1k34").Value
   End If
   '2013/10/22 end
   Text9 = "" & adoacc1k0.Fields("a1k32").Value 'Add by Morgan 2010/6/11
   Text7 = adoacc1k0.Fields("a1k13").Value
   Text21 = adoacc1k0.Fields("a1k14").Value
   Text22 = adoacc1k0.Fields("a1k15").Value
   Text23 = adoacc1k0.Fields("a1k16").Value
   If IsNull(adoacc1k0.Fields("a1k28").Value) Then
      adoquery.CursorLocation = adUseClient
      adoquery.Open "select nvl(pa88, nvl(pa75, pa26)) as No from patent where pa01 = '" & Text7 & "' and pa02 = '" & Text21 & "' and pa03 = '" & Text22 & "' and pa04 = '" & Text23 & "' union " & _
                    "select nvl(tm56, nvl(tm44, tm23)) as No from trademark where tm01 = '" & Text7 & "' and tm02 = '" & Text21 & "' and tm03 = '" & Text22 & "' and tm04 = '" & Text23 & "' union " & _
                    "select nvl(lc26, nvl(lc22, lc11)) as No from lawcase where lc01 = '" & Text7 & "' and lc02 = '" & Text21 & "' and lc03 = '" & Text22 & "' and lc04 = '" & Text23 & "' union " & _
                    "select nvl(sp37, nvl(sp26, sp08)) as No from servicepractice where sp01 = '" & Text7 & "' and sp02 = '" & Text21 & "' and sp03 = '" & Text22 & "' and sp04 = '" & Text23 & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adoquery.RecordCount <> 0 Then
         If IsNull(adoquery.Fields(0).Value) Then
            Text8 = MsgText(601)
         Else
            Text8 = adoquery.Fields(0).Value
         End If
      Else
         Text8 = MsgText(601)
      End If
      adoquery.Close
   Else
      Text8 = adoacc1k0.Fields("a1k28").Value
   End If
   If IsNull(adoacc1k0.Fields("a1k27").Value) Then
      Text6 = Text8
   Else
      Text6 = adoacc1k0.Fields("a1k27").Value
   End If
   'Add By Sindy 2011/01/06 請款匯率
   If Not IsNull(adoacc1k0.Fields("a1k18").Value) Then
      If adoacc1k0.Fields("a1k18").Value = "USD" Then
         Text10 = adoacc1k0.Fields("a1k10").Value
      Else
         '抓請款匯率檔
         strSql = "SELECT DNR03 FROM DebitNoteRate WHERE DNR01='" & adoacc1k0.Fields("a1k18").Value & "' and DNR02<= " & adoacc1k0.Fields("a1k02").Value & " Order By DNR02 Desc "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            Text10 = "" & RsTemp.Fields("DNR03")
         End If
      End If
   End If
   '2011/01/06 End
   
   'Added by Morgan 2012/12/20
   If Not IsNull(adoacc1k0.Fields("a1k33")) Then
      'Added by Morgan 2012/12/20
      '順序不可變更(a1k33=listindex+1)
      Me.Combo4.Clear
      Me.Combo4.AddItem "純台幣", 0
      Me.Combo4.AddItem "台幣+外幣合計", 1
      Me.Combo4.AddItem "純外幣", 2
      Me.Combo4.AddItem "外幣+美金合計", 3
      Me.Combo4.ListIndex = 1 '預設 2:台幣+外幣合計
      'end 2012/12/20
      Combo4.ListIndex = Val(adoacc1k0.Fields("a1k33")) - 1
   End If
End Sub

'*************************************************
'  合計顯示
'
'*************************************************
Public Sub SumShow()
'Dim dblRate As Double
   
   Text14 = MsgText(601): Text13 = MsgText(601)
   'Add By Sindy 2010/3/29
   If "" & Trim(adoacc1k0.Fields("a1k18").Value) <> "USD" Then
      'dblRate = PUB_GetUSXRate_1("" & adoacc1k0.Fields("a1k02").Value, "" & Trim(adoacc1k0.Fields("a1k18").Value))
      'Modify By Sindy 2013/6/24
      'Text13 = Format(((Val(Val("" & adoacc1k0.Fields("a1k11").Value) - Val("" & adoacc1k0.Fields("a1k06").Value) * Val("" & adoacc1k0.Fields("a1k10").Value)) * 100 * 100) \ (dblRate * 100)) / 100, FAmount)
      '2013/7/31 modify by sonia
      'Text13 = Format(((Val(Val("" & adoacc1k0.Fields("a1k11").Value) - Val("" & adoacc1k0.Fields("a1k06").Value)) * 100 * 100) \ (dblRate * 100)) / 100, FAmount)
      Text13 = Format(Val("" & adoacc1k0.Fields("a1k08").Value) - Val("" & adoacc1k0.Fields("a1k31").Value), FAmount)
      'end 2013/7/31
      '2013/6/24 END
      Label11.Caption = "外幣金額 " & "" & Trim(adoacc1k0.Fields("a1k18").Value)
   '2010/3/29 End
   Else
      'Add by Morgan 2006/8/2
      'Modify By Sindy 2013/6/24
      'Text13 = Format(Val("" & adoacc1k0.Fields("a1k08").Value) - Val("" & adoacc1k0.Fields("a1k06").Value), FAmount)
      Text13 = Format(Val("" & adoacc1k0.Fields("a1k08").Value) - Val("" & adoacc1k0.Fields("a1k31").Value), FAmount)
      '2013/6/24 END
   End If
   
   'Add by Morgan 2006/8/18
   If adoadodc1.EOF Then
      MsgBox "舊系統資料無法顯示完整單據內容！"
      Exit Sub
   End If
   'end 2006/8/18
   
   adoadodc1.MoveFirst
   Do While Not adoadodc1.EOF
      Text14 = Val(Text14) + Val("" & adoadodc1("a1l05")) - Val("" & adoadodc1("a1l07"))
      adoadodc1.MoveNext
   Loop
   Exit Sub
   'end 2006/8/2
'2009/5/13 cancel by sonia 已無用
'Dim douAmount As Double
'Dim douDiscount As Double
'
'   adoaccsum.CursorLocation = adUseClient
'   adoaccsum.Open "select sum(a1l05), sum(a1l07), sum((a1l05 - a1l07) / a1k10) from acc1l0, acc1k0 where a1l01 = a1k01 and a1l01 = '" & Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
'   If adoaccsum.RecordCount <> 0 Then
'      If IsNull(adoaccsum.Fields(0).Value) Then
'         douAmount = 0
'      Else
'         douAmount = Val(adoaccsum.Fields(0).Value)
'      End If
'      If IsNull(adoaccsum.Fields(1).Value) Then
'         douDiscount = 0
'      Else
'         douDiscount = Val(adoaccsum.Fields(1).Value)
'      End If
'      If IsNull(adoaccsum.Fields(2).Value) = False Then
'         Text13 = Format(Val(adoaccsum.Fields(2).Value), FAmount)
'      Else
'         Text13 = 0
'      End If
'      Text14 = douAmount - douDiscount
'   Else
'      Text13 = MsgText(601)
'      Text14 = MsgText(601)
'   End If
'   adoaccsum.Close
End Sub

