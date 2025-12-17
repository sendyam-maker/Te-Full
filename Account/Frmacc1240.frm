VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc1240 
   AutoRedraw      =   -1  'True
   Caption         =   "本所案號帳目查詢"
   ClientHeight    =   5460
   ClientLeft      =   50
   ClientTop       =   340
   ClientWidth     =   9490
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5460
   ScaleWidth      =   9490
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
      Left            =   360
      TabIndex        =   4
      Top             =   888
      Width           =   3984
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
      Height          =   385
      Left            =   4920
      TabIndex        =   6
      Top             =   900
      Width           =   1785
   End
   Begin VB.TextBox Text12 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   7665
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   5064
      Width           =   1335
   End
   Begin VB.TextBox Text11 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   4980
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   5064
      Width           =   1335
   End
   Begin VB.TextBox Text10 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   2070
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   5064
      Width           =   1335
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   7665
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   4692
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   4980
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   4692
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2070
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4692
      Width           =   1335
   End
   Begin VB.CommandButton cmdCrtRct 
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
      Height          =   385
      Left            =   6960
      TabIndex        =   7
      Top             =   900
      Width           =   2050
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc1240.frx":0000
      Height          =   3132
      Left            =   60
      TabIndex        =   19
      Top             =   1476
      Width           =   9096
      _ExtentX        =   16069
      _ExtentY        =   5521
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
      Caption         =   "本所案號帳目查詢"
      ColumnCount     =   16
      BeginProperty Column00 
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
      BeginProperty Column01 
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
      BeginProperty Column02 
         DataField       =   "a0k20"
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
      BeginProperty Column03 
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
      BeginProperty Column04 
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
      BeginProperty Column05 
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
      BeginProperty Column06 
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
      BeginProperty Column07 
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
      BeginProperty Column08 
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
      BeginProperty Column09 
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
      BeginProperty Column10 
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
      BeginProperty Column11 
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
      BeginProperty Column12 
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
      BeginProperty Column13 
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
      BeginProperty Column14 
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
      BeginProperty Column15 
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
            Alignment       =   2
            ColumnWidth     =   530.079
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   950.173
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1069.795
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   319.748
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   360
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1120.252
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            ColumnWidth     =   1450.205
         EndProperty
         BeginProperty Column08 
            Alignment       =   1
         EndProperty
         BeginProperty Column09 
            Alignment       =   1
         EndProperty
         BeginProperty Column10 
            Alignment       =   1
         EndProperty
         BeginProperty Column11 
            Alignment       =   1
         EndProperty
         BeginProperty Column12 
            Alignment       =   1
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   1370.268
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   4330.205
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   1390.11
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   312
      Left            =   -168
      Top             =   4944
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
      Left            =   3160
      MaxLength       =   2
      TabIndex        =   3
      Top             =   160
      Width           =   372
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
      Height          =   315
      Left            =   2920
      MaxLength       =   1
      TabIndex        =   2
      Top             =   160
      Width           =   252
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
      Height          =   315
      Left            =   2080
      MaxLength       =   6
      TabIndex        =   1
      Top             =   160
      Width           =   852
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
      Height          =   252
      Left            =   7200
      TabIndex        =   5
      Top             =   160
      Width           =   1212
   End
   Begin VB.TextBox Text2 
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
      Left            =   5280
      TabIndex        =   16
      Top             =   160
      Width           =   1572
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
      Height          =   315
      Left            =   1500
      MaxLength       =   3
      TabIndex        =   0
      Top             =   160
      Width           =   592
   End
   Begin VB.Label lblPS 
      Caption         =   "列：N 未列印收據，Z 不列印收據，# 已開INVOICE"
      ForeColor       =   &H000000C0&
      Height          =   216
      Left            =   360
      TabIndex        =   27
      Top             =   1248
      Width           =   4260
   End
   Begin MSForms.TextBox Text3 
      Height          =   315
      Left            =   1500
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   525
      Width           =   6645
      VariousPropertyBits=   671105049
      BackColor       =   14737632
      Size            =   "11721;556"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "未收金額："
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
      Left            =   6468
      TabIndex        =   26
      Top             =   5064
      Width           =   1248
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "退費金額："
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
      Left            =   3660
      TabIndex        =   25
      Top             =   5064
      Width           =   1248
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "銷帳金額："
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
      Left            =   828
      TabIndex        =   24
      Top             =   5064
      Width           =   1248
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "扣繳額："
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
      Left            =   6468
      TabIndex        =   23
      Top             =   4692
      Width           =   1248
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "已收金額："
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
      Left            =   3660
      TabIndex        =   22
      Top             =   4692
      Width           =   1248
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "應收金額："
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
      Left            =   828
      TabIndex        =   21
      Top             =   4692
      Width           =   1248
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "合 計"
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
      Left            =   120
      TabIndex        =   20
      Top             =   4692
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   4944
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "案件名稱"
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
      Left            =   360
      TabIndex        =   17
      Top             =   520
      Width           =   1452
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "申請國家"
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
      Left            =   4200
      TabIndex        =   15
      Top             =   160
      Width           =   972
   End
   Begin VB.Label Label1 
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
      Left            =   360
      TabIndex        =   14
      Top             =   160
      Width           =   972
   End
End
Attribute VB_Name = "Frmacc1240"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/16 Form2.0已修改
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/11/30 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/28 日期欄已修改
Option Explicit

Public adocase As New ADODB.Recordset
Public adonation As New ADODB.Recordset
Public adoacctmp05 As New ADODB.Recordset
Public adoacc0m0 As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Public frmCall As Form 'Added by Morgan 2015/12/7
Public m_strUserOffice As String 'Added by Lydia 2020/03/30 收據作業:另外傳入所別
Dim strSql As String
Dim strType As String

Private Sub Command1_Click()
'Added by Lydia 2016/01/20 點選呼叫"收據抬頭修改"
   If Adodc1.Recordset.State = 1 Then
      If Adodc1.Recordset.RecordCount = 0 Then
         Exit Sub
      End If
      strItemNo = Adodc1.Recordset.Fields("RNo").Value
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
   Else
      MsgBox "請先按 F12 查詢並點選單據資料..."
   End If
'end 2016/01/20
End Sub

Private Sub Command2_Click()
   
   'Add by Morgan 2005/3/25
   If Adodc1.Recordset.State = adStateClosed Then Exit Sub

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
         Frmacc1222.Show
         Me.Enabled = False
      Case "I"
         Frmacc1224.Show
         Me.Enabled = False
   End Select
End Sub

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
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   'Modify by Amy 2023/10/06 W9500 H5700
   Me.Width = 9615
   '20140115START Modify By eric
   'Me.Height = 5400
   Me.Height = 5925
   '20140115END
   'end 2023/10/06
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath2)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   OpenTable
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
   'Added by Lydia 2020/03/30 收據作業:另外傳入所別
   If m_strUserOffice = "" Then m_strUserOffice = pub_strUserOffice
   
   'Added by Lydia 2016/01/20 分所不可使用"收據抬頭修改"
   'Modified by Lydia 2020/03/30
   'If pub_strUserOffice = "1" Then
   If m_strUserOffice = "1" Then
      Command1.Visible = True
   Else
      Command1.Visible = False
   End If
   'end 2016/01/20
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
'   Acctmp05Delete
'   adoacctmp05.Close
   
   StatusClear
   strFormName = MsgText(601)
   
   'Added by Morgan 2015/12/7
   If Not frmCall Is Nothing Then
      strFormLink = Me.Name
      If Adodc1.Recordset.State = adStateOpen Then
         If Adodc1.Recordset.RecordCount > 0 Then
            If Left("" & Adodc1.Recordset.Fields("RNo").Value, 1) = "E" Then
               frmCall.Tag = Adodc1.Recordset.Fields("RNo").Value
            End If
         End If
      End If
      Forms(0).Toolbar1.Enabled = True
      frmCall.Enabled = True
      frmCall.SetFocus
      strSaveConfirm = MsgText(3)
      tool2_enabled
   Else
   'end 2015/12/7
   
      KeyEnter vbKeyEscape
      MenuEnabled
      
   End If 'Added by Morgan 2015/12/7
   Set Frmacc1240 = Nothing
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
   CloseIme
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'*************************************************
'  查詢資料表(專利基本檔及商標基本檔)
'
'*************************************************
Private Sub QueryTable()
Dim strSql As String
   
On Error GoTo Checking
   Select Case Text1
      Case "P", "CFP", "FCP"
         If Text1 <> MsgText(601) Then
            strSql = " and pa01 = '" & Text1 & "'"
         End If
         If Text6 <> MsgText(601) Then
            strSql = strSql & " and pa02 = '" & Text6 & "'"
         End If
         If Text7 <> MsgText(601) Then
            strSql = strSql & " and pa03 = '" & Text7 & "'"
         End If
         If Text8 <> MsgText(601) Then
            strSql = strSql & " and pa04 = '" & Text8 & "'"
         End If
         If strSql <> MsgText(601) Then
            strSql = " where" & Mid(strSql, 5, Len(strSql) - 4)
            adocase.CursorLocation = adUseClient
            adocase.Open "select * from patent" & strSql, adoTaie, adOpenStatic, adLockReadOnly
      If adocase.RecordCount <> 0 Then
               PatentShow
            End If
            adocase.Close
         End If
      Case "T", "CFT", "FCT"
         If Text1 <> MsgText(601) Then
            strSql = " and tm01 = '" & Text1 & "'"
         End If
         If Text6 <> MsgText(601) Then
            strSql = strSql & " and tm02 = '" & Text6 & "'"
         End If
         If Text7 <> MsgText(601) Then
            strSql = strSql & " and tm03 = '" & Text7 & "'"
         End If
         If Text8 <> MsgText(601) Then
            strSql = strSql & " and tm04 = '" & Text8 & "'"
         End If
         If strSql <> MsgText(601) Then
            strSql = " where" & Mid(strSql, 5, Len(strSql) - 4)
            adocase.CursorLocation = adUseClient
            adocase.Open "select * from trademark" & strSql, adoTaie, adOpenStatic, adLockReadOnly
            If adocase.RecordCount <> 0 Then
               TrademarkShow
            End If
            adocase.Close
         End If
      Case "L", "CFL", "FCL"
         If Text1 <> MsgText(601) Then
            strSql = " and lc01 = '" & Text1 & "'"
         End If
         If Text6 <> MsgText(601) Then
            strSql = strSql & " and lc02 = '" & Text6 & "'"
         End If
         If Text7 <> MsgText(601) Then
            strSql = strSql & " and lc03 = '" & Text7 & "'"
         End If
         If Text8 <> MsgText(601) Then
            strSql = strSql & " and lc04 = '" & Text8 & "'"
         End If
         If strSql <> MsgText(601) Then
            strSql = " where" & Mid(strSql, 5, Len(strSql) - 4)
            adocase.CursorLocation = adUseClient
            adocase.Open "select * from lawcase" & strSql, adoTaie, adOpenStatic, adLockReadOnly
            If adocase.RecordCount <> 0 Then
               LawCaseShow
            End If
            adocase.Close
         End If
      Case "LA"
         If Text1 <> MsgText(601) Then
            strSql = " and hc01 = '" & Text1 & "'"
         End If
         If Text6 <> MsgText(601) Then
            strSql = strSql & " and hc02 = '" & Text6 & "'"
         End If
         If Text7 <> MsgText(601) Then
            strSql = strSql & " and hc03 = '" & Text7 & "'"
         End If
         If Text8 <> MsgText(601) Then
            strSql = strSql & " and hc04 = '" & Text8 & "'"
         End If
         If strSql <> MsgText(601) Then
            strSql = " where" & Mid(strSql, 5, Len(strSql) - 4)
            adocase.CursorLocation = adUseClient
            adocase.Open "select * from hirecase" & strSql, adoTaie, adOpenStatic, adLockReadOnly
            If adocase.RecordCount <> 0 Then
               HireCaseShow
            End If
            adocase.Close
         End If
      Case Else
         If Text1 <> MsgText(601) Then
            strSql = " and sp01 = '" & Text1 & "'"
         End If
         If Text6 <> MsgText(601) Then
            strSql = strSql & " and sp02 = '" & Text6 & "'"
         End If
         If Text7 <> MsgText(601) Then
            strSql = strSql & " and sp03 = '" & Text7 & "'"
         End If
         If Text8 <> MsgText(601) Then
            strSql = strSql & " and sp04 = '" & Text8 & "'"
         End If
         If strSql <> MsgText(601) Then
            strSql = " where" & Mid(strSql, 5, Len(strSql) - 4)
            adocase.CursorLocation = adUseClient
            adocase.Open "select * from servicepractice" & strSql, adoTaie, adOpenStatic, adLockReadOnly
            If adocase.RecordCount <> 0 Then
               ServiceShow
            End If
            adocase.Close
         End If
   End Select
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  顯示資料表(專利基本檔)
'
'*************************************************
Private Sub PatentShow()
   adonation.CursorLocation = adUseClient
   adonation.Open "select nvl(na03, na04) from nation where na01 = '" & adocase.Fields("pa09").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adonation.RecordCount <> 0 Then
      If IsNull(adonation.Fields(0).Value) Then
         Text2 = MsgText(601)
      Else
         Text2 = adonation.Fields(0).Value
      End If
   Else
      Text2 = MsgText(601)
   End If
   If IsNull(adocase.Fields("pa05").Value) Then
      Text3 = MsgText(601)
   Else
      Text3 = adocase.Fields("pa05").Value
   End If
'20140115START REMARK By eric
'   If IsNull(adocase.Fields("pa06").Value) Then
'      Text4 = MsgText(601)
'   Else
'      Text4 = adocase.Fields("pa06").Value
'   End If
'   If IsNull(adocase.Fields("pa07").Value) Then
'      Text5 = MsgText(601)
'   Else
'      Text5 = adocase.Fields("pa07").Value
'   End If
'20140115END   REMARK By eric
   adonation.Close
End Sub

'*************************************************
'  顯示資料表(商標基本檔)
'
'*************************************************
Private Sub TrademarkShow()
   adonation.CursorLocation = adUseClient
   adonation.Open "select nvl(na03, na04) from nation where na01 = '" & adocase.Fields("tm10").Value & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
   If adonation.RecordCount <> 0 Then
      If IsNull(adonation.Fields(0).Value) Then
         Text2 = MsgText(601)
      Else
         Text2 = adonation.Fields(0).Value
      End If
   Else
      Text2 = MsgText(601)
   End If
   If IsNull(adocase.Fields("tm05").Value) Then
      Text3 = MsgText(601)
   Else
      Text3 = adocase.Fields("tm05").Value
   End If
'20140115START REMARK By eric
'   If IsNull(adocase.Fields("tm06").Value) Then
'      Text4 = MsgText(601)
'   Else
'      Text4 = adocase.Fields("tm06").Value
'   End If
'   If IsNull(adocase.Fields("tm07").Value) Then
'      Text5 = MsgText(601)
'   Else
'      Text5 = adocase.Fields("tm07").Value
'   End If
'20140115END   REMARK By eric

   adonation.Close
End Sub

'*************************************************
'  顯示資料表(法務基本檔)
'
'*************************************************
Private Sub LawCaseShow()
   adonation.CursorLocation = adUseClient
   adonation.Open "select nvl(na03, na04) from nation where na01 = '" & adocase.Fields("lc15").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adonation.RecordCount <> 0 Then
      If IsNull(adonation.Fields(0).Value) Then
         Text2 = MsgText(601)
      Else
         Text2 = adonation.Fields(0).Value
      End If
   Else
      Text2 = MsgText(601)
   End If
   If IsNull(adocase.Fields("lc05").Value) Then
      Text3 = MsgText(601)
   Else
      Text3 = adocase.Fields("lc05").Value
   End If
'20140115START REMARK By eric
'   If IsNull(adocase.Fields("lc06").Value) Then
'      Text4 = MsgText(601)
'   Else
'      Text4 = adocase.Fields("lc06").Value
'   End If
'   If IsNull(adocase.Fields("lc07").Value) Then
'      Text5 = MsgText(601)
'   Else
'      Text5 = adocase.Fields("lc07").Value
'   End If
'20140115END   REMARK By eric

   adonation.Close
End Sub

'*************************************************
'  顯示資料表(顧問基本檔)
'
'*************************************************
Private Sub HireCaseShow()
   Text2 = MsgText(601)
   If IsNull(adocase.Fields("hc05").Value) Then
      Text3 = MsgText(601)
   Else
      Text3 = adocase.Fields("hc05").Value
   End If
'20140115START REMARK By eric
'   If IsNull(adocase.Fields("hc06").Value) Then
'      Text4 = MsgText(601)
'   Else
'      Text4 = adocase.Fields("hc06").Value
'   End If
'   If IsNull(adocase.Fields("hc07").Value) Then
'      Text5 = MsgText(601)
'   Else
'      Text5 = adocase.Fields("hc07").Value
'   End If
'20140115END   REMARK By eric

End Sub

'*************************************************
'  顯示資料表(服務業務基本檔)
'
'*************************************************
Private Sub ServiceShow()
   adonation.CursorLocation = adUseClient
   adonation.Open "select nvl(na03, na04) from nation where na01 = '" & adocase.Fields("sp09").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adonation.RecordCount <> 0 Then
      If IsNull(adonation.Fields(0).Value) Then
         Text2 = MsgText(601)
      Else
         Text2 = adonation.Fields(0).Value
      End If
   Else
      Text2 = MsgText(601)
   End If
   If IsNull(adocase.Fields("sp05").Value) Then
      Text3 = MsgText(601)
   Else
      Text3 = adocase.Fields("sp05").Value
   End If
'20140115START REMARK By eric
'   If IsNull(adocase.Fields("sp06").Value) Then
'      Text4 = MsgText(601)
'   Else
'      Text4 = adocase.Fields("sp06").Value
'   End If
'   If IsNull(adocase.Fields("sp07").Value) Then
'      Text5 = MsgText(601)
'   Else
'      Text5 = adocase.Fields("sp07").Value
'   End If
'20140115END   REMARK By eric

   adonation.Close
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

Private Sub Text6_GotFocus()
   TextInverse Text6
End Sub

Private Sub Text6_Validate(Cancel As Boolean)
   If Text6 = "" Then
      Exit Sub
   End If
   If Len(Text6) < 6 Then
      MsgBox MsgText(172), , MsgText(5)
      Cancel = True
      Text6.SetFocus
      Exit Sub
   End If
   Text7 = "0"
   Text8 = "00"
'   Screen.MousePointer = vbHourglass
'   QueryTable
'   AdodcRefresh
'   Screen.MousePointer = vbDefault
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
   '92.7.17 cancel by sonia
   'If strCon9 = MsgText(602) Then
   '   Cancel = True
   '   Text6.SetFocus
   '   TextInverse Text6
   'End If
End Sub

Private Sub Text7_GotFocus()
   TextInverse Text7
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text7_Validate(Cancel As Boolean)
'   If Text6 = "" Then
'      Exit Sub
'   End If
'   QueryTable
'   AdodcRefresh
End Sub

'*************************************************
'  清除顯示
'
'*************************************************
Private Sub FormClear()
   Text2 = ""
   Text3 = ""
'20140115START REMARK By eric
'   Text4 = ""
'   Text5 = ""
'20140115END   REMARK By eric
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoadodc1.CursorLocation = adUseClient
   'Modify by Morgan 2011/8/23 改從 0j0 抓 cp
   'Modified by Morgan 2011/12/27 取消 a0j20
   '20140115 Modify By eric 增加 a0k32
   'Modified by Lydia 2025/05/23 增加# 已開INVOICE => decode(a0k40,null,a0k32,'#')
   'Modified by Lydia 2025/06/23 +INVOICE編號a0k40
   adoadodc1.Open "select a0k05, a0k02, a0k01, a0k20, a0k04, a0j02, cp09, a0j21,decode(a0k40,null,a0k32,'#') a0k32, getcp10desc(cp01,cp10,a0j04) cp10N, (a0j09 + a0j10) as RAmount, cp75 as EAmount, cp76, cp77, cp78 as BAmount, cp79 as NAmount, a0k40 from acc0k0, acc0j0, caseprogress where a0k03 = '1' and a0j13(+) = a0k01 and cp09(+) = a0j01", adoTaie, adOpenStatic, adLockReadOnly
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

Private Sub Text8_GotFocus()
   TextInverse Text8
End Sub

'*************************************************
'  刪除資料表之記錄(智權人員帳款查詢暫存檔)
'
'*************************************************
Private Sub Acctmp05Delete()
   adoTaie.Execute "delete from acctmp05"
End Sub

'*************************************************
'  重新整理 Adodc 之資料
'
'*************************************************
Public Sub AdodcRefresh()
Dim strUnion As String
Dim StrSQLa As String

On Error GoTo Checking
   
    strSql = ""
   If Text1 <> "" And Text6 <> "" And Text7 <> "" And Text8 <> "" Then
      strSql = " and a0j02 = '" & Text1 & Text6 & Text7 & Text8 & "'"
   End If
   
'Modify by Morgan 2005/3/28 改抓基本檔申請人所別控制
'      'Add By Cheng 2004/01/14
'      '若非北所員工, 只能列印該所資料
'      strSQLA = ""
'      If pub_strUserOffice <> "1" Then
'          strSQLA = strSQLA & " And CU13=ST01(+) And ''||ST06='" & pub_strUserOffice & "' "
'      Else
'          strSQLA = strSQLA & " And A0K20=ST01(+) "
'      End If
'      'End
   StrSQLa = StrSQLa & " And A0K20=ST01(+) "
'2005/3/28 end

   If adoadodc1.State = adStateOpen Then
      adoadodc1.Close
   End If
   adoadodc1.CursorLocation = adUseClient
   'strUnion = "select a0k03, a0k04, a0k11, a0k02 as RDate, a0k01 as RNo, a0k20, a0j02, a0j01, a0j21, a0j20, nvl(cp16, 0) as RAmount, 0 as EAmount, decode(a0k30, 'Y', nvl(cp16, 0)*0.1, (nvl(cp16, 0) - nvl(cp17, 0))*0.1) as DAmount, 0 as CAmount, 0 as BAmount, nvl(cp79, 0) as NAmount, a0k01 from caseprogress, acc0k0, acc0j0 where cp60 = a0k01 and cp09 = a0j01 and (a0k09 is null or a0k09 = 0) and a0k01 is not null and cp01 = '" & Text1 & "' and cp02 = '" & Text6 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text8 & "'"
    'Modify By Cheng 2003/05/15
'   strUnion = "select a0k03, a0k04, a0k11, a0k02 as RDate, a0k01 as RNo, a0k20, a0j02, a0j01, a0j21, a0j20, nvl(cp16, 0) as RAmount, 0 as EAmount, 0 as DAmount, 0 as CAmount, 0 as BAmount, nvl(cp79, 0) as NAmount, a0k01 from caseprogress, acc0k0, acc0j0 where cp60 = a0k01 and cp09 = a0j01 and (a0k09 is null or a0k09 = 0) and a0k01 is not null and cp01 = '" & Text1 & "' and cp02 = '" & Text6 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text8 & "'"
'   strUnion = strUnion & " union select a0k03, a0k04, a0k11, a0l02 as RDate, a0l01 as RNo, a0k20, a0j02, a0j01, a0j21, a0j20, 0 as RAmount, (a1u04 + a1u05) as EAmount, a1u06 as DAmount, 0 as CAmount, 0 as BAmount, 0 as NAmount, a0k01 from acc1u0, acc0l0, acc0k0, acc0j0, caseprogress where a1u01 = a0l01 and a1u02 = a0k01 and a1u02 = a0j13 and a1u03 = a0j01 and a1u03 = cp09 and (a0k09 is null or a0k09 = 0) and a0l01 is not null and cp01 = '" & Text1 & "' and cp02 = '" & Text6 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text8 & "'"
'   strUnion = strUnion & " union select a0k03, a0k04, a0k11, a0s03 as RDate, a0s01 as RNo, a0k20, a0j02, a0j01, a0j21, a0j20, 0 as Ramount, 0 as EAmount, 0 as Damount, a0s05 as CAmount, a0s06 + a0s07 as BAmount, 0 as NAmount, a0k01 from acc0j0, acc0s0, acc0k0 where a0j13 = a0s02 and a0j13 = a0k01 and (a0k09 is null or a0k09 = 0) and a0s01 is not null" & strSQL
    'Modify By Cheng 2004/01/14
'   strUnion = "select a0k03, a0k04, a0k11, a0k02 as RDate, a0k01 as RNo, a0k20||' '||ST02 As a0k20, a0j02, a0j01, a0j21, a0j20, nvl(cp16, 0) as RAmount, 0 as EAmount, 0 as DAmount, 0 as CAmount, 0 as BAmount, nvl(cp79, 0) as NAmount, a0k01 from caseprogress, acc0k0, acc0j0, Staff where cp60 = a0k01 and cp09 = a0j01 and (a0k09 is null or a0k09 = 0) and a0k01 is not null And a0k20=ST01(+) and cp01 = '" & Text1 & "' and cp02 = '" & Text6 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text8 & "'"
'   strUnion = strUnion & " union select a0k03, a0k04, a0k11, a0l02 as RDate, a0l01 as RNo, a0k20||' '||ST02 As a0k20, a0j02, a0j01, a0j21, a0j20, 0 as RAmount, (nvl(a1u04, 0) + nvl(a1u05, 0)) as EAmount, a1u06 as DAmount, 0 as CAmount, 0 as BAmount, 0 as NAmount, a0k01 from acc1u0, acc0l0, acc0k0, acc0j0, caseprogress, Staff where a1u01 = a0l01 and a1u02 = a0k01 and a1u02 = a0j13 and a1u03 = a0j01 and a1u03 = cp09 and (a0k09 is null or a0k09 = 0) and a0l01 is not null And a0k20=ST01(+) and cp01 = '" & Text1 & "' and cp02 = '" & Text6 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text8 & "'"
'   strUnion = strUnion & " union select a0k03, a0k04, a0k11, a0s03 as RDate, a0s01 as RNo, a0k20||' '||ST02 As a0k20, a0j02, a0j01, a0j21, a0j20, 0 as Ramount, 0 as EAmount, 0 as Damount, (nvl(a1u07, 0)+nvl(a1u09, 0)) as CAmount, (nvl(a1u08, 0)+nvl(a1u10, 0)) as BAmount, 0 as NAmount, a0k01 from acc0j0, acc1u0, acc0s0, acc0k0, Staff where a0j01 = a1u03 and a1u01 = a0s01 and a0j13 = a0k01 and (a0k09 is null or a0k09 = 0) and a0s01 is not null And a0k20=ST01(+) " & strSQL
'   strUnion = strUnion & " union select a0k03, a0k04, a0k11, a0k02 as RDate, decode(a1v18, '1', '收款扣繳', '補扣繳') as RNo, a0k20||' '||ST02 As a0k20, a0j02, a0j01, a0j21, a0j20, 0 as Ramount, 0 as EAmount, a1v06 as Damount, 0 as CAmount, 0 as BAmount, 0 as NAmount, a0k01 from acc0j0, acc1v0, acc0k0, Staff where a0j01 = a1v01 and a0j13 = a0k01 and (a0k09 is null or a0k09 = 0) and a1v06 <> 0 And a0k20=ST01(+) and a1v18 is null" & strSQL
   
   'Modify by Morgan 2011/8/23 改從 0j0 抓 cp
   'Modify by Morgan 2011/10/27 考慮拆收據情形改先寫暫存
   'strUnion = "select a0k03, a0k04, a0k11, a0k02 as RDate, a0k01 as RNo, a0k20||' '||ST02 As a0k20, a0j02, a0j01, a0j21, a0j20, nvl(cp16, 0) as RAmount, 0 as EAmount, 0 as DAmount, 0 as CAmount, 0 as BAmount, nvl(cp79, 0) as NAmount, a0k01 from caseprogress, acc0j0, acc0k0, Staff, Customer where a0j13 = a0k01(+) and cp09 = a0j01(+) and (a0k09 is null or a0k09 = 0) and a0k01 is not null And substr(A0K03,1,8)=CU01(+) And substr(A0K03,9,1)=CU02(+) and cp01 = '" & Text1 & "' and cp02 = '" & Text6 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text8 & "' " & StrSQLa
   'strUnion = strUnion & " union select a0k03, a0k04, a0k11, a0l02 as RDate, a0l01 as RNo, a0k20||' '||ST02 As a0k20, a0j02, a0j01, a0j21, a0j20, 0 as RAmount, (nvl(a1u04, 0) + nvl(a1u05, 0)) as EAmount, a1u06 as DAmount, 0 as CAmount, 0 as BAmount, 0 as NAmount, a0k01 from acc1u0, acc0l0, acc0k0, acc0j0, caseprogress, Staff, Customer where a1u01 = a0l01 and a1u02 = a0k01 and a1u02 = a0j13 and a1u03 = a0j01 and a1u03 = cp09 and (a0k09 is null or a0k09 = 0) and a0l01 is not null And substr(A0K03,1,8)=CU01(+) And substr(A0K03,9,1)=CU02(+) and cp01 = '" & Text1 & "' and cp02 = '" & Text6 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text8 & "' " & StrSQLa
   'strUnion = strUnion & " union select a0k03, a0k04, a0k11, a0s03 as RDate, a0s01 as RNo, a0k20||' '||ST02 As a0k20, a0j02, a0j01, a0j21, a0j20, 0 as Ramount, 0 as EAmount, 0 as Damount, (nvl(a1u07, 0)+nvl(a1u09, 0)) as CAmount, (nvl(a1u08, 0)+nvl(a1u10, 0)) as BAmount, 0 as NAmount, a0k01 from acc0j0, acc1u0, acc0s0, acc0k0, Staff, Customer where a0j01 = a1u03 and a1u01 = a0s01 and a0j13 = a0k01 and (a0k09 is null or a0k09 = 0) and a0s01 is not null And substr(A0K03,1,8)=CU01(+) And substr(A0K03,9,1)=CU02(+) " & strSql & StrSQLa
   'strUnion = strUnion & " union select a0k03, a0k04, a0k11, a0k02 as RDate, decode(a1v18, '1', '收款扣繳', '補扣繳') as RNo, a0k20||' '||ST02 As a0k20, a0j02, a0j01, a0j21, a0j20, 0 as Ramount, 0 as EAmount, a1v06 as Damount, 0 as CAmount, 0 as BAmount, 0 as NAmount, a0k01 from acc0j0, acc1v0, acc0k0, Staff, Customer where a0j01 = a1v01 and a0j13 = a0k01 and (a0k09 is null or a0k09 = 0) and a1v06 <> 0 And substr(A0K03,1,8)=CU01(+) And substr(A0K03,9,1)=CU02(+) and a1v18 is null" & strSql & StrSQLa
   
   adoTaie.Execute "delete ACCTMP08 where T05='" & Me.Name & "' and T14='" & strUserNum & "'"
   'Modified by Lydia 2016/04/11 舊收據無a0j01
   strUnion = "select a0k01,NVL(a0j01,' ') a0j01,'" & Me.Name & "',a0k01,'" & strUserNum & "' T14 from acc0j0, acc0k0" & _
      " where a0k01(+)=a0j13 and (a0k09 is null or a0k09 = 0) and a0k01 is not null" & strSql
      
   adoTaie.Execute "insert into ACCTMP08(T01,T02,T05,T06,T14) " & strUnion, intI
   
   'Added by Lydia 2025/07/25 排除未達客戶付款週期之應收帳款
   If ChkBillDate.Value = 1 Then
      Call PUB_ProcAcctmp08(Me.Name, strUserNum)
   End If
   'end 2025/07/25
   
   '更新金額欄位
   strSql = "update ACCTMP08 set T08=(select nvl(a0j09, 0)+nvl(a0j10, 0) from acc0j0 where a0j13=T01 and a0j01=T02)" & _
      ",(T09,T10,T11,T12)=(select nvl(sum(a1u04),0)+nvl(sum(a1u05),0) T09,nvl(sum(a1u06),0) T10" & _
      ",nvl(sum(a1u07),0)+nvl(sum(a1u09),0) T11,nvl(sum(a1u08),0)+nvl(sum(a1u10),0) T12 " & _
      " from acc1u0 where a1u02=T01 and a1u03=T02) where T05='" & Me.Name & "' and T14='" & strUserNum & "'"
   adoTaie.Execute strSql, intI
   
   '", T08 RAmount, 0 EAmount, 0 DAmount, T11 CAmount" & _
      ", T12 BAmount, T08-T09-T11+T12 NAmount, a0k01"

   
   '20140115START REMARK By eric
   'Modified by Morgan 2011/12/27 取消 a0j20,a0j21
 
   ' strUnion = "select a0k03, a0k04, a0k11, a0k02 as RDate, a0k01 as RNo, a0k20||' '||ST02 As a0k20" & _
   '   ", a0j02, a0j01, na03, getcp10desc(cp01,cp10,a0j04) cp10N, T08 as RAmount, 0 as EAmount, 0 as DAmount, 0 as CAmount" & _
   '   ", 0 as BAmount, T08-T09-T11+T12 as NAmount, a0k01" & _
   '   " from ACCTMP08, acc0j0, acc0k0, Staff,caseprogress,nation" & _
   '   " where T05='" & Me.Name & "' and T14='" & strUserNum & "' and a0j13(+)=T01 and a0j01(+)=T02 and a0k01(+)=T01 and st01(+)=a0k20 and cp09(+)=a0j01 and na01(+)=a0j04"

   ' strUnion = strUnion & " union select a0k03, a0k04, a0k11, a0l02 as RDate, a0l01 as RNo, a0k20||' '||ST02 As a0k20" & _
   '   ", a0j02, a0j01, na03, getcp10desc(cp01,cp10,a0j04) cp10N, 0 as RAmount, (nvl(a1u04, 0) + nvl(a1u05, 0)) as EAmount, nvl(a1u06,0) as DAmount" & _
   '   ", 0 as CAmount, 0 as BAmount, 0 as NAmount, a0k01" & _
   '   " from ACCTMP08, acc0j0, acc0k0,acc1u0, acc0l0,Staff,caseprogress,nation" & _
   '   " where T05='" & Me.Name & "' and T14='" & strUserNum & "' and a0j13(+)=T01 and a0j01(+)=T02 and a0k01(+)=T01 and st01(+)=a0k20" & _
   '   " and a1u02(+)=a0j13 and a1u03(+)=a0j01 and substr(a1u01,1,1)='F' and a0l01(+)=a1u01 and cp09(+)=a0j01 and na01(+)=a0j04"

   ' strUnion = strUnion & " union select a0k03, a0k04, a0k11, a0s03 as RDate, a0s01 as RNo, a0k20||' '||ST02 As a0k20" & _
   '   ", a0j02, a0j01, na03, getcp10desc(cp01,cp10,a0j04) cp10N, 0 as Ramount, 0 as EAmount, 0 as Damount, (nvl(a1u07, 0)+nvl(a1u09, 0)) as CAmount" & _
   '   ", (nvl(a1u08, 0)+nvl(a1u10, 0)) as BAmount, 0 as NAmount, a0k01" & _
   '   " from ACCTMP08, acc0j0, acc0k0,acc1u0, acc0s0,Staff,caseprogress,nation" & _
   '   " where T05='" & Me.Name & "' and T14='" & strUserNum & "' and a0j13(+)=T01 and a0j01(+)=T02 and a0k01(+)=T01 and st01(+)=a0k20" & _
   '   " and a1u02(+)=a0j13 and a1u03(+)=a0j01 and substr(a1u01,1,1)='I' and a0s01(+)=a1u01 and cp09(+)=a0j01 and na01(+)=a0j04"

   ' strUnion = strUnion & " union select a0k03, a0k04, a0k11, a0k02 as RDate, decode(a1v18, '1', '收款扣繳', '補扣繳') as RNo" & _
   '   ", a0k20||' '||ST02 As a0k20, a0j02, a0j01, na03, getcp10desc(cp01,cp10,a0j04) cp10N, 0 as Ramount, 0 as EAmount, a1v06 as Damount" & _
   '   ", 0 as CAmount, 0 as BAmount, 0 as NAmount, a0k01" & _
   '   " from ACCTMP08, acc0j0, acc0k0,acc1v0,Staff,caseprogress,nation" & _
   '   " where T05='" & Me.Name & "' and T14='" & strUserNum & "' and a0j13(+)=T01 and a0j01(+)=T02 and a0k01(+)=T01 and st01(+)=a0k20" & _
   '   " and a1v01(+)=a0j01 and a1v02(+)=a0j13 and a1v06>0 and a1v18 is null and cp09(+)=a0j01 and na01(+)=a0j04"
    'End 2011/10/27
   '20140115END REMARK By eric
   '20140115START Modify by Eric 增加 a0k32 ,axc01
   'Modified by Lydia 2025/05/23 增加# 已開INVOICE => decode(a0k40,null,a0k32,'#')
   'Modified by Lydia 2025/06/23 +INVOICE編號a0k40
   'Modify By Sindy 2025/9/10 ST02 As a0k20 => ST02||getlos04(a0j01,1) as a0k20
   strUnion = "select a0k03, a0k04, a0k11, a0k02 as RDate, a0k01 as RNo, decode(a0k40,null,a0k32,'#') a0k32, axc01 , ST02||getlos04(a0j01,1) As a0k20" & _
      ", a0j02, a0j01, na03, getcp10desc(cp01,cp10,a0j04) cp10N, T08 as RAmount, 0 as EAmount, 0 as DAmount, 0 as CAmount" & _
      ", 0 as BAmount, T08-T09-T11+T12 as NAmount, a0k01, a0k40" & _
      " from ACCTMP08, acc0j0, acc0k0, Staff,caseprogress,nation,acc431" & _
      " where T05='" & Me.Name & "' and T14='" & strUserNum & "' and a0j13(+)=T01 and a0j01(+)=T02 and a0k01(+)=T01 and st01(+)=a0k20 and cp09(+)=a0j01 and na01(+)=a0j04" & _
      " and axc02(+)=a0k01 "
  
   'Modified by Lydia 2025/05/23 增加# 已開INVOICE => decode(a0k40,null,a0k32,'#')
   'Modified by Lydia 2025/06/23 +INVOICE編號a0k40
   'Modify By Sindy 2025/9/10 ST02 As a0k20 => ST02||getlos04(a0j01,1) as a0k20
   strUnion = strUnion & " union select a0k03, a0k04, a0k11, a0l02 as RDate, a0l01 as RNo, decode(a0k40,null,a0k32,'#') a0k32, axc01, ST02||getlos04(a0j01,1) As a0k20" & _
      ", a0j02, a0j01, na03, getcp10desc(cp01,cp10,a0j04) cp10N, 0 as RAmount, (nvl(a1u04, 0) + nvl(a1u05, 0)) as EAmount, nvl(a1u06,0) as DAmount" & _
      ", 0 as CAmount, 0 as BAmount, 0 as NAmount, a0k01, a0k40" & _
      " from ACCTMP08, acc0j0, acc0k0,acc1u0, acc0l0,Staff,caseprogress,nation,acc431" & _
      " where T05='" & Me.Name & "' and T14='" & strUserNum & "' and a0j13(+)=T01 and a0j01(+)=T02 and a0k01(+)=T01 and st01(+)=a0k20" & _
      " and a1u02(+)=a0j13 and a1u03(+)=a0j01 and substr(a1u01,1,1)='F' and a0l01(+)=a1u01 and cp09(+)=a0j01 and na01(+)=a0j04" & _
      " and axc02(+)=a0k01 "
  
   'Modified by Lydia 2025/05/23 增加# 已開INVOICE => decode(a0k40,null,a0k32,'#')
   'Modified by Lydia 2025/06/23 +INVOICE編號a0k40
   'Modify By Sindy 2025/9/10 ST02 As a0k20 => ST02||getlos04(a0j01,1) as a0k20
   strUnion = strUnion & " union select a0k03, a0k04, a0k11, a0s03 as RDate, a0s01 as RNo, decode(a0k40,null,a0k32,'#') a0k32, axc01, ST02||getlos04(a0j01,1) As a0k20" & _
      ", a0j02, a0j01, na03, getcp10desc(cp01,cp10,a0j04) cp10N, 0 as Ramount, 0 as EAmount, 0 as Damount, (nvl(a1u07, 0)+nvl(a1u09, 0)) as CAmount" & _
      ", (nvl(a1u08, 0)+nvl(a1u10, 0)) as BAmount, 0 as NAmount, a0k01, a0k40" & _
      " from ACCTMP08, acc0j0, acc0k0,acc1u0, acc0s0,Staff,caseprogress,nation,acc431" & _
      " where T05='" & Me.Name & "' and T14='" & strUserNum & "' and a0j13(+)=T01 and a0j01(+)=T02 and a0k01(+)=T01 and st01(+)=a0k20" & _
      " and a1u02(+)=a0j13 and a1u03(+)=a0j01 and substr(a1u01,1,1)='I' and a0s01(+)=a1u01 and cp09(+)=a0j01 and na01(+)=a0j04" & _
      " and axc02(+)=a0k01 "
  
   '2014/12/8 modify by sonia 補扣繳不可顯示收據日期 CFT013014(E10328594)
   'strUnion = strUnion & " union select a0k03, a0k04, a0k11, a0k02 as RDate, decode(a1v18, '1', '收款扣繳', '補扣繳') as RNo" & _
      " ,  a0k32, axc01, ST02 As a0k20, a0j02, a0j01, na03, getcp10desc(cp01,cp10,a0j04) cp10N, 0 as Ramount, 0 as EAmount, a1v06 as Damount" & _
      ", 0 as CAmount, 0 as BAmount, 0 as NAmount, a0k01" & _
      " from ACCTMP08, acc0j0, acc0k0,acc1v0,Staff,caseprogress,nation, acc431" & _
      " where T05='" & Me.Name & "' and T14='" & strUserNum & "' and a0j13(+)=T01 and a0j01(+)=T02 and a0k01(+)=T01 and st01(+)=a0k20" & _
      " and a1v01(+)=a0j01 and a1v02(+)=a0j13 and a1v06>0 and a1v18 is null and cp09(+)=a0j01 and na01(+)=a0j04" & _
      " and axc02(+)=a0k01 "
   'Modified by Lydia 2025/05/23 增加# 已開INVOICE => decode(a0k40,null,a0k32,'#')
   'Modified by Lydia 2025/06/23 +INVOICE編號a0k40
   'Modify By Sindy 2025/9/10 ST02 As a0k20 => ST02||getlos04(a0j01,1) as a0k20
   strUnion = strUnion & " union select a0k03, a0k04, a0k11, decode(a1v18, '1',a0k02,NULL) as RDate, decode(a1v18, '1', '收款扣繳', '補扣繳') as RNo" & _
      " ,  decode(a0k40,null,a0k32,'#') a0k32, axc01, ST02||getlos04(a0j01,1) As a0k20, a0j02, a0j01, na03, getcp10desc(cp01,cp10,a0j04) cp10N, 0 as Ramount, 0 as EAmount, a1v06 as Damount" & _
      ", 0 as CAmount, 0 as BAmount, 0 as NAmount, a0k01, a0k40" & _
      " from ACCTMP08, acc0j0, acc0k0,acc1v0,Staff,caseprogress,nation, acc431" & _
      " where T05='" & Me.Name & "' and T14='" & strUserNum & "' and a0j13(+)=T01 and a0j01(+)=T02 and a0k01(+)=T01 and st01(+)=a0k20" & _
      " and a1v01(+)=a0j01 and a1v02(+)=a0j13 and a1v06>0 and a1v18 is null and cp09(+)=a0j01 and na01(+)=a0j04" & _
      " and axc02(+)=a0k01 "
   '20140115END Modify by Eric  增加 a0k32 axc01
  
    
   adoadodc1.Open strUnion & " order by a0k01 asc, RNo asc", adoTaie, adOpenStatic, adLockReadOnly
   Adodc1.Recordset.Requery
   strCon9 = ""
     
   If Adodc1.Recordset.RecordCount = 0 Then
      Adodc1.Recordset.Close
      MsgBox MsgText(28), , MsgText(5)
      strCon9 = MsgText(602)
      FormClear
      Exit Sub
   End If
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
Public Sub KeyDefine(KeyCode As Integer)
   Select Case KeyCode
      Case vbKeyF12
         If Text7 = "" Then
            Text7 = "0"
         End If
         If Text8 = "" Then
            Text8 = "00"
         End If
         If FormCheck Then
            Screen.MousePointer = vbHourglass
            
            QueryTable
            'Modify by Morgan 2005/3/28 加所別控制
            'AdodcRefresh
            If Adodc1.Recordset.State = adStateOpen Then Adodc1.Recordset.Close
            DataGrid1.Refresh
            Erase strExc: strExc(1) = Text1: strExc(2) = Text6: strExc(3) = Text7: strExc(4) = Text8
            'Modified by Lydia 2020/03/30
            'If PUB_CheckCaseZone(strExc, pub_strUserOffice, "1") = True Then
            If PUB_CheckCaseZone(strExc, m_strUserOffice, "1") = True Then
               AdodcRefresh
            End If
            '2005/3/28 end
            AccShow                            '20140115 add by eric
            Screen.MousePointer = vbDefault
            Exit Sub
         Else
            MsgBox MsgText(181), , MsgText(5)
         End If
   End Select
   KeyEnter KeyCode
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub
'20140115START add by eric
Public Sub AccShow()
Dim douRamt As Double
Dim douEAmt As Double
Dim douDAmt As Double
Dim douCAmt As Double
Dim douBAmt As Double
Dim douNAmt As Double
Dim strQ As String

   douRamt = 0
   douEAmt = 0
   douDAmt = 0
   douCAmt = 0
   douBAmt = 0
   douNAmt = 0
   
   'Modify by Amy 2015/04/27 瑞婷:查詢很慢
'   If adoadodc1.State = adStateOpen Then
'      Set adoaccsum = adoadodc1.Clone
'      Do While adoaccsum.EOF = False
'         If IsNull(adoaccsum.Fields("RAmount").Value) = False Then
'            douRamt = douRamt + Val(adoaccsum.Fields("RAmount").Value)
'         End If
'         If IsNull(adoadodc1.Fields("EAmount").Value) = False Then
'            douEAmt = douEAmt + Val(adoaccsum.Fields("EAmount").Value)
'         End If
'         If IsNull(adoaccsum.Fields("DAmount").Value) = False Then
'            douDAmt = douDAmt + Val(adoaccsum.Fields("DAmount").Value)
'         End If
'         If IsNull(adoadodc1.Fields("CAmount").Value) = False Then
'            douCAmt = douCAmt + Val(adoaccsum.Fields("CAmount").Value)
'         End If
'         If IsNull(adoaccsum.Fields("BAmount").Value) = False Then
'            douBAmt = douBAmt + Val(adoaccsum.Fields("BAmount").Value)
'         End If
'         If IsNull(adoadodc1.Fields("NAmount").Value) = False Then
'            douNAmt = douNAmt + Val(adoaccsum.Fields("NAmount").Value)
'         End If
'         adoaccsum.MoveNext
'      Loop
'      If adoaccsum.RecordCount <> 0 Then
'         adoaccsum.MoveFirst
'      End If
'      adoaccsum.Close
'   End If
   
   strQ = "Select sum(T08) as RAmount,sum(T09) as EAmount,sum(T10) as DAmount,sum(T11) as CAmount, sum(T12) as BAmount, " & _
            "sum(T08)-sum(T09)-sum(T11)+sum(T12) as NAmount From acctmp08 Where T05='" & Me.Name & "' and T14='" & strUserNum & "' "
   If adoaccsum.State = adStateOpen Then adoaccsum.Close
   adoaccsum.Open strQ, cnnConnection, adOpenStatic, adLockReadOnly
   If IsNull(adoaccsum.Fields("RAmount").Value) = False Then
        douRamt = Val(adoaccsum.Fields("RAmount").Value)
   End If
    If IsNull(adoaccsum.Fields("EAmount").Value) = False Then
        douEAmt = Val(adoaccsum.Fields("EAmount").Value)
   End If
    If IsNull(adoaccsum.Fields("DAmount").Value) = False Then
        douDAmt = Val(adoaccsum.Fields("DAmount").Value)
   End If
    If IsNull(adoaccsum.Fields("CAmount").Value) = False Then
        douCAmt = Val(adoaccsum.Fields("CAmount").Value)
   End If
   If IsNull(adoaccsum.Fields("BAmount").Value) = False Then
        douBAmt = Val(adoaccsum.Fields("BAmount").Value)
   End If
   If IsNull(adoaccsum.Fields("NAmount").Value) = False Then
        douNAmt = Val(adoaccsum.Fields("NAmount").Value)
   End If
   'end 2015/04/27
   Text4 = Format(douRamt, FDollar)
   Text5 = Format(douEAmt, FDollar)
   Text9 = Format(douDAmt, FDollar)
   Text10 = Format(douCAmt, FDollar)
   Text11 = Format(douBAmt, FDollar)
   Text12 = Format(douNAmt, FDollar)

End Sub
'20140115END add by eric

'20140115START Add By eric
Private Sub cmdCrtRct_Click()
 
   If Adodc1.Recordset.State = 1 Then
      If Adodc1.Recordset.RecordCount = 0 Then
         Exit Sub
      End If
      If Adodc1.Recordset.Fields("a0k11").Value = "J" And Left(Adodc1.Recordset.Fields("RNo").Value, 1) = "E" And IsNull(Adodc1.Recordset.Fields("axc01").Value) Then
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
         MsgBox "若 J 公司的 E 單據請款單無發票,才可開立發票!!"
      End If
   Else
      MsgBox "請先按 F12 查詢並點選單據資料..."
   End If
   
End Sub
'20140121END

Private Sub Text8_Validate(Cancel As Boolean)
'   If Text6 = "" Then
'      Exit Sub
'   End If
'   QueryTable
'   AdodcRefresh
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
   If Text1 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Text6 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Text7 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Text8 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   FormCheck = False
End Function



