VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc1230 
   AutoRedraw      =   -1  'True
   Caption         =   "智權人員帳款查詢"
   ClientHeight    =   5400
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9264
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5400
   ScaleWidth      =   9264
   Begin VB.CheckBox ChkBillDate 
      Caption         =   "排除未達客戶付款週期之應收帳款"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5376
      TabIndex        =   8
      Top             =   1200
      Width           =   3792
   End
   Begin VB.TextBox Text3 
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
      Left            =   6096
      TabIndex        =   7
      Text            =   "1"
      Top             =   888
      Width           =   400
   End
   Begin VB.CommandButton Command1 
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
      Height          =   300
      Left            =   7152
      TabIndex        =   11
      Top             =   528
      Width           =   1785
   End
   Begin VB.TextBox txtAmt 
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
      Index           =   0
      Left            =   588
      TabIndex        =   26
      Top             =   4848
      Width           =   1000
   End
   Begin VB.TextBox txtAmt 
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
      Index           =   1
      Left            =   2076
      TabIndex        =   25
      Top             =   4848
      Width           =   1000
   End
   Begin VB.TextBox txtAmt 
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
      Index           =   2
      Left            =   3588
      TabIndex        =   24
      Top             =   4848
      Width           =   1000
   End
   Begin VB.TextBox txtAmt 
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
      Index           =   3
      Left            =   5100
      TabIndex        =   23
      Top             =   4848
      Width           =   1000
   End
   Begin VB.TextBox txtAmt 
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
      Index           =   4
      Left            =   6612
      TabIndex        =   22
      Top             =   4848
      Width           =   1000
   End
   Begin VB.TextBox txtAmt 
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
      Index           =   5
      Left            =   8112
      TabIndex        =   21
      Top             =   4848
      Width           =   1000
   End
   Begin VB.CommandButton Command3 
      Caption         =   "請款單開立發票"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7152
      TabIndex        =   10
      Top             =   84
      Width           =   1785
   End
   Begin VB.CheckBox Check1 
      Caption         =   "含未列印收據"
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
      Left            =   4956
      TabIndex        =   3
      Top             =   504
      Value           =   1  '核取
      Width           =   1725
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc1230.frx":0000
      Height          =   3288
      Left            =   156
      TabIndex        =   19
      Top             =   1512
      Width           =   8856
      _ExtentX        =   15600
      _ExtentY        =   5800
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   18
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
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "智權人員帳款查詢"
      ColumnCount     =   17
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
      BeginProperty Column03 
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
      BeginProperty Column04 
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
      BeginProperty Column05 
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
      BeginProperty Column06 
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
      BeginProperty Column09 
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
      BeginProperty Column12 
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
      BeginProperty Column13 
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
      BeginProperty Column14 
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
      BeginProperty Column15 
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
      BeginProperty Column16 
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
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         Size            =   345
         BeginProperty Column00 
            Alignment       =   2
            ColumnWidth     =   492.095
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   972.284
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   348.095
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   408.189
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   2784.189
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1391.811
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1344.189
         EndProperty
         BeginProperty Column09 
            Alignment       =   2
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1272.189
         EndProperty
         BeginProperty Column11 
            Alignment       =   1
         EndProperty
         BeginProperty Column12 
            Alignment       =   1
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
            Alignment       =   1
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   312
      Left            =   216
      Top             =   1392
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
   Begin VB.CommandButton Command2 
      Caption         =   "單據內容"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5808
      TabIndex        =   9
      Top             =   84
      Width           =   1212
   End
   Begin VB.TextBox Text6 
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
      Left            =   2736
      TabIndex        =   6
      Text            =   "1"
      Top             =   888
      Width           =   400
   End
   Begin VB.TextBox Text4 
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
      Left            =   1056
      TabIndex        =   4
      Top             =   888
      Width           =   400
   End
   Begin VB.TextBox Text5 
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
      Left            =   1776
      TabIndex        =   5
      Top             =   888
      Width           =   400
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1056
      TabIndex        =   0
      Top             =   168
      Width           =   1572
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1056
      TabIndex        =   1
      Top             =   528
      Width           =   1572
      _ExtentX        =   2773
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.8
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
      Left            =   2976
      TabIndex        =   2
      Top             =   528
      Width           =   1572
      _ExtentX        =   2773
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.8
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
      Left            =   144
      TabIndex        =   35
      Top             =   1248
      Width           =   4260
   End
   Begin MSForms.TextBox Text2 
      Height          =   300
      Left            =   2616
      TabIndex        =   20
      Top             =   168
      Width           =   1812
      VariousPropertyBits=   671105049
      BackColor       =   14737632
      Size            =   "3201;529"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "(1.抬頭+日期 2.收據編號)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   6576
      TabIndex        =   34
      Top             =   924
      Width           =   2652
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "排序"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   5616
      TabIndex        =   33
      Top             =   924
      Width           =   612
   End
   Begin VB.Label lblAmt 
      BackStyle       =   0  '透明
      Caption         =   "應收"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   0
      Left            =   156
      TabIndex        =   32
      Top             =   4872
      Width           =   492
   End
   Begin VB.Label lblAmt 
      BackStyle       =   0  '透明
      Caption         =   "已收"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   1
      Left            =   1644
      TabIndex        =   31
      Top             =   4872
      Width           =   492
   End
   Begin VB.Label lblAmt 
      BackStyle       =   0  '透明
      Caption         =   "扣繳"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   2
      Left            =   3144
      TabIndex        =   30
      Top             =   4872
      Width           =   492
   End
   Begin VB.Label lblAmt 
      BackStyle       =   0  '透明
      Caption         =   "銷帳"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   3
      Left            =   4644
      TabIndex        =   29
      Top             =   4872
      Width           =   492
   End
   Begin VB.Label lblAmt 
      BackStyle       =   0  '透明
      Caption         =   "退費"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   4
      Left            =   6156
      TabIndex        =   28
      Top             =   4872
      Width           =   492
   End
   Begin VB.Label lblAmt 
      BackStyle       =   0  '透明
      Caption         =   "未收"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   5
      Left            =   7668
      TabIndex        =   27
      Top             =   4872
      Width           =   492
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   -144
      Top             =   4992
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "資料"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   2256
      TabIndex        =   18
      Top             =   924
      Width           =   972
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "(1.未收 2.收款 3.往來)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   3216
      TabIndex        =   17
      Top             =   924
      Width           =   2652
   End
   Begin VB.Label Label7 
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
      Height          =   252
      Left            =   1536
      TabIndex        =   16
      Top             =   888
      Width           =   252
   End
   Begin VB.Label Label4 
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
      Height          =   252
      Left            =   96
      TabIndex        =   15
      Top             =   924
      Width           =   972
   End
   Begin VB.Label Label5 
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
      Height          =   252
      Left            =   2736
      TabIndex        =   14
      Top             =   528
      Width           =   252
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "往來日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   96
      TabIndex        =   13
      Top             =   528
      Width           =   972
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "智權人員"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   96
      TabIndex        =   12
      Top             =   168
      Width           =   972
   End
End
Attribute VB_Name = "Frmacc1230"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/16 Form2.0已修改
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/11/30 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/28 日期欄已修改
Option Explicit

Public adoacc0m0 As New ADODB.Recordset
Public adoacctmp05 As New ADODB.Recordset
Public adostaff As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Dim strSql As String
Dim strSQL1 As String
Dim strSQL2 As String
Dim strType As String

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
End Sub
'end 2016/01/19

Private Sub Command2_Click()
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
      If Adodc1.Recordset.Fields("a0k11").Value = "J" And _
         Left(Adodc1.Recordset.Fields("RNo").Value, 1) = "E" Then
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
   'Modify by Amy 2025/10/30 W9330/H5640
   Me.Width = 9360
   Me.Height = 5846
   'end 2025/10/30
   'Modify by Amy 2023/10/06 原(lngWidth - Me.Width) / 2,調整切畫面不用移-瑞婷
   Me.Move 0, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath2)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat
   OpenTable
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
   
   'Add By Sindy 2014/1/3
   If PUB_GetST06(strUserNum) = "1" Then
      Command3.Visible = True
      'Added by Lydia 2016/01/20
      Command1.Visible = True
   Else
      Command3.Visible = False
      'Added by Lydia 2016/01/20 分所不可使用"收據抬頭修改"
      Command1.Visible = False
   End If
   '2014/1/3 END

End Sub

Private Sub Form_Unload(Cancel As Integer)
'   Acctmp05Delete
'   adoacctmp05.Close
   StatusClear
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc1230 = Nothing
End Sub

Private Sub Text1_Change()
   If Text1 = MsgText(601) Then
      Exit Sub
   End If
   Text2 = StaffQuery(Text1)
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
   CloseIme
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoadodc1.CursorLocation = adUseClient
   'Modify by Morgan 2011/8/23 改從 0j0 抓 cp
   'Modified by Morgan 2011/12/27 取消 a0j20,a0j21
   adoadodc1.Open "select a0k11, a0k02, a0k01, a0k03, a0k04, a0k20, a0j02, a0j01, cp09, na03, getcp10desc(cp01,cp10,a0j04) cp10N, (a0j09 + a0j10) as RAmount, cp75 as EAmount, cp76, cp77, cp78 as BAmount, cp79 as NAmount from acc0k0, acc0j0,caseprogress,nation where a0k03 = '1' and a0j13(+) = a0k01 and cp09(+) = a0j01 and na01(+)=a0j04", adoTaie, adOpenStatic, adLockReadOnly
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
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

On Error GoTo Checking
   strSql = ""
   strSQL1 = ""
   strSQL2 = ""
   If Text1 <> MsgText(601) Then
      strSql = " and a0k20 = '" & Text1 & "'"
      strSQL1 = " and a0k20 = '" & Text1 & "'"
      strSQL2 = " and a0k20 = '" & Text1 & "'"
   End If
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
   strType = ""
   
   '2012/10/15 add by sonia +不包含未列印收據故加 and a0k32 is null條件
   If Me.Check1.Value = 0 Then
       strSql = strSql & " and A0K32 is null"
       strSQL1 = strSQL1 & " and A0K32 is null"
       strSQL2 = strSQL2 & " and A0K32 is null"
   End If
   
   Select Case Text6
      Case "1"
         'Modify by Morgan 2006/7/12
         'strType = " and a0k17 < a0k06 and cp79>0"
         strType = " and (a0k06+a0k07) > (nvl(a0k17, 0)+nvl(a0k18, 0)) and cp79>0"
      Case "2"
         'Modify by Morgan 2006/7/12
         'strType = " and a0k17 >= a0k06"
         strType = " and (nvl(a0k17, 0)+nvl(a0k18, 0))>0"
   End Select
    '若非北所員工, 只能看該所資料
    If pub_strUserOffice <> "1" Then
        strSql = strSql & " and ''||ST06='" & pub_strUserOffice & "' "
        strSQL1 = strSQL1 & " and ''||ST06='" & pub_strUserOffice & "' "
        strSQL2 = strSQL2 & " and ''||ST06='" & pub_strUserOffice & "' "
    End If
    
   If adoadodc1.State = adStateOpen Then
      adoadodc1.Close
   End If
   adoadodc1.CursorLocation = adUseClient
   
   
   'Modify by Morgan 2011/10/27 考慮拆收據情形改先寫暫存
   ''收據
   ''Modify by Morgan 2011/8/23 改從 0j0 抓 cp
   'strUnion = "select a0k03, a0k04, a0k11, a0k02 as RDate, a0k01 as RNo, a0k20, a0j02, a0j01, a0j21, a0j20" & _
      ", nvl(cp16, 0) as RAmount, 0 as EAmount, 0 as DAmount, nvl(cp77, 0) as CAmount, nvl(cp78, 0) as BAmount" & _
      ", (a0j09 + a0j10) - nvl(cp75, 0) as NAmount, a0k01" & _
      " from acc0k0, acc0j0, caseprogress, Staff" & _
      " where (a0k09 is null or a0k09 = 0) and a0j13 (+)=a0k01 and cp09(+) = a0j01 " & _
      " and ST01(+)=A0K20 " & strSql & strType
   ''收款
   ''Modify by Morgan 2011/8/11 已扣繳改抓 acc1v0 資料,原a0k14欄位保留再使用
   'strUnion = strUnion & " union" & _
      " select a0k03, a0k04, a0k11, a0l02 as RDate, a0l01 as RNo, a0k20, a0j02, a0j01, a0j21, a0j20" & _
      ", 0 as RAmount, (nvl(a1u04, 0) + nvl(a1u05, 0)) as EAmount, a1v06 as DAmount, 0 as CAmount, 0 as BAmount" & _
      ", 0 as NAmount, a0k01 from acc0k0, acc1u0, acc0l0, acc0j0, caseprogress, Staff,acc1v0" & _
      " where (a0k09 is null or a0k09 = 0) and a1u02(+) = a0k01 and substr(a1u01,1,1)='F'" & _
      " and a0l01(+)=a1u01 and a0j01 (+)=a1u03 and a1u03 = cp09(+)" & _
      " and a1v01(+)=cp09" & _
      " and ST01(+)=A0K20 " & strSQL1 & strType
   
   'If Text6 = "1" Or Text6 = "3" Then
      ''銷帳退費
      ''Modify by Morgan 2011/8/23 改從 0j0 抓 cp
      'strUnion = strUnion & " union" & _
         " select a0k03, a0k04, a0k11, a0s03 as RDate, a0s01 as RNo, a0k20, cp01||cp02||cp03||cp04, '', '', ''" & _
         ", 0 as Ramount, 0 as EAmount, 0 as Damount, a0s05 as CAmount, (nvl(a0s06, 0) + nvl(a0s07, 0)) as BAmount" & _
         ", 0 as NAmount, a0k01 from acc0s0, acc0k0,acc0j0, caseprogress, Staff" & _
         " where (a0k09 is null or a0k09 = 0) and a0s02(+) = a0k01" & _
         " and a0j13 (+)=a0k01 and cp09(+)=a0j01 and a0s01 is not null" & _
         " and ST01(+)=A0K20 " & strSQL2 & strType
   'End If
   adoTaie.Execute "delete ACCTMP08 where T05='" & Me.Name & "' and T14='" & strUserNum & "'"
   
   'Modified by Lydia 2016/04/11 舊收據無a0j01
   strUnion = "select a0k01,NVL(a0j01,' ') a0j01,'" & Me.Name & "',a0k01,'" & strUserNum & "'" & _
      " from acc0k0, acc0j0, caseprogress, Staff" & _
      " where (a0k09 is null or a0k09 = 0) and a0j13(+)=a0k01 and cp09(+) = a0j01 " & _
      " and ST01(+)=A0K20 " & strSql & strType
   'Add by Amy 2022/07/26 有下智權人員抓此人之案源資料
   If Trim(Text1) <> MsgText(601) Then
        strUnion = strUnion & " Union " & _
            "Select a0k01,NVL(a0j01,' ') a0j01,'" & Me.Name & "',a0k01,'" & strUserNum & "'" & _
            " From acc0k0, acc0j0, caseprogress, Staff, LawOfficeSource" & _
            " Where (a0k09 is null or a0k09 = 0) and a0j13(+)=a0k01 and los06(+) = a0j01 And los01 is not null " & _
            " And cp09(+)=los06 And ST01(+)=los04 " & Replace(strSql, "a0k20", "los04") & strType
   End If
   
   'Modified by Lydia 2016/04/11 舊收據無a0j01
   'modify by sonia 2016/9/26 加and a0j13(+) = a1u02,否則F10507303會抓到二筆ACC0J0
   strUnion = strUnion & " union select a0k01, NVL(a0j01,' ') a0j01,'" & Me.Name & "',a0l01,'" & strUserNum & "'" & _
      " from acc0k0, acc1u0, acc0l0, acc0j0, caseprogress, Staff" & _
      " where (a0k09 is null or a0k09 = 0) and a1u02(+) = a0k01 and substr(a1u01,1,1)='F'" & _
      " and a0l01(+)=a1u01 and a0j01(+)=a1u03 and a0j13(+) = a1u02 and a1u03 = cp09(+)" & _
      " and ST01(+)=A0K20 " & strSQL1 & strType
   'Add by Amy 2022/07/26 有下智權人員抓此人之案源資料
   If Trim(Text1) <> MsgText(601) Then
        strUnion = strUnion & " Union " & _
            "Select a0k01, NVL(a0j01,' ') a0j01,'" & Me.Name & "',a0l01,'" & strUserNum & "'" & _
            " From acc0k0, acc1u0, acc0l0, acc0j0, caseprogress, Staff ,LawOfficeSource" & _
            " Where (a0k09 is null or a0k09 = 0) and a1u02(+) = a0k01 and substr(a1u01,1,1)='F'" & _
            " and a0l01(+)=a1u01 and a0j01(+)=a1u03 and a0j13(+) = a1u02 and a1u03 = Los06(+) And los01 is not null" & _
            " And los06=cp09(+) and ST01(+)=los04 " & Replace(strSQL1, "a0k20", "los04") & strType
   End If
      
   If Text6 = "1" Or Text6 = "3" Then    '銷帳
      'Modified by Lydia 2016/04/11 舊收據無a0j01
      strUnion = strUnion & " union select a0k01, NVL(a0j01,' ') a0j01,'" & Me.Name & "',a0s01,'" & strUserNum & "'" & _
         " from acc0k0,acc0s0,acc0j0, caseprogress, Staff" & _
         " where (a0k09 is null or a0k09 = 0) and a0k10 is not null" & _
         " and a0s02(+)=a0k01 and a0s01 is not null" & _
         " and a0j13(+)=a0k01 and cp09(+)=a0j01" & _
         " and ST01(+)=A0K20 " & strSQL2 & strType
         'Add by Amy 2022/07/26 有下智權人員抓此人之案源資料
        If Trim(Text1) <> MsgText(601) Then
            strUnion = strUnion & " Union " & _
                "Select a0k01, NVL(a0j01,' ') a0j01,'" & Me.Name & "',a0s01,'" & strUserNum & "'" & _
                " From acc0k0,acc0s0,acc0j0, caseprogress, Staff, LawOfficeSource" & _
                " Where (a0k09 is null or a0k09 = 0) and a0k10 is not null" & _
                " and a0s02(+)=a0k01 and a0s01 is not null" & _
                " and a0j13(+)=a0k01 and los04(+)=a0j01 And los01 is not null" & _
                " And los04=cp09(+) and ST01(+)=los04 " & Replace(strSQL2, "a0k20", "los04") & strType
        End If
   End If
   
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
   
   '刪除銷帳金額為0的資料
   strSql = "delete ACCTMP08 where T05='" & Me.Name & "' and T14='" & strUserNum & "' and T06 like 'I%' and T11=0"
   adoTaie.Execute strSql, intI
   
   '去除拆收據已收齊的資料
   If Text6 = "1" Then
      strSql = "delete ACCTMP08 where T05='" & Me.Name & "' and T14='" & strUserNum & "' and T08-T09-T11+T12=0"
      adoTaie.Execute strSql, intI
   End If
   
   'Modified by Morgan 2011/12/27 取消 a0j20,a0j21
   
   '收據
   'Modified by Lydia 2025/05/23 增加# 已開INVOICE => decode(a0k40,null,a0k32,'#')
   strUnion = "select a0k03, a0k04, a0k11, a0k02 as RDate, a0k01 as RNo,decode(a0k40,null,a0k32,'#') a0k32,axc01, a0k20, a0j02, a0j01, na03, getcp10desc(cp01,cp10,a0j04) cp10N" & _
      ", T08 RAmount, T09 EAmount, T10 DAmount, T11 CAmount" & _
      ", T12 BAmount, T08-T09-T11+T12 NAmount, a0k01" & _
      " from ACCTMP08, acc0j0,acc0k0,caseprogress,nation,acc431" & _
      " where T05='" & Me.Name & "' and T14='" & strUserNum & "' and T06=T01 and a0j13(+)=T01 and a0j01(+)=T02" & _
      " and a0k01(+)=T01 and cp09(+)=a0j01 and na01(+)=a0j04 and axc02(+)=T01"
      
   '收款
   strUnion = strUnion & " union" & _
      " select a0k03, a0k04, a0k11, a0l02 as RDate, a0l01 as RNo,'' a0k32,'' axc01, a0k20, a0j02, a0j01, na03, getcp10desc(cp01,cp10,a0j04) cp10N" & _
      ", 0 as RAmount, (nvl(a1u04, 0) + nvl(a1u05, 0)) as EAmount, nvl(a1u06,0) as DAmount, 0 as CAmount, 0 as BAmount" & _
      ", 0 as NAmount, a0k01" & _
      " from ACCTMP08, acc0j0, acc0k0, acc1u0, acc0l0,caseprogress,nation" & _
      " where T05='" & Me.Name & "' and T14='" & strUserNum & "' and substr(T06,1,1)='F' and a0j13(+)=T01 and a0j01(+)=T02" & _
      " and a0k01(+)=T01 and a1u01(+)=T06 and a1u02(+)=T01 and a1u03(+)=T02 and a0l01(+)=T06 and cp09(+)=a0j01 and na01(+)=a0j04"
      
   If Text6 = "1" Or Text6 = "3" Then
      '銷帳退費
      strUnion = strUnion & " union" & _
         " select a0k03, a0k04, a0k11, a0s03 as RDate, a0s01 as RNo,'' a0k32,'' axc01, a0k20,a0j02,a0j01, na03, getcp10desc(cp01,cp10,a0j04) cp10N" & _
         ", 0 as Ramount, 0 as EAmount, 0 as Damount,nvl(a1u07, 0)+nvl(a1u09, 0) as CAmount" & _
         ", nvl(a1u08, 0)+nvl(a1u10, 0) as BAmount, 0 as NAmount, a0k01" & _
         " from ACCTMP08, acc0j0, acc0k0, acc1u0, acc0s0,caseprogress,nation" & _
         " where T05='" & Me.Name & "' and T14='" & strUserNum & "' and substr(T06,1,1)='I' and a0j13(+)=T01 and a0j01(+)=T02" & _
         " and a0k01(+)=T01 and a1u01(+)=T06 and a1u02(+)=T01 and a1u03(+)=T02 and a0S01(+)=T06 and cp09(+)=a0j01 and na01(+)=a0j04"
   End If
   'end 2011/10/27
   'Modified by Lydia 2016/01/26 + 排序選項
   'strUnion = strUnion & " order by a0k01 asc, RNo asc"
   If Text3 = "1" Then
      strUnion = strUnion & " order by a0k04 asc, RDate asc, RNo asc"
   Else
      strUnion = strUnion & " order by RNo asc,a0j01 asc"
   End If
   
   adoadodc1.Open strUnion, adoTaie, adOpenStatic, adLockReadOnly
   Adodc1.Recordset.Requery
   Calculate 'Add by Sindy 2013/12/31
   If Adodc1.Recordset.RecordCount = 0 Then
      Adodc1.Recordset.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   End If
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'Add by Sindy 2013/12/31
Private Sub Calculate()
   Dim dblAmt(5) As Double
   Set RsTemp = Adodc1.Recordset.Clone
   With RsTemp
   If .RecordCount > 0 Then
      Do While Not .EOF
'         Select Case Left("" & .Fields("RNo"), 1)
'            Case "E"
               dblAmt(0) = dblAmt(0) + Val("" & .Fields("RAmount"))
               dblAmt(5) = dblAmt(5) + Val("" & .Fields("NAmount"))
'            Case "F"
               dblAmt(1) = dblAmt(1) + Val("" & .Fields("EAmount"))
               dblAmt(2) = dblAmt(2) + Val("" & .Fields("DAmount"))
'            Case "I"
               dblAmt(3) = dblAmt(3) + Val("" & .Fields("CAmount"))
               dblAmt(4) = dblAmt(4) + Val("" & .Fields("BAmount"))
'         End Select
         .MoveNext
      Loop
   End If
   End With
   For intI = 0 To 5
      txtAmt(intI) = Format(dblAmt(intI), "#,##0")
   Next
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
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
'  功能鍵定義
'
'*************************************************
Public Sub KeyDefine(KeyCode As Integer)
   Select Case KeyCode
      Case vbKeyF12
         If FormCheck Then
            Screen.MousePointer = vbHourglass
            AdodcRefresh
            Screen.MousePointer = vbDefault
            Exit Sub
         Else
            MsgBox MsgText(181), , MsgText(5)
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
   If Text1 <> MsgText(601) Then
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
   'Added by Lydia 2016/01/26
   If Text3 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   FormCheck = False
End Function

''2012/10/15 add by sonia
'Private Sub Text7_GotFocus()
'    TextInverse Text7
'End Sub
'
'Private Sub Text7_KeyPress(KeyAscii As Integer)
'    KeyAscii = UpperCase(KeyAscii)
'    Select Case KeyAscii
'    Case 89, 8
'    Case Else
'        KeyAscii = 0
'    End Select
'End Sub
''2012/10/15 end
'Added by Lydia 2016/01/26
Private Sub Text3_GotFocus()
   TextInverse Text3
End Sub
