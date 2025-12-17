VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc4240 
   AutoRedraw      =   -1  'True
   Caption         =   "科目明細查詢(對沖)"
   ClientHeight    =   5388
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9480
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5388
   ScaleWidth      =   9480
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   1572
      TabIndex        =   0
      Top             =   90
      Width           =   620
   End
   Begin VB.TextBox Text14 
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
      Left            =   3672
      TabIndex        =   7
      Top             =   690
      Width           =   1572
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   5676
      TabIndex        =   4
      Top             =   396
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
      Left            =   7452
      TabIndex        =   5
      Top             =   396
      Width           =   1572
      _ExtentX        =   2794
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc4240.frx":0000
      Height          =   3360
      Left            =   240
      TabIndex        =   24
      Top             =   1440
      Width           =   9000
      _ExtentX        =   15875
      _ExtentY        =   5927
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   18
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.8
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
      Caption         =   "科目明細資料"
      ColumnCount     =   12
      BeginProperty Column00 
         DataField       =   "ax201"
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
         DataField       =   "a0205"
         Caption         =   "傳票日期"
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
         DataField       =   "ax202"
         Caption         =   "傳票號碼"
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
         DataField       =   "ax206"
         Caption         =   "借方金額"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "ax207"
         Caption         =   "貸方金額"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "ax212"
         Caption         =   "摘要"
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
         DataField       =   "ax214"
         Caption         =   "對沖代號(本所案號)"
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
         DataField       =   "ax208"
         Caption         =   "對沖代號(客)"
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
         DataField       =   "cust"
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
      BeginProperty Column09 
         DataField       =   "ax209"
         Caption         =   "對沖代號(業)"
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
         DataField       =   "salesname"
         Caption         =   "智權人員名稱"
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
         DataField       =   "ax213"
         Caption         =   "對沖代號(其他)"
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
            ColumnWidth     =   527.811
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   1031.811
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1344.189
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   3539.906
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   2039.811
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1379.906
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   3047.811
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1428.095
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   2111.811
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1620.284
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   312
      Left            =   240
      Top             =   1176
      Visible         =   0   'False
      Width           =   972
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
   Begin VB.TextBox Text12 
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
      Left            =   2520
      TabIndex        =   9
      Top             =   990
      Width           =   1572
   End
   Begin VB.TextBox Text11 
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
      Height          =   324
      Left            =   7440
      TabIndex        =   22
      Top             =   4896
      Width           =   1500
   End
   Begin VB.TextBox Text10 
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
      Left            =   3636
      TabIndex        =   20
      Top             =   90
      Width           =   1572
   End
   Begin VB.TextBox Text9 
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
      Left            =   3000
      TabIndex        =   1
      Top             =   90
      Width           =   612
   End
   Begin VB.TextBox Text8 
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
      Height          =   300
      Left            =   2652
      TabIndex        =   19
      Top             =   390
      Width           =   1932
   End
   Begin VB.TextBox Text7 
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
      Left            =   1572
      TabIndex        =   3
      Top             =   390
      Width           =   1092
   End
   Begin VB.TextBox Text4 
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
      Height          =   324
      Left            =   4512
      TabIndex        =   13
      Top             =   4896
      Width           =   1500
   End
   Begin VB.TextBox Text3 
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
      Height          =   324
      Left            =   2760
      TabIndex        =   12
      Top             =   4872
      Width           =   1500
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
      Left            =   7200
      TabIndex        =   8
      Top             =   690
      Width           =   1572
   End
   Begin VB.TextBox Text2 
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
      Left            =   1812
      TabIndex        =   6
      Top             =   690
      Width           =   1572
   End
   Begin MSForms.TextBox Text13 
      Height          =   300
      Left            =   6240
      TabIndex        =   10
      Top             =   996
      Width           =   1572
      VariousPropertyBits=   679493659
      Size            =   "2773;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text15 
      Height          =   312
      Left            =   5880
      TabIndex        =   2
      Top             =   96
      Width           =   2172
      VariousPropertyBits=   679493659
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "摘要                               (模糊查詢)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Left            =   5292
      TabIndex        =   31
      Top             =   96
      Width           =   3828
   End
   Begin VB.Label Label14 
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
      Left            =   3468
      TabIndex        =   30
      Top             =   696
      Width           =   132
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '透明
      Caption         =   "對沖代號(其他)"
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
      Left            =   4692
      TabIndex        =   29
      Top             =   996
      Width           =   1572
   End
   Begin VB.Label Label12 
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
      Left            =   7296
      TabIndex        =   28
      Top             =   372
      Width           =   132
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '透明
      Caption         =   "傳票日期"
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
      Left            =   4692
      TabIndex        =   27
      Top             =   420
      Width           =   972
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "貸"
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
      Left            =   4272
      TabIndex        =   26
      Top             =   4896
      Width           =   252
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "借"
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
      Left            =   2520
      TabIndex        =   25
      Top             =   4896
      Width           =   252
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "對沖代號(本所案號)"
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
      Left            =   360
      TabIndex        =   23
      Top             =   996
      Width           =   2172
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "餘額"
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
      Left            =   6756
      TabIndex        =   21
      Top             =   4896
      Width           =   612
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   1296
      Left            =   240
      Top             =   36
      Width           =   9000
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
      Height          =   252
      Left            =   960
      TabIndex        =   18
      Top             =   4896
      Width           =   612
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "對沖代號(業)"
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
      Left            =   5652
      TabIndex        =   17
      Top             =   696
      Width           =   1452
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   60
      Top             =   2856
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "對沖代號(客)"
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
      Left            =   360
      TabIndex        =   16
      Top             =   696
      Width           =   1452
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "會計科目"
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
      Left            =   360
      TabIndex        =   15
      Top             =   420
      Width           =   972
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "部門別"
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
      Left            =   2292
      TabIndex        =   14
      Top             =   120
      Width           =   732
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "公司別"
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
      Left            =   360
      TabIndex        =   11
      Top             =   120
      Width           =   732
   End
End
Attribute VB_Name = "Frmacc4240"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/13 Form2.0已修改 Text15/DataGrid1/Text13(1110607改)
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/3 日期欄已修改
Option Explicit
Public adoaccsum As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset

'Add by Amy 2020/03/31
Private Sub Combo1_GotFocus()
    TextInverse Combo1
    CloseIme
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Combo1_Validate(Cancel As Boolean)
    If Combo1 = MsgText(601) Then Exit Sub
    
    If InStr(GetBookKeepCmp, Combo1) = 0 Then
        MsgBox Label3 & MsgText(63), , MsgText(5)
        Cancel = True
        Combo1.SetFocus
        Exit Sub
    End If
End Sub
'end 2020/03/31

Private Sub DataGrid1_DblClick()
   If DataGrid1.row >= 0 Then
      '20140124START Modify By eric
      If DataGrid1.Columns(2).Text <> "" Then
         Load Frmacc4221
         Frmacc4221.p_stA0202 = DataGrid1.Columns(2).Text
         Frmacc4221.p_stA0201 = DataGrid1.Columns(0).Text 'Added by Morgan 204/2/25
         Frmacc4221.QueryTable
         Set Frmacc4221.p_oForm = Me
         Me.Enabled = False
      End If
      'If DataGrid1.Columns(1).Text <> "" Then
      '   Load Frmacc4221
      '   Frmacc4221.p_stA0202 = DataGrid1.Columns(1).Text
      '   Frmacc4221.QueryTable
      '   Set Frmacc4221.p_oForm = Me
      '   Me.Enabled = False
      'End If
      '20140124END
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call PUB_SaveTrackMode(0, KeyCode)  ' Add by Amy 2021/12/13 Form2.0 記錄鍵盤傳入順序
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
   Me.Width = 9700 'Modify by Amy 2023/07/19 原:9500
   Me.Height = 5950 'Modify by Amy 2023/07/19 原:5500
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath2)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   Call Pub_SetCboCmpNo(Combo1, True) 'Add by Amy 2020/03/31 公司別下拉
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat
   OpenTable
   '20140124START Remark By eric    修改公司別為可輸入
   'Text5 = "1"
   '20140124END
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   StatusClear
   strTrackMode = "" 'Add by Amy 2021/12/13 Form2.0 記錄鍵盤傳入順序(清除)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc4240 = Nothing
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text12_GotFocus()
   TextInverse Text12
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text12_Validate(Cancel As Boolean)
   Text12 = CaseNoZero(Text12)
End Sub

Private Sub Text13_GotFocus()
   TextInverse Text13
End Sub

'Modify by Amy 2022/05/07 原:Integer
Private Sub Text13_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
   If Len(Text2) = 6 Then
      Text2 = AfterZero(Text2)
   End If
   'Add By Cheng 2004/05/13
   '對沖代號(客)迄帶對沖代號(客)起的資料
   'Modify by Morgan 2004/11/22 預設尾碼999
   'Modify By Sindy 2014/8/11 999=>ZZZ
   'If Text2.Text <> "" Then Me.Text14.Text = Left(Me.Text2.Text, 6) & "999"
   If Text2.Text <> "" Then Me.Text14.Text = Left(Me.Text2.Text, 6) & "ZZZ"
End Sub

Private Sub Text14_GotFocus()
'   If (Text14 = "") Then Text14 = Text2
   TextInverse Text14
End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text14_Validate(Cancel As Boolean)
   If Len(Text14) = 6 Then
      Text14 = AfterZero(Text14)
   End If
End Sub
'Mark by Amy 2020/03/31 改下拉
'Private Sub Text5_Change()
'   '20140124START Modify By eric
'   'If Text5 = MsgText(601) Then
'   '   Exit Sub
'   'End If
'   If Text5.Text <> "1" And Text5.Text <> "J" And Text5.Text <> "" Then
'      MsgBox "公司別僅可為 1 或 J 或空白  ! (1.台一 J.智權 空白.全部)"
'      Text5.Text = ""
'      Text5.SetFocus
'      Exit Sub
'   End If
'   '20140124END
'
'   'Text6 = A0802Query(Text5)    '2010/3/10 cancel by sonia
'End Sub

'Private Sub Text5_GotFocus()
'   TextInverse Text5
'   '20140124START Modify By eric
'   CloseIme
'   '20140124END
'End Sub

'Private Sub Text5_LostFocus()
'   If Text5.Text <> "1" And Text5.Text <> "J" And Text5.Text <> "" Then
'      MsgBox "公司別僅可為 1 或 J 或空白  ! (1.台一 J.智權 空白.全部)"
'      Text5.Text = ""
'      Text5.SetFocus
'      Exit Sub
'   End If
'End Sub

'20140124ADD By eric
'Private Sub Text5_KeyPress(KeyAscii As Integer)
'   KeyAscii = UpperCase(KeyAscii)
'End Sub
'end 2020/03/31

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoadodc1.CursorLocation = adUseClient
   Select Case strAccount
      Case "2"
      '20140124START Modify By eric
      '   adoadodc1.Open "select ax308 as ax208, ax309 as ax209, ax314 as ax214, a0305 as a0205, ax302 as ax202, ax307 as ax207, ax312 as ax312 from acc031, acc030 where ax301 = a0301 and ax302 = a0302 and ax302 = '1' order by ax302 asc, AX303 ASC", adoTaie, adOpenStatic, adLockReadOnly
         adoadodc1.Open "select ax301 as ax201, ax308 as ax208, ax309 as ax209, ax314 as ax214, a0305 as a0205, ax302 as ax202, ax307 as ax207, ax312 as ax312 from acc031, acc030 where ax301 = a0301 and ax302 = a0302 and ax302 = '1' order by ax302 asc, AX303 ASC", adoTaie, adOpenStatic, adLockReadOnly
      '20140124END
      Case Else
         adoadodc1.Open "select * from acc021, acc020 where ax201 = a0201 and ax202 = a0202 and ax202 = '1' order by ax202 asc,AX203 ASC", adoTaie, adOpenStatic, adLockReadOnly
   End Select
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  查詢資料表(傳票資料)，並計算及顯示合計金額
'
'*************************************************
Public Sub QueryTable()
Dim douDebit, douCredit As Double, strSql As String
Dim StrSQLa As String

On Error GoTo Checking
   If adoadodc1.State = adStateOpen Then
      adoadodc1.Close
   End If
   If adoaccsum.State = adStateOpen Then
      adoaccsum.Close
   End If
   adoadodc1.CursorLocation = adUseClient
   adoaccsum.CursorLocation = adUseClient
   Select Case strAccount
      Case "2"
         '20140124START Add By eric
         'Modify by Amy 2020/03/31 改下拉
         'If Text5 <> MsgText(601) Then
         If Trim(Combo1) <> MsgText(601) Then
            strSql = " and ax301 = '" & Combo1 & "'"
         End If
         '20140124END
         If Text9 <> MsgText(601) Then
            strSql = strSql & " and ax304 = '" & Text9 & "'"
         Else
      '      strSQL = strSQL & " and (ax204 = '" & MsgText(55) & "' or ax204 = '' or ax204 is null)"
         End If
         'Add By Sindy 2010/3/9
         If Text15 <> MsgText(601) Then
            strSql = strSql & " and INSTR(AX312,'" & Text15 & "') > 0 "
         End If
         '2010/3/9 End
         If Text7 <> MsgText(601) Then
            strSql = strSql & " and ax305 = '" & Text7 & "'"
         End If
         If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
            strSql = strSql & " and a0305 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
         End If
         If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
            strSql = strSql & " and a0305 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
         End If
         
         'Modify by Morgan 2003/11/27
'         If Text2 <> MsgText(601) Then
'            strSQL = strSQL & " and ax308 = '" & Text2 & "'"
'         End If
         If Text2 <> MsgText(601) Then
            strSql = strSql & " and ax308 >= '" & Text2 & "'"
         End If
         
         If Text14 <> MsgText(601) Then
            strSql = strSql & " and ax308 <= '" & Text14 & "'"
         End If
         
         'End
         
         If Text1 <> MsgText(601) Then
            strSql = strSql & " and ax309 = '" & Text1 & "'"
         End If
         If Text12 <> MsgText(601) Then
            strSql = strSql & " and ax314 = '" & Text12 & "'"
         End If
         If Text13 <> MsgText(601) Then
            strSql = strSql & " and ax313 = '" & Text13 & "'"
         End If
         '20140124START Modify By eric
         adoadodc1.Open "select ax301 as ax201, ax308 as ax208, ax309 as ax209, ax314 as ax214, a0305 as a0205, ax302 as ax202, ax306 as ax206, ax307 as ax207, ax312 as ax212, cu04 as cust, sales.st02 as salesname, ax303 as ax203, ax313 as ax213 from acc031, acc030, customer, staff sales where ax301 = a0301 and ax302 = a0302 and substr(ax308, 1, 1) = 'X' and substr(ax308, 1, 8) = cu01 and ax309 = sales.st01 (+)" & strSql & _
                        " union select ax301 as ax201, ax308 as ax208, ax309 as ax209, ax314 as ax214, a0305 as a0205, ax302 as ax202, ax306 as ax206, ax307 as ax207, ax312 as ax212, fa04 as cust, sales.st02 as salesname, ax303 as ax203, ax313 as ax213 from acc031, acc030, fagent, staff sales where ax301 = a0301 and ax302 = a0302 and substr(ax308, 1, 1) = 'Y' and substr(ax308, 1, 8) = fa01 and ax309 = sales.st01 (+)" & strSql & _
                        " union select ax301 as ax201, ax308 as ax208, ax309 as ax209, ax314 as ax214, a0305 as a0205, ax302 as ax202, ax306 as ax206, ax307 as ax207, ax312 as ax212, a0i02 as cust, sales.st02 as salesname, ax303 as ax203, ax313 as ax213 from acc031, acc030, acc0i0, staff sales where ax301 = a0301 and ax302 = a0302 and ax308 = a0i01 and ax309 = sales.st01 (+)" & strSql & _
                        " union select ax301 as ax201, ax308 as ax208, ax309 as ax209, ax314 as ax214, a0305 as a0205, ax302 as ax202, ax306 as ax206, ax307 as ax207, ax312 as ax212, staff.st02 as cust, sales.st02 as salesname, ax303 as ax203, ax313 as ax213 from acc031, acc030, staff, staff sales where ax301 = a0301 and ax302 = a0302 and ax308 = staff.st01 and ax309 = sales.st01 (+)" & strSql & _
                        " union select ax301 as ax201, ax308 as ax208, ax309 as ax209, ax314 as ax214, a0305 as a0205, ax302 as ax202, ax306 as ax206, ax307 as ax207, ax312 as ax212, '' as cust, sales.st02 as salesname, ax303 as ax203, ax313 as ax213 from acc031, acc030, staff sales where ax301 = a0301 and ax302 = a0302 and ax308 is null and ax309 = sales.st01 (+)" & strSql & _
                        " order by a0205 asc, ax201 asc, ax202 asc, ax203 asc", adoTaie, adOpenStatic, adLockReadOnly
         'adoadodc1.Open "select ax308 as ax208, ax309 as ax209, ax314 as ax214, a0305 as a0205, ax302 as ax202, ax306 as ax206, ax307 as ax207, ax312 as ax212, cu04 as cust, sales.st02 as salesname, ax303 as ax203, ax313 as ax213 from acc031, acc030, customer, staff sales where ax301 = a0301 and ax302 = a0302 and substr(ax308, 1, 1) = 'X' and substr(ax308, 1, 8) = cu01 and ax309 = sales.st01 (+)" & strSql & " union " & _
         '               "select ax308 as ax208, ax309 as ax209, ax314 as ax214, a0305 as a0205, ax302 as ax202, ax306 as ax206, ax307 as ax207, ax312 as ax212, fa04 as cust, sales.st02 as salesname, ax303 as ax203, ax313 as ax213 from acc031, acc030, fagent, staff sales where ax301 = a0301 and ax302 = a0302 and substr(ax308, 1, 1) = 'Y' and substr(ax308, 1, 8) = fa01 and ax309 = sales.st01 (+)" & strSql & _
         '               " union select ax308 as ax208, ax309 as ax209, ax314 as ax214, a0305 as a0205, ax302 as ax202, ax306 as ax206, ax307 as ax207, ax312 as ax212, a0i02 as cust, sales.st02 as salesname, ax303 as ax203, ax313 as ax213 from acc031, acc030, acc0i0, staff sales where ax301 = a0301 and ax302 = a0302 and ax308 = a0i01 and ax309 = sales.st01 (+)" & strSql & _
         '               " union select ax308 as ax208, ax309 as ax209, ax314 as ax214, a0305 as a0205, ax302 as ax202, ax306 as ax206, ax307 as ax207, ax312 as ax212, staff.st02 as cust, sales.st02 as salesname, ax303 as ax203, ax313 as ax213 from acc031, acc030, staff, staff sales where ax301 = a0301 and ax302 = a0302 and ax308 = staff.st01 and ax309 = sales.st01 (+)" & strSql & _
         '               " union select ax308 as ax208, ax309 as ax209, ax314 as ax214, a0305 as a0205, ax302 as ax202, ax306 as ax206, ax307 as ax207, ax312 as ax212, '' as cust, sales.st02 as salesname, ax303 as ax203, ax313 as ax213 from acc031, acc030, staff sales where ax301 = a0301 and ax302 = a0302 and ax308 is null and ax309 = sales.st01 (+)" & strSql & _
         '               " order by ax202 asc", adoTaie, adOpenStatic, adLockReadOnly
         '20140124END
                        
         adoaccsum.Open "select sum(ax306), sum(ax307) from acc031, acc030 where ax301 = a0301 and ax302 = a0302" & strSql, adoTaie, adOpenStatic, adLockReadOnly
      Case Else
         'Modify by Amy 2020/03/31 改下拉
         'If Text5 <> MsgText(601) Then
         If Trim(Combo1) <> MsgText(601) Then
            '2005/11/23 MODIFY BY SONIA
            'strSQL = " and ax201 = '" & Text5 & "'"
            strSql = " and ax201||'' = '" & Combo1 & "'"
         End If
         If Text9 <> MsgText(601) Then
            strSql = strSql & " and ax204 = '" & Text9 & "'"
         Else
      '      strSQL = strSQL & " and (ax204 = '" & MsgText(55) & "' or ax204 = '' or ax204 is null)"
         End If
         'Add By Sindy 2010/3/9
         If Text15 <> MsgText(601) Then
            strSql = strSql & " and INSTR(AX212,'" & Text15 & "') > 0 "
         End If
         '2010/3/9 End
         If Text7 <> MsgText(601) Then
            '2005/11/23 MODIFY BY SONIA
            'strSQL = strSQL & " and ax205 = '" & Text7 & "'"
            strSql = strSql & " and ax205||'' = '" & Text7 & "'"
         End If
         If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
            strSql = strSql & " and a0205 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
         End If
         If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
            strSql = strSql & " and a0205 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
         End If
         
         'Modify by Morgan 2003/11/27
'         If Text2 <> MsgText(601) Then
'            strSQL = strSQL & " and ax208 = '" & Text2 & "'"
'         End If

         If Text2 <> MsgText(601) Then
            strSql = strSql & " and ax208 >= '" & Text2 & "'"
         End If
         
         If Text14 <> MsgText(601) Then
            strSql = strSql & " and ax208 <= '" & Text14 & "'"
         End If
         
         'End
         
         If Text1 <> MsgText(601) Then
            strSql = strSql & " and ax209 = '" & Text1 & "'"
         End If
         If Text12 <> MsgText(601) Then
            '2005/11/23 MODIFY BY SONIA
            'strSQL = strSQL & " and ax214 = '" & Text12 & "'"
            '2010/4/9 MODIFY BY SONIA TF母案領土延伸分開算,CFP僅EPC子案與母案一起算,接續案或集體設計個別算
            'If Mid(Text12, 1, 2) = "TF" Then
            '   strSQL = strSQL & " and AX214 >= '" & Mid(Text12, 1, Len(Text12) - 4) & "0000' AND AX214 <= '" & Mid(Text12, 1, Len(Text12) - 4) & "9999' "
            If Mid(Text12, 1, 2) = "TF" Then
               strSql = strSql & " and AX214 >= '" & Mid(Text12, 1, Len(Text12) - 3) & "000' AND AX214 <= '" & Mid(Text12, 1, Len(Text12) - 3) & "999' "
            Else
               strSql = strSql & " and AX214 >= '" & Mid(Text12, 1, Len(Text12) - 2) & "00' AND AX214 <= '" & Mid(Text12, 1, Len(Text12) - 2) & "99' "
            End If
            '2005/11/23 END
         End If
         If Text13 <> MsgText(601) Then
            strSql = strSql & " and ax213 = '" & Text13 & "'"
         End If
         
        '20140124START Modify By eric
        'Modify By Cheng 2004/04/14
'         adoadodc1.Open "select ax205, ax208, ax209, ax214, a0205, ax202, ax206, ax207, ax212, cu04 as cust, sales.st02 as salesname, ax203, ax213 from acc021, acc020, customer, staff sales where ax201 = a0201 and ax202 = a0202 and substr(ax208, 1, 8) = cu01(+) and substr(ax208, 9, 1) = cu02(+) and ax209 = sales.st01 (+)" & strSQL & " union " & _
'                        "select ax205, ax208, ax209, ax214, a0205, ax202, ax206, ax207, ax212, fa04 as cust, sales.st02 as salesname, ax203, ax213 from acc021, acc020, fagent, staff sales where ax201 = a0201 and ax202 = a0202 and substr(ax208, 1, 8) = fa01(+) and substr(ax208, 9, 1) = fa02(+) and ax209 = sales.st01 (+)" & strSQL & _
'                        " union select ax205, ax208, ax209, ax214, a0205, ax202, ax206, ax207, ax212, a0i02 as cust, sales.st02 as salesname, ax203, ax213 from acc021, acc020, acc0i0, staff sales where ax201 = a0201 and ax202 = a0202 and ax208 = a0i01 and ax209 = sales.st01 (+) " & strSQL & _
'                        " union select ax205, ax208, ax209, ax214, a0205, ax202, ax206, ax207, ax212, staff.st02 as cust, sales.st02 as salesname, ax203, ax213 from acc021, acc020, staff, staff sales where ax201 = a0201 and ax202 = a0202 and ax208 = staff.st01 and ax209 = sales.st01 (+)" & strSQL & _
'                        " union select ax205, ax208, ax209, ax214, a0205, ax202, ax206, ax207, ax212, '' as cust, sales.st02 as salesname, ax203, ax213 from acc021, acc020, staff sales where ax201 = a0201 and ax202 = a0202 and ax208 is null and ax209 = sales.st01 (+)" & strSQL & _
'                        " order by a0205 asc, ax202 asc", adoTaie, adOpenStatic, adLockReadOnly
         '2007/9/4 modify by sonia D096084565 2401 因F5584同時存在於廠商檔及員工檔會造成資料重覆
         'strSQLA = "select ax205, ax208, ax209, ax214, a0205, ax202, ax206, ax207, ax212, cu04 as cust, sales.st02 as salesname, ax203, ax213 from acc021, acc020, customer, staff sales where ax201 = a0201 and ax202 = a0202 and substr(ax208, 1, 8) = cu01(+) and substr(ax208, 9, 1) = cu02(+) and ax209 = sales.st01 (+)" & strSQL & " And Decode(Substr(AX208,1,1),'X',AX208, NULL, AX208,'Y')=AX208 " &
         '               " union select ax205, ax208, ax209, ax214, a0205, ax202, ax206, ax207, ax212, fa04 as cust, sales.st02 as salesname, ax203, ax213 from acc021, acc020, fagent, staff sales where ax201 = a0201 and ax202 = a0202 and substr(ax208, 1, 8) = fa01(+) and substr(ax208, 9, 1) = fa02(+) and ax209 = sales.st01 (+)" & strSQL & "And Decode(Substr(AX208,1,1),'Y',AX208, NULL, AX208,'X')=AX208 " & _
         '               " union select ax205, ax208, ax209, ax214, a0205, ax202, ax206, ax207, ax212, a0i02 as cust, sales.st02 as salesname, ax203, ax213 from acc021, acc020, acc0i0, staff sales where ax201 = a0201 and ax202 = a0202 and ax208 = a0i01 and ax209 = sales.st01 (+) " & strSQL & _
         '               " union select ax205, ax208, ax209, ax214, a0205, ax202, ax206, ax207, ax212, staff.st02 as cust, sales.st02 as salesname, ax203, ax213 from acc021, acc020, staff, staff sales where ax201 = a0201 and ax202 = a0202 and ax208 = staff.st01 and ax209 = sales.st01 (+)" & strSQL & _
         '               " union select ax205, ax208, ax209, ax214, a0205, ax202, ax206, ax207, ax212, '' as cust, sales.st02 as salesname, ax203, ax213 from acc021, acc020, staff sales where ax201 = a0201 and ax202 = a0202 and ax208 is null and ax209 = sales.st01 (+)" & strSQL & _
         '               " order by a0205 asc, ax202 asc"
         '2010/4/2 MODIFY BY SONIA 加排序條件AX203 asc
     '    StrSQLa = "select ax205, ax208, ax209, ax214, a0205, ax202, ax206, ax207, ax212, cu04 as cust, sales.st02 as salesname, ax203, ax213 from acc021, acc020, customer, staff sales where ax201 = a0201 and ax202 = a0202 and substr(ax208, 1, 8) = cu01(+) and substr(ax208, 9, 1) = cu02(+) and ax209 = sales.st01 (+)" & strSql & " And Decode(Substr(AX208,1,1),'X',AX208, NULL, AX208,'Y')=AX208 " & _
     '                   " union select ax205, ax208, ax209, ax214, a0205, ax202, ax206, ax207, ax212, fa04 as cust, sales.st02 as salesname, ax203, ax213 from acc021, acc020, fagent, staff sales where ax201 = a0201 and ax202 = a0202 and substr(ax208, 1, 8) = fa01(+) and substr(ax208, 9, 1) = fa02(+) and ax209 = sales.st01 (+)" & strSql & "And Decode(Substr(AX208,1,1),'Y',AX208, NULL, AX208,'X')=AX208 " & _
     '                   " union select ax205, ax208, ax209, ax214, a0205, ax202, ax206, ax207, ax212, a0i02 as cust, sales.st02 as salesname, ax203, ax213 from acc021, acc020, acc0i0, staff sales where ax201 = a0201 and ax202 = a0202 and ax208 = a0i01 and ax209 = sales.st01 (+) " & strSql & " And Substr(AX208,1,1)='V' " & _
     '                   " union select ax205, ax208, ax209, ax214, a0205, ax202, ax206, ax207, ax212, staff.st02 as cust, sales.st02 as salesname, ax203, ax213 from acc021, acc020, staff, staff sales where ax201 = a0201 and ax202 = a0202 and ax208 = staff.st01 and ax209 = sales.st01 (+)" & strSql & " And Substr(AX208,1,1)='F' " & _
     '                   " union select ax205, ax208, ax209, ax214, a0205, ax202, ax206, ax207, ax212, '' as cust, sales.st02 as salesname, ax203, ax213 from acc021, acc020, staff sales where ax201 = a0201 and ax202 = a0202 and ax208 is null and ax209 = sales.st01 (+)" & strSql & _
     '                   " order by a0205 asc, ax202 asc, ax203 asc"
         '2007/9/4 end
         StrSQLa = "select ax201, ax205, ax208, ax209, ax214, a0205, ax202, ax206, ax207, ax212, cu04 as cust, sales.st02 as salesname, ax203, ax213 from acc021, acc020, customer, staff sales where ax201 = a0201 and ax202 = a0202 and substr(ax208, 1, 8) = cu01(+) and substr(ax208, 9, 1) = cu02(+) and ax209 = sales.st01 (+)" & strSql & " And Decode(Substr(AX208,1,1),'X',AX208, NULL, AX208,'Y')=AX208 " & _
                        " union select ax201, ax205, ax208, ax209, ax214, a0205, ax202, ax206, ax207, ax212, fa04 as cust, sales.st02 as salesname, ax203, ax213 from acc021, acc020, fagent, staff sales where ax201 = a0201 and ax202 = a0202 and substr(ax208, 1, 8) = fa01(+) and substr(ax208, 9, 1) = fa02(+) and ax209 = sales.st01 (+)" & strSql & " And Decode(Substr(AX208,1,1),'Y',AX208, NULL, AX208,'X')=AX208 " & _
                        " union select ax201, ax205, ax208, ax209, ax214, a0205, ax202, ax206, ax207, ax212, a0i02 as cust, sales.st02 as salesname, ax203, ax213 from acc021, acc020, acc0i0, staff sales where ax201 = a0201 and ax202 = a0202 and ax208 = a0i01 and ax209 = sales.st01 (+) " & strSql & " And Substr(AX208,1,1)='V' " & _
                        " union select ax201, ax205, ax208, ax209, ax214, a0205, ax202, ax206, ax207, ax212, staff.st02 as cust, sales.st02 as salesname, ax203, ax213 from acc021, acc020, staff, staff sales where ax201 = a0201 and ax202 = a0202 and ax208 = staff.st01 and ax209 = sales.st01 (+)" & strSql & " And Substr(AX208,1,1)='F' " & _
                        " union select ax201, ax205, ax208, ax209, ax214, a0205, ax202, ax206, ax207, ax212, '' as cust, sales.st02 as salesname, ax203, ax213 from acc021, acc020, staff sales where ax201 = a0201 and ax202 = a0202 and ax208 is null and ax209 = sales.st01 (+)" & strSql & _
                        " order by a0205 asc, ax201 asc, ax202 asc, ax203 asc"
         '20140124END
         
         adoadodc1.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
         adoaccsum.Open "select sum(ax206), sum(ax207) from acc021, acc020 where ax201 = a0201 and ax202 = a0202" & strSql, adoTaie, adOpenStatic, adLockReadOnly
   End Select
   Adodc1.Recordset.Requery
   
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0).Value) Then
         Text3 = MsgText(601)
         douDebit = 0
      Else
         Text3 = Format(adoaccsum.Fields(0).Value, FDollar)
         douDebit = Val(adoaccsum.Fields(0).Value)
      End If
      If IsNull(adoaccsum.Fields(1).Value) Then
         Text4 = MsgText(601)
         douCredit = 0
      Else
         Text4 = Format(adoaccsum.Fields(1).Value, FDollar)
         douCredit = Val(adoaccsum.Fields(1).Value)
      End If
      '2005/12/16 MODIFY BY SONIA
      'Text11 = Format(douDebit - douCredit, FDollar)
      If GetDebitCredit(Text7) = "1" Then
         Text11 = Format(douDebit - douCredit, FDollar)
      Else
         Text11 = Format(douCredit - douDebit, FDollar)
      End If
   Else
      Text3 = MsgText(601)
      Text4 = MsgText(601)
      Text11 = MsgText(601)
   End If
   adoaccsum.Close
   If Adodc1.Recordset.RecordCount = 0 Then
      Adodc1.Recordset.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   '2014/3/3 ADD BY SONIA 辜說要直接跳到最後面
   Else
      Adodc1.Recordset.MoveLast
   '2014/3/3 END
   End If
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

Private Sub Text7_Change()
   If Text7 = MsgText(601) Then
      Exit Sub
   End If
   Text8 = A0102Query(Text7)
End Sub

Private Sub Text7_GotFocus()
   TextInverse Text7
End Sub

Private Sub Text9_Change()
   If Text9 = MsgText(601) Then
      Exit Sub
   End If
   Text10 = A0902Query(Text9)
End Sub

Private Sub Text9_GotFocus()
   TextInverse Text9
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Public Sub KeyDefine(KeyCode As Integer)
   Call PUB_SaveTrackMode(1, KeyCode) 'Add by Amy 2021/12/13 Form2.0
    
    'Add by Amy 2021/12/13 Form2.0控制Function鍵：記錄鍵盤傳入順序，判斷是否可執行
    If PUB_ChkTrackMode = False Then
        Exit Sub
    End If
   'end 2021/12/13
 
   Select Case KeyCode
      Case vbKeyF12
         If FormCheck Then
            Screen.MousePointer = vbHourglass
            Text12 = CaseNoZero(Text12)
            QueryTable
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
   'Add by Amy 200/03/31 +公司別判斷
   Dim bolCancel As Boolean
   
   If Trim(Combo1) <> MsgText(601) Then
      Call Combo1_Validate(bolCancel)
      If bolCancel = False Then
        FormCheck = True
        Exit Function
      End If
   End If
   'end 2020/03/31
   If Text9 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Text7 <> MsgText(601) Then
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
   If Text2 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Text1 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Text12 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Text13 <> MsgText(601) Then
      'edit by nickc 2007/02/08
      'formchek = True
      FormCheck = True
      Exit Function
   End If
   FormCheck = False
End Function

'Add By Sindy 2010/3/9
Private Sub Text15_GotFocus()
   TextInverse Text15
End Sub


