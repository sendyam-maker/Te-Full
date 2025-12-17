VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc41f0 
   AutoRedraw      =   -1  'True
   Caption         =   "結餘保留放出產生傳票"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9405
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5115
   ScaleWidth      =   9405
   Begin VB.CommandButton Command1 
      Caption         =   "產生傳票"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7680
      TabIndex        =   9
      Top             =   120
      Width           =   1350
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5400
      TabIndex        =   3
      Top             =   510
      Width           =   495
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1335
      TabIndex        =   0
      Top             =   135
      Width           =   612
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc41f0.frx":0000
      Height          =   3250
      Left            =   240
      TabIndex        =   10
      Top             =   1380
      Width           =   9000
      _ExtentX        =   15875
      _ExtentY        =   5741
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
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
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "智權人員結餘點數資料"
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "Name"
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
      BeginProperty Column01 
         DataField       =   "T1"
         Caption         =   "大陸商標"
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
      BeginProperty Column02 
         DataField       =   "P1"
         Caption         =   "大陸專利"
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
      BeginProperty Column03 
         DataField       =   "T2"
         Caption         =   "國外商標"
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
         DataField       =   "P2"
         Caption         =   "國外專利"
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
         DataField       =   "CFL"
         Caption         =   "ＣＦＬ"
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
      BeginProperty Column06 
         DataField       =   "TOTAL"
         Caption         =   "合計"
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
      BeginProperty Column07 
         DataField       =   "Dept"
         Caption         =   "區別"
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
         DataField       =   "ID"
         Caption         =   "員編"
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
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column01 
            Alignment       =   1
            ColumnWidth     =   1184.882
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1244.976
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1230.236
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1124.787
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column07 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column08 
            Object.Visible         =   0   'False
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
      Left            =   240
      Top             =   1560
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   556
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
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1320
      TabIndex        =   1
      Top             =   510
      Width           =   1335
      _ExtentX        =   2355
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
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   300
      Left            =   3030
      TabIndex        =   2
      Top             =   510
      Width           =   1335
      _ExtentX        =   2355
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
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "Frmacc41f0.frx":0015
      Height          =   405
      Left            =   240
      TabIndex        =   12
      Top             =   4680
      Width           =   9000
      _ExtentX        =   15875
      _ExtentY        =   714
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      BackColor       =   -2147483624
      ColumnHeaders   =   0   'False
      HeadLines       =   1
      RowHeight       =   20
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
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "Name"
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
      BeginProperty Column01 
         DataField       =   "T1"
         Caption         =   "大陸商標"
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
      BeginProperty Column02 
         DataField       =   "P1"
         Caption         =   "大陸專利"
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
      BeginProperty Column03 
         DataField       =   "T2"
         Caption         =   "國外商標"
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
         DataField       =   "P2"
         Caption         =   "國外專利"
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
         DataField       =   "CFL"
         Caption         =   "ＣＦＬ"
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
      BeginProperty Column06 
         DataField       =   "TOTAL"
         Caption         =   "合計"
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
      BeginProperty Column07 
         DataField       =   "Dept"
         Caption         =   "區別"
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
         DataField       =   "ID"
         Caption         =   "員編"
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
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column01 
            Alignment       =   1
            ColumnWidth     =   1184.882
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1244.976
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1230.236
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1124.787
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column07 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column08 
            Object.Visible         =   0   'False
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   315
      Left            =   120
      Top             =   4440
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
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "傳票迄日為空：預設今天"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   390
      TabIndex        =   14
      Top             =   1110
      Width           =   2505
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "傳票起日為空：預設1年前的今天"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   390
      TabIndex        =   13
      Top             =   870
      Width           =   3200
   End
   Begin MSForms.TextBox Text5 
      Height          =   300
      Left            =   1950
      TabIndex        =   4
      Top             =   135
      Width           =   2415
      VariousPropertyBits=   679493661
      BackColor       =   14737632
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "PS: 不在空白表格名單內之人員, 請產生傳票後自行處理"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   3330
      TabIndex        =   11
      Top             =   1050
      Width           =   6000
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Top             =   4680
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "資料別           (1日期區間餘額 2空白)"
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
      Left            =   4560
      TabIndex        =   8
      Top             =   510
      Width           =   4335
   End
   Begin VB.Label Label3 
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
      Left            =   360
      TabIndex        =   7
      Top             =   135
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "傳票日期"
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
      Left            =   360
      TabIndex        =   6
      Top             =   495
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
      Left            =   2760
      TabIndex        =   5
      Top             =   495
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   1305
      Left            =   240
      Top             =   45
      Width           =   9000
   End
End
Attribute VB_Name = "Frmacc41f0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/11/01 Form2.0已修改 Text5/DataGrid1
'Create by Amy  2013/07/23
Option Explicit

Public adoadodc1 As New ADODB.Recordset
Public adoSum As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
Dim strSql As String
Dim i As Integer
Dim IsFirst As Boolean '是否按下F12查詢
Dim dgRow As Integer '記錄目前更新那筆的Datagrid 的row
Dim dgFirstRow As Integer '記錄目前更新那筆Datagrid的第一筆
Dim RsSum As New ADODB.Recordset

Private Sub Command1_Click()
   'Add by Amy 2013/12/12
   If FormCheck = False Then
      Exit Sub
   End If
   'end 2013/12/12
   If adoquery.State = adStateOpen Then
      adoquery.Close
   End If
   adoquery.CursorLocation = adUseClient
   '2014/1/22 modify by sonia 加公司別
   'adoquery.Open "Select a0b10 From acc0b0 Where a0b10 = '01'", adoTaie, adOpenStatic, adLockReadOnly
   adoquery.Open "Select a0b10 From acc0b0 Where a0b04 = '" & Text4 & "' and a0b10 = '01'", adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      MsgBox MsgText(197), , MsgText(5)
      adoquery.Close
      Exit Sub
   End If
   adoquery.Close
   '2014/1/22 modify by sonia 加公司別
   'adoTaie.Execute "update acc0b0 set a0b10 = '01'"
   adoTaie.Execute "update acc0b0 set a0b10 = '01' Where a0b04 = '" & Text4 & "'"
   Screen.MousePointer = vbHourglass
   Command1.Enabled = False
   IsFirst = False
   TransferTable
   Screen.MousePointer = vbDefault
   '2014/1/22 modify by sonia 加公司別
   'adoTaie.Execute "update acc0b0 set a0b10 = null"
   adoTaie.Execute "update acc0b0 set a0b10 = null Where a0b04 = '" & Text4 & "'"
 
End Sub

Private Sub DataGrid1_AfterColUpdate(ByVal ColIndex As Integer)
 Dim j As Integer
 
    If IsFirst Then
       Dim FieldName As String
       Dim strUpd As String
       Dim SumCol6 As Double
              
       If Adodc1.Recordset.RecordCount = 0 Then
            Exit Sub
       End If
       
       If Right(DataGrid1.Columns(7), 1) = "T" Or DataGrid1.Columns(7) = "ZZZ" Then
            Adodc1.Recordset.Requery
            DataGrid1.Scroll 0, dgFirstRow - 1
            DataGrid1.row = Val(dgRow)
            Exit Sub
       End If
       
       FieldName = ""
      
       With DataGrid1
          .Columns(6).Text = CDbl(IIf(.Columns(1).Text = "", "0", .Columns(1))) + CDbl(IIf(.Columns(2).Text = "", "0", .Columns(2))) + _
          CDbl(IIf(.Columns(3).Text = "", "0", .Columns(3))) + CDbl(IIf(.Columns(4).Text = "", "0", .Columns(4))) + CDbl(IIf(.Columns(5).Text = "", "0", .Columns(5)))
          
          Adodc1.Recordset.UpdateBatch
       
        Select Case ColIndex
            Case 1
                FieldName = "R42304"
            Case 2
                FieldName = "R42305"
            Case 3
                FieldName = "R42306"
            Case 4
                FieldName = "R42307"
            Case 5
                FieldName = "R42308"
       End Select
       
       If Left(.Columns(7), 1) <> "X" Then
            strSql = .Columns(7)
       Else
            strSql = Left(.Columns(7), 2)
       End If
   
        strExc(0) = "Select * From (Select NVL(sum(" & FieldName & "),0) as T,NVL(sum(R42309),0) as SumT From accrpt423 where R42302='" & .Columns(7) & "' And R42301='" & strUserNum & "')," & _
                          "(Select NVL(sum(" & FieldName & "),0) as SX,NVL(sum(R42309),0) as SumSX From accrpt423 where substr(R42302,1,2)='" & Left(.Columns(7), 2) & "' And SubStr(R42302,length(R42302))<>'T'  And R42301='" & strUserNum & "')," & _
                          "(Select NVL(sum(" & FieldName & "),0) as Total,NVL(sum(R42309),0) as SumTot From accrpt423 Where SubStr(R42302,length(R42302))<>'T' And R42302<>'ZZZ' And R42301='" & strUserNum & "') "
       
       If adoSum.State = adStateOpen Then adoSum.Close
       adoSum.CursorLocation = adUseClient
       adoSum.Open strExc(0), adoTaie, adOpenStatic, adLockReadOnly
        
       If Left(.Columns(7), 1) <> "X" Then
            '更新區合計
             strUpd = "Update accrpt423 set " & FieldName & "=" & Val(adoSum.Fields("T")) & ",R42309=" & Val(adoSum.Fields("SumT")) & " Where R42302='" & .Columns(7) & "T'  And R42301='" & strUserNum & "' "
             adoTaie.Execute strUpd
       End If
       
       '更新所(或其他)合計
        strUpd = "Update accrpt423 set " & FieldName & "=" & Val(adoSum.Fields("SX")) & ",R42309=" & Val(adoSum.Fields("SumSX")) & " Where SubStr(R42302,1,3)='" & Left(.Columns(7), 2) & "T'  And R42301='" & strUserNum & "' "
        adoTaie.Execute strUpd
        
       '更新總合計
        strUpd = "Update accrpt423 set " & FieldName & "=" & Val(adoSum.Fields("Total")) & ",R42309=" & Val(adoSum.Fields("SumTot")) & " Where R42302='ZZZ'  And R42301='" & strUserNum & "' "
        adoTaie.Execute strUpd
      
      Adodc1.Recordset.Requery
      
      
      .Scroll 0, dgFirstRow - 1 '設定更新那筆的第一筆
      .row = Val(dgRow) '跳至更新的那一筆
      Command1.Enabled = True
     
      End With
      
    Else
        DataGrid1.Columns(ColIndex).Text = ""
    End If
    Call SumShow 'Add by Amy 2019/01/14
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
   Dim intCounter As Integer
   
   dgRow = DataGrid1.row
   dgFirstRow = DataGrid1.FirstRow
   
   Select Case KeyCode
      Case vbKeyReturn
         Select Case DataGrid1.col
            Case 1
               SendKeys "{RIGHT}"
            Case 2
               SendKeys "{RIGHT}"
            Case 3
               SendKeys "{RIGHT}"
            Case 4
               SendKeys "{RIGHT}"
            Case 5
               SendKeys "{DOWN}"
               For intCounter = 1 To 4
                  SendKeys "{LEFT}"
               Next intCounter
         End Select
   End Select
End Sub

Private Sub Form_Activate()
    'Modify by Amy 2013/12/12
    'MaskEdBox1.SetFocus
    Text4.SetFocus
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
   Me.Width = 9500
   Me.Height = 5500
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
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
   'Modify by Amy 2013/12/12 要可輸入J公司
   'Text4 = "1"
   'Text4.Enabled = False '目前只有1公司,且程式也沒有過濾
   'end 2013/12/12

   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   StatusClear
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set adoSum = Nothing
   Set Frmacc41f0 = Nothing
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoadodc1.CursorLocation = adUseClient
   strExc(0) = "select * from acc021, acc010, acc020 where acc021.ax205 = acc010.a0101 and acc021.ax201 = acc020.a0201 and acc021.ax202 = acc020.a0202 and ax201 = '" & Text4 & "' and a0205 >= " & Val(FCDate(MaskEdBox1.Text)) & " and a0205 <= " & Val(FCDate(MaskEdBox2.Text)) & " order by a0205 desc, ax202 asc, ax203 asc"
   adoadodc1.Open strExc(0), adoTaie, adOpenStatic, adLockReadOnly

   Set Adodc1.Recordset = adoadodc1
   
   'Add by Amy 2019/01/14 增加合計
   If RsSum.State = adStateOpen Then RsSum.Close
   adoSum.CursorLocation = adUseClient
   strExc(0) = "Select '' as Dept,'全所合計' as Name,'' T1,'' P1,'' T2,'' P2,'' CFL,'' TOTAL,'' ID From Dual "
   adoSum.Open strExc(0), adoTaie, adOpenStatic, adLockReadOnly
   Set Adodc2.Recordset = adoSum
   'end 2019/01/14
    
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  查詢資料表(傳票資料)
'
'*************************************************
Public Sub QueryTable()
Dim StrSQLa As String
On Error GoTo Checking
   strSql = ""
   If adoadodc1.State = adStateOpen Then
      adoadodc1.Close
   End If
   adoadodc1.CursorLocation = adUseClient
 
   'Modify by Amy 2015/01/05 改抓1年傳票有資料的名單資料 原:Left(strSrvDate(2), 3) & "/01/01"
   If Text1 = "2" And (MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29)) Then
      MaskEdBox1.Text = ChangeTStringToTDateString(strSrvDate(1) - 19120000)
   End If
   If Text1 = "2" And (MaskEdBox2.Text = MsgText(601) Or MaskEdBox2.Text = MsgText(29)) Then
      MaskEdBox2.Text = ChangeTStringToTDateString(strSrvDate(2))
   End If
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
       strSql = strSql & " and a0205 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
   End If
   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
        strSql = strSql & " and a0205 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
   End If
   
   StrSQLa = " Having (Nvl(sum(decode(ax205, '249101', ax207-ax206, 0)),0)<>0 Or Nvl(sum(decode(ax205, '249102', ax207-ax206 , 0)),0)<>0 Or Nvl(sum(decode(ax205, '249103', ax207-ax206 , 0)),0)<>0 Or Nvl(sum(decode(ax205, '249104', ax207-ax206, 0)),0)<>0 Or Nvl(sum(decode(ax205, '249105', ax207-ax206, 0)),0)<>0) "
   'Modify by Amy 2014/01/10 +公司別 ax201= '" & Text4 & "'
   'Modify by Amy 2015/01/08 修正資料抓太慢 ax201= '" & Text4 & "' 改抓a0201
   If Text1 = "2" Then
       strSql = "select st15 as Dept, st01 as ID, st02 as Name, '' as T1, '' as P1, '' as T2, '' as P2, '' as CFL,'' as TOTAL from acc021, acc020, staff where a0201= '" & Text4 & "' And ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and substr(st15, 1, 1) = 'S' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & " group by st15, st01, st02" & StrSQLa & _
                   " union select a0901||'T' as Dept, a0901||'T' as ID, a0902 as Name, '' as T1, '' as P1, '' as T2, '' as P2,'' as CFL, '' as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where a0201= '" & Text4 & "' And ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and substr(st15, 1, 1) = 'S' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & " group by a0901, a0902" & StrSQLa & _
                   " union select 'S1T' as Dept, 'S1T' as ID, '北所' as Name, '' as T1, '' as P1, '' as T2, '' as P2, '' as CFL,'' as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where a0201= '" & Text4 & "' And ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and substr(st15, 1, 2) = 'S1' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & StrSQLa & _
                   " union select 'S2T' as Dept, 'S2T' as ID, '中所' as Name, '' as T1, '' as P1, '' as T2, '' as P2, '' as CFL,'' as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where a0201= '" & Text4 & "' And ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and substr(st15, 1, 2) = 'S2' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & StrSQLa & _
                   " union select 'S3T' as Dept, 'S3T' as ID, '南所' as Name, '' as T1, '' as P1, '' as T2, '' as P2, '' as CFL,'' as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where a0201= '" & Text4 & "' And ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and substr(st15, 1, 2) = 'S3' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & StrSQLa & _
                   " union select 'S4T' as Dept, 'S4T' as ID, '高所' as Name, '' as T1, '' as P1, '' as T2, '' as P2, '' as CFL,'' as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where a0201= '" & Text4 & "' And ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and substr(st15, 1, 2) = 'S4' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & StrSQLa & _
                   " union select 'S9T' as Dept, 'S9T' as ID, '廣東' as Name, '' as T1, '' as P1, '' as T2, '' as P2, '' as CFL,'' as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where a0201= '" & Text4 & "' And ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and substr(st15, 1, 2) = 'S9' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & StrSQLa & _
                   " union select 'X01' as Dept, st01 as ID, st02 as Name, '' as T1, '' as P1, '' as T2, '' as P2, '' as CFL,'' as TOTAL from acc021, acc020, staff where a0201= '" & Text4 & "' And ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and substr(st15, 1, 1) <> 'S' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & " group by st01, st02" & StrSQLa & _
                   " union select 'X0T' as Dept, 'X0T' as ID, '其他合計' as Name, '' as T1, '' as P1, '' as T2, '' as P2, '' as CFL,'' as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where a0201= '" & Text4 & "' And ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and substr(st15, 1, 1) <> 'S' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & StrSQLa & _
                   " union select 'ZZZ' as Dept, 'ZZZ' as ID, '全所合計' as Name, '' as T1, '' as P1, '' as T2, '' as P2, '' as CFL,'' as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where a0201= '" & Text4 & "' And ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & StrSQLa & " ORDER BY DEPT,ID"
   Else
       strSql = "select st15 as Dept, st02 as Name, sum(decode(ax205, '249101', ax207-ax206 , 0)) as T1, sum(decode(ax205, '249102', ax207-ax206 , 0)) as P1, sum(decode(ax205, '249103', ax207-ax206 , 0)) as T2, sum(decode(ax205, '249104', ax207-ax206 , 0)) as P2, sum(decode(ax205, '249105', ax207-ax206, 0)) as CFL, sum(ax207-ax206) as TOTAL, st01 as ID from acc021, acc020, staff where a0201= '" & Text4 & "' And ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and substr(st15, 1, 1) = 'S' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & " group by st15, st01, st02" & StrSQLa & _
                   " union select a0901||'T' as Dept, a0902 as Name, sum(decode(ax205, '249101', ax207-ax206 , 0)) as T1, sum(decode(ax205, '249102', ax207-ax206 , 0)) as P1, sum(decode(ax205, '249103', ax207-ax206 , 0)) as T2, sum(decode(ax205, '249104', ax207-ax206 , 0)) as P2, sum(decode(ax205, '249105', ax207-ax206 , 0)) as CFL, sum(ax207-ax206) as TOTAL, a0901||'T' as ID from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where a0201= '" & Text4 & "' And ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and substr(st15, 1, 1) = 'S' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & " group by a0901, a0902" & StrSQLa & _
                   " union select 'S1T' as Dept, '北所' as Name, sum(decode(ax205, '249101', ax207-ax206, 0)) as T1, sum(decode(ax205, '249102', ax207-ax206 , 0)) as P1, sum(decode(ax205, '249103', ax207-ax206 , 0)) as T2, sum(decode(ax205, '249104', ax207-ax206 , 0)) as P2, sum(decode(ax205, '249105', ax207-ax206 , 0)) as CFL, sum(ax207-ax206) as TOTAL, 'S1T' as ID from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where a0201= '" & Text4 & "' And ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and substr(st15, 1, 2) = 'S1' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & StrSQLa & _
                   " union select 'S2T' as Dept, '中所' as Name, sum(decode(ax205, '249101', ax207-ax206 , 0)) as T1, sum(decode(ax205, '249102', ax207-ax206 , 0)) as P1, sum(decode(ax205, '249103', ax207-ax206, 0)) as T2, sum(decode(ax205, '249104', ax207-ax206 , 0)) as P2, sum(decode(ax205, '249105', ax207-ax206 , 0)) as CFL, sum(ax207-ax206) as TOTAL, 'S2T' as ID from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where a0201= '" & Text4 & "' And ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and substr(st15, 1, 2) = 'S2' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & StrSQLa & _
                   " union select 'S3T' as Dept, '南所' as Name, sum(decode(ax205, '249101', ax207-ax206 , 0)) as T1, sum(decode(ax205, '249102', ax207-ax206 , 0)) as P1, sum(decode(ax205, '249103', ax207-ax206 , 0)) as T2, sum(decode(ax205, '249104', ax207-ax206 , 0)) as P2, sum(decode(ax205, '249105', ax207-ax206 , 0)) as CFL, sum(ax207-ax206 ) as TOTAL, 'S3T' as ID from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where a0201= '" & Text4 & "' And ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and substr(st15, 1, 2) = 'S3' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & StrSQLa & _
                   " union select 'S4T' as Dept, '高所' as Name, sum(decode(ax205, '249101', ax207-ax206 , 0)) as T1, sum(decode(ax205, '249102', ax207-ax206 , 0)) as P1, sum(decode(ax205, '249103', ax207-ax206 , 0)) as T2, sum(decode(ax205, '249104', ax207-ax206 , 0)) as P2, sum(decode(ax205, '249105', ax207-ax206 , 0)) as CFL, sum(ax207-ax206) as TOTAL, 'S4T' as ID from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where a0201= '" & Text4 & "' And ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and substr(st15, 1, 2) = 'S4' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & StrSQLa & _
                   " union select 'S9T' as Dept, '廣東' as Name, sum(decode(ax205, '249101', ax207-ax206 , 0)) as T1, sum(decode(ax205, '249102', ax207-ax206 , 0)) as P1, sum(decode(ax205, '249103', ax207-ax206 , 0)) as T2, sum(decode(ax205, '249104', ax207-ax206 , 0)) as P2, sum(decode(ax205, '249105', ax207-ax206 , 0)) as CFL, sum(ax207-ax206) as TOTAL, 'S9T' as ID from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where a0201= '" & Text4 & "' And ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and substr(st15, 1, 2) = 'S9' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & StrSQLa & _
                   " union select 'X01' as Dept, st02 as Name, sum(decode(ax205, '249101', ax207-ax206 , 0)) as T1, sum(decode(ax205, '249102', ax207-ax206, 0)) as P1, sum(decode(ax205, '249103', ax207-ax206 , 0)) as T2, sum(decode(ax205, '249104', ax207-ax206 , 0)) as P2, sum(decode(ax205, '249105', ax207-ax206 , 0)) as CFL, sum(ax207-ax206) as TOTAL, st01 as ID from acc021, acc020, staff where a0201= '" & Text4 & "' And ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and substr(st15, 1, 1) <> 'S' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & " group by st01, st02" & StrSQLa & _
                   " union select 'X0T' as Dept, '其他合計' as Name, sum(decode(ax205, '249101', ax207-ax206 , 0)) as T1, sum(decode(ax205, '249102', ax207-ax206 , 0)) as P1, sum(decode(ax205, '249103', ax207-ax206 , 0)) as T2, sum(decode(ax205, '249104', ax207-ax206, 0)) as P2, sum(decode(ax205, '249105', ax207-ax206 , 0)) as CFL, sum(ax207-ax206) as TOTAL, 'X0T' as ID from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where a0201= '" & Text4 & "' And ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and substr(st15, 1, 1) <> 'S' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & StrSQLa & _
                   " union select 'ZZZ' as Dept, '全所合計' as Name, sum(decode(ax205, '249101', ax207-ax206 , 0)) as T1, sum(decode(ax205, '249102', ax207-ax206 , 0)) as P1, sum(decode(ax205, '249103', ax207-ax206 , 0)) as T2, sum(decode(ax205, '249104', ax207-ax206 , 0)) as P2, sum(decode(ax205, '249105', ax207-ax206 , 0)) as CFL, sum(ax207-ax206) as TOTAL, 'ZZZ' as ID from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where a0201= '" & Text4 & "' And ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & StrSQLa & " ORDER BY DEPT,ID"
   End If
   'end 2014/01/10
   adoadodc1.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
   
   adoTaie.Execute "delete from accrpt423 Where R42301='" & strUserNum & "' "
   Adodc1.Recordset.Requery
   If Adodc1.Recordset.RecordCount = 0 Then
      Adodc1.Recordset.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   Else
      adoadodc1.MoveFirst
      Do While adoadodc1.EOF = False
         adoTaie.Execute "insert into accrpt423 (R42301,R42302,R42303,R42304,R42305,R42306,R42307,R42308,R42309,R42310) Values ('" & strUserNum & "', '" & adoadodc1.Fields("Dept").Value & "', '" & adoadodc1.Fields("Name").Value & "', " & _
         IIf(IsNull(adoadodc1.Fields("T1").Value), "Null", adoadodc1.Fields("T1").Value) & ", " & IIf(IsNull(adoadodc1.Fields("P1").Value), "Null", adoadodc1.Fields("P1").Value) & ", " & IIf(IsNull(adoadodc1.Fields("T2").Value), "Null", adoadodc1.Fields("T2").Value) & ", " & IIf(IsNull(adoadodc1.Fields("P2").Value), "Null", adoadodc1.Fields("P2").Value) & ", " & IIf(IsNull(adoadodc1.Fields("CFL").Value), "Null", adoadodc1.Fields("CFL").Value) & ", " & IIf(IsNull(adoadodc1.Fields("TOTAL").Value), "Null", adoadodc1.Fields("TOTAL").Value) & ",'" & adoadodc1.Fields("ID") & "')"
         adoadodc1.MoveNext
      Loop
      If Text1 = "2" Then
          If adoadodc1.State = adStateOpen Then adoadodc1.Close
          adoadodc1.CursorLocation = adUseClient
          strExc(0) = "select  r42302 as Dept,r42303 as Name,r42304 T1,r42305 P1,r42306 T2,r42307 P2,r42308 CFL,r42309 TOTAL,r42310 ID From accrpt423 Where r42301='" & strUserNum & "' Order by Dept,ID"
          adoadodc1.Open strExc(0), adoTaie, adOpenDynamic, adLockBatchOptimistic
         
          DataGrid1.Enabled = True
          Adodc1.Recordset.Requery
          MaskEdBox1.Text = MsgText(29)
          MaskEdBox2.Text = MsgText(29)
      Else
        Command1.Enabled = True
      End If
      'Add by Amy 2019/01/14
      Call SumShow
   End If
   
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

Private Sub Text1_Change()
    If Text1 = MsgText(601) Then
        Exit Sub
    End If
    If Text1 = 1 Then
        MaskEdBox1.SetFocus
    ElseIf Text1 = 2 Then
        MaskEdBox2.SetFocus
    End If
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text4_Change()
   If Text4 = MsgText(601) Then
      Text5 = ""
      Exit Sub
   Else
        'Modify by Amy 2020/04/24
        'If Text4 <> "1" And Text4 <> "J" Then
        If InStr(GetBookKeepCmp, Text4) = 0 Then
            Text5 = ""
            MsgBox MsgText(63), , MsgText(5)
            Exit Sub
        End If
   End If
   Text5 = A0802Query(Text4)
End Sub

'Add by Amy 2013/12/12
Private Sub Text4_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub
'end 2013/12/12

Private Sub Text4_GotFocus()
   TextInverse Text4
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Public Sub KeyDefine(KeyCode As Integer)
   Select Case KeyCode
      Case vbKeyF12
         If FormCheck = True Then
            Screen.MousePointer = vbHourglass
            IsFirst = True
            QueryTable
            Screen.MousePointer = vbDefault
            Exit Sub
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
   'Add by Amy 2013/12/12
   If Trim(Text4) = MsgText(601) Then
        FormCheck = False
        MsgBox "公司別不可為空值！", , MsgText(5)
        Text4.SetFocus
        TextInverse Text4
        Exit Function
    Else
        'Modify by Amy 2020/04/24
        'If Text4 <> "1" And Text4 <> "J" Then
        If InStr(GetBookKeepCmp, Text4) = 0 Then
            FormCheck = False
            Text5 = ""
            MsgBox MsgText(63), , MsgText(5)
            Text4.SetFocus
            TextInverse Text4
            Exit Function
        End If
    End If
   'end 2013/12/12
   If Trim(Text1) = MsgText(601) Then
      FormCheck = False
      MsgBox MsgText(181), , MsgText(5)
      Text1.SetFocus
      TextInverse Text1
      Exit Function
   End If
   If Trim(Text1) <> "1" And Trim(Text1) <> "2" Then
      FormCheck = False
      MsgBox "資料別錯誤，請重新輸入!", , MsgText(5)
      Text1.SetFocus
      TextInverse Text1
      Exit Function
   End If
   
   If Trim(Text1) = "1" Then
     If MaskEdBox1.Text = MsgText(29) Then
          FormCheck = False
          MsgBox "資料別為1時，傳票日期必需輸入!", , MsgText(5)
          MaskEdBox1.SetFocus
          Exit Function
     End If
     If MaskEdBox2.Text = MsgText(29) Then
         FormCheck = False
         MsgBox "資料別為1時，傳票日期必需輸入!", , MsgText(5)
         MaskEdBox2.SetFocus
         Exit Function
     End If
   End If
   FormCheck = True
End Function

'*************************************************
'  轉入傳票檔 acc020 acc021
'
'*************************************************
Private Sub TransferTable()
Dim SeqNo As Integer
Dim strToday As String, stra0202 As String
Dim AX205_D As String, AX205_C, CreditDept As String, AX212 As String
Dim strSN01 As String, strST06 As String
Dim strSave As String
Dim strIns As String, strA1R01 As String 'Add by Amy 2013/12/12
    On Error GoTo ErrHand
    
    adoTaie.BeginTrans
    strToday = ChangeTStringToTDateString(strSrvDate(2))
    SeqNo = 0
    'Add by Amy 2013/12/12
    If Text4 = "J" Then
        strA1R01 = MsgText(819) 'JD
    'Add by Amy 2020/04/24
    ElseIf Text4 = "L" Then
        strA1R01 = MsgText(820) 'LD
    Else
        strA1R01 = MsgText(801) 'D
    End If
    'end 2013/12/12
    With adoadodc1
        stra0202 = AccAutoNo(strA1R01, 4, Val(Year(strToday)), Val(Month(strToday))) '傳票編號
        
        '回寫acc1r0
        strSave = AccSaveAutoNo(strA1R01, Right(stra0202, 4), Mid(stra0202, 2, 3), Mid(stra0202, 5, 2))
         
        '新增主檔 Acc020
        strIns = "Insert Into Acc020 (a0201,a0202,a0205,a0206,a0207,a0208) " & _
                    "Values ('" & Text4 & "', '" & stra0202 & "', " & Val(strSrvDate(2)) & "," & Val(strSrvDate(2)) & ",to_char(sysdate,'HH24MISS'),'" & strUserNum & "') "
        adoTaie.Execute strIns
     
        '新增交檔 Acc021
        .MoveFirst
        Do While .EOF = False
          strSN01 = GetSalesData(.Fields("ID"), strST06)
          If Right(.Fields("Dept"), 1) <> "T" And .Fields("Dept") <> "ZZZ" Then
            For i = 2 To 6
                If Not IsNull(.Fields(i)) Then
                    If Val(.Fields(i)) > 0 Then
                        AX205_D = "24910" & i - 1 '借方科目
                        Select Case i
                            Case 2 '商標(大陸)
                                AX205_C = "410103" '貸方科目
                                CreditDept = "T" '貸方部門別
                            Case 3 '專利(大陸)
                                AX205_C = "411103"
                                CreditDept = "P"
                            Case 4 '商標(國外)
                                AX205_C = "412101"    'modify by sonia 2016/1/5 4121改412101
                                CreditDept = "CFT"
                            Case 5 '專利(國外)
                                AX205_C = "413101"    'modify by sonia 2016/1/5 4131改413101
                                CreditDept = "CFP"
                            Case 6 'CFL
                                'modify by sonia 2016/1/5 CFL改CFT法務收入
                                'AX205_C = "416102"
                                'CreditDept = "FCL"
                                AX205_C = "412102"
                                CreditDept = "CFT"
                                'end 2016/1/5
                        End Select
                    
                    '新增借方
                    strIns = "Insert Into Acc021 (ax201,ax202,ax203,ax204,ax205,ax206,ax207,ax209,ax212,ax213) " & _
                                "Values ('" & Text4 & "', '" & stra0202 & "', '" & ZeroBeforeNo(CStr(SeqNo), 3) & "', 'TOT','" & AX205_D & "', " & Val(.Fields(i)) & ", 0, '" & .Fields("ID") & "', '" & strSN01 & "', '" & strST06 & "') "
                    adoTaie.Execute strIns
                    SeqNo = SeqNo + 1
                    
                    '新增貸方
                    strIns = "Insert Into Acc021 (ax201,ax202,ax203,ax204,ax205,ax206,ax207,ax209,ax212,ax213) " & _
                                "Values ('" & Text4 & "', '" & stra0202 & "', '" & ZeroBeforeNo(CStr(SeqNo), 3) & "', '" & CreditDept & "','" & AX205_C & "',  0, " & Val(.Fields(i)) & ",'" & .Fields("ID") & "', '" & strSN01 & "', '" & strST06 & "') "
                    adoTaie.Execute strIns
                    SeqNo = SeqNo + 1
                   End If
                End If
            Next i
          End If
          adoadodc1.MoveNext
        Loop
    End With
    adoTaie.CommitTrans
    MsgBox "已產生傳票,傳票號碼 " & stra0202, , MsgText(21)
   Exit Sub
   
ErrHand:
   adoTaie.RollbackTrans
   If Err.Number = 0 Then
      Exit Sub
   End If
   adoSum.Close
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  2017/10/05以sales員編抓取ST02及ST06 原:抓SN02
'  ST02沒資料顯示員編/結餘，ST06顯示"結餘北/中/南/高"
'*************************************************
Private Function GetSalesData(ByVal p_ST01 As String, ByRef strST06) As String

On Error GoTo ErrHnd
    GetSalesData = ""
   
    'Moidify by Amy 2017/10/05 簡稱Table
    'strSql = "Select NVL(SN01,'Null') SN01,ST06 From Salesno,Staff Where ST01='" & p_ST01 & "' And St01=SN02(+)"
    strSql = "Select ST02,ST06 From Staff Where ST01='" & p_ST01 & "' "
       
    CheckOC3
    With AdoRecordSet3
       .CursorLocation = adUseClient
       .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
       If .RecordCount = 0 Then
          GetSalesData = p_ST01 & "/結餘"
       Else
          'GetSalesData = IIf(.Fields("SN01") = "Null", p_ST01, .Fields("SN01")) & "/結餘"
          GetSalesData = "" & .Fields("ST02") & "/結餘"
    'end 2017/10/05
          Select Case Val(.Fields("ST06"))
             Case 1
                 strST06 = "北"
             Case 2
                 strST06 = "中"
             Case 3
                 strST06 = "南"
             Case 4
                 strST06 = "高"
             Case Else
                 strST06 = ""
          End Select
          
          If strST06 <> "" Then
             If p_ST01 = "M0100" Then
                 strST06 = "結餘總"
             Else
                 strST06 = "結餘" & strST06
             End If
         End If
       End If
    End With
   
ErrHnd:
    If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Function

'Add by Amy 2019/01/14 增加合計
Private Sub SumShow()
    Dim strQ As String
    
    strExc(0) = "select  'ZZZ' as Dept,'全所合計' as Name,r42304 T1,r42305 P1,r42306 T2,r42307 P2,r42308 CFL,r42309 TOTAL,r42310 ID From accrpt423 Where r42301='" & strUserNum & "' And r42302='ZZZ' "
    If RsSum.State = adStateOpen Then RsSum.Close
    RsSum.CursorLocation = adUseClient
    RsSum.Open strExc(0), adoTaie, adOpenStatic, adLockReadOnly
    RsSum.MoveFirst
    Set Adodc2.Recordset = RsSum
    Adodc2.Recordset.Requery
End Sub
