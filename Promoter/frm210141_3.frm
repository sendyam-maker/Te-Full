VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm210141_3 
   BorderStyle     =   1  '單線固定
   Caption         =   "繳款輸入-繳款記錄刪除-明細"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9390
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   9390
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "繳款確認(&O)"
      Height          =   400
      Index           =   0
      Left            =   6840
      TabIndex        =   11
      Top             =   60
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.TextBox txtTot 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H8000000F&
      Height          =   270
      Index           =   2
      Left            =   3645
      TabIndex        =   4
      Top             =   3420
      Width           =   1050
   End
   Begin VB.TextBox txtTot 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H8000000F&
      Height          =   270
      Index           =   3
      Left            =   5220
      TabIndex        =   3
      Top             =   3420
      Width           =   1050
   End
   Begin VB.TextBox txtTot 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H8000000F&
      Height          =   270
      Index           =   4
      Left            =   6795
      TabIndex        =   2
      Top             =   3420
      Width           =   915
   End
   Begin VB.TextBox txtTot 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H8000000F&
      Height          =   270
      Index           =   5
      Left            =   8235
      TabIndex        =   1
      Top             =   3420
      Width           =   1050
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Index           =   1
      Left            =   8100
      TabIndex        =   0
      Top             =   60
      Width           =   1200
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frm210141_3.frx":0000
      Height          =   2835
      Left            =   90
      TabIndex        =   5
      Top             =   510
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   5001
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   14
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
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
      ColumnCount     =   12
      BeginProperty Column00 
         DataField       =   "公司別"
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
         DataField       =   "收據編號"
         Caption         =   "收據編號"
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
         DataField       =   "單據日期"
         Caption         =   "單據日期"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "ee/mm/dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "本所案號"
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
      BeginProperty Column04 
         DataField       =   "案件性質"
         Caption         =   "案件性質"
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
      BeginProperty Column05 
         DataField       =   "國別"
         Caption         =   "國別"
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
      BeginProperty Column06 
         DataField       =   "服務費"
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
      BeginProperty Column07 
         DataField       =   "規費"
         Caption         =   "規費"
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
         DataField       =   "扣繳金額"
         Caption         =   "扣繳金額"
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
         DataField       =   "案件名稱"
         Caption         =   "案件名稱"
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
         DataField       =   "a0k03"
         Caption         =   "客戶編號"
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
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
            Alignment       =   2
            ColumnWidth     =   675.213
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   1244.976
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   884.976
         EndProperty
         BeginProperty Column03 
            Locked          =   -1  'True
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column04 
            Locked          =   -1  'True
            ColumnWidth     =   840.189
         EndProperty
         BeginProperty Column05 
            Locked          =   -1  'True
            ColumnWidth     =   615.118
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            ColumnWidth     =   659.906
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   645.165
         EndProperty
         BeginProperty Column08 
            Alignment       =   1
            ColumnAllowSizing=   -1  'True
            Locked          =   -1  'True
            ColumnWidth     =   854.929
         EndProperty
         BeginProperty Column09 
            Locked          =   -1  'True
            ColumnWidth     =   2624.882
         EndProperty
         BeginProperty Column10 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column11 
            Locked          =   -1  'True
            ColumnWidth     =   2564.788
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   3870
      Top             =   150
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
   Begin VB.Label lblMemo 
      AutoSize        =   -1  'True
      Caption         =   "說明：收據附加符號: ◎未列印 ＊已開發票"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   1
      Left            =   90
      TabIndex        =   10
      Top             =   270
      Width           =   3375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "服務費"
      Height          =   180
      Index           =   7
      Left            =   3060
      TabIndex        =   9
      Top             =   3465
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "規費"
      Height          =   180
      Index           =   8
      Left            =   4815
      TabIndex        =   8
      Top             =   3465
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "扣繳"
      Height          =   180
      Index           =   9
      Left            =   6390
      TabIndex        =   7
      Top             =   3465
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "總計"
      Height          =   180
      Index           =   10
      Left            =   7830
      TabIndex        =   6
      Top             =   3465
      Width           =   360
   End
End
Attribute VB_Name = "frm210141_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2021/12/28 Form2.0已修改(DataGrid1改Fonts)
'Created by Morgan 2013/12/3
'Memo by Lydia 2019/07/01 表單名稱:智權人員繳款資料輸入=>繳款輸入
Option Explicit

Public m_Return As Integer 'Added by Morgan 2017/12/18

Private Sub cmdOK_Click(Index As Integer)
   m_Return = Index 'Added by Morgan 2017/12/18
   Unload Me
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm210141_3 = Nothing
End Sub
