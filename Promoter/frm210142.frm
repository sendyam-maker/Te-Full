VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm210142 
   Caption         =   "繳款查詢及簽收查詢"
   ClientHeight    =   4630
   ClientLeft      =   60
   ClientTop       =   350
   ClientWidth     =   9530
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   4635
   ScaleMode       =   0  '使用者自訂
   ScaleWidth      =   9525.001
   Begin VB.CommandButton cmdCall 
      Caption         =   "簽收資料查詢"
      Height          =   405
      Left            =   7440
      TabIndex        =   26
      Top             =   1110
      Width           =   1485
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Index           =   5
      Left            =   2320
      MaxLength       =   9
      TabIndex        =   11
      Top             =   960
      Width           =   1100
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Index           =   4
      Left            =   1080
      MaxLength       =   9
      TabIndex        =   10
      Top             =   960
      Width           =   1100
   End
   Begin VB.TextBox txtTot 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H8000000F&
      Height          =   270
      Left            =   3585
      TabIndex        =   22
      Top             =   4300
      Width           =   1050
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   375
      Left            =   8400
      TabIndex        =   15
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton CmdDetail 
      Caption         =   "明細(&T)"
      Height          =   375
      Left            =   7440
      TabIndex        =   14
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "查詢(&F)"
      Default         =   -1  'True
      Height          =   375
      Left            =   7440
      TabIndex        =   13
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Index           =   3
      Left            =   6780
      MaxLength       =   2
      TabIndex        =   9
      Top             =   600
      Width           =   420
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Index           =   2
      Left            =   6480
      MaxLength       =   1
      TabIndex        =   8
      Top             =   600
      Width           =   276
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Index           =   1
      Left            =   5160
      MaxLength       =   6
      TabIndex        =   7
      Top             =   600
      Width           =   1236
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Index           =   0
      Left            =   4560
      MaxLength       =   3
      TabIndex        =   6
      Top             =   600
      Width           =   612
   End
   Begin VB.TextBox txtSales 
      Height          =   300
      Left            =   4560
      MaxLength       =   6
      TabIndex        =   2
      Top             =   240
      Width           =   915
   End
   Begin VB.TextBox txtSalesArea1 
      Height          =   300
      Left            =   2160
      TabIndex        =   1
      Top             =   240
      Width           =   915
   End
   Begin VB.TextBox txtSalesArea 
      Height          =   300
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   915
   End
   Begin VB.TextBox txtCloseDate 
      Height          =   300
      Index           =   0
      Left            =   1080
      MaxLength       =   7
      TabIndex        =   4
      Top             =   600
      Width           =   915
   End
   Begin VB.TextBox txtCloseDate 
      Height          =   300
      Index           =   1
      Left            =   2160
      MaxLength       =   7
      TabIndex        =   5
      Top             =   600
      Width           =   915
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frm210142.frx":0000
      Height          =   2625
      Left            =   60
      TabIndex        =   16
      Top             =   1600
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   4639
      _Version        =   393216
      AllowUpdate     =   0   'False
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
      ColumnCount     =   24
      BeginProperty Column00 
         DataField       =   "Rdate"
         Caption         =   "繳款日期"
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
         DataField       =   "Rtime"
         Caption         =   "時間"
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
      BeginProperty Column02 
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
      BeginProperty Column03 
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
      BeginProperty Column04 
         DataField       =   "A4414"
         Caption         =   "出納"
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
         DataField       =   "recYN"
         Caption         =   "收款"
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
         DataField       =   "pnt"
         Caption         =   "點數"
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
      BeginProperty Column07 
         DataField       =   "A4405"
         Caption         =   "票據金額"
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
         DataField       =   "A4406"
         Caption         =   "電匯金額"
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
         DataField       =   "A4408"
         Caption         =   "現金"
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
         DataField       =   "A4409"
         Caption         =   "抵暫收款"
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
      BeginProperty Column11 
         DataField       =   "A4430"
         Caption         =   "其他"
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
      BeginProperty Column12 
         DataField       =   "A4410"
         Caption         =   "溢收款"
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
         DataField       =   "A4411"
         Caption         =   "手續費"
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
      BeginProperty Column14 
         DataField       =   "A4422"
         Caption         =   "補扣繳"
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
         DataField       =   "A4426"
         Caption         =   "外幣"
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
      BeginProperty Column16 
         DataField       =   "A4431"
         Caption         =   "其他備註"
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
         DataField       =   "A4412"
         Caption         =   "智權人員備註"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column18 
         DataField       =   "ST02"
         Caption         =   "操作人員"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column19 
         DataField       =   "axd04"
         Caption         =   ""
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
      BeginProperty Column20 
         DataField       =   "A4401"
         Caption         =   ""
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
      BeginProperty Column21 
         DataField       =   "A4402"
         Caption         =   ""
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
      BeginProperty Column22 
         DataField       =   "A4403"
         Caption         =   ""
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
      BeginProperty Column23 
         DataField       =   "a0k01"
         Caption         =   ""
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
            ColumnAllowSizing=   -1  'True
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   849.381
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   479.937
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   909.444
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   909.444
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   469.738
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   469.738
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   599.496
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            ColumnWidth     =   889.612
         EndProperty
         BeginProperty Column08 
            Alignment       =   1
            ColumnWidth     =   949.675
         EndProperty
         BeginProperty Column09 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   769.486
         EndProperty
         BeginProperty Column10 
            Alignment       =   1
            ColumnAllowSizing=   0   'False
            Locked          =   -1  'True
            ColumnWidth     =   849.381
         EndProperty
         BeginProperty Column11 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   729.822
         EndProperty
         BeginProperty Column12 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   709.423
         EndProperty
         BeginProperty Column13 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   669.759
         EndProperty
         BeginProperty Column14 
            Alignment       =   1
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   669.759
         EndProperty
         BeginProperty Column15 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   649.927
         EndProperty
         BeginProperty Column16 
            Locked          =   -1  'True
            ColumnWidth     =   1069.234
         EndProperty
         BeginProperty Column17 
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   1859.119
         EndProperty
         BeginProperty Column18 
            Alignment       =   2
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   899.811
         EndProperty
         BeginProperty Column19 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column20 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column21 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column22 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column23 
            Object.Visible         =   0   'False
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   120
      Top             =   4440
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   547
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
   Begin MSForms.ComboBox Combo3 
      Height          =   336
      Left            =   4560
      TabIndex        =   3
      Top             =   240
      Width           =   1920
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "3387;593"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtA0k04 
      Height          =   330
      Left            =   4560
      TabIndex        =   12
      Top             =   960
      Width           =   2640
      VariousPropertyBits=   671105051
      Size            =   "4657;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblSalesName 
      Height          =   300
      Left            =   5550
      TabIndex        =   27
      Top             =   240
      Width           =   1470
      VariousPropertyBits=   27
      Size            =   "2593;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Line Line4 
      X1              =   6119.788
      X2              =   7079.284
      Y1              =   720.778
      Y2              =   720.778
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收據抬頭："
      Height          =   180
      Index           =   2
      Left            =   3600
      TabIndex        =   25
      Top             =   1020
      Width           =   900
   End
   Begin VB.Line Line3 
      X1              =   2159.866
      X2              =   2430.724
      Y1              =   1080.165
      Y2              =   1080.165
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "客戶編號："
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   24
      Top             =   1020
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "點數合計"
      Height          =   180
      Index           =   7
      Left            =   2775
      TabIndex        =   23
      Top             =   4300
      Width           =   720
   End
   Begin VB.Label Label4 
      Caption         =   "說明：收據附加符號: ◎未列印 ＊已開發票"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   1360
      Width           =   3615
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Left            =   3600
      TabIndex        =   20
      Top             =   645
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   180
      Left            =   3600
      TabIndex        =   19
      Top             =   285
      Width           =   900
   End
   Begin VB.Line Line2 
      X1              =   1919.992
      X2              =   2159.866
      Y1              =   360.389
      Y2              =   360.389
   End
   Begin VB.Line Line1 
      X1              =   1919.992
      X2              =   2189.851
      Y1              =   720.778
      Y2              =   720.778
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "繳款日期："
      Height          =   180
      Left            =   120
      TabIndex        =   18
      Top             =   640
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "業 務 區  ："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   17
      Top             =   280
      Width           =   900
   End
End
Attribute VB_Name = "frm210142"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/07/27 智權-調整財務系統(20200909) 'Memo by Lydia 2021/08/27 上線
'原「繳款資料查詢」標題修改為「繳款查詢及簽收查詢」。
'加入”簽收資料查詢”功能鍵，按此功能鍵後，進入財務處系統的界面，須再鍵入員工代號及密碼。
'end 2021/07/27
'Memo by Lydia 2021/07/16 改成Form2.0 ; lblSalesName、txtA0k04、DataGrid1改字型=新細明體-ExtB
'Memo by Lydia 2019/07/01 表單名稱:智權人員繳款資料查詢=>繳款資料查詢
'Created by Eric 2013/12/31
Option Explicit

Dim adoquery As ADODB.Recordset
Dim bolShowMsgBox As Boolean, bolSelData As Boolean
Dim stST05 As String, stST15 As String
'Add by Amy 2014/05/21
Dim bolSpecMan As Boolean  '是否為特殊設定檔人員
Dim strSpecCode As String '特殊設定檔設定代號
Dim m_strListPer As String 'Add By Sindy 2020/6/11
'Add By Sindy 2023/6/12
Dim arrID
Dim bolAreaMan As Boolean '下拉選單有區主管
'2023/6/12 END


'Added by Lydia 2021/07/09 加入”簽收資料查詢”功能鍵
Private Sub cmdCall_Click()
    Me.Hide
    If frm210106_1.setNextForm = "" Then
       Call frm210106.SetParent(Me)
       frm210106.Show
    Else
       Call frm210106_1.setCaller(frm210106, Me)
       frm210106_1.Show
    End If
    
End Sub

Private Sub cmdDetail_Click()
Dim stVTB11 As String, stVTB22 As String
Dim iCol    As Integer
Dim rtNo    As String
Dim strCon  As String
Dim Role    As String
Dim PayDate As Long
Dim PayTime As Long
Dim dblVal(3) As Double
     
   Role = Adodc1.Recordset("a4401")                '智權人員
   PayDate = Adodc1.Recordset("a4402")             '繳款日期
   PayTime = Adodc1.Recordset("a4403")             '繳款時間

   'Modified by Morgan 2015/7/23 +公司別
   'Modified by Lydia 2018/06/05 修改顯示案件性質 '020',CPM04,CPM03 => '000',CPM03,CPM04
   'Modify by Amy 2020/04/09 公司別抓acc080簡稱 原:decode(a0k11,'1','商標','2','專利','J','智權',a0k11)
   'Modified by Lydia 2023/11/13 開立INVOICE，不列印收據;decode(nvl(a0k19,0),0,'◎')=> decode(a0k32,'Z','',decode(nvl(a0k19,0),0,'◎'))
   strExc(0) = "select sqldatet(a0k02) 單據日期" & _
       ",cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) 本所案號" & _
       ",decode(a0j04,'000',cpm03,cpm04) 案件性質" & _
       ",na03 國別,axd06 服務費,axd07 規費,axd08 扣繳金額,a0k01||decode(a0k32,'Z','',decode(nvl(a0k19,0),0,'◎'))||decode(AXC01,null,'','＊') 收據編號" & _
       ",nvl(tm05,nvl(pa05,nvl(lc05,nvl(sp05,hc06)))) 案件名稱,a0k03,a0k04,a0820 公司別" & _
       " from ACC441,ACC0J0,acc0k0,acc431,Acc080,caseprogress,casepropertymap,nation" & _
       ",trademark,patent,lawcase,servicepractice,hirecase" & _
       " where A0J01(+)=AXD05 AND A0J13(+)=AXD04 And a0k11=a0801(+)" & _
       " and axd01='" & Role & "' and axd02='" & PayDate & "' and axd03='" & PayTime & "'" & _
       " and a0k01(+)=a0j13 and axc02(+)=a0j13 and cp09(+)=a0j01" & _
       " and cpm01(+)=cp01 and cpm02(+)=cp10 and na01(+)=a0j04" & _
       " and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04" & _
       " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
       " and lc01(+)=cp01 and lc02(+)=cp02 and lc03(+)=cp03 and lc04(+)=cp04" & _
       " and sp01(+)=cp01 and sp02(+)=cp02 and sp03(+)=cp03 and sp04(+)=cp04" & _
       " and hc01(+)=cp01 and hc02(+)=cp02 and hc03(+)=cp03 and hc04(+)=cp04" & _
       " order by a0k02,a0j13,a0j01"
                                   
   intI = 1
   Set adoquery = ClsLawReadRstMsg(intI, strExc(0))
    
   If adoquery.RecordCount > 0 Then
      With frm210141_3
      'Modified by Lydia 2019/06/27
      'frm210141_3.Caption = "智權人員繳款資料查詢-繳款資料明細"
      frm210141_3.Caption = "繳款資料查詢-繳款資料明細"
      'Modify by Amy 2014/06/10 +FormName 改暫存TB
      'Set .Adodc1.Recordset = PUB_CreateRecordset(adoquery)
      Set .Adodc1.Recordset = PUB_CreateRecordset(adoquery, , , , .Name)
         With adoquery
           .MoveFirst
           Do While Not .EOF
              dblVal(1) = dblVal(1) + Val("" & .Fields("服務費"))
              dblVal(2) = dblVal(2) + Val("" & .Fields("規費"))
              dblVal(3) = dblVal(3) + Val("" & .Fields("扣繳金額"))
              .MoveNext
           Loop
         End With
         
         .txtTot(2) = Format(dblVal(1), "#,##0")
         .txtTot(3) = Format(dblVal(2), "#,##0")
         .txtTot(4) = Format(dblVal(3), "#,##0")
         .txtTot(5) = Format(dblVal(1) + dblVal(2) - dblVal(3), "#,##0")
         .Show vbModal
      End With
   End If
   
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdSearch_Click()
Dim bolCancel As Boolean 'Add By Sindy 2020/6/12
Dim intErrCol As Integer
   
   'Add By Sindy 2023/6/12
   If Combo3.Visible = True Then
      bolCancel = False
      Call Combo3_Validate(bolCancel)
      If bolCancel = True Then
         Exit Sub
      End If
   End If
   'Add by Amy 2020/03/25 +有下拉選單
   If Combo3.Visible = True Then
      Call Combo3_LostFocus 'Add By Sindy 2020/7/15 讓人員按Enter,須再啟動此函數,txtSales欄位值才會置換
      If Combo3 = MsgText(601) Then
          Call Combo3_Validate(bolCancel)
          If bolCancel = True Then
              Combo3.SetFocus
              Exit Sub
          End If
      ElseIf txtSales = MsgText(601) Then
          txtSales = Mid(Combo3, 1, Val(InStr(Combo3, " ")) - 1)
      End If
   End If
   Call txtSales_Validate(bolCancel)
   If bolCancel = True Then
      'Modify by Amy 2020/03/25 +有下拉選單
      If Combo3.Visible = True Then
         Combo3.SetFocus
      'Modified by Lydia 2021/05/20 排除隱藏
      'ElseIf txtSales.Enabled = True Then
      ElseIf txtSales.Enabled = True And txtSales.Visible = True Then
         txtSales.SetFocus
         txtSales_GotFocus
      End If
      Exit Sub
   End If
   '2023/6/12 END
'   'Add By Sindy 2020/6/12
'   Call txtSales_Validate(bolCancel)
'   If bolCancel = True Then
'      txtSales.SetFocus
'      txtSales_GotFocus
'      Exit Sub
'   End If
'   '2020/6/12 END
   
   'Modify By Sindy 2020/7/29 檢查部門欄位
   If PUB_ChkFormSalesDept(strUserNum, txtSales, txtSalesArea, txtSalesArea1, intErrCol) = False Then
      If intErrCol = 0 Then
         txtSales.SetFocus
         txtSales_GotFocus
      ElseIf intErrCol = 1 Then
         txtSalesArea.SetFocus
         txtSalesArea_GotFocus
      Else
         txtSalesArea1.SetFocus
         txtSalesArea1_GotFocus
      End If
      Exit Sub
   End If
   
'   'add by sonia 2016/12/21 柄佑82026可看中所全部或自已
'   If strUserNum = "82026" Then
'      If txtSalesArea <> "P11" Or txtSalesArea1 <> "P11" Then
'         If txtSalesArea < "S2" Or txtSalesArea > "S29" Then
'            MsgBox "業務區起始條件錯誤！只可查中所業務區", vbExclamation
'            txtSalesArea.SetFocus
'            txtSalesArea_GotFocus
'            Exit Sub
'         End If
'         If txtSalesArea1 < "S2" Or txtSalesArea1 > "S29" Then
'            MsgBox "業務區迄止條件錯誤！只可查中所業務區", vbExclamation
'            txtSalesArea1.SetFocus
'            txtSalesArea1_GotFocus
'            Exit Sub
'         End If
'      Else
'         If Trim(txtSales) <> strUserNum Then
'            MsgBox "查專利處工程師時只可查自己的資料", vbExclamation
'            txtSales.SetFocus
'            txtSales_GotFocus
'            Exit Sub
'         End If
'      End If
'   End If
'   'Added by Lydia 2020/06/15 簡協理可看所有智權人員
'   If strUserNum = "69005" Then
'      If Left(txtSalesArea, 1) <> "S" Then
'         MsgBox "業務區起始條件錯誤！只可查智權部", vbExclamation
'         txtSalesArea.SetFocus
'         txtSalesArea_GotFocus
'         Exit Sub
'      End If
'      If Left(txtSalesArea1, 1) <> "S" Then
'         MsgBox "業務區迄止條件錯誤！只可查智權部", vbExclamation
'         txtSalesArea1.SetFocus
'         txtSalesArea1_GotFocus
'         Exit Sub
'      End If
'   End If
'   'end 2020/06/15
'
'   If txtSalesArea > txtSalesArea1 Then
'      MsgBox "業務區範圍條件錯誤！", vbExclamation
'      txtSalesArea.SetFocus
'      txtSalesArea_GotFocus
'      Exit Sub
'   End If
'   'end 2016/12/21
   
   Screen.MousePointer = vbHourglass
   doQuery

   bolSelData = False
   bolShowMsgBox = True
   Screen.MousePointer = vbDefault
  
End Sub

Private Sub doQuery(Optional pNoMsg As Boolean = False)
Dim strCon As String, bolOK As Boolean
Dim stVTB1 As String, stVTB2 As String
Dim dblVal(3) As Double
Dim opt As Integer
Dim SaleArea As String
Dim SaleArea1 As String
Dim TMk01 As String
Dim TMk02 As String
Dim TMk03 As String
Dim TMk04 As String
Dim txtCloseDate0  As Long
Dim txtCloseDate1  As Long
Dim txtCloseDateSt As Long
Dim txtCloseDateEd As Long

   SaleArea = UCase(txtSalesArea.Text)
   SaleArea1 = UCase(txtSalesArea1.Text)
   
   If txtCloseDate(0).Text <> "" Then
       txtCloseDate0 = txtCloseDate(0).Text
   End If
   If txtCloseDate(1).Text <> "" Then
       txtCloseDate1 = txtCloseDate(1).Text
   End If
   TMk01 = UCase(Text1(0).Text)
   TMk02 = UCase(Text1(1).Text)
   If UCase(Text1(2).Text) = "" Then
      TMk03 = "0"
   Else
      TMk03 = UCase(Text1(2).Text)
   End If
   
   If UCase(Text1(3).Text) = "" Then
      TMk04 = "00"
   Else
      TMk04 = UCase(Text1(3).Text)
   End If
     
      
   '業務區與智權人員欄位不可同時空白;繳款日期不可均空白
   If SaleArea = "" And SaleArea1 = "" Then
      If txtSales = "" Then
         MsgBox "業務區與智權人員欄位不可同時空白！", vbInformation
         txtSalesArea.SetFocus
         Exit Sub
      End If
   End If
          
   If txtCloseDate1 = 0 And txtCloseDate0 = 0 Then
      MsgBox "繳款日期不可均空白！", vbInformation
      txtCloseDate(0).SetFocus
      Exit Sub
   End If

   
   'ACC440/ACC441內日期格式為西元年月日 故轉換使用者輸入日期
   If txtCloseDate0 <> 0 Then
       txtCloseDateSt = txtCloseDate0 + 19110000
   End If
   If txtCloseDate1 <> 0 Then
       txtCloseDateEd = txtCloseDate1 + 19110000
   End If
      
   strExc(0) = "" 'Added by Lydia 2017/09/20
   
'Move by Lydia 2017/09/20 從"組查詢條件"下面移來
   'Modify by Amy 2014/05/21
   '區別
   'Add by Amy 2019/02/13 +總經理業務工作代理人,可處理總經理員工編號
   If bolSpecMan = True And (InStr(strSpecCode, "A8") > 0 Or InStr(strSpecCode, "總經理業務工作代理人") > 0) And txtSales <> strUserNum Then
        '開放專利處部份智權同仁資料給彥葶代為處理,不考慮業務區(因彥葶與開放的智權同仁業務區不同)
   Else
        If SaleArea <> "" Then
            strExc(0) = strExc(0) & " and s2.st15 >= '" & SaleArea & "' "
        End If
        If SaleArea1 <> "" Then
            strExc(0) = strExc(0) & " and s2.st15 <= '" & SaleArea1 & "' "
        End If
   End If
   
   'Modify by Amy 2014/05/21
   'Mark by 2023/01/11 移到下方; 改從主Query判斷智權人員編號
   ''智權人員
   'If txtSales <> "" Then
   '   strExc(0) = strExc(0) & " and x.a4401 = '" & txtSales & "'"
   ''智權人員 為空
   'Else
   '     If bolSpecMan = True And InStr(strSpecCode, "A8") > 0 Then
   '         'A2023彥葶登入,未輸智權人員-設定查A7人員
   '         strExc(0) = strExc(0) & " and x.a4401 in ('" & Replace(Pub_GetSpecMan("A7"), ";", "','") & "') "
   '     End If
   'End If
   'end 2014/05/21
    
   If txtCloseDateSt <> 0 Then
      strExc(0) = strExc(0) & " and x.a4402 >= '" & txtCloseDateSt & "' "
   End If
   If txtCloseDateEd <> 0 Then
      strExc(0) = strExc(0) & " and x.a4402 <= '" & txtCloseDateEd & "' "
   End If
   
   strCon = Replace(Replace(strExc(0), "x.a44", "axd"), "s2.", "") 'Added by Lydia 2017/09/20

   'Move by 2023/01/11 改從主Query判斷智權人員編號
   'Modify by Amy 2014/05/21
   '智權人員
   If txtSales <> "" Then
      'Modified by Lydia 2023/01/11 開放客戶現在的智權人員也可以查詢
      'strExc(0) = strExc(0) & " and x.a4401 = '" & txtSales & "'"
      strExc(0) = strExc(0) & " and (x.a4401 = '" & txtSales & "' or cu13='" & txtSales & "')"
      'end 2023/01/11
   '智權人員 為空
   Else
        If bolSpecMan = True And InStr(strSpecCode, "A8") > 0 Then
            'A2023彥葶登入,未輸智權人員-設定查A7人員
            'Modified by Lydia 2023/01/11 開放客戶現在的智權人員也可以查詢
            'strExc(0) = strExc(0) & " and x.a4401 in ('" & Replace(Pub_GetSpecMan("A7"), ";", "','") & "') "
            strExc(0) = strExc(0) & " and (x.a4401 in ('" & Replace(Pub_GetSpecMan("A7"), ";", "','") & "') or cu13 in ('" & Replace(Pub_GetSpecMan("A7"), ";", "','") & "'))"
        End If
   End If
   'end ---'Move by 2023/01/11
   
   'Modified by Lydia 2017/11/16 增加條件:客戶和收據抬頭
   If Text1(4).Text <> "" Then
      strExc(0) = strExc(0) & " and a0k03>=" & CNULL(Text1(4))
   End If
   If Text1(5).Text <> "" Then
      strExc(0) = strExc(0) & " and a0k03<=" & CNULL(Text1(5))
   End If
   If txtA0k04.Text <> "" Then
      strExc(0) = strExc(0) & " and instr(a0k04," & CNULL(txtA0k04) & ") > 0 "
   End If
   'end 2017/11/16
   
   If TMk01 <> "" And TMk02 <> "" And TMk03 <> "" And TMk04 <> "" Then
      strExc(0) = strExc(0) & " and exists(select * from ACC441 ,CASEPROGRESS    " & _
                              "            Where axd01 = X.a4401 And axd02 = X.a4402 And axd03 = X.a4403 " & _
                              "            AND CP09(+)=AXD05                                             " & _
                              "            AND CP01='" & TMk01 & "' AND CP02='" & TMk02 & "' AND CP03='" & TMk03 & "' AND CP04='" & TMk04 & "')    "
   End If
'end 2017/09/20

   '組查詢條件
   'modify by sonia 2014/7/2 原a4404已不用故取消,原位置改放點數
   'strExc(0) = "select distinct sqldatet(a4402) rdate,sqltime(a4403) rtime,x.a4414,DECODE(NVL(x.a4416,'N'),'N','N','Y') recYN, x.a4404,x.a4405, " & _
               " DECODE(x.a4406,0,x.a4407,x.a4406) a4406,x.a4408,x.a4409,x.a4410,x.a4411,x.a4422,x.a4426,x.a4412,s.st02,x.a4401,x.a4402,x.a4403 " & _
               " from acc440 x , staff s,staff s2 " & _
               " where  s.st01(+) = x.a4417 " & _
               " and  s2.st01(+) = x.a4401 "
   'Modified by Morgan 2015/7/15 +A4430,A4431
   'Modified by Morgan 2015/10/2 出納應顯示中文
   'Modified by Lydia 2017/09/20 增加收據抬頭,只抓明細第1筆收據
   'strExc(0) = "select sqldatet(a4402) rdate,sqltime(a4403) rtime,max(s3.st02) a4414,DECODE(NVL(x.a4416,'N'),'N','N','Y') recYN,sum(axd06)/1000 pnt,x.a4405, " & _
               " DECODE(x.a4406,0,x.a4407,x.a4406) a4406,x.a4408,x.a4409,x.a4410,x.a4411,x.a4422,x.a4426,x.a4412,s.st02,x.a4401,x.a4402,x.a4403,max(x.A4430) A4430,max(x.A4431) A4431 " & _
               " from acc440 x ,staff s,staff s2,staff s3,acc441 " & _
               " where  s.st01(+) = x.a4417 and s2.st01(+) = x.a4401 and x.a4401=axd01(+) and x.a4402=axd02(+) and x.a4403=axd03(+) and s3.st01(+)=x.a4414 "
   'Modified by Lydia 2017/11/16 顯示客戶編號,出納=Y
   'strExc(0) = "select sqldatet(a4402) rdate,sqltime(a4403) rtime,substr(a0k04,1,8) a0k04,max(s3.st02) a4414,DECODE(NVL(x.a4416,'N'),'N','N','Y') recYN,sum(axd06)/1000 pnt,x.a4405, " & _
               " DECODE(x.a4406,0,x.a4407,x.a4406) a4406,x.a4408,x.a4409,x.a4410,x.a4411,x.a4422,x.a4426,x.a4412,s.st02,x.a4401,x.a4402,x.a4403,max(x.A4430) A4430,max(x.A4431) A4431 " & _
               " from acc440 x ,staff s,staff s2,staff s3,acc441,acc0k0, " & _
               " (select axd01 n01,axd02 n02,axd03 n03,min(axd04) n04 from acc441,staff where axd01= st01(+) " & strCon & " group by axd01,axd02,axd03) vtb01 " & _
               " where  s.st01(+) = x.a4417 and s2.st01(+) = x.a4401 and x.a4401=axd01(+) and x.a4402=axd02(+) and x.a4403=axd03(+) and s3.st01(+)=x.a4414 " & _
               " and x.a4401=n01 and x.a4402=n02 and x.a4403=n03 and n04=a0k01(+) " & strExc(0)
   'Modified by Lydia 2023/01/11 開放客戶現在的智權人員也可以查詢=>增加Customer 條件and substr(a0k03,1,8)=cu01(+) and substr(a0k03,9,1)=cu02(+)
   strExc(0) = "select sqldatet(a4402) rdate,substr(sqltime(a4403),1,5) rtime,a0k03,substr(a0k04,1,8) a0k04,decode(x.a4414,null,'','Y') a4414,DECODE(NVL(x.a4416,'N'),'N','N','Y') recYN,sum(axd06)/1000 pnt,x.a4405, " & _
               " DECODE(x.a4406,0,x.a4407,x.a4406) a4406,x.a4408,x.a4409,x.a4410,x.a4411,x.a4422,x.a4426,x.a4412,s.st02,x.a4401,x.a4402,x.a4403,max(x.A4430) A4430,max(x.A4431) A4431 " & _
               " from acc440 x ,staff s,staff s2,staff s3,acc441,acc0k0,customer, " & _
               " (select axd01 n01,axd02 n02,axd03 n03,min(axd04) n04 from acc441,staff where axd01= st01(+) " & strCon & " group by axd01,axd02,axd03) vtb01 " & _
               " where  s.st01(+) = x.a4417 and s2.st01(+) = x.a4401 and x.a4401=axd01(+) and x.a4402=axd02(+) and x.a4403=axd03(+) and s3.st01(+)=x.a4414 " & _
               " and x.a4401=n01 and x.a4402=n02 and x.a4403=n03 and n04=a0k01(+) and substr(a0k03,1,8)=cu01(+) and substr(a0k03,9,1)=cu02(+) " & strExc(0)
               
   'add by sonia 2014/7/2 加點數欄
   'Modified by Lydia 2017/09/20 + substr(a0k04,1,8)
   'Modified by Lydia 2017/11/16 顯示客戶編號,出納=Y
   'strExc(0) = strExc(0) & " group by sqldatet(a4402),sqltime(a4403),substr(a0k04,1,8),x.a4414,DECODE(NVL(x.a4416,'N'),'N','N','Y'),x.a4405, " & _
                           " DECODE(x.a4406,0,x.a4407,x.a4406),x.a4408,x.a4409,x.a4410,x.a4411,x.a4422,x.a4426,x.a4412,s.st02,x.a4401,x.a4402,x.a4403"
   strExc(0) = strExc(0) & " group by sqldatet(a4402),substr(sqltime(a4403),1,5),a0k03,substr(a0k04,1,8),decode(x.a4414,null,'','Y'),DECODE(NVL(x.a4416,'N'),'N','N','Y'),x.a4405, " & _
                           " DECODE(x.a4406,0,x.a4407,x.a4406),x.a4408,x.a4409,x.a4410,x.a4411,x.a4422,x.a4426,x.a4412,s.st02,x.a4401,x.a4402,x.a4403"

   
   'Added by Lydia 2017/09/20 +排序
   strExc(0) = strExc(0) & " order by x.a4401,x.a4402,x.a4403 "
   
   intI = 1
   Set adoquery = ClsLawReadRstMsg(intI, strExc(0))
   
   FormReset
   'Modify by Amy 2014/06/10 +FormName 改暫存TB
   'Set Adodc1.Recordset = PUB_CreateRecordset(adoquery)
   Set Adodc1.Recordset = PUB_CreateRecordset(adoquery, , , , Me.Name)
   
   CmdDetail.Enabled = True
   
   If adoquery.RecordCount <= 0 Then
      If pNoMsg = False Then
         CmdDetail.Enabled = False
         MsgBox "無符合資料！", vbExclamation
      End If
   'ADD BY SONIA 2014/7/2 加點數合計
   Else
      dblVal(1) = 0
      With adoquery
        .MoveFirst
        Do While Not .EOF
           dblVal(1) = dblVal(1) + Val("" & .Fields("pnt"))
           .MoveNext
        Loop
      End With
         
         txtTot = Format(dblVal(1), "##,##0.###")
   'END 2014/7/2
   End If
     
End Sub

Private Sub FormReset()
      
   If Not Adodc1.Recordset Is Nothing Then
      If Adodc1.Recordset.State = 1 Then
         Adodc1.Recordset.Close
         DataGrid1.Refresh
      End If
   End If
End Sub

Private Sub Form_Load()
    
   'Modified by Lydia 2107/09/20
   'frm210142.Height = 4545
   'frm210142.Width = 9480
   'Modified by Lydia 2017/11/16
   'frm210142.Height = 4800
   'frm210142.Width = 9480
   'end 2017/09/20
   Me.Height = 5220
   Me.Width = 9660

   
   MoveFormToCenter Me
   bolShowMsgBox = False
   
   stST15 = PUB_GetStaffST15(strUserNum, 1)
   stST05 = PUB_GetST05(strUserNum)
   
   bolAreaMan = False 'Add By Sindy 2023/6/12
   'Modify By Sindy 2020/7/28 設定員編,部門,所別權限
   Call PUB_SetFormSaleDept(strUserNum, , txtSalesArea, txtSalesArea1, txtSales, bolSpecMan, strSpecCode)
   'Add By Sindy 2023/6/12
   '檢查當時是否需要為他人職代
   Combo3.Clear
   If txtSales <> strUserNum And txtSales <> "" Then
      Combo3.AddItem txtSales & " " & GetPrjSalesNM(txtSales)
   End If
   Combo3.AddItem strUserNum & " " & strUserName
   Call Pub_SetForOthersEmpCombo(strUserNum, Combo3, False, m_strListPer)
   If m_strListPer = "" Then
      Combo3.Visible = False
   Else
      'Add by Amy 2020/03/25 判斷下拉選單是否有區主管
      If InStr(m_strListPer, GetDeptMan(stST15)) > 0 Then
         bolAreaMan = True
      End If
      Combo3.Visible = True
      Combo3.ListIndex = 0
      'Added by Lydia 2021/05/20 Form 2.0物件無法覆蓋Form 1.0
      txtSales.Visible = False
      lblSalesName.Visible = False
   End If
   '2023/6/12 END
   
'   txtSalesArea.Enabled = False
'   txtSalesArea1.Enabled = False
'   txtSales.Enabled = False
'   Select Case strUserNum
'       '外商陳經理可看全所CFT,FCT,S,CFC
'      Case "68005"
'         txtSalesArea.Enabled = True
'         txtSalesArea1.Enabled = True
'         txtSales.Enabled = True
'         txtSalesArea = stST15
'         txtSalesArea1 = stST15
''cancel by sonia 2014/6/9
''       '蔣律師可看中所全部
''      Case "79037"
''         txtSalesArea.Enabled = True
''         txtSalesArea1.Enabled = True
''         txtSales.Enabled = True
''         txtSalesArea = "S2"
''         txtSalesArea1 = "S29"
''end 2014/6/9
'       '小真,杜副總可看全部
'       'modify by sonia 2014/6/9 +美珍77027
'       'Modify by Amy 2015/02/04 拿掉美珍 改寫至特殊人員(總經理業務工作代理人員)
'      Case "65001", "68006"
'         txtSalesArea.Enabled = True
'         txtSalesArea1.Enabled = True
'         txtSales.Enabled = True
'
'         If strUserNum = "68006" Then
'            txtSalesArea = "S"
'            txtSalesArea1 = "S99"
'         End If
'      '杜燕文,劉大愛可看S31
'      'modify by sonia 2016/8/22劉大愛78007改為蘇嫄媛79053
'      Case "74018", "79053"
'         txtSalesArea = "S31"
'         txtSalesArea1 = "S31"
'         txtSales.Enabled = True
'
'      '王協理可看專利處
'      Case "71011"
'         txtSalesArea = "P10"
'         txtSalesArea1 = "P19"
'         txtSales.Enabled = True
'      '葉經理可看商標處
'      'modify by sonia 2016/2/24 +69008
'      Case "67002", "69008"
'         txtSalesArea = "P20"
'         txtSalesArea1 = "P29"
'         txtSales.Enabled = True
'      'add by sonia 2016/12/21 柄佑可看中所全部但業務區仍預設自已部門
'      Case "82026"
'         txtSalesArea.Enabled = True
'         txtSalesArea1.Enabled = True
'         txtSales.Enabled = True
'         txtSalesArea = stST15
'         txtSalesArea1 = stST15
'         txtSales = strUserNum
'      'end 2016/12/21
'      Case Else
'         Select Case stST05
'            '電腦中心,財務,總經理看全部
'            '2015/7/28 MODIFY BY SONIA +主任秘書(等級08)可看全部
'            Case "00", "01", "08"
'               txtSalesArea.Enabled = True
'               txtSalesArea1.Enabled = True
'               txtSales.Enabled = True
'
'            '各區主管
'            Case "SM"
'               txtSalesArea.Locked = True
'               txtSalesArea1.Locked = True
'               If strUserNum = "71003" Then
'                  txtSalesArea = "S23"
'                  txtSalesArea1 = "S23"
'                  txtSalesArea.Locked = False
'                  txtSalesArea1.Locked = False
'                  txtSalesArea.Enabled = True
'                  txtSalesArea1.Enabled = True
'               Else
'               If strUserNum = "69005" Then
'                  txtSalesArea = "S15"
'                  txtSalesArea1 = "S15"
'                  txtSalesArea.Locked = False
'                  txtSalesArea1.Locked = False
'                  txtSalesArea.Enabled = True
'                  txtSalesArea1.Enabled = True
'               Else
'                  txtSalesArea = stST15
'                  txtSalesArea1 = stST15
'               End If
'               End If
'               txtSales.Enabled = True
'            Case "21", "26", "28"
'               txtSalesArea = stST15
'               txtSalesArea1 = stST15
'               txtSales = strUserNum
'               txtSales.Enabled = True
'            'Added by Morgan 2015/10/2 管理部分所人員開放可看該所資料
'            Case "NM"
'               txtSalesArea = "S" & pub_strUserOffice
'               txtSalesArea1 = "S" & pub_strUserOffice & "9"
'               txtSales = strUserNum
'               txtSales.Enabled = True
'
'            '其他只能看自己
'            Case Else
'               txtSalesArea = stST15
'               txtSalesArea1 = stST15
'               txtSales = strUserNum
'         End Select
'   End Select
'
'   '若操作人員的ST05=SA且在職員工的ST52有該編號存在,此類人員稱之為帶人主管,開放智權人員欄位可輸入
'   If PUB_GetST05Limits(strUserNum) = True Then
'      txtSales.Enabled = True
'   End If
'   '中三區人員可單獨下68096看期限
'   If PUB_GetStaffST15(strUserNum, "1") = "S23" Then
'      txtSales.Enabled = True
'   End If
'   '北五區人員可單獨下10051看期限
'   If PUB_GetStaffST15(strUserNum, "1") = "S15" Then
'      txtSales.Enabled = True
'   End If
'
'   'Add by Amy 2015/02/04 +總經理業務工作代理人員
'   If CheckLevel(strUserNum, "總經理業務工作代理人員") = True Then
'       bolSpecMan = True
'       strSpecCode = "總經理業務工作代理人員"
'   'Modify  by Amy 2014/05/21 開放專利處部份智權同仁資料給彥葶代為處理
'   ElseIf CheckLevel(strUserNum, "A8") = True Then
'        bolSpecMan = True
'        strSpecCode = "A8"
'   End If
'
'   If bolSpecMan = True Then
'        'Add by Amy 2015/02/04 +總經理業務工作代理人員
'        If InStr(strSpecCode, "總經理業務工作代理人員") > 0 Then
'            txtSalesArea.Enabled = True: txtSalesArea = ""
'            txtSalesArea1.Enabled = True: txtSalesArea1 = ""
'            txtSales.Enabled = True
'        End If
'        If InStr(strSpecCode, "A8") > 0 Then txtSales.Enabled = True: txtSales = ""
'   Else
'        txtSales = strUserNum
'   End If
'   'end 2014/05/21

   bolSelData = False
   '從 Unload 移來(因畫面沒離開時沒寫Log會造成逾時重新登入後重複執行)
   PUB_AddExcuteLog Me.Name
   
   CmdDetail.Enabled = False
 
'   'Add By Sindy 2016/5/6 記錄原操作人可以查詢的業務區及所別
'   'txtZone.Tag = txtZone
'   txtSalesArea.Tag = txtSalesArea
'   txtSalesArea1.Tag = txtSalesArea1
'   '2016/5/6 END
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm210142 = Nothing
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
   CloseIme
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtSales_Change()
   If Len(txtSales) > 4 Then
      lblSalesName = GetStaffName(txtSales, True)
   Else
      lblSalesName = ""
   End If
End Sub

Private Sub txtSales_GotFocus()
   TextInverse txtSales
   CloseIme
   If Combo3.Enabled = True And Combo3.Visible = True Then Combo3.SetFocus 'Add By Sindy 2023/6/12
End Sub

Private Sub txtSales_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtSales_LostFocus()
   If Trim(txtSales) = "" Then
       lblSalesName = ""
   End If
End Sub

Private Sub txtSalesArea_GotFocus()
   TextInverse txtSalesArea
   CloseIme
End Sub

Private Sub txtCloseDate_GotFocus(Index As Integer)
   TextInverse txtCloseDate(Index)
   CloseIme
End Sub

Private Sub txtSalesArea_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtSalesArea1_GotFocus()
   TextInverse txtSalesArea1
   CloseIme
End Sub

Private Sub txtSalesArea1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtSales_Validate(Cancel As Boolean)
   'Modify By Sindy 2024/9/18 智權部:使用智權人員ID查詢權限; 改為共用函數
   If PUB_txtSales_Limit(txtSales, m_strListPer, , txtSalesArea, txtSalesArea1, _
                         bolSpecMan, strSpecCode, lblSalesName) = False Then
      If txtSales.Visible = True Then 'Added by Lydia 2021/05/20 排除隱藏
        txtSales.SetFocus
        txtSales_GotFocus
      End If 'Added by Lydia 2021/05/20
      Cancel = True
      Exit Sub
   End If
   '2024/9/18 END
End Sub

'Added by Lydia 2017/11/16 增加客戶編號
Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
   If Trim(Text1(Index).Text) = "" Then Exit Sub
   
   Text1(Index).Text = Trim(Text1(Index).Text)
   
   Select Case Index
       Case 1
          If Text1(0).Text <> "" Then
             Text1(2).Text = "0"
             Text1(3).Text = "00"
          End If
       Case 4 '客戶編號-起
          Text1(4).Text = Mid(Text1(4).Text & String(9, "0"), 1, 9)
          If Right(Text1(4).Text, 3) = "000" Then
             Text1(5).Text = Mid(Text1(4).Text, 1, 6) & "ZZZ"
          Else
             Text1(5).Text = Text1(4).Text
          End If
       Case 5 '客戶編號-止
          If Text1(4).Text <> "" And Text1(5).Text < Text1(5).Text Then
             MsgBox "客戶編號止不可小於起始編號 !", vbExclamation
             Text1(5).SetFocus
             Text1_GotFocus 5
             Cancel = True
             Exit Sub
          End If
   End Select
End Sub

Private Sub txtA0k04_GotFocus()
   TextInverse txtA0k04
End Sub

'Add By Sindy 2023/6/12
Private Sub Combo3_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub Combo3_LostFocus()
   If Trim(Combo3) <> "" And Trim(Combo3) <> "全部" Then
      arrID = Split(Combo3, " ")
      txtSales = arrID(0)
      Combo3 = txtSales & " " & GetPrjSalesNM(txtSales)
   ElseIf Trim(Combo3) <> "全部" Then
      txtSales = ""
   End If
End Sub
Private Sub Combo3_Validate(Cancel As Boolean)
Dim strEmp As String
Dim stTmp As String 'Add by Amy 2020/03/25
   
   If Combo3 <> "" And Trim(Combo3) <> "全部" Then
      'Add by Amy 2020/03/25 只能輸入下拉選單中已有的人員
      stTmp = Combo3
      '直接輸員編未串名字會錯
      If InStr(stTmp, " ") > 0 Then
        stTmp = Mid(stTmp, 1, Val(InStr(stTmp, " ")) - 1)
      Else
        stTmp = Combo3
      End If
      'Modify By Sindy 2020/6/15 Mark
'      If InStr(m_strListPer, stTmp) = 0 And stTmp <> strUserNum And Pub_StrUserSt03 <> "M51" Then
'         MsgBox "不可輸入下拉選單以外的人員！"
'         Cancel = True
'         Combo3.SetFocus
'         Exit Sub
'      End If
      'end 2020/03/25
      arrID = Split(Combo3, " ")
      txtSales = arrID(0)
      lblSalesName.Caption = GetStaffName(txtSales, True)
      If lblSalesName.Caption = "" Then
         MsgBox "智權人員輸入錯誤！", vbCritical
         Combo3.SetFocus
         Cancel = True
      End If
      Combo3 = txtSales & " " & GetPrjSalesNM(txtSales)
   'Modify By Sindy 2024/8/5 mark; 因 txtSales_Validate 會檢查相關的權限
   Else
      txtSales = ""
'   'Modify by Amy 2023/05/09 +st05
'   ElseIf Combo3 = MsgText(601) And stST05 <> "00" And stST05 <> "01" And stST05 <> "08" Then
'        'Add by Amy 2020/03/25 下拉選單無區主管智權人員不可為空
'        'Modify By Sindy 2020/7/14
'        'If bolAreaMan = False And Pub_StrUserSt03 <> "M51" Then
'        'Modify By Sindy 2023/9/21 開放杜協理權限 + And Not (Mid(txtSalesArea, 1, 1) = "S" And Mid(txtSalesArea1, 1, 1) = "S" And InStr(Pub_GetSpecMan("全所智權部主管"), strUserNum) > 0)
'        If (GetDeptMan(txtSalesArea) <> strUserNum Or GetDeptMan(txtSalesArea1) <> strUserNum) _
'            And Pub_StrUserSt03 <> "M51" _
'            And Not (Mid(txtSalesArea, 1, 1) = "S" And Mid(txtSalesArea1, 1, 1) = "S" And InStr(Pub_GetSpecMan("全所智權部主管"), strUserNum) > 0) Then
'        '2020/7/14 END
'           MsgBox "非區主管職代智權人員不可空白！"
'           Cancel = True
'           Combo3.SetFocus
'           Exit Sub
'        End If
'        'end 2020/03/25
   '2024/8/5 END
   End If
   'end 2016/6/7
End Sub
'2023/6/12 END
