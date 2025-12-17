VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc11g0 
   AutoRedraw      =   -1  'True
   Caption         =   "扣繳稅款沖轉作業"
   ClientHeight    =   5470
   ClientLeft      =   60
   ClientTop       =   350
   ClientWidth     =   9410
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5470
   ScaleWidth      =   9410
   Begin VB.TextBox txtSales 
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
      Left            =   1200
      MaxLength       =   6
      TabIndex        =   3
      Top             =   930
      Width           =   1065
   End
   Begin VB.CommandButton cmdSearch1 
      Caption         =   "相似尋找"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   3592
      TabIndex        =   22
      Top             =   960
      Width           =   1100
   End
   Begin VB.CommandButton cmdLikeSearch 
      Caption         =   "相似搜尋"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   8115
      TabIndex        =   19
      Top             =   255
      Width           =   1100
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "尋找"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   8115
      TabIndex        =   18
      Top             =   600
      Width           =   1100
   End
   Begin VB.TextBox txtCustNo 
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
      Index           =   1
      Left            =   3120
      MaxLength       =   9
      TabIndex        =   2
      Top             =   555
      Width           =   1572
   End
   Begin VB.TextBox txtCustNo 
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
      Index           =   0
      Left            =   1200
      MaxLength       =   9
      TabIndex        =   1
      Top             =   555
      Width           =   1572
   End
   Begin VB.TextBox Text4 
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
      Height          =   300
      Left            =   8424
      TabIndex        =   17
      Top             =   5040
      Width           =   735
   End
   Begin VB.TextBox Text3 
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
      Height          =   300
      Left            =   2112
      TabIndex        =   12
      Top             =   1740
      Width           =   3492
   End
   Begin VB.CommandButton Command2 
      Caption         =   "收回金額輸入"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   7650
      TabIndex        =   7
      Top             =   960
      Width           =   1560
   End
   Begin VB.ComboBox Combo1 
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
      Left            =   1272
      TabIndex        =   5
      Top             =   1740
      Width           =   852
   End
   Begin VB.TextBox Text2 
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
      Left            =   1200
      MaxLength       =   3
      TabIndex        =   4
      Top             =   1305
      Width           =   612
   End
   Begin VB.TextBox Text5 
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
      Height          =   300
      Left            =   7536
      TabIndex        =   11
      Top             =   5055
      Width           =   735
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
      Height          =   300
      Left            =   5928
      TabIndex        =   10
      Top             =   5070
      Width           =   735
   End
   Begin VB.TextBox Text9 
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
      Height          =   300
      Left            =   5028
      TabIndex        =   9
      Top             =   5070
      Width           =   735
   End
   Begin VB.TextBox Text8 
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
      Height          =   300
      Left            =   3432
      TabIndex        =   8
      Top             =   5070
      Width           =   735
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "Frmacc11g0.frx":0000
      Height          =   2895
      Left            =   0
      TabIndex        =   6
      Top             =   2070
      Width           =   9390
      _ExtentX        =   16563
      _ExtentY        =   5098
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
      ColumnCount     =   13
      BeginProperty Column00 
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
      BeginProperty Column01 
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
      BeginProperty Column02 
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
         DataField       =   "a1v11"
         Caption         =   "沖轉稅額"
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
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
            Alignment       =   2
            ColumnWidth     =   629.858
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   629.858
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   920.126
         EndProperty
         BeginProperty Column03 
            Locked          =   -1  'True
            ColumnWidth     =   840.189
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   810.142
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   840.189
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   819.78
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   840.189
         EndProperty
         BeginProperty Column08 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   819.78
         EndProperty
         BeginProperty Column09 
            Alignment       =   1
            ColumnWidth     =   840.189
         EndProperty
         BeginProperty Column10 
            Alignment       =   1
            ColumnWidth     =   819.78
         EndProperty
         BeginProperty Column11 
            Locked          =   -1  'True
            ColumnWidth     =   1569.827
         EndProperty
         BeginProperty Column12 
            Locked          =   -1  'True
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
      Left            =   0
      Top             =   1950
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
   Begin MSForms.ComboBox cboTitle 
      Height          =   300
      Left            =   1200
      TabIndex        =   0
      Top             =   225
      Width           =   6780
      VariousPropertyBits=   679495707
      BackColor       =   12648447
      DisplayStyle    =   3
      Size            =   "13652;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblSales 
      Height          =   330
      Left            =   2325
      TabIndex        =   24
      Top             =   960
      Width           =   1155
      VariousPropertyBits=   19
      Caption         =   "LblFM2"
      Size            =   "2037;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "智權人員"
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
      Left            =   225
      TabIndex        =   23
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "客戶編號"
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
      Left            =   225
      TabIndex        =   21
      Top             =   585
      Width           =   975
   End
   Begin VB.Label Label7 
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
      Left            =   2880
      TabIndex        =   20
      Top             =   585
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   -75
      Top             =   5010
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "收據抬頭"
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
      Left            =   225
      TabIndex        =   16
      Top             =   240
      Width           =   975
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
      Left            =   315
      TabIndex        =   15
      Top             =   1740
      Width           =   975
   End
   Begin VB.Label Label5 
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
      Left            =   1125
      TabIndex        =   14
      Top             =   5070
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   1575
      Left            =   105
      Top             =   150
      Width           =   9210
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "扣繳年度"
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
      Left            =   225
      TabIndex        =   13
      Top             =   1320
      Width           =   975
   End
End
Attribute VB_Name = "Frmacc11g0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/11/30 Form2.0已修改
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/11/26 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit

Public adoquery As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Public adocheck As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Dim intY As Integer
'Dim strSql As String
Dim strRNo As String
Dim intCconfirm As Integer
Dim strSql0k As String, strSql1k As String, str1vCon As String
Const m_ACC1V0 = "a1v01,a1v02,a1v03,a1v04,a1v05,a1v06,a1v07,a1v08,a1v09,a1v10,a1v11,a1v12,a1v13"


Private Sub cboTitle_GotFocus()
   'edit by nickc 2007/06/11  切換輸入法改用API
   'cboTitle.IMEMode = 1
   OpenIme
End Sub

Private Sub Command2_Click()
   FeeShow
   If intCconfirm = vbCancel Then
      If strRNo <> "" Then
         strRNo = "(" & Mid(strRNo, 1, Len(strRNo) - 1) & ")"
         adoTaie.Execute "update acc1v0 set a1v08 = null where a1v08 is not null and a1v01 in " & strRNo
         strRNo = ""
      End If
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   
'Modify by Morgan 2003/12/15
'   If Text1 <> MsgText(601) Then
'      strItemNo = Text1 & Text2
'   Else
'      MsgBox MsgText(10), , MsgText(5)
'      Text1.SetFocus
'      Exit Sub
'   End If
   If cboTitle.Text <> MsgText(601) Then
      strItemNo = cboTitle.Text & Text2
   Else
      MsgBox MsgText(10), , MsgText(5)
      cboTitle.SetFocus
      Exit Sub
   End If
'End 2003/12/15
   
   'Modify by Morgan 2005/5/3 要加調整稅款
   'If Text4 <> MsgText(601) And Val(Text4) <> 0 Then
   '   strCon1 = Val(Text4)
   If Text4 <> MsgText(601) And Val(Text4) <> 0 And Val(Text4) + Val(Text5) > 0 Then
      strCon1 = Val(Text4) + Val(Text5)
      
   Else
      MsgBox MsgText(115), , MsgText(5)
      Exit Sub
   End If
   adocheck.CursorLocation = adUseClient
   'Modify By Sindy 2016/1/6 + acc1k0
   'modify by sonia 2021/1/28 +a0k11抓公司別
   adocheck.Open "select a0k03, a0k20, a0k01, a0k11 from acc0k0 where a0k01 = '" & Adodc1.Recordset.Fields("a1v02").Value & "'" & _
                 " union " & _
                 "select a1k28, cp13, a1k01, a1k37 from acc1k0,caseprogress where a1k01 = '" & Adodc1.Recordset.Fields("a1v02").Value & "' and cp09='" & Adodc1.Recordset.Fields("a1v01").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adocheck.RecordCount <> 0 Then
      If IsNull(adocheck.Fields(0).Value) Then
         strCon2 = MsgText(601)
      Else
         strCon2 = adocheck.Fields(0).Value
      End If
      If IsNull(adocheck.Fields(1).Value) Then
         strCon3 = MsgText(601)
      Else
         strCon3 = adocheck.Fields(1).Value
      End If
      If IsNull(adocheck.Fields(2).Value) Then
         strCon4 = MsgText(601)
      Else
         strCon4 = adocheck.Fields(2).Value
      End If
   End If
   'add by sonia 2021/1/28
   Frmacc11g1.strCompNo = "1"
   If "" & adocheck.Fields("a0k11").Value = "L" Then
      Frmacc11g1.strCompNo = "" & adocheck.Fields("a0k11").Value
   End If
   'end 2021/1/28
   adocheck.Close
   If strRNo <> "" Then
      If Right(strRNo, 1) = "," Then
         strCon5 = "(" & Mid(strRNo, 1, Len(strRNo) - 1) & ")"
      Else
         strCon5 = "(" & strRNo & ")"
      End If
   End If
   strFormLink = Name
   tool3_enabled
   Frmacc11g1.Show
   Me.Enabled = False
   Screen.MousePointer = vbDefault
End Sub

Private Sub DataGrid2_Click()
   Select Case DataGrid2.col
      Case 0
         If DataGrid2.Columns(0).Text = MsgText(602) Then
            SendKeys "{BACKSPACE}"
            SendKeys "{DEL}"
        Else
            SendKeys "{Y}"
            SendKeys "{ENTER}"
         End If
   End Select
End Sub

Private Sub DataGrid2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   Select Case DataGrid2.col
      Case 0
         If KeyAscii = 89 Then
            intY = KeyAscii
         End If
   End Select
End Sub

Private Sub Form_Activate()
   If strItemNo <> MsgText(601) Then
      AdodcRefresh
      SumShow
   End If
   strFormName = Name
   strFormLink = MsgText(601)
   strItemNo = ""
   strCon1 = ""
   strCon2 = ""
   strCon3 = ""
   strCon4 = ""
   strCon6 = ""
   strCon7 = ""
   'Add By Cheng 2003/06/30
   '預設扣繳年度
   'Modify by Morgan 2004/4/8
   '預設年度改判斷4月
   'If Val(Right(ServerDate, 4)) >= 501 Then
   If Val(Right(strSrvDate(2), 4)) >= 401 Then
      Text2 = strSrvDate(2) \ 10000
   Else
      Text2 = strSrvDate(2) \ 10000 - 1
   End If
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(157) & " / " & MsgText(98)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(157) & " / " & MsgText(98)
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 9645
   Me.Height = 6060
   'Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   MoveFormToCenter Me
   lblSales.Caption = ""
   Image1 = LoadPicture(strBackPicPath1)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   strItemNo = MsgText(601)
   OpenTable
   adoquery.CursorLocation = adUseClient
   adoquery.Open "select a0801 from acc080 order by a0801 asc", adoTaie, adOpenStatic, adLockReadOnly
   Do While adoquery.EOF = False
      Combo1.AddItem adoquery.Fields(0).Value
      adoquery.MoveNext
   Loop
   adoquery.Close
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(157) & " / " & MsgText(98)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If strRNo <> "" Then
      strRNo = "(" & Mid(strRNo, 1, Len(strRNo) - 1) & ")"
      adoTaie.Execute "update acc1v0 set a1v08 = null where a1v08 is not null and a1v01 in " & strRNo
      strRNo = ""
   End If
   StatusClear
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc11g0 = Nothing
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
'   strSql = strSql & " and a0k11<>'J'" 'Add By Sindy 2013/12/30
'   Select Case Combo1
'      Case "1", "2"
'         strSql = strSql & " and a0k11 in ('1', '2')"
'      Case Else
'         strSql = strSql & " and a0k11 = '" & Combo1 & "'"
'   End Select
   adoadodc1.CursorLocation = adUseClient
   'adoadodc1.Open "select * from acc0k0, acc1v0 where a0k01 = a1v02 and a1v16 is null and a1v15 is null and a0k01 = 'Z'" & strSql & " order by a1v03 asc, a1v01 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoadodc1.Open "select * from acc0k0, acc1v0 where a0k01 = a1v02 and a1v16 is null and a1v15 is null and a0k01 = 'Z'" & " order by a1v03 asc, a1v01 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Set Adodc1.Recordset = adoadodc1
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
         SumShow
   End Select
   KeyEnter KeyCode
End Sub

'*************************************************
'  重新整理 Adodc 之資料
'
'*************************************************
Public Sub AdodcRefresh(Optional ByVal bolExact As Boolean = True)
   
On Error GoTo Checking
   
   Screen.MousePointer = vbHourglass
   strSql0k = MsgText(601)
   strSql1k = MsgText(601)
   str1vCon = MsgText(601)
   
'Modify by Morgan 2003/12/17
'   If Text1 <> MsgText(601) Then
'      'strSQL = " and instr(a0k04, '" & Text1 & "') > 0"
'      strSQL = " and a0k04 = '" & Text1 & "'"
'   End If
   '收據抬頭
   If cboTitle.Text <> MsgText(601) Then
      If bolExact = True Then
         strSql0k = " and a0k04= '" & cboTitle.Text & "'"
         strSql1k = " and a1k35= '" & cboTitle.Text & "'"
      Else
         strSql0k = " and instrb(a0k04, '" & cboTitle.Text & "') = 1"
         strSql1k = " and instrb(a1k35, '" & cboTitle.Text & "') = 1"
      End If
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
'End 2003/12/17
   
   '扣繳年度
   If Text2 <> MsgText(601) Then
      'strSql0k = strSql0k & " and a0k16 = " & Val(Text2) & ""
      str1vCon = str1vCon & " and a1v09 = " & Val(Text2) & ""
   End If
   strSql0k = strSql0k & " and a0k11<>'J'" 'Add By Sindy 2013/12/30
   
   '公司別
   If Combo1 <> MsgText(601) Then
      Select Case Combo1
         Case "1", "2"
            'strSql0k = strSql0k & " and a0k11 in ('1', '2')"
            str1vCon = str1vCon & " and a1v03 in ('1', '2')"
         Case Else
            'strSql0k = strSql0k & " and a0k11 = '" & Combo1 & "'"
            str1vCon = str1vCon & " and a1v03 = '" & Combo1 & "'"
      End Select
   End If
   adoadodc1.Close
   adoadodc1.CursorLocation = adUseClient
   'adoadodc1.Open "select * from acc0k0, acc1v0 where a0k01 = a1v02 and a1v10 < 0 and a1v16 is null" & StrSQL & " order by a1v03 asc, a1v01 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   'Modify By Sindy 2015/11/10
   'adoadodc1.Open "select * from acc0k0, acc1v0 where a0k01 = a1v02 and a1v16 is null and a1v15 is null and a1v06 <> 0" & strSql & " order by a1v03 asc, a1v01 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoadodc1.Open "select " & m_ACC1V0 & " from acc0k0, acc1v0 where a0k01 = a1v02 and a1v16 is null and a1v15 is null and a1v06 <> 0" & str1vCon & strSql0k & _
                  " union select " & m_ACC1V0 & " from acc1k0, acc1v0, caseprogress where a1k01 = a1v02 and a1v01=cp09(+) and cp09 is not null and a1v16 is null and a1v15 is null and a1v06 <> 0" & str1vCon & strSql1k & _
                  " order by a1v03 asc, a1v01 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   '2015/11/10 END
   Adodc1.Recordset.Requery
   'If Adodc1.Recordset.RecordCount <> 0 Then
   '   If IsNull(Adodc1.Recordset.Fields("a0k04").Value) = False Then
   '      Text1 = Trim(Adodc1.Recordset.Fields("a0k04").Value)
   '   End If
   'End If
   If adoadodc1.State = adStateOpen Then
      If adoadodc1.RecordCount = 0 Then
         MsgBox MsgText(28), , MsgText(5)
      End If
   Else
      MsgBox MsgText(28), , MsgText(5)
   End If
   Screen.MousePointer = vbDefault
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(157) & " / " & MsgText(98)
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  產生扣繳明細資料
'
'*************************************************
'Removed by Morgan 2011/11/24 一定會有 1v0 的資料,無須再重新產生,若將來有需要重新產生則要改逐筆抓收款資料
'Public Sub ProcessData(Optional ByVal bolExact As Boolean = True)
'   Screen.MousePointer = vbHourglass
'   strSql = MsgText(601)
'
''Modify by Morgan 2003/12/17
''   If Text1 <> MsgText(601) Then
''      strSQL = " and instrb(a0k04, '" & Text1 & "') = 1"
''   End If
'
'   If cboTitle.Text <> MsgText(601) Then
'      If bolExact = True Then
'         strSql = " and a0k04= '" & cboTitle.Text & "'"
'      Else
'         strSql = " and instrb(a0k04, '" & cboTitle.Text & "') = 1"
'      End If
'   End If
'
'   If txtCustNo(0) <> "" Then
'      strSql = strSql & " and a0k03>='" & txtCustNo(0).Text & "'"
'   End If
'   If txtCustNo(1) <> "" Then
'      strSql = strSql & " and a0k03<='" & txtCustNo(1).Text & "'"
'   End If
'
'   If txtSales <> "" Then
'      strSql = strSql & " and a0k20||''='" & txtSales & "'"
'   End If
''End 2003/12/17
'
'   If Text2 <> MsgText(601) Then
'      strSql = strSql & " and a0k16 = " & Val(Text2) & ""
'   End If
'   If Combo1 <> MsgText(601) Then
'      Select Case Combo1
'         Case "1", "2"
'            strSql = strSql & " and a0k11 in ('1', '2')"
'         Case Else
'            strSql = strSql & " and a0k11 = '" & Combo1 & "'"
'      End Select
'   End If
'   adoquery.CursorLocation = adUseClient
'   adoquery.Open "select cp09, cp60, a0k11, decode(a0k30, 'Y', (nvl(cp16, 0) - nvl(cp77, 0)) * 0.1, (nvl(cp16, 0) - nvl(cp17, 0)) * 0.1) as TAmount, cp16, cp75, a0k16, a0j20, a0j21, a0k04, decode(a0k30, 'Y', (nvl(cp16, 0) - nvl(cp77, 0)) * 0.1 - nvl(cp76, 0), (nvl(cp16, 0) - nvl(cp17, 0)) * 0.1  - nvl(cp76, 0)) as RAmount, cp76, a0k13 from acc0k0, caseprogress, acc0j0 where a0k01 = cp60 (+) and a0k01 = a0j13 (+) and cp09 is not null" & strSql & " order by a0k11 asc, cp09 asc", adoTaie, adOpenStatic, adLockReadOnly
'   If adoquery.RecordCount <> 0 Then
'      If IsNull(adoquery.Fields("a0k04").Value) = False Then
''Remove by Morgan 2003/12/17
'''Modify by Morgan 2003/12/15
'''         Text1 = Trim(adoquery.Fields("a0k04").Value)
''         cboTitle.Clear
''         cboTitle.AddItem Trim(adoquery.Fields("a0k04").Value)
''         cboTitle.ListIndex = 0
'''End 2003/12/15
''Remove 2003/12/17
'      End If
'   End If
'   Do While adoquery.EOF = False
'      With adoquery
'         adocheck.CursorLocation = adUseClient
'         adocheck.Open "select a1v01 from acc1v0 where a1v01 = '" & .Fields("cp09").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
'         If adocheck.RecordCount = 0 Then
'            adoTaie.Execute "insert into acc1v0 values ('" & .Fields("cp09").Value & "', '" & .Fields("cp60").Value & "', '" & .Fields("a0k11").Value & "', " & .Fields("TAmount").Value & ", '" & IIf(IsNull(.Fields("a0k13").Value) = True Or .Fields("a0k13").Value = "", "N", "Y") & "', " & IIf(IsNull(.Fields("cp76").Value), 0, .Fields("cp76").Value) & ", " & Val(.Fields("RAmount").Value) & ", null, " & .Fields("a0k16").Value & ", 0, 0, '" & .Fields("a0j20").Value & "', '" & .Fields("a0j21").Value & "', null, null, null, null, " & IIf(IsNull(.Fields("cp76").Value) = True Or .Fields("cp76").Value = 0, "null", "'1'") & ")"
'         End If
'         adocheck.Close
'      End With
'      adoquery.MoveNext
'   Loop
'   adoquery.Close
'   adoTaie.Execute "delete from acc1v0 where a1v06 = 0 and a1v07 = 0"
'   Screen.MousePointer = vbDefault
'End Sub

'*************************************************
'  計算並顯示合計
'
'*************************************************
Public Sub SumShow()
   adoaccsum.CursorLocation = adUseClient
   'adoaccsum.Open "select sum(a1v04), sum(a1v06), sum(a1v07), sum(a1v11), sum(a1v10) from acc0k0, acc1v0 where a0k01 = a1v02 and a1v10 < 0 and a1v16 is null" & StrSQL, adoTaie, adOpenStatic, adLockReadOnly
   'Modify By Sindy 2015/11/10
   'adoaccsum.Open "select sum(a1v04), sum(a1v06), sum(a1v07), sum(a1v11), sum(a1v10) from acc0k0, acc1v0 where a0k01 = a1v02 and a1v16 is null and a1v15 is null and a1v06 <> 0" & strSql, adoTaie, adOpenStatic, adLockReadOnly
   adoaccsum.Open "select sum(a1v04), sum(a1v06), sum(a1v07), sum(a1v11), sum(a1v10) from (" & _
                  "select a0k01, a1v01, a1v02, a1v04, a1v06, a1v07, a1v11, a1v10 from acc0k0, acc1v0 where a0k01 = a1v02 and a1v16 is null and a1v15 is null and a1v06 <> 0" & str1vCon & strSql0k & _
                  " union select a1k01, a1v01, a1v02, a1v04, a1v06, a1v07, a1v11, a1v10 from acc1k0, acc1v0, caseprogress where a1k01 = a1v02 and a1v01=cp09(+) and cp09 is not null and a1v16 is null and a1v15 is null and a1v06 <> 0" & str1vCon & strSql1k & _
                  ")", adoTaie, adOpenStatic, adLockReadOnly
   '2015/11/10 END
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
      If IsNull(adoaccsum.Fields(2).Value) Then
         Text10 = "0"
      Else
         Text10 = adoaccsum.Fields(2).Value
      End If
      If IsNull(adoaccsum.Fields(4).Value) Then
         Text5 = "0"
      Else
         Text5 = adoaccsum.Fields(4).Value
      End If
   Else
      Text8 = "0"
      Text9 = "0"
      Text10 = "0"
      Text4 = "0"
      Text5 = "0"
   End If
   adoaccsum.Close
   adoaccsum.CursorLocation = adUseClient
   'Modify By Sindy 2015/11/10
   'adoaccsum.Open "select sum(a1v11) from acc0k0, acc1v0 where a0k01 = a1v02 and a1v08 is not null and a1v15 is null" & strSql, adoTaie, adOpenStatic, adLockReadOnly
   adoaccsum.Open "select sum(a1v11) from (" & _
                  "select a0k01, a1v01, a1v02, a1v11 from acc0k0, acc1v0 where a0k01 = a1v02 and a1v08 is not null and a1v15 is null" & str1vCon & strSql0k & _
                  " union select a1k01, a1v01, a1v02, a1v11 from acc1k0, acc1v0, caseprogress where a1k01 = a1v02 and a1v01=cp09(+) and cp09 is not null and a1v08 is not null and a1v15 is null" & str1vCon & strSql1k & _
                  ")", adoTaie, adOpenStatic, adLockReadOnly
   '2015/11/10 END
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0).Value) Then
         Text4 = "0"
      Else
         Text4 = adoaccsum.Fields(0).Value
      End If
   Else
      Text4 = "0"
   End If
   adoaccsum.Close
End Sub

Private Sub Combo1_Click()
   Text3 = A0802Query(Combo1)
   AdodcRefresh
   SumShow
End Sub

Private Sub Combo1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub DataGrid2_AfterColUpdate(ByVal ColIndex As Integer)
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   Select Case DataGrid2.col
      Case 0
         If DataGrid2.Columns(0).Text = MsgText(602) Then
            strRNo = strRNo & "'" & DataGrid2.Columns(2).Text & "',"
            'Add By Sindy 2015/11/10
            'Modify By Sindy 2017/3/17 只做請款單Acc1k0=X編號的
            If Left(DataGrid2.Columns(3).Text, 1) = "X" Then
            '2017/3/17 END
               strCon6 = cboTitle.Text '收據抬頭
               strCon7 = DataGrid2.Columns(8).Text '扣繳年度
            End If
            '2015/11/10 END
         End If
   End Select
   If DataGrid2.Columns(0).Text = MsgText(602) Then
      If (Val(DataGrid2.Columns(6).Value) + Val(DataGrid2.Columns(7).Value)) > Val(DataGrid2.Columns(4).Value) Then
         MsgBox MsgText(122), , MsgText(5)
         DataGrid2.Columns(6).Value = 0
      End If
      DataGrid2.Columns(10).Value = DataGrid2.Columns(6).Value
   Else
      DataGrid2.Columns(10).Value = 0
   End If
   Adodc1.Recordset.UpdateBatch
   Select Case DataGrid2.col
      Case 8
         adoTaie.Execute "update acc0k0 set a0k16 = " & Val(Adodc1.Recordset.Fields("a1v09").Value) & " where a0k01 = '" & Adodc1.Recordset.Fields("a1v02").Value & "'"
         adoTaie.Execute "update acc0m0 set a0m07 = " & Val(Adodc1.Recordset.Fields("a1v09").Value) & " where a0m02 = '" & Adodc1.Recordset.Fields("a1v02").Value & "'"
      'Add by Morgan 2004/4/6
      '修改扣繳時要更新cp76
      Case 6
         'Removed by Morgan 2011/11/9 改在11g1更新(考慮拆收據問題且也要和1u0資料一致)
         'adoTaie.Execute "update caseprogress set cp76 = " & Val(Adodc1.Recordset.Fields("a1v06").Value) & " where cp09 = '" & Adodc1.Recordset.Fields("a1v01").Value & "'"
   End Select
   SumShow
End Sub

Private Sub DataGrid2_GotFocus()
   DataGrid2.col = 0
End Sub

Private Sub DataGrid2_KeyUp(KeyCode As Integer, Shift As Integer)
Dim intCounter As Integer

   Select Case KeyCode
      Case vbKeyReturn
         Select Case DataGrid2.col
            Case 0
               For intCounter = 1 To 10
                  SendKeys "{RIGHT}"
               Next intCounter
            Case 10
               SendKeys "{DOWN}"
               For intCounter = 1 To 10
                  SendKeys "{LEFT}"
               Next intCounter
         End Select
   End Select
   KeyDefine KeyCode
End Sub

'*************************************************
'  計算並顯示退費資料
'
'*************************************************
Public Sub FeeShow()
   'Modify by Morgna 2005/5/4
   'intCconfirm = MsgBox(MsgText(125) & Format(Val(Text4), DDollar), vbOKCancel + vbDefaultButton1, MsgText(5))
   intCconfirm = MsgBox(MsgText(125) & Format(Val(Text4) + Val(Text5), DDollar), vbOKCancel + vbDefaultButton1, MsgText(5))
End Sub

Private Sub Text2_GotFocus()
   'edit by nickc 2007/06/11  切換輸入法改用API
   'Text2.IMEMode = 2
   CloseIme
   TextInverse Text2
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
'Add by Morgan 2003/12/16
Private Sub cboTitle_KeyPress(KeyAscii As MSForms.ReturnInteger)
   If txtCustNo(0) <> "" Or txtCustNo(1) <> "" Or cboTitle.ListCount > 0 Then
      txtCustNo(0) = "": txtCustNo(1) = ""
      txtSales = "": lblSales = ""
      cboTitle.Clear
   End If
End Sub
'Add by Morgan 2003/12/15
Private Sub cboTitle_Validate(Cancel As Boolean)
   If CheckLen(Label1, cboTitle, 100) = MsgText(603) Then
      Cancel = True
      Exit Sub
   End If
   'edit by nickc 2007/06/11  切換輸入法改用API
   If Cancel = False Then CloseIme
End Sub
'Add by Morgan 2003/12/15
Private Sub cmdLikeSearch_Click()
   
   If cboTitle.Text = "" Then
      MsgBox "請輸入收據抬頭！", vbCritical
   Else
      txtCustNo(0) = "": txtCustNo(1) = ""  'add by sonia 2023/11/7 輸第二筆時要清前一筆的客戶編號
      'Modify by Morgan 2007/10/2 改呼叫共用函數
      'AddItem2CboTitle
      'Modify by Sindy 2013/12/30
      PUB_AddItem2CboTitle cboTitle, txtCustNo(0), txtCustNo(1), Text2, True
      'end 2007/10/2
   End If
   
End Sub

''Add by Morgan 2003/12/15
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

'Add by Morgan 2003/12/15
Private Sub cmdSearch_Click()
   If CheckLen(Label1, cboTitle, 100) = MsgText(603) Then
      'edit by nickc 2007/02/08
      'MsgBox "收據抬頭不可空白！", vbCriticalv
      MsgBox "收據抬頭不可空白！", vbCritical
      Exit Sub
   End If
   If cboTitle = "" Then
      Exit Sub
   End If
   'Add By Sindy 2013/12/30
   If Combo1.Text = "J" Then
      MsgBox "公司別不可為 J 公司！", vbCritical
      Combo1.SetFocus
      Exit Sub
   End If
   '2013/12/30 END
   
   'ProcessData 'Removed by Morgan 2011/11/24 一定會有 1v0 的資料,無須再重新產生,若將來有需要重新產生則要改逐筆抓收款資料
   AdodcRefresh
   SumShow
End Sub

'Add by Morgan 2003/12/17
Private Sub cmdSearch1_Click()
   If CheckLen(Label1, cboTitle, 100) = MsgText(603) Then
      'edit by nickc 2007/02/08
      'MsgBox "收據抬頭不可空白！", vbCriticalv
      MsgBox "收據抬頭不可空白！", vbCritical
      Exit Sub
   End If
   If cboTitle = "" Then
      Exit Sub
   End If
   'Add By Sindy 2013/12/30
   If Combo1.Text = "J" Then
      MsgBox "公司別不可為 J 公司！", vbCritical
      Combo1.SetFocus
      Exit Sub
   End If
   '2013/12/30 END
   
   'ProcessData False'Removed by Morgan 2011/11/24 一定會有 1v0 的資料,無須再重新產生,若將來有需要重新產生則要改逐筆抓收款資料
   AdodcRefresh False
   SumShow
End Sub

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

'Remove by Morgan 2003/12/15
'Private Sub Text1_GotFocus()
'   TextInverse Text1
'   StatusView MsgText(65) & "100"
'   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(157) & " / " & MsgText(98)
'End Sub
'
'Private Sub Text1_LostFocus()
'   StatusView MsgText(601)
'End Sub
'
'Private Sub Text1_Validate(Cancel As Boolean)
'   If CheckLen(Label1, Text1, 100) = MsgText(603) Then
'      Cancel = True
'      Exit Sub
'   End If
'   ProcessData
'   AdodcRefresh
'   SumShow
'End Sub
