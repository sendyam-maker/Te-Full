VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Frmacc11l1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "收文金額分配查詢"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8730
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   8730
   Begin VB.ComboBox Combo3 
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
      Left            =   5700
      TabIndex        =   3
      Top             =   90
      Width           =   2772
   End
   Begin VB.ComboBox Combo2 
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
      Left            =   2580
      TabIndex        =   2
      Top             =   90
      Width           =   2772
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
      ItemData        =   "Frmacc11l1.frx":0000
      Left            =   180
      List            =   "Frmacc11l1.frx":0002
      TabIndex        =   1
      Top             =   90
      Width           =   2052
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
      Left            =   405
      Top             =   450
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc11l1.frx":0004
      Height          =   4545
      Left            =   90
      TabIndex        =   0
      Top             =   510
      Width           =   8565
      _ExtentX        =   15108
      _ExtentY        =   8017
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      BackColor       =   -2147483624
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   18
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   11.25
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
      Caption         =   "收文金額分配資料"
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "RDate"
         Caption         =   "收文日"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "CaseNo"
         Caption         =   "本所案號"
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
         DataField       =   "RecNo"
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
         DataField       =   "Property"
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
      BeginProperty Column04 
         DataField       =   "Fee"
         Caption         =   "費用"
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
         DataField       =   "OFee"
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
      BeginProperty Column06 
         DataField       =   "Point"
         Caption         =   "點數"
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
         DataField       =   "Pay"
         Caption         =   "已收款"
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
            ColumnWidth     =   884.976
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1560.189
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   824.882
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   764.787
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            ColumnWidth     =   555.024
         EndProperty
         BeginProperty Column07 
            Alignment       =   2
            ColumnWidth     =   764.787
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "~"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5460
      TabIndex        =   5
      Top             =   90
      Width           =   135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "="
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2340
      TabIndex        =   4
      Top             =   90
      Width           =   135
   End
End
Attribute VB_Name = "Frmacc11l1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/14 Form2.0已修改
'Memo By Sonia 2012/12/4 智權人員欄已修改
'Create by Morgan 2011/4/6
Option Explicit

Public adoadodc1 As New ADODB.Recordset


Private Sub Combo2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Combo2_Validate(Cancel As Boolean)
   Combo3 = Combo2
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Form_Load()
   '表單初始化
   PUB_InitForm Me, Me.Width, Me.Height, strBackPicPath2

   strCon1 = "收文日"
   Combo1.AddItem MsgText(31)
   Combo1.AddItem strCon1
   Combo1 = MsgText(31)
   OpenTable
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Adodc1.Recordset.RecordCount <> 0 Then
      strItemNo = Adodc1.Recordset.Fields("RecNo").Value
   Else
      strItemNo = MsgText(601)
   End If
   StatusClear
   tool1_enabled
   Frmacc11l0.Enabled = True
   Frmacc11l0.m_bRefresh = True
   Frmacc11l0.Show
   Set Frmacc11l1 = Nothing
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)
   Select Case KeyCode
      Case vbKeyF12
         QueryData
   End Select
   KeyEnter KeyCode
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub


Private Function GetQuerySQL() As String
   GetQuerySQL = "select distinct substrb(' '||sqldatet(cp05),-9) RDate" & _
      ",cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) CaseNo" & _
      ",cp09 RecNo,cpm03 Property,cp16 Fee,cp17 OFee,cp18 Point,decode(nvl(cp79,0),0,'Y') Pay" & _
      " From acc0n0, caseprogress, casepropertymap" & _
      " where cp09(+)=a0n01 and cpm01(+)=cp01 and cpm02(+)=cp10"

End Function

'*************************************************
'  扣單作廢資料查詢
'
'*************************************************
Private Sub QueryData()
   Dim stCon As String, stSQL As String
   
On Error GoTo Checking
   
   If adoadodc1.State <> adStateClosed Then adoadodc1.Close
   adoadodc1.CursorLocation = adUseClient
   
   stSQL = GetQuerySQL
   
   If Combo1 <> MsgText(31) Then
      Select Case Combo1
         Case strCon1
            strCondition = "RDate"
      End Select
      
      If Combo2 <> MsgText(601) Then
         If Combo3 = MsgText(601) Then
            stCon = stCon & " and " & strCondition & " = '" & Combo2 & "'"
         Else
            stCon = stCon & " and " & strCondition & " >= '" & Combo2 & "'"
         End If
      End If
      
      If Combo3 <> MsgText(601) Then
         stCon = stCon & " and " & strCondition & " <= '" & Combo3 & "'"
      End If
   End If
   
   If strCondition <> "" Then
      stSQL = stSQL & " order by " & strCondition & " asc"
   End If
   
   adoadodc1.Open stSQL, adoTaie, adOpenStatic, adLockReadOnly
   Adodc1.Recordset.Requery
   If Adodc1.Recordset.RecordCount = 0 Then
      MsgBox MsgText(33), , MsgText(5)
   End If
   Exit Sub
   
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   strExc(0) = GetQuerySQL & " and rownum<1"
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open strExc(0), adoTaie, adOpenStatic, adLockReadOnly
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

