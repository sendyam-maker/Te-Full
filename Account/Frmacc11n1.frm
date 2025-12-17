VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Frmacc11n1 
   AutoRedraw      =   -1  'True
   Caption         =   "客戶電匯資料查詢"
   ClientHeight    =   5160
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   8850
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5160
   ScaleWidth      =   8850
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   1590
      MaxLength       =   200
      TabIndex        =   0
      Top             =   120
      Width           =   4110
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc11n1.frx":0000
      Height          =   4095
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   8295
      _ExtentX        =   14623
      _ExtentY        =   7214
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   18
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   11.5
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
      Caption         =   "客戶電匯資料"
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "custname"
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
      BeginProperty Column01 
         DataField       =   "cu04"
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
      BeginProperty Column02 
         DataField       =   "st02"
         Caption         =   "智權人員"
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
         DataField       =   "cu155"
         Caption         =   "電匯資料"
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
      BeginProperty Column04 
         DataField       =   "cu13"
         Caption         =   "cu13"
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
            ColumnAllowSizing=   0   'False
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2560.252
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1349.858
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   7570.205
         EndProperty
         BeginProperty Column04 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
      Left            =   270
      Top             =   480
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
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "註：按ESC鍵，即可離開查詢，進入維護作業！"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   300
      TabIndex        =   4
      Top             =   4740
      Width           =   5265
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "模糊比對"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5730
      TabIndex        =   3
      Top             =   150
      Width           =   960
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "電匯資料："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   315
      TabIndex        =   2
      Top             =   150
      Width           =   1275
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   4800
      Visible         =   0   'False
      Width           =   132
   End
End
Attribute VB_Name = "Frmacc11n1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/14 Form2.0已修改
'Memo By Sonia 2012/12/4 智權人員欄已修改
'Create by Sindy 2012/8/29
Option Explicit

Public adoacc0i0 As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Dim m_ST06 As String
Dim m_PrevForm As Form  'Add by Amy 2017/12/07前畫面


'Add by Amy 2017/12/07
Public Sub DataGrid1_DblClick()
    Dim j As Integer
    Dim strA2330 As String
    Dim strTmp, strGet As String
    
    If TypeName(m_PrevForm) = "Nothing" Then Exit Sub
    '帶入資料
    strA2330 = DataGrid1.Columns(3)
    '只取特取字有的字串
    If InStr(strA2330, ",") > 0 Then
        strTmp = Split(strA2330, ",")
        For j = LBound(strTmp) To UBound(strTmp)
            If InStr(strTmp(j), Text1) > 0 Then
                strGet = strGet & "," & strTmp(j)
            End If
        Next j
        strA2330 = Mid(strGet, 2)
    End If
    m_PrevForm.txtA2304 = DataGrid1.Columns(0)
    m_PrevForm.txtCustomer = DataGrid1.Columns(1)
    m_PrevForm.txtA2303 = DataGrid1.Columns(4)
    m_PrevForm.txtSales = DataGrid1.Columns(2)
    m_PrevForm.txtA2330 = strA2330
    Unload Me
End Sub
'end 2017/12/07

Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
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
   'Modify by Amy 2023/10/05 原:W:8850/H:5400
   Me.Width = 8970
   Me.Height = 5625
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath2)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   
   strCompanyNo = MsgText(601)
   m_ST06 = PUB_GetST06(strUserNum)
   OpenTable
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
   If Adodc1.Recordset.RecordCount <> 0 Then
      strCompanyNo = Adodc1.Recordset.Fields("custname").Value
   Else
      strCompanyNo = MsgText(601)
   End If
   StatusClear
    'Modify by Amy 2017/12/07
    If TypeName(m_PrevForm) <> "Nothing" Then
        m_PrevForm.Enabled = True
        m_PrevForm.Show
    Else
        tool1_enabled
        Frmacc11n0.Enabled = True
        Frmacc11n0.Show
    End If
   'end 2017/12/07
   Set Frmacc11n1 = Nothing
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
'Modify by Amy 2017/12/07 原:Private
Public Sub KeyDefine(KeyCode As Integer)
   Select Case KeyCode
      Case vbKeyF12
         Acc0i0Query
   End Select
   KeyEnter KeyCode
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

'*************************************************
'  客戶資料查詢
'
'*************************************************
Private Sub Acc0i0Query()
Dim strCon1 As String, strCon2 As String
   
On Error GoTo Checking
   
   Screen.MousePointer = vbHourglass
   strCon1 = "": strCon2 = ""
   If Text1 <> "" Then
      strCon1 = strCon1 & " and instr(NLS_Upper(CU155),'" & UCase(ChgSQL(Text1)) & "')>0"
      strCon2 = strCon2 & " and instr(NLS_Upper(FA114),'" & UCase(ChgSQL(Text1)) & "')>0"
   End If
   If m_ST06 <> "1" Then '若為分所人員不顯示代理人資料
      strCon2 = strCon2 & " and fa01='null' and fa02='null'"
   End If
   'Modify  by Amy 2017/12/07 +cu13
   strSql = "select cu01||cu02 as custname,NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) as cu04,st02,cu155,cu13 from customer,staff where cu155>' ' and cu155 is not null and cu13=st01(+)" & strCon1 & _
            " Union" & _
            " select fa01||fa02 as custname,NVL(fa04,NVL(fa05||fa63||fa64||fa65,fa06)) as cu04,' ',fa114 as cu155,'' as cu13 from fagent where fa114>' ' and fa114 is not null" & strCon2 & _
            " order by custname asc"
   adoadodc1.Close
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
   Adodc1.Recordset.Requery
   If Adodc1.Recordset.RecordCount = 0 Then
      MsgBox MsgText(33), , MsgText(5)
   End If
   Screen.MousePointer = vbDefault
   Exit Sub
   
Checking:
   Screen.MousePointer = vbDefault
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
Dim strCon2 As String
   
On Error GoTo Checking
   
   adoadodc1.CursorLocation = adUseClient
   strCon2 = ""
   If m_ST06 <> "1" Then '若為分所人員不顯示代理人資料
      strCon2 = strCon2 & " and fa01='null' and fa02='null'"
   End If
   'Modify by Amy 2017/12/07 +cu13
   strSql = "select cu01||cu02 as custname,NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) as cu04,st02,cu155,cu13 from customer,staff where cu155>' ' and cu155 is not null and cu13=st01(+)" & _
            " Union" & _
            " select fa01||fa02 as custname,NVL(fa04,NVL(fa05||fa63||fa64||fa65,fa06)) as cu04,' ',fa114 as cu155,'' as cu13 from fagent where fa114>' ' and fa114 is not null" & strCon2 & _
            " order by custname asc"
   adoadodc1.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

Private Sub Text1_Click()
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

Private Sub Text1_GotFocus()
   InverseTextBox Text1
   OpenIme
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

'Add by Amy 2017/12/07
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

