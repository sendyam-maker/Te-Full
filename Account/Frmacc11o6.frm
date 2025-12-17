VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc11o6 
   AutoRedraw      =   -1  'True
   Caption         =   "特殊發票客戶資料查詢"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9150
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5550
   ScaleWidth      =   9150
   Begin VB.TextBox Text2 
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
      Left            =   6360
      MaxLength       =   1
      TabIndex        =   5
      Top             =   435
      Width           =   560
   End
   Begin VB.TextBox txtCustNo 
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
      Left            =   2385
      MaxLength       =   9
      TabIndex        =   3
      Top             =   780
      Width           =   1240
   End
   Begin VB.TextBox txtCustNo 
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
      Left            =   3950
      MaxLength       =   9
      TabIndex        =   4
      Top             =   780
      Width           =   1240
   End
   Begin VB.TextBox txtArea 
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
      Left            =   2385
      MaxLength       =   3
      TabIndex        =   0
      Top             =   90
      Width           =   675
   End
   Begin VB.TextBox txtArea 
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
      Left            =   3405
      MaxLength       =   3
      TabIndex        =   1
      Top             =   90
      Width           =   675
   End
   Begin VB.TextBox txtSales 
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
      Left            =   2385
      MaxLength       =   6
      TabIndex        =   2
      Top             =   435
      Width           =   890
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc11o6.frx":0000
      Height          =   3890
      Left            =   0
      TabIndex        =   6
      Top             =   1140
      Width           =   8900
      _ExtentX        =   15690
      _ExtentY        =   6853
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   16
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "a0902"
         Caption         =   "業務區"
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
         DataField       =   "st02"
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
      BeginProperty Column02 
         DataField       =   "cu144"
         Caption         =   "特殊發票"
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
         DataField       =   "custno"
         Caption         =   "客戶編號"
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
         DataField       =   "cu04"
         Caption         =   "客戶名稱"
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
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1035.213
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1425.26
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   7574.741
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   320
      Left            =   90
      Top             =   1080
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
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "特殊發票："
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
      Left            =   5250
      TabIndex        =   14
      Top             =   470
      Width           =   1130
   End
   Begin VB.Label Label4 
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
      Height          =   260
      Left            =   3710
      TabIndex        =   13
      Top             =   810
      Width           =   260
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "客戶編號："
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
      Left            =   1170
      TabIndex        =   12
      Top             =   810
      Width           =   1185
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
      Left            =   3150
      TabIndex        =   11
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "業 務 區："
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
      Left            =   1170
      TabIndex        =   10
      Top             =   120
      Width           =   1185
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "智權人員："
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
      Left            =   1170
      TabIndex        =   9
      Top             =   465
      Width           =   1185
   End
   Begin MSForms.Label lblSales 
      Height          =   260
      Left            =   3330
      TabIndex        =   8
      Top             =   465
      Width           =   1430
      VariousPropertyBits=   19
      Caption         =   "lblSales"
      Size            =   "2522;459"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "註：按ESC鍵，即可離開查詢，進入維護作業！"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   290
      Left            =   300
      TabIndex        =   7
      Top             =   5070
      Width           =   5270
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   4800
      Visible         =   0   'False
      Width           =   132
   End
End
Attribute VB_Name = "Frmacc11o6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/08 Form2.0已修改
'Create by Sindy 2013/12/13
Option Explicit

Public adoacc0i0 As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset


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
   'Modify by Amy    2023/10/06 原W:9045/H:5700
   Me.Width = 9270
   Me.Height = 6015
   'Modify by Amy    2023/10/06 原 (lngWidth - Me.Width) 切畫面不需再調整,改畫面左移-瑞婷
   Me.Move 0, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath2)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   
   lblSales = ""
   strCompanyNo = ""
   OpenTable
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Adodc1.Recordset.RecordCount <> 0 Then
      strCompanyNo = Adodc1.Recordset.Fields("custno").Value
   Else
      strCompanyNo = MsgText(601)
   End If
   StatusClear
   tool3_enabled
   'Forms(0).Toolbar1.Buttons.Item(7).Enabled = True
   Forms(0).Toolbar1.Buttons.Item(9).Enabled = True
   Frmacc11o5.Enabled = True
   Frmacc11o5.Show
   Set Frmacc11o6 = Nothing
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)
   Select Case KeyCode
      Case vbKeyF12
         Acc0i0Query
   End Select
   KeyEnter KeyCode
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

'*************************************************
'  不開發票客戶查詢
'
'*************************************************
Private Sub Acc0i0Query()
Dim strCon As String
   
On Error GoTo Checking
   
   Screen.MousePointer = vbHourglass
   
   strCon = ""
   If txtSales <> "" Then
      strCon = strCon & " and cu13='" & txtSales & "'"
   End If
   If txtArea(0) <> "" Then
      strCon = strCon & " and cu12>='" & txtArea(0) & "'"
   End If
   If txtArea(1) <> "" Then
      strCon = strCon & " and cu12<='" & txtArea(1) & "'"
   End If
   If txtCustNo(0) <> "" Then
      strCon = strCon & " and cu01||cu02>='" & txtCustNo(0) & "'"
   End If
   If txtCustNo(1) <> "" Then
      strCon = strCon & " and cu01||cu02<='" & txtCustNo(1) & "'"
   End If
   'Add By Sindy 2023/9/6
   If Text2 <> "" Then
      strCon = strCon & " and cu144='" & Text2 & "'"
   End If
   '2023/9/6 END
   
   'Modify By Sindy 2023/9/4 cu144='N' ==> cu144 is not null
   '                         +,cu144
   strSql = "select a0902,st02,cu144,cu01||cu02 as custno,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) as cu04 from customer,staff,acc090" & _
            " where cu144 is not null and cu02='0'" & _
            " and cu13=st01(+)" & _
            " and cu12=a0901(+)" & strCon & _
            " order by cu12,cu13,cu01"
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
   
On Error GoTo Checking
   
   adoadodc1.CursorLocation = adUseClient
   'Modify By Sindy 2023/9/4 cu144='N' ==> cu144 is not null
   '                         +,cu144
   strSql = "select a0902,st02,cu144,cu01||cu02 as custno,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) as cu04 from customer,staff,acc090" & _
            " where cu144 is not null and cu02='0'" & _
            " and cu13=st01(+)" & _
            " and cu12=a0901(+)" & _
            " order by cu12,cu13,cu01"
   adoadodc1.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

Private Sub Text2_GotFocus()
   InverseTextBox Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   Text2 = Trim(Text2)
   'Modify By Sindy 2023/9/4 + And Text2 <> "A" And Text2 <> "B"
   If Text2 <> "" And Text2 <> "N" And Text2 <> "A" And Text2 <> "B" Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "請輸入 N 或 A 或 B 或 空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      Text2_GotFocus
   End If
End Sub

Private Sub txtSales_Click()
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

Private Sub txtSales_GotFocus()
   InverseTextBox txtSales
   OpenIme
End Sub

Private Sub txtSales_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

Private Sub txtSales_Validate(Cancel As Boolean)
   lblSales = ""
   If txtSales <> "" Then
      lblSales = GetPrjSalesNM(txtSales)
   End If
End Sub

Private Sub txtCustNo_GotFocus(Index As Integer)
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

Private Sub txtCustNo_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

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

Private Sub txtArea_GotFocus(Index As Integer)
   CloseIme
   If Index = 1 Then
      If txtArea(0) <> "" And txtArea(1) = "" Then
         txtArea(1) = txtArea(0)
      Else
         TextInverse txtArea(Index)
      End If
   Else
      TextInverse txtArea(Index)
   End If
End Sub

Private Sub txtArea_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
