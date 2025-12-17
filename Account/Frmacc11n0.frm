VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc11n0 
   AutoRedraw      =   -1  'True
   Caption         =   "客戶電匯資料"
   ClientHeight    =   5340
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   8820
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5340
   ScaleWidth      =   8820
   Begin VB.CommandButton Command1 
      Caption         =   "修改"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7860
      TabIndex        =   6
      Top             =   930
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Height          =   300
      Left            =   2580
      Picture         =   "Frmacc11n0.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   210
      Width           =   350
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc11n0.frx":0102
      Height          =   3375
      Left            =   210
      TabIndex        =   7
      Top             =   1680
      Width           =   8325
      _ExtentX        =   14676
      _ExtentY        =   5944
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
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
      ColumnCount     =   4
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
            ColumnWidth     =   1450.205
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2649.827
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1349.858
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   7490.268
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text1 
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
      Left            =   1320
      MaxLength       =   9
      TabIndex        =   0
      Top             =   210
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
      Left            =   0
      Top             =   3000
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
   Begin MSForms.TextBox Text5 
      Height          =   330
      Left            =   2190
      TabIndex        =   5
      Top             =   1290
      Width           =   5620
      VariousPropertyBits=   671105049
      MaxLength       =   200
      Size            =   "9913;582"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text4 
      Height          =   330
      Left            =   1320
      TabIndex        =   4
      Top             =   930
      Width           =   6490
      VariousPropertyBits=   671105049
      MaxLength       =   200
      Size            =   "11448;582"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text3 
      Height          =   330
      Left            =   1320
      TabIndex        =   3
      Top             =   570
      Width           =   1935
      VariousPropertyBits=   671105049
      Size            =   "3413;582"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text2 
      Height          =   330
      Left            =   4080
      TabIndex        =   2
      Top             =   210
      Width           =   4332
      VariousPropertyBits=   671105049
      Size            =   "7223;529"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "新增其他電匯資料"
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
      Left            =   360
      TabIndex        =   13
      Top             =   1320
      Width           =   1875
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "註：不同匯款名稱以 , 分開即可"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   3810
      TabIndex        =   12
      Top             =   660
      Width           =   3825
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
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
      Height          =   210
      Left            =   360
      TabIndex        =   11
      Top             =   615
      Width           =   900
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '透明
      Caption         =   "電匯資料"
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
      Left            =   360
      TabIndex        =   10
      Top             =   990
      Width           =   1665
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   4800
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   1545
      Left            =   1320
      Top             =   120
      Width           =   8295
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "客戶名稱"
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
      Left            =   3120
      TabIndex        =   9
      Top             =   255
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "客戶代號"
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
      Left            =   360
      TabIndex        =   8
      Top             =   255
      Width           =   900
   End
End
Attribute VB_Name = "Frmacc11n0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/14 Form2.0已修改
'Memo By Sonia 2012/12/4 智權人員欄已修改
'Create by Sindy 2012/8/29
Option Explicit

Public adoadodc1 As New ADODB.Recordset
Dim m_ST06 As String


Private Sub Command1_Click()
   Text4.Enabled = True
   Text4.SetFocus
End Sub

Private Sub Command3_Click()
Dim Rs As New ADODB.Recordset
   
   If Text1 = MsgText(601) Then
      Exit Sub
   Else
      Text1 = Left(Text1 & "000000000", 9)
   End If
   If strSaveConfirm = MsgText(3) Then '新增狀態時
      '先檢查在多筆視窗中是否已有電匯資料,若有不可再新增
      Adodc1.Recordset.Find "custname = '" & Text1 & "'", 0, adSearchForward, 1
      If Adodc1.Recordset.EOF = False Then
         FormShow
         RecordShow
         Exit Sub
      End If
      Rs.CursorLocation = adUseClient
      If Left(Text1, 1) = "X" Then
         '讀取客戶資料
         Rs.Open "select cu01||cu02 as custname,NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) as cu04,st02,cu155 from customer,staff where cu13=st01(+) and cu01='" & Left(Text1, 8) & "' and cu02='" & Right(Text1, 1) & "'", adoTaie, adOpenStatic, adLockReadOnly
      ElseIf Left(Text1, 1) = "Y" Then
         Rs.Open "select fa01||fa02 as custname,NVL(fa04,NVL(fa05||fa63||fa64||fa65,fa06)) as cu04,' ',fa114 from fagent where fa01='" & Left(Text1, 8) & "' and fa02='" & Right(Text1, 1) & "'", adoTaie, adOpenStatic, adLockReadOnly
      End If
      If Rs.RecordCount <> 0 Then
         '存在帶出資料
         Text1 = Rs.Fields(0)
         Text2 = "" & Rs.Fields(1)
         Text3 = "" & Rs.Fields(2)
         Text4 = "" & Rs.Fields(3)
         Text5.SetFocus
      Else
         '不存在顯示警示訊息
         If Left(Text1, 1) = "X" Then
            MsgBox "無此客戶編號！", , MsgText(5)
         ElseIf Left(Text1, 1) = "Y" Then
            MsgBox "無此代理人編號！", , MsgText(5)
         End If
         strControlButton = MsgText(602)
         Text1.SetFocus
         Exit Sub
      End If
      Rs.Close
   Else
      Adodc1.Recordset.Find "custname = '" & Text1 & "'", 0, adSearchForward, 1
      If Adodc1.Recordset.EOF = False Then
         FormShow
         RecordShow
      Else
         MsgBox MsgText(33), , MsgText(5)
         If Adodc1.Recordset.RecordCount > 0 Then
            Adodc1.Recordset.MoveFirst
         End If
      End If
   End If
End Sub

Private Sub Command3_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Command3_Click
         Exit Sub
   End Select
   KeyEnter KeyCode
End Sub

Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub DataGrid1_SelChange(Cancel As Integer)
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   FormShow
   RecordShow
End Sub

Private Sub Form_Activate()
   strFormName = Name
   If strCompanyNo = MsgText(601) Then
      Exit Sub
   End If
   If Adodc1.Recordset.RecordCount <> 0 Then
      Adodc1.Recordset.MoveFirst
   End If
   Adodc1.Recordset.Find "custname = '" & strCompanyNo & "'", 0, adSearchForward, 1
   If Adodc1.Recordset.EOF = False Then
      FormShow
      RecordShow
   End If
   strCompanyNo = MsgText(601)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
Dim intCounter As Integer
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   'Modify by Amy 2023/10/05 原W:8850/H:5500
   Me.Width = 8940
   Me.Height = 5805
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath1)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   'strCompanyNo = MsgText(601)
   m_ST06 = PUB_GetST06(strUserNum)
   
   OpenTable
   If Adodc1.Recordset.RecordCount <> 0 Then
      RecordShow
   End If
   
   Call FormDisabled
   Text1.Text = "X"
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
      Cancel = 1
      Exit Sub
   End If
   StatusClear
   strConTitle = MsgText(601)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc11n0 = Nothing
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
   strSql = "select cu01||cu02 as custname,NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) as cu04,st02,cu155 from customer,staff where cu155>' ' and cu155 is not null and cu13=st01(+)" & _
            " Union" & _
            " select fa01||fa02 as custname,NVL(fa04,NVL(fa05||fa63||fa64||fa65,fa06)) as cu04,' ',fa114 as cu155 from fagent where fa114>' ' and fa114 is not null" & strCon2 & _
            " order by custname asc"
   adoadodc1.Open strSql, adoTaie, adOpenDynamic, adLockBatchOptimistic
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  顯示資料表(廠商資料)
'
'*************************************************
Public Sub FormShow()
   Text1 = Adodc1.Recordset.Fields("custname").Value
   strControlButton = MsgText(601)
   If IsNull(Adodc1.Recordset.Fields("cu04").Value) Then
      Text2 = MsgText(601)
   Else
      Text2 = Adodc1.Recordset.Fields("cu04").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("st02").Value) Then
      Text3 = MsgText(601)
   Else
      Text3 = Adodc1.Recordset.Fields("st02").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("cu155").Value) Then
      Text4 = MsgText(601)
   Else
      Text4 = Adodc1.Recordset.Fields("cu155").Value
   End If
   Text5 = ""
End Sub

'*************************************************
'  關閉分錄欄位輸入狀態
'
'*************************************************
Public Sub FormDisabled()
   'Text1.Enabled = False
   Text4.Enabled = False
   Text5.Enabled = False
   Command1.Enabled = False
End Sub

'*************************************************
'  開啟分錄欄位輸入狀態
'
'*************************************************
Public Sub FormEnabled()
   'If strSaveConfirm = MsgText(3) Then '新增狀態時
      'Text1.Enabled = True
   'End If
   Text4.Enabled = False
   Text5.Enabled = True
   Command1.Enabled = True
End Sub

'*************************************************
'  重新整理 Adodc 之資料
'
'*************************************************
Public Sub AdodcRefresh()
Dim strCon2 As String
   
On Error GoTo Checking
   
   adoadodc1.Close
   adoadodc1.CursorLocation = adUseClient
   strCon2 = ""
   If m_ST06 <> "1" Then '若為分所人員不顯示代理人資料
      strCon2 = strCon2 & " and fa01='null' and fa02='null'"
   End If
   strSql = "select cu01||cu02 as custname,NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) as cu04,st02,cu155 from customer,staff where cu155>' ' and cu155 is not null and cu13=st01(+)" & _
            " Union" & _
            " select fa01||fa02 as custname,NVL(fa04,NVL(fa05||fa63||fa64||fa65,fa06)) as cu04,' ',fa114 as cu155 from fagent where fa114>' ' and fa114 is not null" & strCon2 & _
            " order by custname asc"
   adoadodc1.Open strSql, adoTaie, adOpenDynamic, adLockBatchOptimistic
   Adodc1.Recordset.Requery
   If Adodc1.Recordset.RecordCount <> 0 Then
      If Text1 <> MsgText(601) Then
         Adodc1.Recordset.Find "custname = '" & Text1 & "'", 0, adSearchForward, 1
         If Adodc1.Recordset.EOF = False Then
            FormShow
            RecordShow
         End If
      End If
   End If
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  顯示筆數
'
'*************************************************
Public Sub RecordShow()
   Frmacc0000.StatusBar1.Panels(2).Text = Adodc1.Recordset.Bookmark & MsgText(35) & Adodc1.Recordset.RecordCount
End Sub

Private Sub Text1_GotFocus()
Dim intPos As Integer
   'TextInverse Text1
   'Modify By Sindy 2012/9/17
   '將游標停在最後一個字的後面
   With Me.Text1
      If Len("" & .Text) > 0 Then
         .SelStart = Len(.Text)
         .SelLength = 0
      End If
   End With
   '2012/9/17 End
   CloseIme
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Public Sub Text1_Validate(Cancel As Boolean)
   Cancel = False
   If Text1.Text <> "" And Text1.Text <> "X" Then
      Text1 = Left(Text1 & "000000000", 9)
      If m_ST06 = "1" And Left(Text1.Text, 1) <> "X" And Left(Text1.Text, 1) <> "Y" Then
         MsgBox "客戶編號只可輸入X或Y !!"
         Text1.SetFocus
         Cancel = True
      ElseIf m_ST06 <> "1" And Left(Text1.Text, 1) <> "X" Then
         MsgBox "客戶編號只可輸入X !!"
         Text1.SetFocus
         Cancel = True
      End If
   End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_KeyUp(KeyCode As MSForms.ReturnInteger, Shift As Integer)
   KeyEnter Val(KeyCode)
End Sub

Private Sub Text3_KeyUp(KeyCode As MSForms.ReturnInteger, Shift As Integer)
   KeyEnter Val(KeyCode)
End Sub

Private Sub Text4_GotFocus()
Dim intPos As Integer
   'TextInverse Text4
   'Modify By Sindy 2012/9/17
   '將游標停在最後一個字的後面
   With Me.Text4
      If Len("" & .Text) > 0 Then
         .SelStart = Len(.Text)
         .SelLength = 0
      End If
   End With
   '2012/9/17 End
   OpenIme
End Sub

Private Sub Text4_KeyUp(KeyCode As MSForms.ReturnInteger, Shift As Integer)
   'Add by Amy 2017/12/13 改輸入分號按Enter 會重讀資料(因設Command3的Default設True)
   Dim bolCancel As Boolean
   Text4_Validate (bolCancel)
   If bolCancel = True Then Exit Sub
   'end 2017/12/13
   KeyEnter CInt(KeyCode) 'Modify By Sindy 2025/5/8 +CInt
End Sub

'Add by Amy 2017/12/13
Public Sub Text4_Validate(Cancel As Boolean)
    If strSaveConfirm = MsgText(601) Then Exit Sub
    
    If InStr(Me.Text4, ";") > 0 Then
        MsgBox "不可輸入分號!!"
        Cancel = True
    End If
End Sub

Public Sub Text5_Validate(Cancel As Boolean)
    If strSaveConfirm = MsgText(601) Then Exit Sub
    
    If InStr(Me.Text5, ";") > 0 Then
        MsgBox "不可輸入分號!!"
        Cancel = True
    End If
End Sub
'end 2017/12/13

Private Sub Text5_GotFocus()
   TextInverse Text5
   OpenIme
End Sub

Private Sub Text5_KeyUp(KeyCode As MSForms.ReturnInteger, Shift As Integer)
   'Add by Amy 2017/12/13 改輸入分號按Enter 會重讀資料(因設Command3的Default設True)
   Dim bolCancel As Boolean
   Text5_Validate (bolCancel)
   If bolCancel = True Then Exit Sub
   'end 2017/12/13
   KeyEnter CInt(KeyCode) 'Modify By Sindy 2025/5/8 +CInt
End Sub

'Add By Sindy 2012/8/29
Public Sub Frmacc11n0_Clear()
   With Frmacc11n0
      .Text1 = "X"
      .Text2 = ""
      .Text3 = ""
      .Text4 = ""
      .Text5 = ""
      .Text1.SetFocus
   End With
End Sub
'2012/8/29 End

'Add By Sindy 2012/8/29
Public Sub Frmacc11n0_Delete()
On Error GoTo Checking
   With Frmacc11n0
      If .Text1 = MsgText(601) Or Trim(.Text1) = "X" Or _
         (Left(Trim(.Text1), 1) <> "X" And Left(Trim(.Text1), 1) <> "Y") Then
         MsgBox "尚未查出欲刪除的資料 !", , MsgText(5)
         strControlButton = MsgText(602)
         Exit Sub
      End If
      
      If Left(Trim(.Text1), 1) = "X" Then
         If DeleteCheck("select cu01 from customer where cu01='" & Left(.Text1, 8) & "' and cu02='" & Mid(.Text1, 9, 1) & "'") = MsgText(603) Then
            Exit Sub
         End If
         adoTaie.Execute "update customer set cu155=null where cu01='" & Left(.Text1, 8) & "' and cu02='" & Mid(.Text1, 9, 1) & "'"
      ElseIf Left(Trim(.Text1), 1) = "Y" Then
         If DeleteCheck("select fa01 from fagent where fa01='" & Left(.Text1, 8) & "' and fa02='" & Mid(.Text1, 9, 1) & "'") = MsgText(603) Then
            Exit Sub
         End If
         adoTaie.Execute "update fagent set fa114=null where fa01='" & Left(.Text1, 8) & "' and fa02='" & Mid(.Text1, 9, 1) & "'"
      End If
      .AdodcRefresh
      If .Adodc1.Recordset.RecordCount <> 0 Then
         .Adodc1.Recordset.MoveFirst
         .RecordShow
      Else
         StatusClear
      End If
   End With
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'Add By Sindy 2012/8/29
Public Sub Frmacc11n0_First()
   With Frmacc11n0
      If .Adodc1.Recordset.RecordCount <> 0 Then
         .Adodc1.Recordset.MoveFirst
         .FormShow
         .RecordShow
      End If
   End With
End Sub

'Add By Sindy 2012/8/29
Public Sub Frmacc11n0_Last()
   With Frmacc11n0
      If .Adodc1.Recordset.RecordCount <> 0 Then
         .Adodc1.Recordset.MoveLast
         .FormShow
         .RecordShow
      End If
   End With
End Sub

'Add By Sindy 2012/8/29
Public Sub Frmacc11n0_Next()
   With Frmacc11n0
      If .Adodc1.Recordset.EOF = False Then
         .Adodc1.Recordset.MoveNext
         If .Adodc1.Recordset.EOF Then
            .Adodc1.Recordset.MoveLast
            MsgBox MsgText(8), , MsgText(5)
         End If
         .FormShow
         .RecordShow
      End If
   End With
End Sub

'Add By Sindy 2012/8/29
Public Sub Frmacc11n0_Previous()
   With Frmacc11n0
      If .Adodc1.Recordset.BOF = False Then
         .Adodc1.Recordset.MovePrevious
         If .Adodc1.Recordset.BOF Then
            .Adodc1.Recordset.MoveFirst
            MsgBox MsgText(7), , MsgText(5)
         End If
         .FormShow
         .RecordShow
      End If
   End With
End Sub

'Add By Sindy 2012/8/29
Public Sub Frmacc11n0_Save()
Dim strText As String
   
   On Error GoTo Checking
   
   'Modify by Amy 2017/12/13 原存檔前檢查改至FormCheck
   With Frmacc11n0
      '更新DB資料
      If Trim(.Text4) <> "" Then
         strText = Trim(.Text4) & IIf(Trim(.Text5) <> "", "," & Trim(.Text5), "")
      Else
         strText = Trim(.Text5)
      End If
      adoTaie.BeginTrans
      If Left(.Text1, 1) = "X" Then
         strSql = "update customer " & _
                  "set cu155='" & strText & "' " & _
                  "where cu01='" & Left(.Text1, 8) & "' and cu02='" & Right(.Text1, 1) & "' "
      ElseIf Left(.Text1, 1) = "Y" Then
         strSql = "update fagent " & _
                  "set fa114='" & strText & "' " & _
                  "where fa01='" & Left(.Text1, 8) & "' and fa02='" & Right(.Text1, 1) & "' "
      End If
      adoTaie.Execute strSql
      adoTaie.CommitTrans
      
      .AdodcRefresh
      .FormDisabled
Checking:
   If Err.Number = 0 Then
      Exit Sub
   Else
      adoTaie.RollbackTrans
   End If
   MsgBox Err.Description, , MsgText(5)
   End With
End Sub

'Add by Amy 2017/12/13 將原於 frmacc11n0_Save裡的檢查拆出來
Public Function FormCheck() As Boolean
    Dim bolCancel As Boolean
    
    FormCheck = False
   
    If Text1 = MsgText(601) Then
        MsgBox MsgText(10), , MsgText(5)
        strControlButton = MsgText(602)
        Text1.SetFocus
        Exit Function
    End If
    If Text2 = MsgText(601) Then
        MsgBox "無此客戶編號！", , MsgText(5)
        Text1.SetFocus
        Exit Function
    End If
    Call Text1_Validate(bolCancel)
    If bolCancel = True Then
         Text1.SetFocus
         Exit Function
    End If
    If Text4.Enabled = False And Text5 = MsgText(601) Then
         MsgBox MsgText(10), , MsgText(5)
         Text5.SetFocus
         Exit Function
    End If
    Call Text4_Validate(bolCancel)
    If bolCancel = True Then
        Text4.SetFocus
        Exit Function
    End If
    Call Text5_Validate(bolCancel)
    If bolCancel = True Then
        Text5.SetFocus
        Exit Function
    End If
    
    FormCheck = True
End Function
