VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc21v0 
   AutoRedraw      =   -1  'True
   Caption         =   "客戶會計師資料維護"
   ClientHeight    =   3648
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8928
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3648
   ScaleWidth      =   8928
   Begin VB.CommandButton cmdSave 
      Caption         =   "存檔"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6420
      TabIndex        =   10
      Top             =   120
      Width           =   765
   End
   Begin VB.CommandButton Command3 
      Height          =   300
      Left            =   8520
      Picture         =   "Frmacc21v0.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   630
      Visible         =   0   'False
      Width           =   350
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
      Left            =   1800
      Top             =   330
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
   Begin MSForms.TextBox textA4901_C 
      Height          =   345
      Left            =   1485
      TabIndex        =   2
      Top             =   600
      Width           =   1365
      VariousPropertyBits=   671105049
      MaxLength       =   20
      Size            =   "2408;609"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textA4914 
      Height          =   900
      Left            =   1485
      TabIndex        =   9
      Top             =   2550
      Width           =   7005
      VariousPropertyBits=   -1466941413
      MaxLength       =   200
      ScrollBars      =   2
      Size            =   "12356;1587"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textA4913 
      Height          =   345
      Left            =   1485
      TabIndex        =   8
      Top             =   2160
      Width           =   7005
      VariousPropertyBits=   671105051
      MaxLength       =   80
      Size            =   "12356;609"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textA4912 
      Height          =   345
      Left            =   4560
      TabIndex        =   4
      Top             =   990
      Width           =   3915
      VariousPropertyBits=   671105051
      MaxLength       =   40
      Size            =   "6906;609"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textA4905 
      Height          =   345
      Left            =   1485
      TabIndex        =   7
      Top             =   1770
      Width           =   7005
      VariousPropertyBits=   671105051
      MaxLength       =   100
      Size            =   "12356;609"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textA4901 
      Height          =   345
      Left            =   1485
      TabIndex        =   0
      Top             =   600
      Width           =   7005
      VariousPropertyBits=   671105049
      MaxLength       =   100
      Size            =   "12356;609"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textA4903 
      Height          =   345
      Left            =   1485
      TabIndex        =   5
      Top             =   1380
      Width           =   2055
      VariousPropertyBits=   671105051
      MaxLength       =   20
      Size            =   "3625;609"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textA4904 
      Height          =   345
      Left            =   4560
      TabIndex        =   6
      Top             =   1380
      Width           =   2055
      VariousPropertyBits=   671105051
      MaxLength       =   20
      Size            =   "3625;609"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textA4902 
      Height          =   345
      Left            =   1485
      TabIndex        =   3
      Top             =   990
      Width           =   2055
      VariousPropertyBits=   671105051
      MaxLength       =   30
      Size            =   "3625;609"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label63 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "備註："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   765
      TabIndex        =   19
      Top             =   2610
      Width           =   720
   End
   Begin VB.Label Label63 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "地址："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   765
      TabIndex        =   18
      Top             =   2220
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "事務所："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   3600
      TabIndex        =   17
      Top             =   1050
      Width           =   960
   End
   Begin VB.Label lblA4901_C 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2880
      TabIndex        =   16
      Top             =   660
      Width           =   75
   End
   Begin VB.Label Label63 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "E-mail："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   3
      Left            =   570
      TabIndex        =   15
      Top             =   1830
      Width           =   915
   End
   Begin VB.Label lblA4901 
      BackStyle       =   0  '透明
      Caption         =   "收據抬頭："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   270
      TabIndex        =   14
      Top             =   660
      Width           =   1215
   End
   Begin VB.Label Label63 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "電話："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   9
      Left            =   765
      TabIndex        =   13
      Top             =   1440
      Width           =   720
   End
   Begin VB.Label Label63 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "傳真："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   10
      Left            =   3840
      TabIndex        =   12
      Top             =   1440
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "會計師姓名："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   10
      Left            =   45
      TabIndex        =   11
      Top             =   1050
      Width           =   1440
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   180
      Top             =   2370
      Visible         =   0   'False
      Width           =   135
   End
End
Attribute VB_Name = "Frmacc21v0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/02 改成Form2.0 ;所有TextBox
'Create by Sindy 2016/10/21
Option Explicit

Dim adoadodc1 As New ADODB.Recordset
Dim strSaveConfirm As String
Dim bolCallMe As Boolean


Private Sub CmdSave_Click()
Dim Cancel As Boolean
Dim strA4901 As String
Dim strUpdType As String
   
On Error GoTo Checking
   
   With Frmacc21v0
      If textA4901.Visible = True Then
         If .textA4901 = MsgText(601) Then
            MsgBox MsgText(10), , MsgText(5)
            strControlButton = MsgText(602)
            .textA4901.SetFocus
            Exit Sub
         End If
      Else
         If .textA4901_C = MsgText(601) Then
            MsgBox MsgText(10), , MsgText(5)
            strControlButton = MsgText(602)
            .textA4901_C.SetFocus
            Exit Sub
         End If
      End If
'      If .textA4902 = MsgText(601) And _
'         .textA4912 = MsgText(601) Then
'         MsgBox "姓名 或 事務所至少輸入一項！", , MsgText(5)
'         strControlButton = MsgText(602)
'         .textA4902.SetFocus
'         Exit Sub
'      End If
'      If .textA4903 = MsgText(601) And _
'         .textA4904 = MsgText(601) And _
'         .textA4905 = MsgText(601) Then
'         MsgBox "電話、傳真、E-Mail至少要輸入一項！", , MsgText(5)
'         strControlButton = MsgText(602)
'         .textA4903.SetFocus
'         Exit Sub
'      End If
      
      If textA4901.Visible = True Then
         Call textA4901_Validate(Cancel)
         If Cancel = True Then
            strControlButton = MsgText(602)
            .textA4901.SetFocus
            Exit Sub
         End If
      Else
         Call textA4901_C_Validate(Cancel)
         If Cancel = True Then
            strControlButton = MsgText(602)
            .textA4901_C.SetFocus
            Exit Sub
         End If
      End If
      Call textA4902_Validate(Cancel)
      If Cancel = True Then
         strControlButton = MsgText(602)
         .textA4902.SetFocus
         Exit Sub
      End If
      Call textA4903_Validate(Cancel)
      If Cancel = True Then
         strControlButton = MsgText(602)
         .textA4903.SetFocus
         Exit Sub
      End If
      Call textA4904_Validate(Cancel)
      If Cancel = True Then
         strControlButton = MsgText(602)
         .textA4904.SetFocus
         Exit Sub
      End If
      Call textA4905_Validate(Cancel)
      If Cancel = True Then
         strControlButton = MsgText(602)
         .textA4905.SetFocus
         Exit Sub
      End If
      Call textA4912_Validate(Cancel)
      If Cancel = True Then
         strControlButton = MsgText(602)
         .textA4912.SetFocus
         Exit Sub
      End If
      Call textA4913_Validate(Cancel)
      If Cancel = True Then
         strControlButton = MsgText(602)
         .textA4913.SetFocus
         Exit Sub
      End If
      Call textA4914_Validate(Cancel)
      If Cancel = True Then
         strControlButton = MsgText(602)
         .textA4914.SetFocus
         Exit Sub
      End If
      'Added by Lydia 2021/12/02 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
      If PUB_ChkUniText(Me, , True, "TextBox") = False Then
           Exit Sub
      End If
    
      '更新DB資料
      strUpdType = ""
      '刪除
      If Trim(.textA4902) = "" And Trim(.textA4903) = "" And _
         Trim(.textA4904) = "" And Trim(.textA4905) = "" And _
         Trim(.textA4912) = "" And Trim(.textA4913) = "" And _
         Trim(.textA4914) = "" Then
         strUpdType = "D"
      '新增
      ElseIf strSaveConfirm = MsgText(3) Then
         strUpdType = "A"
      '修改
      Else
         strUpdType = "M"
      End If
      
      If textA4901.Visible = True Then
         strA4901 = textA4901
         If strUpdType = "D" Then
            strExc(0) = "select A4221 from ACC420 where A4201='" & ChgSQL(strA4901) & "' and A4221='2'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               MsgBox "收據抬頭:" & strA4901 & "「繳款書寄件處」為會計師，不可刪除此筆資料！", , MsgText(5)
               strControlButton = MsgText(602)
               Exit Sub
            End If
         End If
      Else
         strA4901 = textA4901_C
         If strUpdType = "D" Then
            strExc(0) = "select cu169 from customer where cu01='" & Left(strA4901, 8) & "' and cu02='" & Mid(strA4901, 9, 1) & "' and cu169='2'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               MsgBox "客戶編號:" & strA4901 & "「繳款書寄件處」為會計師，不可刪除此筆資料！", , MsgText(5)
               strControlButton = MsgText(602)
               Exit Sub
            End If
         End If
      End If
      
      adoTaie.BeginTrans
      'Modify By Sindy 2016/11/7 + a4912,a4913,a4914
      '新增
      If strUpdType = "A" Then
         strSql = "insert into acc490(a4901,a4902,a4903,a4904,a4905,a4912,a4913,a4914) " & _
                  "values(" & CNULL(strA4901) & "," & CNULL(.textA4902) & "," & CNULL(.textA4903) & _
                  "," & CNULL(.textA4904) & "," & CNULL(.textA4905) & "," & CNULL(.textA4912) & _
                  "," & CNULL(.textA4913) & "," & CNULL(.textA4914) & ")"
      '刪除
      ElseIf strUpdType = "D" Then
         strSql = "delete from acc490 where a4901='" & strA4901 & "'"
      '修改
      Else
         strSql = "update acc490 " & _
                  "set a4902=" & CNULL(.textA4902) & _
                     ",a4903=" & CNULL(.textA4903) & _
                     ",a4904=" & CNULL(.textA4904) & _
                     ",a4905=" & CNULL(.textA4905) & _
                     ",a4912=" & CNULL(.textA4912) & _
                     ",a4913=" & CNULL(.textA4913) & _
                     ",a4914=" & CNULL(.textA4914) & _
                  " where a4901='" & strA4901 & "' "
      End If
      adoTaie.Execute strSql
      
      'Modify By Sindy 2018/1/5
      '修改資料時, 若有變更名稱的編號資料也要一併更新,
      '包含會計師資料
      If Left(strA4901, 1) = "X" Then 'X客戶編號
         strSql = "update acc490 " & _
                     "set a4902=" & CNULL(.textA4902) & _
                        ",a4903=" & CNULL(.textA4903) & _
                        ",a4904=" & CNULL(.textA4904) & _
                        ",a4905=" & CNULL(.textA4905) & _
                        ",a4912=" & CNULL(.textA4912) & _
                        ",a4913=" & CNULL(.textA4913) & _
                        ",a4914=" & CNULL(.textA4914) & _
                     " where substr(a4901,1,8)='" & Left(strA4901, 8) & "' and a4901<>'" & strA4901 & "'"
         adoTaie.Execute strSql
      End If
      '2018/1/5 END
      
      adoTaie.CommitTrans
      
      textA4902.Tag = textA4902.Text
      textA4903.Tag = textA4903.Text
      textA4904.Tag = textA4904.Text
      textA4905.Tag = textA4905.Text
      textA4912.Tag = textA4912.Text
      textA4913.Tag = textA4913.Text
      textA4914.Tag = textA4914.Text
   End With
   Unload Me
   
Checking:
   If Err.Number = 0 Then
      Exit Sub
   Else '-2147168237
      adoTaie.RollbackTrans
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

Public Sub Command3_Click()
Dim Rs As New ADODB.Recordset
   
   If textA4901.Visible = True Then
      If textA4901 = MsgText(601) Then
         Exit Sub
      Else
         strCompanyNo = textA4901
      End If
   ElseIf textA4901_C.Visible = True Then
      If textA4901_C = MsgText(601) Then
         Exit Sub
      Else
         strCompanyNo = textA4901_C
      End If
   End If
   If strSaveConfirm = MsgText(3) Then '新增狀態時
      '先檢查在多筆視窗中是否已有資料,若有不可再新增
      Adodc1.Recordset.Find "a4901 = '" & strCompanyNo & "'", 0, adSearchForward, 1
      If Adodc1.Recordset.EOF = False Then
         FormShow
         RecordShow
         Exit Sub
      End If
   Else
      Adodc1.Recordset.Find "a4901 = '" & strCompanyNo & "'", 0, adSearchForward, 1
      If Adodc1.Recordset.EOF = False Then
         FormShow
         RecordShow
         If bolCallMe = True Then
            Frmacc0000.Toolbar1.Buttons.Item(5).Enabled = True '修改
         End If
      Else
         MsgBox MsgText(33), , MsgText(5)
         If bolCallMe = True Then
            strSaveConfirm = "A"
            FormEnabled
            Frmacc21v0_Clear
            If textA4901.Visible = True Then
               textA4901 = strCompanyNo
            Else
               textA4901_C = strCompanyNo
            End If
            tool2_enabled
         End If
         If Adodc1.Recordset.RecordCount > 0 Then
            Adodc1.Recordset.MoveFirst
         End If
      End If
   End If
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
   
   'Modified by Lydia 2021/12/07 改成模組
'   Me.Icon = LoadPicture(strIcoPath)
''   Me.Height = 5700
''   Me.Width = 9045
'   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2 + 900
'   Image1 = LoadPicture(strBackPicPath1)
'   sglWidth = Image1.Width
'   sglHeight = Image1.Height
'   For intX = 0 To Int(ScaleWidth / sglWidth)
'       For intY = 0 To Int(ScaleHeight / sglHeight)
'           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
'       Next
'   Next
   PUB_InitForm Me, Me.Width, Me.Height, strBackPicPath1
   'end 2021/12/07
   
   'OpenTable
   'Call FormDisabled
   'FormEnabled
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If textA4902.Tag <> textA4902.Text Or _
      textA4903.Tag <> textA4903.Text Or _
      textA4904.Tag <> textA4904.Text Or _
      textA4905.Tag <> textA4905.Text Or _
      textA4912.Tag <> textA4912.Text Or _
      textA4913.Tag <> textA4913.Text Or _
      textA4914.Tag <> textA4914.Text Then
      If MsgBox("資料有異動確定要結束嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
         Cancel = True
         Exit Sub
      End If
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim strCU11 As String
   
'   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
'      Cancel = 1
'      Exit Sub
'   End If
'   If bolCallMe = False Then
'      If Adodc1.Recordset.RecordCount <> 0 Then
'         strCompanyNo = Adodc1.Recordset.Fields("a4901").Value
'      Else
'         strCompanyNo = MsgText(601)
'      End If
'   End If
'
'   StatusClear
'
'   If UCase(strUserLevel) = UCase("Frmacc44t0") Then
'      Frmacc44t0.Show
'      Frmacc44t0.cmdQuery_Click
'      tool3_enabled
'   End If
'
'   strUserLevel = MsgText(601)
'   strConTitle = MsgText(601)
'   strFormName = MsgText(601)
'   KeyEnter vbKeyEscape
'   MenuEnabled
   
   'Add By Sindy 2024/9/4
   If Frmacc21v0.Tag <> "" Then
      strFormName = Frmacc21v0.Tag
   End If
   '2024/9/4 END
   
   Set Frmacc21v0 = Nothing
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Public Sub OpenTable()
   
On Error GoTo Checking
   
   adoadodc1.CursorLocation = adUseClient
   strSql = "select * from acc490 where a4901='" & IIf(textA4901_C.Text <> "", textA4901_C.Text, ChgSQL(textA4901.Text)) & "' order by a4907,a4908 asc"
   adoadodc1.Open strSql, adoTaie, adOpenDynamic, adLockBatchOptimistic
   Set Adodc1.Recordset = adoadodc1
   If Adodc1.Recordset.RecordCount = 0 Then
      strSaveConfirm = "A"
   Else
      Call FormShow
   End If
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  顯示資料表
'
'*************************************************
Public Sub FormShow()
   strControlButton = MsgText(601)
   
   If IsNull(Adodc1.Recordset.Fields("A4901").Value) Then
      If textA4901.Visible = True Then
         textA4901 = MsgText(601)
      Else
         textA4901_C = MsgText(601)
      End If
   Else
      If textA4901.Visible = True Then
         textA4901 = Adodc1.Recordset.Fields("A4901").Value
      Else
         textA4901_C = Adodc1.Recordset.Fields("A4901").Value
      End If
      If UCase(strUserLevel) <> UCase("Frmacc44t0") Then
         'cmdCopy.Visible = True '查詢出資料時,才顯示此按鈕
      End If
   End If
   textA4901.Tag = textA4901.Text
   textA4901_C.Tag = textA4901_C.Text
   '會計師姓名
   If IsNull(Adodc1.Recordset.Fields("A4902").Value) Then
      textA4902 = MsgText(601)
   Else
      textA4902 = Adodc1.Recordset.Fields("A4902").Value
   End If
   textA4902.Tag = textA4902.Text
   '電話
   If IsNull(Adodc1.Recordset.Fields("A4903").Value) Then
      textA4903 = MsgText(601)
   Else
      textA4903 = Adodc1.Recordset.Fields("A4903").Value
   End If
   textA4903.Tag = textA4903.Text
   '傳真
   If IsNull(Adodc1.Recordset.Fields("A4904").Value) Then
      textA4904 = MsgText(601)
   Else
      textA4904 = Adodc1.Recordset.Fields("A4904").Value
   End If
   textA4904.Tag = textA4904.Text
   'E-Mail
   If IsNull(Adodc1.Recordset.Fields("A4905").Value) Then
      textA4905 = MsgText(601)
   Else
      textA4905 = Adodc1.Recordset.Fields("A4905").Value
   End If
   textA4905.Tag = textA4905.Text
   '事務所
   If IsNull(Adodc1.Recordset.Fields("A4912").Value) Then
      textA4912 = MsgText(601)
   Else
      textA4912 = Adodc1.Recordset.Fields("A4912").Value
   End If
   textA4912.Tag = textA4912.Text
   '地址
   If IsNull(Adodc1.Recordset.Fields("A4913").Value) Then
      textA4913 = MsgText(601)
   Else
      textA4913 = Adodc1.Recordset.Fields("A4913").Value
   End If
   textA4913.Tag = textA4913.Text
   '備註
   If IsNull(Adodc1.Recordset.Fields("A4914").Value) Then
      textA4914 = MsgText(601)
   Else
      textA4914 = Adodc1.Recordset.Fields("A4914").Value
   End If
   textA4914.Tag = textA4914.Text
End Sub

'*************************************************
'  關閉分錄欄位輸入狀態
'
'*************************************************
Public Sub FormDisabled()
   textA4901.Enabled = True
   textA4901_C.Enabled = True
   textA4902.Enabled = False
   textA4903.Enabled = False
   textA4904.Enabled = False
   textA4905.Enabled = False
   textA4912.Enabled = False
   textA4913.Enabled = False
   textA4914.Enabled = False
End Sub

'*************************************************
'  開啟分錄欄位輸入狀態
'
'*************************************************
Public Sub FormEnabled()
'   If strSaveConfirm = MsgText(3) Then '新增狀態時
'      If textA4901.Visible = True Then
'         textA4901.Enabled = True
'      Else
'         textA4901_C.Enabled = True
'      End If
'   Else
'      If textA4901.Visible = True Then
'         If textA4901 <> "" Then
'            strCompanyNo = textA4901
'            textA4901.Enabled = False
'            'Me.cmdCopy.Visible = False
'         Else
'            textA4901.Enabled = True
'         End If
'      Else
'         If textA4901_C <> "" Then
'            strCompanyNo = textA4901_C
'            textA4901_C.Enabled = False
'            'Me.cmdCopy.Visible = False
'         Else
'            textA4901_C.Enabled = True
'         End If
'      End If
'   End If
'   If bolCallMe = True Then
'      If textA4901.Visible = True Then
'         textA4901 = strCompanyNo
'         textA4901.Enabled = False
'      Else
'         textA4901_C = strCompanyNo
'         textA4901_C.Enabled = False
'      End If
'   End If
   
   textA4902.Enabled = True
   textA4903.Enabled = True
   textA4904.Enabled = True
   textA4905.Enabled = True
   textA4912.Enabled = True
   textA4913.Enabled = True
   textA4914.Enabled = True
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
   strSql = "select * from acc490 order by a4907,a4908 asc"
   adoadodc1.Open strSql, adoTaie, adOpenDynamic, adLockBatchOptimistic
   Adodc1.Recordset.Requery
   If Adodc1.Recordset.RecordCount <> 0 Then
      If textA4901.Visible = True Then
         If textA4901 <> MsgText(601) Then
            Adodc1.Recordset.Find "a4901 = '" & textA4901 & "'", 0, adSearchForward, 1
            If Adodc1.Recordset.EOF = False Then
               FormShow
               RecordShow
            End If
         End If
      Else
         If textA4901_C <> MsgText(601) Then
            Adodc1.Recordset.Find "a4901 = '" & textA4901_C & "'", 0, adSearchForward, 1
            If Adodc1.Recordset.EOF = False Then
               FormShow
               RecordShow
            End If
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

Public Sub Frmacc21v0_Clear()
   With Frmacc21v0
      '.cmdCopy.Visible = False
      .textA4901 = ""
      .textA4901_C = ""
      .textA4902 = ""
      .textA4903 = ""
      .textA4904 = ""
      .textA4905 = ""
      .textA4912 = ""
      .textA4913 = ""
      .textA4914 = ""
      If textA4901.Visible = True Then
         If textA4901.Enabled = True Then
            .textA4901.SetFocus
         End If
      Else
         If textA4901_C.Enabled = True Then
            .textA4901_C.SetFocus
         End If
      End If
      If bolCallMe = True Then
         If textA4901.Visible = True Then
            textA4901 = strCompanyNo
            textA4901.Enabled = False
         Else
            textA4901_C = strCompanyNo
            textA4901_C.Enabled = False
         End If
      End If
   End With
End Sub

Public Function Frmacc21v0_Delete() As Boolean
On Error GoTo Checking
   With Frmacc21v0
'      If .textA4203.Enabled = False And Trim(.textA4203) <> "" And Trim(.textA4201) <> "" Then
'      Else
'         MsgBox "尚未查出欲刪除的資料 !", , MsgText(5)
'         strControlButton = MsgText(602)
'         Exit Function
'      End If
      If textA4901.Visible = True Then
         adoTaie.Execute "delete from acc490 where a4901='" & ChgSQL(.textA4901) & "'"
      Else
         adoTaie.Execute "delete from acc490 where a4901='" & .textA4901_C & "'"
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
      Exit Function
   End If
   MsgBox Err.Description, , MsgText(5)
End Function

Public Sub Frmacc21v0_First()
   With Frmacc21v0
      If .Adodc1.Recordset.RecordCount <> 0 Then
         .Adodc1.Recordset.MoveFirst
         .FormShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc21v0_Last()
   With Frmacc21v0
      If .Adodc1.Recordset.RecordCount <> 0 Then
         .Adodc1.Recordset.MoveLast
         .FormShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc21v0_Next()
   With Frmacc21v0
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

Public Sub Frmacc21v0_Previous()
   With Frmacc21v0
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

'Public Sub Frmacc21v0_Save()
'Dim Cancel As Boolean
'
'   On Error GoTo Checking
'
'   With Frmacc21v0
'      If textA4901.Visible = True Then
'         If .textA4901 = MsgText(601) Then
'            MsgBox MsgText(10), , MsgText(5)
'            strControlButton = MsgText(602)
'            .textA4901.SetFocus
'            Exit Sub
'         End If
'      Else
'         If .textA4901_C = MsgText(601) Then
'            MsgBox MsgText(10), , MsgText(5)
'            strControlButton = MsgText(602)
'            .textA4901_C.SetFocus
'            Exit Sub
'         End If
'      End If
'      If .textA4902 = MsgText(601) Then
'         MsgBox MsgText(10), , MsgText(5)
'         strControlButton = MsgText(602)
'         .textA4902.SetFocus
'         Exit Sub
'      End If
'      If .textA4903 = MsgText(601) And _
'         .textA4904 = MsgText(601) And _
'         .textA4905 = MsgText(601) Then
'         MsgBox "電話、傳真、E-Mail至少要輸入一項！", , MsgText(5)
'         strControlButton = MsgText(602)
'         .textA4903.SetFocus
'         Exit Sub
'      End If
'
'      If textA4901.Visible = True Then
'         Call textA4901_Validate(Cancel)
'         If Cancel = True Then
'            strControlButton = MsgText(602)
'            .textA4901.SetFocus
'            Exit Sub
'         End If
'      Else
'         Call textA4901_C_Validate(Cancel)
'         If Cancel = True Then
'            strControlButton = MsgText(602)
'            .textA4901_C.SetFocus
'            Exit Sub
'         End If
'      End If
'      Call textA4902_Validate(Cancel)
'      If Cancel = True Then
'         strControlButton = MsgText(602)
'         .textA4902.SetFocus
'         Exit Sub
'      End If
'      Call textA4903_Validate(Cancel)
'      If Cancel = True Then
'         strControlButton = MsgText(602)
'         .textA4903.SetFocus
'         Exit Sub
'      End If
'      Call textA4904_Validate(Cancel)
'      If Cancel = True Then
'         strControlButton = MsgText(602)
'         .textA4904.SetFocus
'         Exit Sub
'      End If
'      Call textA4905_Validate(Cancel)
'      If Cancel = True Then
'         strControlButton = MsgText(602)
'         .textA4905.SetFocus
'         Exit Sub
'      End If
'
'      '更新DB資料
'      If textA4901.Visible = True Then
'         strCompanyNo = textA4901
'      Else
'         strCompanyNo = textA4901_C
'      End If
'
'      adoTaie.BeginTrans
'      If strSaveConfirm = MsgText(3) Then '新增狀態時
'         strSql = "insert into acc490(a4901,a4902,a4903,a4904,a4905) " & _
'                  "values(" & CNULL(strCompanyNo) & "," & CNULL(.textA4902) & "," & CNULL(.textA4903) & _
'                  "," & CNULL(.textA4904) & "," & CNULL(.textA4905) & ")"
'      Else '修改
'         strSql = "update acc490 " & _
'                  "set a4902=" & CNULL(.textA4902) & _
'                     ",a4903=" & CNULL(.textA4903) & _
'                     ",a4904=" & CNULL(.textA4904) & _
'                     ",a4905=" & CNULL(.textA4905) & _
'                  " where a4901='" & strCompanyNo & "' "
'      End If
'      adoTaie.Execute strSql
'
'      adoTaie.CommitTrans
'
'      .AdodcRefresh
'      .FormDisabled
'   End With
'
'Checking:
'   If Err.Number = 0 Then
'      Exit Sub
'   Else '-2147168237
'      adoTaie.RollbackTrans
'   End If
'   MsgBox Err.Description, , MsgText(5)
'End Sub

Private Sub textA4901_GotFocus()
   OpenIme
   TextInverse textA4901
End Sub

Private Sub textA4901_Validate(Cancel As Boolean)
   '剔除跳行符號
   textA4901.Text = PUB_StringFilter(textA4901.Text)
   
   If textA4901.Enabled = False Then Exit Sub
   
   If textA4901.Text = "" Then Exit Sub
   If Not CheckLengthIsOK(textA4901, textA4901.MaxLength) Then
      Cancel = True
   End If
   
   If strSaveConfirm = MsgText(3) Then '新增狀態時,檢查是否有重覆
      If IsRecordExist(textA4901) = True Then
         MsgBox "該筆記錄已存在!!", , MsgText(5)
         textA4901.SetFocus
         Cancel = True
         Exit Sub
      End If
   End If
End Sub

Private Sub textA4901_C_GotFocus()
   OpenIme
   TextInverse textA4901_C
End Sub

Private Sub textA4901_C_Validate(Cancel As Boolean)
   '剔除跳行符號
   textA4901_C.Text = PUB_StringFilter(textA4901_C.Text)
   
   If textA4901_C.Enabled = False Then Exit Sub
   
   If textA4901_C.Text = "" Then Exit Sub
   If Not CheckLengthIsOK(textA4901_C, textA4901_C.MaxLength) Then
      Cancel = True
   End If
   
   If strSaveConfirm = MsgText(3) Then '新增狀態時,檢查是否有重覆
      If IsRecordExist(textA4901_C) = True Then
         MsgBox "該筆記錄已存在!!", , MsgText(5)
         textA4901_C.SetFocus
         Cancel = True
         Exit Sub
      End If
   End If
End Sub

' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal strKEY01 As String) As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   IsRecordExist = False
   
   strSql = "SELECT * FROM acc490 " & _
            "WHERE a4901='" & ChgSQL(strKEY01) & "'"
   
   ' 讀取資料庫
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   ' 檢查讀取的資料筆數
   If rsTmp.RecordCount > 0 Then
      IsRecordExist = True
   Else
      IsRecordExist = False
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

Private Sub textA4902_GotFocus()
   InverseTextBox textA4902
   OpenIme
End Sub
Private Sub textA4902_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If CheckLengthIsOK(textA4902, textA4902.MaxLength) = False Then
      Cancel = True
      textA4902_GotFocus
   End If
   If Cancel = False Then CloseIme
End Sub

Private Sub textA4903_GotFocus()
   InverseTextBox textA4903
   CloseIme
End Sub
Private Sub textA4903_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If CheckLengthIsOK(textA4903, textA4903.MaxLength) = False Then
      Cancel = True
      textA4903_GotFocus
   End If
   If Cancel = False Then CloseIme
End Sub

Private Sub textA4904_GotFocus()
   InverseTextBox textA4904
   CloseIme
End Sub
Private Sub textA4904_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If CheckLengthIsOK(textA4904, textA4904.MaxLength) = False Then
      Cancel = True
      textA4904_GotFocus
   End If
   If Cancel = False Then CloseIme
End Sub

Private Sub textA4905_GotFocus()
   CloseIme
   TextInverse textA4905
End Sub

'Modified by Lydia 2021/12/02 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub textA4905_KeyPress(KeyAscii As MSForms.ReturnInteger)
   'Modified by Lydia 2021/12/02 +val()
   PUB_EMailFilter Val(KeyAscii) 'Email輸入字元檢查
End Sub
Private Sub textA4905_Validate(Cancel As Boolean)
   If textA4905.Enabled = False Then Exit Sub
   
   If textA4905.Text = "" Then Exit Sub
   Cancel = Not PUB_CheckMail(textA4905.Text)
End Sub

'Add By Sindy 2016/11/7
Private Sub textA4912_GotFocus()
   InverseTextBox textA4912
   OpenIme
End Sub
Private Sub textA4912_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If CheckLengthIsOK(textA4912, textA4912.MaxLength) = False Then
      Cancel = True
      textA4912_GotFocus
   End If
   If Cancel = False Then CloseIme
End Sub

Private Sub textA4913_GotFocus()
   OpenIme
   TextInverse textA4913
End Sub
'Modified by Lydia 2021/12/02 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub textA4913_KeyPress(KeyAscii As MSForms.ReturnInteger)
   'Modified by Lydia 2021/12/14 +物件名稱
   KeyAscii = ChangeZIP(KeyAscii, textA4913)
End Sub
Private Sub textA4913_Validate(Cancel As Boolean)
   If textA4913.Enabled = False Then Exit Sub

   If textA4913.Text = "" Then Exit Sub
   textA4913 = ReplaceAddrTW(textA4913) 'Add by Amy 2024/12/25 址址使用貼上不會轉全型
   If Not CheckLengthIsOK(textA4913, textA4913.MaxLength) Then
      Cancel = True
   End If
End Sub

Private Sub textA4914_GotFocus()
   InverseTextBox textA4914
   OpenIme
End Sub
Private Sub textA4914_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If CheckLengthIsOK(textA4914, textA4914.MaxLength) = False Then
      Cancel = True
      textA4914_GotFocus
   End If
   If Cancel = False Then CloseIme
End Sub
'2016/11/7 END
