VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060502 
   BorderStyle     =   1  '單線固定
   Caption         =   "員工外譯代號對照檔維護"
   ClientHeight    =   3768
   ClientLeft      =   1752
   ClientTop       =   1860
   ClientWidth     =   7608
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3768
   ScaleWidth      =   7608
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6885
      Top             =   660
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060502.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060502.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060502.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060502.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060502.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060502.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060502.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060502.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060502.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060502.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060502.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   576
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   7608
      _ExtentX        =   13420
      _ExtentY        =   1016
      ButtonWidth     =   1101
      ButtonHeight    =   974
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "新增"
            Key             =   "keyInsert"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "修改"
            Key             =   "keyUpdate"
            ImageIndex      =   2
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "刪除"
            Key             =   "keyDelete"
            ImageIndex      =   3
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "查詢"
            Key             =   "keyQuery"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "第一筆"
            Key             =   "keyFirst"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "前一筆"
            Key             =   "keyPrevious"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "後一筆"
            Key             =   "keyNext"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "最後筆"
            Key             =   "keyLast"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "確定"
            Key             =   "keyOk"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "取消"
            Key             =   "keyCancel"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "結束"
            Key             =   "keyExit"
            ImageIndex      =   11
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frm060502.frx":20F4
      Height          =   2295
      Left            =   90
      TabIndex        =   3
      Top             =   1440
      Width           =   6195
      _ExtentX        =   10922
      _ExtentY        =   4043
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   14
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
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
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "SName"
         Caption         =   "員工名稱"
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
         DataField       =   "SCode1"
         Caption         =   "員工代碼"
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
         DataField       =   "SCode2"
         Caption         =   "外譯代碼"
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
      BeginProperty Column03 
         DataField       =   "SState"
         Caption         =   "狀態"
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
            ColumnWidth     =   1247.811
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1188.284
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1404.284
         EndProperty
         BeginProperty Column03 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
      Left            =   6300
      Top             =   1590
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
   Begin MSForms.TextBox txtData 
      Height          =   285
      Index           =   2
      Left            =   3420
      TabIndex        =   2
      Top             =   1050
      Width           =   1110
      VariousPropertyBits=   671105051
      MaxLength       =   6
      Size            =   "1958;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtData 
      Height          =   285
      Index           =   1
      Left            =   1125
      TabIndex        =   1
      Top             =   1050
      Width           =   1110
      VariousPropertyBits=   671105051
      MaxLength       =   6
      Size            =   "1958;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtData 
      Height          =   285
      Index           =   0
      Left            =   1125
      TabIndex        =   0
      Top             =   720
      Width           =   2250
      VariousPropertyBits=   671105051
      MaxLength       =   6
      Size            =   "3969;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblName2 
      Height          =   285
      Left            =   4560
      TabIndex        =   8
      Top             =   1050
      Width           =   1380
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2434;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "外譯代碼："
      Height          =   180
      Index           =   2
      Left            =   2475
      TabIndex        =   7
      Top             =   1050
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "員工代碼："
      Height          =   180
      Index           =   1
      Left            =   180
      TabIndex        =   6
      Top             =   1095
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "員工名稱："
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   5
      Top             =   765
      Width           =   900
   End
End
Attribute VB_Name = "frm060502"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/07 Form2.0已修改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/13 日期欄已修改
'Add by Morgan 2007/5/22
Option Explicit

'使用者權限設定
Dim bolInsert As Boolean
Dim bolUpdate As Boolean
Dim bolDelete As Boolean
Dim bolSelect As Boolean
Dim iCurState As Integer '目前狀態: 0=瀏覽 1=新增 2=修改 3=查詢 9=無資料
Dim adoadodc1 As New ADODB.Recordset
Dim bolBarShow As Boolean


Private Sub MoveFirst()
   If adoadodc1.RecordCount > 0 Then
      adoadodc1.MoveFirst
      FormShow
      RecordShow
   End If
End Sub

Private Sub MoveLast()
   If adoadodc1.RecordCount > 0 Then
      adoadodc1.MoveLast
      FormShow
      RecordShow
   End If
End Sub

Private Sub MoveNext()
   If adoadodc1.RecordCount > 0 Then
      adoadodc1.MoveNext
      If Not adoadodc1.EOF Then
         FormShow
         RecordShow
      Else
         adoadodc1.MoveLast
      End If
   End If
End Sub

Private Sub MovePrevious()
   If adoadodc1.RecordCount > 0 Then
      adoadodc1.MovePrevious
      If Not adoadodc1.BOF Then
         FormShow
         RecordShow
      Else
         adoadodc1.MoveFirst
      End If
   End If
End Sub
'*************************************************
'  顯示資料表
'
'*************************************************
Private Sub FormShow()
   If Adodc1.Recordset.EOF And Adodc1.Recordset.BOF Then
      iCurState = 9
      txtData(1).Tag = ""
   Else
   
      'Added by Morgan 2025/11/13 修正查詢不到資料會當掉問題
      If Adodc1.Recordset.EOF And txtData(1).Tag <> "" Then
         strExc(0) = "SCode1='" & txtData(1).Tag & "'"
         Adodc1.Recordset.Find strExc(0), 0, adSearchForward, 1
         If Adodc1.Recordset.EOF Then
            Exit Sub
         End If
      End If
      'end 2025/11/13
      
      For intI = 0 To 2
         txtData(intI) = "" & Adodc1.Recordset.Fields(intI).Value
      Next
      lblName2 = "" & Adodc1.Recordset.Fields(3).Value
      If txtData(0) = lblName2 Then
         lblName2.Visible = False
      Else
         lblName2.Visible = True
      End If
      txtData(1).Tag = txtData(1)
      RecordShow
   End If
End Sub
'*************************************************
'  顯示筆數
'
'*************************************************
Private Sub RecordShow()
   Forms(0).StatusBar1.Panels(2).Text = Adodc1.Recordset.Bookmark & MsgText(35) & Adodc1.Recordset.RecordCount
End Sub
'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "select st1.st02 SName,sim01 SCode1,sim02 SCode2,st2.st02 SName2,decode(st1.st04,'1','在職','離職') SState from Staff_IdMap,staff st1,staff st2 WHERE st1.ST01(+)=sim01 and st2.st01(+)=sim02 order by 2", cnnConnection, adOpenStatic, adLockReadOnly
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

Private Sub AdodcRefresh()
   Adodc1.Recordset.ReQuery
End Sub

Private Sub DataGrid1_SelChange(Cancel As Integer)
   FormShow
End Sub

Private Sub Form_Activate()
   bolBarShow = Forms(0).StatusBar1.Visible
   Forms(0).StatusBar1.Visible = True
End Sub

Private Sub Form_Deactivate()
   Forms(0).StatusBar1.Visible = bolBarShow
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   setAuthority
   OpenTable
   If Adodc1.Recordset.RecordCount > 0 Then
      Adodc1.Recordset.MoveFirst
      FormShow
      RecordShow
      iCurState = 0
   Else
      iCurState = 9
   End If
   SetToolBar
   FormEnable
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm060502 = Nothing
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)
   Select Case KeyCode
      Case vbKeyF2 '新增
         If TBar1.Buttons(1).Enabled = True Then
            Call Tbar1_ButtonClick(TBar1.Buttons(1))
         End If
      Case vbKeyF3 '修改
         If TBar1.Buttons(2).Enabled = True Then
            Call Tbar1_ButtonClick(TBar1.Buttons(2))
         End If
      Case vbKeyF5 '刪除
         If TBar1.Buttons(3).Enabled = True Then
            Call Tbar1_ButtonClick(TBar1.Buttons(3))
         End If
      Case vbKeyF4 '查詢
         If TBar1.Buttons(4).Enabled = True Then
            Call Tbar1_ButtonClick(TBar1.Buttons(4))
         End If
      Case vbKeyHome '第一筆
         If TBar1.Buttons(6).Enabled = True Then
            Call Tbar1_ButtonClick(TBar1.Buttons(6))
         End If
      Case vbKeyPageUp '上一筆
         If TBar1.Buttons(7).Enabled = True Then
            Call Tbar1_ButtonClick(TBar1.Buttons(7))
         End If
      Case vbKeyPageDown '下一筆
         If TBar1.Buttons(8).Enabled = True Then
            Call Tbar1_ButtonClick(TBar1.Buttons(8))
         End If
      Case vbKeyEnd '最後筆
         If TBar1.Buttons(9).Enabled = True Then
            Call Tbar1_ButtonClick(TBar1.Buttons(9))
         End If
      Case vbKeyF9 '存檔
         If TBar1.Buttons(11).Enabled = True Then
            Call Tbar1_ButtonClick(TBar1.Buttons(11))
         End If
      Case vbKeyF10 '取消
         If TBar1.Buttons(12).Enabled = True Then
            Call Tbar1_ButtonClick(TBar1.Buttons(12))
         End If
   End Select
   
End Sub
'使用者權限設定
Private Sub setAuthority()
   bolInsert = IsUserHasRightOfFunction(Me.Name, strAdd, False)
   bolUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
   bolDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   bolSelect = IsUserHasRightOfFunction(Me.Name, strFind, False)
End Sub
'工具列控制
Private Sub SetToolBar()
   For intI = 1 To 13
      TBar1.Buttons(intI).Enabled = False
   Next
   TBar1.Buttons(14).Enabled = True
   Select Case iCurState
      Case 0 '瀏覽
         If bolInsert Then
            TBar1.Buttons(1).Enabled = True
         End If
         If bolUpdate Then
            TBar1.Buttons(2).Enabled = True
         End If
         If bolDelete Then
            TBar1.Buttons(3).Enabled = True
         End If
         If bolSelect Then
            TBar1.Buttons(4).Enabled = True
         End If
         TBar1.Buttons(6).Enabled = True
         TBar1.Buttons(7).Enabled = True
         TBar1.Buttons(8).Enabled = True
         TBar1.Buttons(9).Enabled = True
      Case 1, 2, 4 '1:新增  '2:修改  '4查詢
         TBar1.Buttons(11).Enabled = True
         TBar1.Buttons(12).Enabled = True
      Case 9 '無資料
         If bolInsert Then
            TBar1.Buttons(1).Enabled = True
         End If
   End Select
End Sub

Private Function FormDelete() As Boolean
   
On Error GoTo ErrHnd
   strSql = "delete from Staff_IdMap where sim01='" & txtData(1).Text & "'"
   cnnConnection.Execute strSql, intI
   FormDelete = True
   FormClear
   AdodcRefresh
   FormShow
   Exit Function
   
ErrHnd:
   MsgBox Err.Description, vbCritical
End Function

Public Sub FormClear()
   Dim oObj As Object
   For Each oObj In Me.Controls
      If TypeName(oObj) = "TextBox" Then
         oObj.Text = Empty
      End If
   Next
   lblName2 = ""
End Sub

Private Function FormSave() As Boolean
   If SaveCheck = False Then
      Exit Function
   End If
   
On Error GoTo ErrHnd
   '新增
   If iCurState = 1 Then
      strSql = "INSERT INTO STAFF_IDMAP(SIM01,SIM02)" & _
         " VALUES('" & txtData(1).Text & "','" & txtData(2) & "')"
   '修改
   Else
      strSql = "UPDATE STAFF_IDMAP SET SIM02='" & txtData(2) & "'" & _
         " WHERE SIM01='" & txtData(1) & "'"
   End If
   cnnConnection.Execute strSql, intI
   FormSave = True
   Exit Function
   
ErrHnd:
   If Err.Number = -2147217873 Then
      MsgBox "資料已存在，請改為修改模式作業！"
   Else
      MsgBox Err.Description, vbCritical
   End If
End Function

Private Function SaveCheck() As Boolean
   If txtData(1) = "" Then
      txtData(0) = ""
      MsgBox "員工代碼不可空白！", vbExclamation
      Exit Function
   ElseIf txtData(2) = "" Then
      lblName2 = ""
      MsgBox "外譯代碼不可空白！", vbExclamation
      Exit Function
   Else
      txtData(0) = GetStaffName(txtData(1), True)
      If txtData(0) = "" Then
         MsgBox "員工代碼不存在！", vbExclamation
         txtData(1).SetFocus
         Exit Function
      End If
      If Left(GetStaffDepartment(txtData(1)), 2) = "F5" Then
         MsgBox "員工代碼不可為翻譯人員代碼！", vbExclamation
         txtData(1).SetFocus
         Exit Function
      End If
      lblName2 = GetStaffName(txtData(2), True)
      If lblName2 = "" Then
         MsgBox "外譯代碼不存在！", vbExclamation
         txtData(2).SetFocus
         Exit Function
      End If
      
      If Left(GetStaffDepartment(txtData(2)), 2) <> "F5" Then
         MsgBox "外譯代碼只可為翻譯人員代碼！", vbExclamation
         txtData(2).SetFocus
         Exit Function
      End If
      If txtData(0) <> lblName2 Then
         MsgBox "外譯名稱與員工名稱不同！", vbExclamation
         txtData(2).SetFocus
         Exit Function
      End If
   End If
   SaveCheck = True
End Function
Private Function FormSearch() As Boolean
   If txtData(1) = "" And txtData(2) = "" And txtData(0) = "" Then
      MsgBox "查詢條件不可全部空白！", vbExclamation
   Else
      If txtData(0) <> "" Then
         strExc(0) = "SName='" & txtData(0) & "'"
      ElseIf txtData(1) <> "" Then
         strExc(0) = "SCode1='" & txtData(1) & "'"
      Else
         strExc(0) = "SCode2='" & txtData(2) & "'"
      End If
      Adodc1.Recordset.Find strExc(0), 0, adSearchForward, 1
      If Adodc1.Recordset.EOF = False Then
         FormSearch = True
      Else
         MsgBox "無符合資料！", vbExclamation
      End If
   End If
End Function

Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case Button.Index
      Case 1 '新增
         iCurState = 1
         FormClear
      Case 2 '修改
         iCurState = 2
      Case 3 '刪除
         If MsgBox("是否要刪除此筆資料?", vbCritical + vbYesNo + vbDefaultButton2, "詢問") = vbYes Then
            FormDelete
         End If
      Case 4 '查詢
         iCurState = 4
         FormClear
      Case 6 '第一筆
         MoveFirst
      Case 7 '上一筆
         MovePrevious
      Case 8 '下一筆
         MoveNext
      Case 9 '最後筆
         MoveLast
      Case 11 '確定
         Select Case iCurState
            Case 1, 2 '新增,修改
               'Add by Sindy 2021/12/07 檢查畫面上的物件是否含有Unicode文字
               If PUB_ChkUniText(Me, True, True) = False Then
                  Exit Sub
               End If

               If FormSave = True Then
                  AdodcRefresh
               Else
                  Exit Sub
               End If
         End Select
         If FormSearch = True Then
            iCurState = 0
            FormShow
         End If
      Case 12 '取消
         Select Case iCurState
            Case 1, 2 '新增,修改
               If MsgBox("你並未存檔，確定離開嗎 ?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
                  Exit Sub
               End If
         End Select
         iCurState = 0
         FormShow
      Case 14
      '結束
         If iCurState = 2 Or iCurState = 1 Then
            If MsgBox("你並未存檔，確定離開嗎 ?", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
               Unload Me
               Exit Sub
            End If
         Else
            Unload Me
            Exit Sub
         End If
         
   End Select
   SetToolBar
   FormEnable
End Sub
Public Sub FormEnable()
   Select Case iCurState
      Case 0 '瀏覽
         txtData(0).Locked = True
         txtData(1).Locked = True
         txtData(2).Locked = True
         DataGrid1.Enabled = True
      Case 1 '新增
         txtData(0).Locked = True
         txtData(1).Locked = False
         txtData(2).Locked = False
         lblName2.Visible = True
         DataGrid1.Enabled = False
         txtData(1).SetFocus
      Case 2 '修改
         txtData(0).Locked = True
         txtData(1).Locked = True
         txtData(2).Locked = False
         DataGrid1.Enabled = False
         lblName2.Visible = True
         txtData(2).SetFocus
      Case 4 '查詢
         txtData(0).Locked = False
         txtData(1).Locked = False
         txtData(2).Locked = False
         DataGrid1.Enabled = False
         txtData(0).SetFocus
      Case Else
         txtData(0).Locked = True
         txtData(1).Locked = True
         txtData(2).Locked = True
         DataGrid1.Enabled = False
   End Select
   
End Sub

Private Sub Txtdata_GotFocus(Index As Integer)
   If txtData(Index).Locked = False Then
      TextInverse txtData(Index)
   End If
End Sub

Private Sub txtdata_KeyPress(Index As Integer, KeyAscii As ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Txtdata_Validate(Index As Integer, Cancel As Boolean)
   If Index = 1 Then
      If iCurState = 1 Then
         If txtData(1) <> "" Then
            txtData(0) = GetStaffName(txtData(1), True)
         Else
            txtData(0) = ""
         End If
      End If
   ElseIf Index = 2 Then
      If iCurState = 2 Then
         If txtData(2) <> "" Then
            lblName2 = GetStaffName(txtData(2), True)
         Else
            lblName2 = ""
         End If
      End If
   End If
End Sub
