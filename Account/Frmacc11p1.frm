VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc11p1 
   AutoRedraw      =   -1  'True
   Caption         =   "收據抬頭基本資料查詢"
   ClientHeight    =   5016
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5016
   ScaleWidth      =   8760
   Begin VB.Frame Frame2 
      Height          =   350
      Left            =   6360
      TabIndex        =   8
      Top             =   -100
      Width           =   2100
      Begin VB.OptionButton Option3 
         Caption         =   "字首比對"
         Height          =   180
         Index           =   0
         Left            =   72
         TabIndex        =   10
         Top             =   144
         Width           =   1020
      End
      Begin VB.OptionButton Option3 
         Caption         =   "模糊比對"
         Height          =   180
         Index           =   1
         Left            =   1050
         TabIndex        =   9
         Top             =   144
         Value           =   -1  'True
         Width           =   1020
      End
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      ItemData        =   "Frmacc11p1.frx":0000
      Left            =   210
      List            =   "Frmacc11p1.frx":0002
      TabIndex        =   1
      Top             =   260
      Width           =   2052
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc11p1.frx":0004
      Height          =   3996
      Left            =   60
      TabIndex        =   4
      Top             =   696
      Width           =   8628
      _ExtentX        =   15219
      _ExtentY        =   7049
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   18
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "A4201"
         Caption         =   "收據抬頭"
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
         DataField       =   "A4202"
         Caption         =   "統一編號"
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
         DataField       =   "A4203"
         Caption         =   "地址"
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
         DataField       =   "A4204"
         Caption         =   "電話"
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
         DataField       =   "A4205"
         Caption         =   "傳真"
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
      BeginProperty Column05 
         DataField       =   "st02"
         Caption         =   "智權人員"
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
      BeginProperty Column06 
         DataField       =   "A4207"
         Caption         =   "聯絡人"
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
         DataField       =   "A4208"
         Caption         =   "備註"
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
         Size            =   275
         BeginProperty Column00 
            ColumnWidth     =   1307.906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1284.095
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   984.189
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1247.811
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1031.811
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column06 
         EndProperty
         BeginProperty Column07 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
      Left            =   210
      Top             =   450
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   550
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
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "選「收據抬頭」且起迄條件相同"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   216
      Left            =   240
      TabIndex        =   7
      Top             =   36
      Width           =   3192
   End
   Begin MSForms.ComboBox Combo3 
      Height          =   336
      Left            =   5736
      TabIndex        =   3
      Top             =   260
      Width           =   2808
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "4948;582"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo2 
      Height          =   336
      Left            =   2616
      TabIndex        =   2
      Top             =   260
      Width           =   2808
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "4948;582"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "~"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   5496
      TabIndex        =   6
      Top             =   260
      Width           =   132
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "="
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   2376
      TabIndex        =   5
      Top             =   260
      Width           =   132
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "註：按ESC鍵，即可離開查詢，進入維護作業！"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   300
      TabIndex        =   0
      Top             =   4740
      Width           =   5265
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   4800
      Visible         =   0   'False
      Width           =   132
   End
End
Attribute VB_Name = "Frmacc11p1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/14 Form2.0已修改
'Create by Sindy 2013/12/19
Option Explicit

Public adoacc420 As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset


Private Sub Combo1_Change()
'   SelectScope
End Sub

Private Sub Combo1_Click()
'   SelectScope
End Sub

Private Sub Combo2_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Combo2_Validate(Cancel As Boolean)
   Combo3 = Combo2
End Sub

Private Sub Combo3_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
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
   Me.Width = 8880
   Me.Height = 5430
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath2)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   strCon1 = "收據抬頭"
   strCon2 = "統一編號"
   strCon3 = "智權人員"
   Combo1.AddItem MsgText(31)
   Combo1.AddItem strCon1
   Combo1.AddItem strCon2
   Combo1.AddItem strCon3
   Combo1 = MsgText(31)
   SelectScope
   OpenTable
   'Add by Amy 2024/12/23 電腦中心可選 字首或模糊比對 ex:櫃檯收到支票要電腦中心找「卓異」是誰的客戶
   Frame2.Visible = False
   If Pub_StrUserSt03 = "M51" Then
      Frame2.Visible = True
      Label4 = Label4 & "，可選字首或模糊比對"
      Option3(1).Value = 1
   Else
      Label4 = Label4 & "，以字首比對"
      Option3(0).Value = 1
   End If
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Adodc1.Recordset.RecordCount <> 0 Then
      strCompanyNo = Adodc1.Recordset.Fields("A4201").Value
   Else
      strCompanyNo = MsgText(601)
   End If
   StatusClear
   tool1_enabled
   Frmacc11p0.Enabled = True
   Frmacc11p0.Show
   Set Frmacc11p1 = Nothing
End Sub

'*************************************************
'  搜尋條件範圍值，並代入 Combo2、Combo3 之中
'
'*************************************************
Private Sub SelectScope()
   strCondition = MsgText(601)
   If Combo1 = MsgText(31) Then
      Exit Sub
   End If
   Select Case Combo1
      Case strCon1
         strCondition = "A4201"
      Case strCon2
         strCondition = "A4202"
      Case strCon3
         strCondition = "A4206"
   End Select
   If strCondition = MsgText(601) Then
      Exit Sub
   End If
   Combo2.Clear
   Combo3.Clear
   adoacc420.CursorLocation = adUseClient
   adoacc420.Open "select distinct " & strCondition & " from acc420 order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
   Do While adoacc420.EOF = False
      If IsNull(adoacc420.Fields(0).Value) = False Then
         Combo2.AddItem adoacc420.Fields(0).Value
         Combo3.AddItem adoacc420.Fields(0).Value
      End If
      adoacc420.MoveNext
   Loop
   adoacc420.Close
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)
   Select Case KeyCode
      Case vbKeyF12
         Screen.MousePointer = vbHourglass
         Acc020Query
         Screen.MousePointer = vbDefault
   End Select
   KeyEnter KeyCode
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

'*************************************************
'  傳票資料查詢
'
'*************************************************
Private Sub Acc020Query()
   Dim strWhere As String, strQ As String 'Add by Amy 2023/02/01
   
On Error GoTo Checking
   If adoadodc1.State = adStateOpen Then
      adoadodc1.Close
   End If
   adoadodc1.CursorLocation = adUseClient
   Select Case Combo1
      Case strCon1
         strCondition = "A4201"
      Case strCon2
         strCondition = "A4202"
      Case strCon3
         strCondition = "A4206"
      Case MsgText(31)
         adoadodc1.Open "select acc420.*,st02 from acc420,staff where a4206=st01(+) order by a4210 asc, a4211 asc", adoTaie, adOpenStatic, adLockReadOnly
         Adodc1.Recordset.Requery
         Exit Sub
      Case Else
         Exit Sub
   End Select
   'Modify by Amy 2023/02/01 選「收據抬頭」時,若起迄字相同 以Instr找收據抬頭(a4201) ex:櫃檯收到支票要電腦中心找「卓異」是誰的客戶
   If strCondition = "A4201" And Trim(Combo2) <> MsgText(601) And Trim(Combo3) <> MsgText(601) And Trim(Combo2) = Trim(Combo3) Then
      'Modify by Amy 2024/12/23 +字首或模糊比對
      If Option3(0).Value = True Then
         strWhere = "=1"
      Else
          strWhere = ">0"
      End If
      strWhere = "And InStr(" & strCondition & "," & CNULL(ChgSQL(Combo2)) & ")" & strWhere & " Order by " & strCondition & " asc"
   ElseIf Combo3 = MsgText(601) Then
      'adoadodc1.Open "select acc420.*,st02 from acc420,staff where a4206=st01(+) and " & strCondition & " = '" & ChgSQL(Combo2) & "' order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
      strWhere = "And " & strCondition & " = '" & ChgSQL(Combo2) & "' order by " & strCondition & " asc"
   Else
      If Combo2 = MsgText(601) Then
         'adoadodc1.Open "select acc420.*,st02 from acc420,staff where a4206=st01(+) and " & strCondition & " <= '" & ChgSQL(Combo3) & "' order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
         strWhere = "And " & strCondition & " <= '" & ChgSQL(Combo3) & "' order by " & strCondition & " asc"
      Else
         'adoadodc1.Open "select acc420.*,st02 from acc420,staff where a4206=st01(+) and " & strCondition & " >= '" & ChgSQL(Combo2) & "' and " & strCondition & " <= '" & ChgSQL(Combo3) & "' order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
         strWhere = "And " & strCondition & " >= '" & ChgSQL(Combo2) & "' And " & strCondition & " <= '" & ChgSQL(Combo3) & "' order by " & strCondition & " asc"
      End If
   End If
   strQ = "Select acc420.*,st02 From acc420,staff Where a4206=st01(+) " & strWhere
   adoadodc1.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
   'end 2023/02/01
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
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "select acc420.*,st02 from acc420,staff where a4206=st01(+) and a4201='' order by a4210 asc, a4211 asc", adoTaie, adOpenStatic, adLockReadOnly
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub
