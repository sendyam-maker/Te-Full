VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc1171 
   AutoRedraw      =   -1  'True
   Caption         =   "應付款資料查詢"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5025
   ScaleWidth      =   8760
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
      Height          =   300
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2052
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc1171.frx":0000
      Height          =   4092
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   8292
      _ExtentX        =   14631
      _ExtentY        =   7223
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   20
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "應付款資料"
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "a0o01"
         Caption         =   "應付款單號"
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
         DataField       =   "contect"
         Caption         =   "往來對象"
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
         DataField       =   "a0o04"
         Caption         =   "發票號碼"
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
         DataField       =   "a0o05"
         Caption         =   "入帳日期"
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
      BeginProperty Column04 
         DataField       =   "a0o06"
         Caption         =   "欲處理日期"
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
      BeginProperty Column05 
         DataField       =   "a0o10"
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
         Size            =   344
         BeginProperty Column00 
            ColumnWidth     =   1349.858
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2654.929
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1365.165
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   1124.787
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   1260.284
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   3960
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   312
      Left            =   240
      Top             =   600
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
   Begin MSForms.ComboBox Combo3 
      Height          =   335
      Left            =   5760
      TabIndex        =   5
      Top             =   240
      Width           =   2772
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "4890;591"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo2 
      Height          =   335
      Left            =   2640
      TabIndex        =   1
      Top             =   240
      Width           =   2772
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "4890;591"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
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
      Height          =   252
      Left            =   2400
      TabIndex        =   4
      Top             =   240
      Width           =   132
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
      Height          =   252
      Left            =   5520
      TabIndex        =   3
      Top             =   240
      Width           =   132
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   4800
      Visible         =   0   'False
      Width           =   132
   End
End
Attribute VB_Name = "Frmacc1171"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/11/26 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/28 日期欄已修改
Option Explicit
Public adoacc0o0 As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset

Private Sub Combo1_Change()
'   SelectScope
End Sub

Private Sub Combo1_Click()
'   SelectScope
End Sub

'Modify by Amy 2021/08/20 改Form2.0 原:Integer
Private Sub Combo2_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Combo2_Validate(Cancel As Boolean)
   Combo3 = Combo2
End Sub

'Modify by Amy 2021/08/20 改Form2.0 原:Integer
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
   Me.Width = 8850
   Me.Height = 5400
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath2)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   strCon1 = "應付款單號"
   strCon2 = "往來對象"
   strCon3 = "發票號碼"
   strCon4 = "入帳日期"
   strCon5 = "欲處理日期"
   strCon6 = "中文名稱"
   Combo1.AddItem MsgText(31)
   Combo1.AddItem strCon1
   Combo1.AddItem strCon2
   Combo1.AddItem strCon3
   Combo1.AddItem strCon4
   Combo1.AddItem strCon5
   Combo1.AddItem strCon6
   Combo1 = MsgText(31)
   SelectScope
   OpenTable
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Adodc1.Recordset.RecordCount <> 0 Then
      strItemNo = Adodc1.Recordset.Fields("a0o01").Value
   Else
      strItemNo = MsgText(601)
   End If
   StatusClear
   tool1_enabled
   Frmacc1170.Enabled = True
   Frmacc1170.Show
   Set Frmacc1171 = Nothing
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
         strCondition = "a0o01"
      Case strCon2
         strCondition = "a0o03"
      Case strCon3
         strCondition = "a0o04"
      Case strCon4
         strCondition = "a0o05"
      Case strCon5
         strCondition = "a0o06"
   End Select
   If strCondition = MsgText(601) Then
      Exit Sub
   End If
   Combo2.Clear
   Combo3.Clear
   adoacc0o0.CursorLocation = adUseClient
   adoacc0o0.Open "select distinct " & strCondition & " from acc0o0 order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
   Do While adoacc0o0.EOF = False
      If IsNull(adoacc0o0.Fields(0).Value) = False Then
         Combo2.AddItem adoacc0o0.Fields(0).Value
         Combo3.AddItem adoacc0o0.Fields(0).Value
      End If
      adoacc0o0.MoveNext
   Loop
   adoacc0o0.Close
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)
   Select Case KeyCode
      Case vbKeyF12
         Screen.MousePointer = vbHourglass
         Acc0o0Query
         Screen.MousePointer = vbDefault
   End Select
   KeyEnter KeyCode
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

'*************************************************
'  票據資料查詢
'
'*************************************************
Private Sub Acc0o0Query()
Dim strUnion As String

On Error GoTo Checking
   If adoadodc1.State = adStateOpen Then
      adoadodc1.Close
   End If
   adoadodc1.CursorLocation = adUseClient
   Select Case Combo1
      Case strCon1
         strCondition = "a0o01"
      Case strCon2
         strCondition = "a0o03"
      Case strCon3
         strCondition = "a0o04"
      Case strCon4
         strCondition = "a0o05"
      Case strCon5
         strCondition = "a0o06"
      Case MsgText(31)
         strUnion = "select a0o01, a0o04, a0o05, a0o06, a0o10, a0i02 as contect from acc0o0, acc0i0 where a0o02 = '1' and a0o03 = a0i01"
         strUnion = strUnion & " union select a0o01, a0o04, a0o05, a0o06, a0o10, cu04 as contect from acc0o0, customer where a0o02 = '2' and substr(a0o03, 1, 8) = cu01 and substr(a0o03, 9, 1) = decode(a0o02, '2', cu02)"
         strUnion = strUnion & " union select a0o01, a0o04, a0o05, a0o06, a0o10, st02 as contect from acc0o0, staff where a0o02 = '3' and a0o03 = st01"
         strUnion = strUnion & " union select a0o01, a0o04, a0o05, a0o06, a0o10, '' as contect from acc0o0 where a0o02 = '4' order by contect asc"
         adoadodc1.Open strUnion, adoTaie, adOpenStatic, adLockReadOnly
         Adodc1.Recordset.Requery
         Exit Sub
   End Select
   If Combo1 = strCon6 Then
      strUnion = "select a0o01, a0o04, a0o05, a0o06, a0o10, a0i02 as contect from acc0o0, acc0i0 where a0o02 = '1' and a0o03 = a0i01 and instr(a0i02, '" & Combo2 & "') <> 0"
      strUnion = strUnion & " union select a0o01, a0o04, a0o05, a0o06, a0o10, cu04 as contect from acc0o0, customer where a0o02 = '2' and substr(a0o03, 1, 8) = cu01 and substr(a0o03, 9, 1) = decode(a0o02, '2', cu02) and instr(cu04, '" & Combo2 & "') <> 0"
      strUnion = strUnion & " union select a0o01, a0o04, a0o05, a0o06, a0o10, st02 as contect from acc0o0, staff where a0o02 = '3' and a0o03 = st01  and instr(st02, '" & Combo2 & "') <> 0 order by contect asc"
      'strUnion = strUnion & " union select a0o01, a0o04, a0o05, a0o06, a0o10, '' as contect from acc0o0 where a0o02 = '4' order by contect asc"
      adoadodc1.Open strUnion, adoTaie, adOpenStatic, adLockReadOnly
   Else
      If Combo3 = MsgText(601) Then
         If Combo1 = strCon4 Or Combo1 = strCon5 Then
            strUnion = "select a0o01, a0o04, a0o05, a0o06, a0o10, a0i02 as contect from acc0o0, acc0i0 where a0o02 = '1' and a0o03 = a0i01 and " & strCondition & " = " & Val(Combo2) & ""
            strUnion = strUnion & " union select a0o01, a0o04, a0o05, a0o06, a0o10, cu04 as contect from acc0o0, customer where a0o02 = '2' and substr(a0o03, 1, 8) = cu01 and substr(a0o03, 9, 1) = decode(a0o02, '2', cu02) and " & strCondition & " = " & Val(Combo2) & ""
            strUnion = strUnion & " union select a0o01, a0o04, a0o05, a0o06, a0o10, st02 as contect from acc0o0, staff where a0o02 = '2' and a0o03 = st01 and " & strCondition & " = " & Val(Combo2) & ""
            strUnion = strUnion & " union select a0o01, a0o04, a0o05, a0o06, a0o10, '' as contect from acc0o0 where a0o02 = '4' and " & strCondition & " = " & Val(Combo2) & " order by contect asc"
            adoadodc1.Open strUnion, adoTaie, adOpenStatic, adLockReadOnly
         Else
            strUnion = "select a0o01, a0o04, a0o05, a0o06, a0o10, a0i02 as contect from acc0o0, acc0i0 where a0o02 = '1' and a0o03 = a0i01 and " & strCondition & " = '" & Combo2 & "'"
            strUnion = strUnion & " union select a0o01, a0o04, a0o05, a0o06, a0o10, cu04 as contect from acc0o0, customer where a0o02 = '2' and substr(a0o03, 1, 8) = cu01 and substr(a0o03, 9, 1) = decode(a0o02, '2', cu02) and " & strCondition & " = '" & Combo2 & "'"
            strUnion = strUnion & " union select a0o01, a0o04, a0o05, a0o06, a0o10, st02 as contect from acc0o0, staff where a0o02 = '3' and a0o03 = st01 and " & strCondition & " = '" & Combo2 & "'"
            strUnion = strUnion & " union select a0o01, a0o04, a0o05, a0o06, a0o10, '' as contect from acc0o0 where a0o02 = '4' and " & strCondition & " = '" & Combo2 & "' order by contect asc"
            adoadodc1.Open strUnion, adoTaie, adOpenStatic, adLockReadOnly
         End If
      Else
         If Combo2 = MsgText(601) Then
            If Combo1 = strCon4 Or Combo1 = strCon5 Then
               strUnion = "select a0o01, a0o04, a0o05, a0o06, a0o10, a0i02 as contect from acc0o0, acc0i0 where a0o02 = '1' and a0o03 = a0i01 and " & strCondition & " <= " & Val(Combo3) & ""
               strUnion = strUnion & " union select a0o01, a0o04, a0o05, a0o06, a0o10, cu04 as contect from acc0o0, customer where a0o02 = '2' and substr(a0o03, 1, 8) = cu01 and substr(a0o03, 9, 1) = decode(a0o02, '2', cu02) and " & strCondition & " <= " & Val(Combo3) & ""
               strUnion = strUnion & " union select a0o01, a0o04, a0o05, a0o06, a0o10, st02 as contect from acc0o0, staff where a0o02 = '3' and a0o03 = st01 and " & strCondition & " <= " & Val(Combo3) & ""
               strUnion = strUnion & " union select a0o01, a0o04, a0o05, a0o06, a0o10, '' as contect from acc0o0 where a0o02 = '4' and " & strCondition & " <= " & Val(Combo3) & " order by contect asc"
               adoadodc1.Open strUnion, adoTaie, adOpenStatic, adLockReadOnly
            Else
               strUnion = "select a0o01, a0o04, a0o05, a0o06, a0o10, a0i02 as contect from acc0o0, acc0i0 where a0o02 = '1' and a0o03 = a0i01 and " & strCondition & " <= '" & Combo3 & "'"
               strUnion = strUnion & " union select a0o01, a0o04, a0o05, a0o06, a0o10, cu04 as contect from acc0o0, customer where a0o02 = '2' and substr(a0o03, 1, 8) = cu01 and substr(a0o03, 9, 1) = decode(a0o02, '2', cu02) and " & strCondition & " <= '" & Combo3 & "'"
               strUnion = strUnion & " union select a0o01, a0o04, a0o05, a0o06, a0o10, st02 as contect from acc0o0, staff where a0o02 = '3' and a0o03 = st01 and " & strCondition & " <= '" & Combo3 & "'"
               strUnion = strUnion & " union select a0o01, a0o04, a0o05, a0o06, a0o10, '' as contect from acc0o0 where a0o02 = '4' and " & strCondition & " <= '" & Combo3 & "' order by contect asc"
               adoadodc1.Open strUnion, adoTaie, adOpenStatic, adLockReadOnly
            End If
         Else
            If Combo1 = strCon4 Or Combo1 = strCon5 Then
               strUnion = "select a0o01, a0o04, a0o05, a0o06, a0o10, a0i02 as contect from acc0o0, acc0i0 where a0o02 = '1' and a0o03 = a0i01 and " & strCondition & " >= " & Val(Combo2) & " and " & strCondition & " <= " & Val(Combo3) & ""
               strUnion = strUnion & " union select a0o01, a0o04, a0o05, a0o06, a0o10, cu04 as contect from acc0o0, customer where a0o02 = '2' and substr(a0o03, 1, 8) = cu01 and substr(a0o03, 9, 1) = decode(a0o02, '2', cu02) and " & strCondition & " >= " & Val(Combo2) & " and " & strCondition & " <= " & Val(Combo3) & ""
               strUnion = strUnion & " union select a0o01, a0o04, a0o05, a0o06, a0o10, st02 as contect from acc0o0, staff where a0o02 = '3' and a0o03 = st01 and " & strCondition & " >= " & Val(Combo2) & " and " & strCondition & " <= " & Val(Combo3) & ""
               strUnion = strUnion & " union select a0o01, a0o04, a0o05, a0o06, a0o10, '' as contect from acc0o0 where a0o02 = '4' and " & strCondition & " >= " & Val(Combo2) & " and " & strCondition & " <= " & Val(Combo3) & " order by contect asc"
               adoadodc1.Open strUnion, adoTaie, adOpenStatic, adLockReadOnly
            Else
               strUnion = "select a0o01, a0o04, a0o05, a0o06, a0o10, a0i02 as contect from (select * from acc0o0 where " & strCondition & " >= '" & Combo2 & "' and " & strCondition & " <= '" & Combo3 & "') new, acc0i0 where a0o02 = '1' and a0o03 = a0i01"
               strUnion = strUnion & " union select a0o01, a0o04, a0o05, a0o06, a0o10, cu04 as contect from (select * from acc0o0 where " & strCondition & " >= '" & Combo2 & "' and " & strCondition & " <= '" & Combo3 & "') new, (select cu01, cu02, cu04 from customer) cust where a0o02 = '2' and substr(a0o03, 1, 8) = cu01 and substr(a0o03, 9, 1) = decode(a0o02, '2', cu02)"
               strUnion = strUnion & " union select a0o01, a0o04, a0o05, a0o06, a0o10, st02 as contect from (select * from acc0o0 where " & strCondition & " >= '" & Combo2 & "' and " & strCondition & " <= '" & Combo3 & "') new, staff where a0o02 = '3' and a0o03 = st01"
               strUnion = strUnion & " union select a0o01, a0o04, a0o05, a0o06, a0o10, '' as contect from acc0o0 where a0o02 = '4' and " & strCondition & " >= '" & Combo2 & "' and " & strCondition & " <= '" & Combo3 & "' order by contect asc"
               adoadodc1.Open strUnion, adoTaie, adOpenStatic, adLockReadOnly
            End If
         End If
      End If
   End If
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
   adoadodc1.Open "select * from acc0o0 where a0o01 = '" & Combo2 & "' order by a0o01 asc", adoTaie, adOpenStatic, adLockReadOnly
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub
