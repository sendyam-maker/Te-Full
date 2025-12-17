VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Frmacc4161 
   AutoRedraw      =   -1  'True
   Caption         =   "預算科目查詢"
   ClientHeight    =   4995
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8730
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4995
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
      Left            =   5760
      TabIndex        =   2
      Top             =   240
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
      Left            =   2640
      TabIndex        =   1
      Top             =   240
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
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2052
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc4161.frx":0000
      Height          =   4095
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   7223
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   20
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
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "a0401"
         Caption         =   "年度"
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
         DataField       =   "a0402"
         Caption         =   "月份"
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
         DataField       =   "a0403"
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
      BeginProperty Column03 
         DataField       =   "a0902"
         Caption         =   "部門別"
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
         DataField       =   "a0405"
         Caption         =   "科目代號"
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
         DataField       =   "a0102"
         Caption         =   "科目名稱"
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
            ColumnWidth     =   645.165
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   555.024
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   810.142
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   870.236
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1379.906
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   3795.024
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
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
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Top             =   4800
      Visible         =   0   'False
      Width           =   135
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
      Left            =   5520
      TabIndex        =   5
      Top             =   240
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
      Left            =   2400
      TabIndex        =   4
      Top             =   240
      Width           =   135
   End
End
Attribute VB_Name = "Frmacc4161"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/08/23 Form2.0已修改 DataGrid1
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/3 日期欄已修改
Option Explicit
Public adoacc040 As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset

Private Sub Combo1_Change()
'   SelectScope
End Sub

Private Sub Combo1_Click()
'   SelectScope
End Sub

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
   strCon1 = "公司別"
   strCon2 = "部門別"
   strCon3 = "科目代號"
   Combo1.AddItem MsgText(31)
   Combo1.AddItem strCon1
   Combo1.AddItem strCon2
   Combo1.AddItem strCon3
   Combo1 = MsgText(31)
   SelectScope
   OpenTable
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Adodc1.Recordset.RecordCount <> 0 Then
      strCon1 = Adodc1.Recordset.Fields("a0401").Value
      strCon2 = Adodc1.Recordset.Fields("a0403").Value
      strCon3 = Adodc1.Recordset.Fields("a0404").Value
      strCon4 = Adodc1.Recordset.Fields("a0405").Value
   Else
      strCon1 = MsgText(601)
      strCon2 = MsgText(601)
      strCon3 = MsgText(601)
      strCon4 = MsgText(601)
   End If
   StatusClear
   tool1_enabled
   Select Case strFormLink
      Case "Frmacc4160"
         Frmacc4160.Enabled = True
         Frmacc4160.Show
      Case "Frmacc5200"
         Frmacc5200.Enabled = True
         Frmacc5200.Show
   End Select
   Set Frmacc4161 = Nothing
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
         strCondition = "a0403"
      Case strCon2
         strCondition = "a0404"
      Case strCon3
         strCondition = "a0405"
   End Select
   If strCondition = MsgText(601) Then
      Exit Sub
   End If
   Combo2.Clear
   Combo3.Clear
   adoacc040.CursorLocation = adUseClient
   adoacc040.Open "select distinct " & strCondition & " from acc040 order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
   Do While adoacc040.EOF = False
      If IsNull(adoacc040.Fields(0).Value) = False Then
         Combo2.AddItem adoacc040.Fields(0).Value
         Combo3.AddItem adoacc040.Fields(0).Value
      End If
      adoacc040.MoveNext
   Loop
   adoacc040.Close
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)
   Select Case KeyCode
      Case vbKeyF12
         Screen.MousePointer = vbHourglass
         Acc040Query
         Screen.MousePointer = vbDefault
   End Select
   KeyEnter KeyCode
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

'*************************************************
'  科目預算資料查詢
'
'*************************************************
Private Sub Acc040Query()
On Error GoTo Checking
   adoadodc1.Close
   adoadodc1.CursorLocation = adUseClient
   Select Case strAccount
      Case "2"
         Select Case Combo1
            Case strCon1
               strCondition = "a0503"
            Case strCon2
               strCondition = "a0504"
            Case strCon3
               strCondition = "a0505"
            Case MsgText(31)
               adoadodc1.Open "select a0501 as a0401, A0502 as a0402, a0503 as a0403, a0504 as a0404, a0505 as a0405, a0102, a0902 from acc050, acc010, acc090 where a0505 = a0101 and a0504 = a0901 group by a0501, A0502, a0503, a0504, a0505, a0102, a0902", adoTaie, adOpenStatic, adLockReadOnly
               Adodc1.Recordset.Requery
               Exit Sub
            Case Else
               Exit Sub
         End Select
         If Combo3 = MsgText(601) Then
            adoadodc1.Open "select a0501 as a0401, a0502 as a0402, a0503 as a0403, a0504 as a0404, a0505 as a0405, a0506 as a0406, a0507 as a0407, a0508 as a0408, a0509 as a0409, a0511 as a0411, a0512 as a0412, a0513 as a0413, a0514 as a0414, a0515 as a0415, a0516 as a0416, a0102, a0902 from acc050, acc010, acc090 where a0505 = a0101 and a0504 = a0901 and " & strCondition & " = '" & Combo2 & "' order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
         Else
            If Combo2 = MsgText(601) Then
               adoadodc1.Open "select a0501 as a0401, a0502 as a0402, a0503 as a0403, a0504 as a0404, a0505 as a0405, a0506 as a0406, a0507 as a0407, a0508 as a0408, a0509 as a0409, a0511 as a0411, a0512 as a0412, a0513 as a0413, a0514 as a0414, a0515 as a0415, a0516 as a0416, a0102, a0902 from acc050, acc010, acc090 where a0505 = a0101 and a0504 = a0901 and " & strCondition & " <= '" & Combo3 & "' order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
            Else
               adoadodc1.Open "select a0501 as a0401, a0502 as a0402, a0503 as a0403, a0504 as a0404, a0505 as a0405, a0506 as a0406, a0507 as a0407, a0508 as a0408, a0509 as a0409, a0511 as a0411, a0512 as a0412, a0513 as a0413, a0514 as a0414, a0515 as a0415, a0516 as a0416, a0102, a0902 from acc050, acc010, acc090 where a0505 = a0101 and a0504 = a0901 and " & strCondition & " >= '" & Combo2 & "' and " & strCondition & " <= '" & Combo3 & "' order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
            End If
         End If
      Case Else
         Select Case Combo1
            Case strCon1
               strCondition = "a0403"
            Case strCon2
               strCondition = "a0404"
            Case strCon3
               strCondition = "a0405"
            Case MsgText(31)
               adoadodc1.Open "select a0401, A0402, a0403, a0404, a0405, a0102, a0902 from acc040, acc010, acc090 where a0405 = a0101 and a0404 = a0901 group by a0401, A0402, a0403, a0404, a0405, a0102, a0902", adoTaie, adOpenStatic, adLockReadOnly
               Adodc1.Recordset.Requery
               Exit Sub
            Case Else
               Exit Sub
         End Select
         If Combo3 = MsgText(601) Then
            adoadodc1.Open "select * from acc040, acc010, acc090 where a0405 = a0101 and a0404 = a0901 and " & strCondition & " = '" & Combo2 & "' order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
         Else
            If Combo2 = MsgText(601) Then
               adoadodc1.Open "select * from acc040, acc010, acc090 where a0405 = a0101 and a0404 = a0901 and " & strCondition & " <= '" & Combo3 & "' order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
            Else
               adoadodc1.Open "select * from acc040, acc010, acc090 where a0405 = a0101 and a0404 = a0901 and " & strCondition & " >= '" & Combo2 & "' and " & strCondition & " <= '" & Combo3 & "' order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
            End If
         End If
   End Select
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
   Select Case strAccount
      Case "2"
         adoadodc1.Open "select a0501 as a0401, a0503 as a0403, a0504 as a0404, a0505 as a0405, a0102 from acc050, acc010 where a0505 = a0101 (+) and a0503 = '" & Combo1 & "' group by a0501, a0503, a0504, a0505, a0102", adoTaie, adOpenStatic, adLockReadOnly
      Case Else
         adoadodc1.Open "select a0401, a0403, a0404, a0405, a0102 from acc040, acc010 where a0405 = a0101 (+) and a0403 = '" & Combo1 & "' group by a0401, a0403, a0404, a0405, a0102", adoTaie, adOpenStatic, adLockReadOnly
   End Select
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

