VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc4171 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "每月固定傳票資料查詢"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8730
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   8730
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
      Bindings        =   "Frmacc4171.frx":0000
      Height          =   4095
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   7223
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   17
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
         Name            =   "新細明體-ExtB"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "axd01"
         Caption         =   "公司別"
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
         DataField       =   "axd02"
         Caption         =   "流水號"
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
         DataField       =   "axd03"
         Caption         =   "每月傳票日"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "###/##/##"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "axd04"
         Caption         =   "上次處理日期"
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
         DataField       =   "a0102"
         Caption         =   "會計科目"
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
         DataField       =   "a0d06"
         Caption         =   "借方金額"
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
      BeginProperty Column06 
         DataField       =   "a0d07"
         Caption         =   "貸方金額"
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
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         Size            =   275
         BeginProperty Column00 
            Alignment       =   2
            ColumnWidth     =   780.095
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   810.142
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   1440
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1395.213
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   1035.213
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            ColumnWidth     =   1035.213
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
   Begin MSForms.ComboBox Combo2 
      Height          =   350
      Left            =   2700
      TabIndex        =   1
      Top             =   240
      Width           =   2775
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "4895;617"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo3 
      Height          =   350
      Left            =   5730
      TabIndex        =   2
      Top             =   240
      Width           =   2775
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "4895;617"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
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
Attribute VB_Name = "Frmacc4171"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/1 改成Form2.0 ; Combo2、Combo3、DataGrid1改字型=新細明體-ExtB
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/3 日期欄已修改
Option Explicit
Public adoacc0d1 As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset

Private Sub Combo1_Change()
'   SelectScope
End Sub

Private Sub Combo1_Click()
'   SelectScope
End Sub

'Modified by Lydia 2021/12/01 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub Combo2_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Combo2_Validate(Cancel As Boolean)
   Combo3 = Combo2
End Sub

'Modified by Lydia 2021/12/01 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
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
   strCon1 = "公司別"
   strCon2 = "上次處理日期"
   strCon3 = "會計科目"
   strCon4 = "部門別"
   Combo1.AddItem MsgText(31)
   Combo1.AddItem strCon1
   Combo1.AddItem strCon2
   Combo1.AddItem strCon3
   Combo1.AddItem strCon4
   Combo1 = MsgText(31)
   SelectScope
   OpenTable
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Adodc1.Recordset.RecordCount <> 0 Then
      strCompanyNo = Adodc1.Recordset.Fields("axd01").Value
      strItemNo = Adodc1.Recordset.Fields("axd02").Value
   Else
      strCompanyNo = MsgText(601)
      strItemNo = MsgText(601)
   End If
   StatusClear
   tool1_enabled
   Frmacc4170.Enabled = True
   Frmacc4170.Show
   Set Frmacc4171 = Nothing
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
         strCondition = "axd01"
      Case strCon2
         strCondition = "axd04"
      Case strCon3
         strCondition = "a0d05"
      Case strCon4
         strCondition = "a0d08"
   End Select
   If strCondition = MsgText(601) Then
      Exit Sub
   End If
   Combo2.Clear
   Combo3.Clear
   adoacc0d1.CursorLocation = adUseClient
   adoacc0d1.Open "select distinct " & strCondition & " from acc0d0, acc0d1 where a0d01 = axd01 and a0d02 = axd02 order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
   Do While adoacc0d1.EOF = False
      If IsNull(adoacc0d1.Fields(0).Value) = False Then
         Combo2.AddItem adoacc0d1.Fields(0).Value
         Combo3.AddItem adoacc0d1.Fields(0).Value
      End If
      adoacc0d1.MoveNext
   Loop
   adoacc0d1.Close
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)
   Select Case KeyCode
      Case vbKeyF12
         Screen.MousePointer = vbHourglass
         Acc0d1Query
         Screen.MousePointer = vbDefault
   End Select
   KeyEnter KeyCode
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

'*************************************************
'  傳票資料查詢
'
'*************************************************
Private Sub Acc0d1Query()
On Error GoTo Checking
   adoadodc1.Close
   adoadodc1.CursorLocation = adUseClient
   Select Case Combo1
      Case strCon1
         strCondition = "axd01"
      Case strCon2
         strCondition = "axd04"
      Case strCon3
         strCondition = "a0d05"
      Case strCon4
         strCondition = "a0d08"
      Case MsgText(31)
         adoadodc1.Open "select * from acc0d0, acc0d1, acc010 where a0d01 = axd01 and a0d02 = axd02 and a0d05 = a0101 order by a0d01 asc, a0d02 asc", adoTaie, adOpenStatic, adLockReadOnly
         Adodc1.Recordset.Requery
         Exit Sub
      Case Else
         Exit Sub
   End Select
   If Combo3 = MsgText(601) Then
      If Combo1 = strCon3 Then
         adoadodc1.Open "select * from acc0d0, acc0d1, acc010 where a0d01 = axd01 and a0d02 = axd02 and a0d05 = a0101 and " & strCondition & " = " & Val(Combo2) & " order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
      Else
         adoadodc1.Open "select * from acc0d0, acc0d1, acc010 where a0d01 = axd01 and a0d02 = axd02 and a0d05 = a0101 and " & strCondition & " = '" & Combo2 & "' order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
      End If
   Else
      If Combo2 = MsgText(601) Then
         If Combo1 = strCon3 Then
            adoadodc1.Open "select * from acc0d0, acc0d1, acc010 where a0d01 = axd01 and a0d02 = axd02 and a0d05 = a0101 and " & strCondition & " <= " & Val(Combo3) & " order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
         Else
            adoadodc1.Open "select * from acc0d0, acc0d1, acc010 where a0d01 = axd01 and a0d02 = axd02 and a0d05 = a0101 and " & strCondition & " <= '" & Combo3 & "' order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
         End If
      Else
         If Combo1 = strCon3 Then
            adoadodc1.Open "select * from acc0d0, acc0d1, acc010 where a0d01 = axd01 and a0d02 = axd02 and a0d05 = a0101 and " & strCondition & " >= " & Val(Combo2) & " and " & strCondition & " <= " & Val(Combo3) & " order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
         Else
            adoadodc1.Open "select * from acc0d0, acc0d1, acc010 where a0d01 = axd01 and a0d02 = axd02 and a0d05 = a0101 and " & strCondition & " >= '" & Combo2 & "' and " & strCondition & " <= '" & Combo3 & "' order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
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
   adoadodc1.Open "select * from acc0d0, acc0d1 where a0d01 = axd01 and a0d02 = axd02 and a0d01 = '" & Combo2 & "' order by a0d01 asc, a0d02 asc", adoTaie, adOpenStatic, adLockReadOnly
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub


