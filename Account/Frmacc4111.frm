VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc4111 
   AutoRedraw      =   -1  'True
   Caption         =   "會計科目資料查詢"
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
      Height          =   330
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2052
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc4111.frx":0000
      Height          =   4092
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   8292
      _ExtentX        =   14631
      _ExtentY        =   7223
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
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
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "a0101"
         Caption         =   "科目編號"
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
      BeginProperty Column02 
         DataField       =   "a0109"
         Caption         =   "專用公司別"
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
         DataField       =   "a0103"
         Caption         =   "借/貸方"
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
         DataField       =   "a0104"
         Caption         =   "層級"
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
         DataField       =   "a0702"
         Caption         =   "分攤類別"
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
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3825.071
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column05 
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
      Height          =   330
      Left            =   5760
      TabIndex        =   4
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
      Height          =   330
      Left            =   2640
      TabIndex        =   2
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
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   4800
      Visible         =   0   'False
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
      TabIndex        =   1
      Top             =   240
      Width           =   132
   End
End
Attribute VB_Name = "Frmacc4111"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/08/23 Form2.0已修改 DataGrid1
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/3 日期欄已修改
Option Explicit
Public adoacc010 As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset

Private Sub Combo1_Change()
'   SelectScope
End Sub

Private Sub Combo1_Click()
'   SelectScope
End Sub

'Modify by Amy 2021/08/23 改Form2.0 原:Integer
Private Sub Combo2_KeyPress(KeyAscii As MSForms.ReturnInteger)
    'Add by Amy 2013/12/12
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Combo2_Validate(Cancel As Boolean)
   Combo3 = Combo2
End Sub

'Modify by Amy 2021/08/23 改Form2.0 原:Integer
Private Sub Combo3_KeyPress(KeyAscii As MSForms.ReturnInteger)
    'Add by Amy 2013/12/12
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
   strCon1 = "會計科目編號"
   strCon2 = "會計科目名稱"
   strCon3 = "借/貸方別"
   strCon4 = "會計科目層級"
   strCon5 = "分攤類別"
   strCon6 = "專用公司別" 'Add by Amy 2013/12/12
   Combo1.AddItem MsgText(31)
   Combo1.AddItem strCon1
   Combo1.AddItem strCon2
   Combo1.AddItem strCon3
   Combo1.AddItem strCon4
   Combo1.AddItem strCon5
   Combo1.AddItem strCon6 'Add by Amy 2013/12/12
   Combo1 = MsgText(31)
'   SelectScope
   OpenTable
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Adodc1.Recordset.RecordCount <> 0 Then
      strItemNo = Adodc1.Recordset.Fields("a0101").Value
   Else
      strItemNo = MsgText(601)
   End If
   StatusClear
   tool1_enabled
   Frmacc4110.Enabled = True
   Frmacc4110.Show
   Set Frmacc4111 = Nothing
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
         strCondition = "a0101"
      Case strCon2
         strCondition = "a0102"
      Case strCon3
         strCondition = "a0103"
      Case strCon4
         strCondition = "a0104"
      Case strCon5
         strCondition = "a0105"
      Case Else
         Exit Sub
   End Select
   Combo2.Clear
   Combo3.Clear
   adoacc010.CursorLocation = adUseClient
   adoacc010.Open "select distinct " & strCondition & " from acc010 order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
   Do While adoacc010.EOF = False
      If IsNull(adoacc010.Fields(0).Value) = False Then
         Combo2.AddItem adoacc010.Fields(0).Value
         Combo3.AddItem adoacc010.Fields(0).Value
      End If
      adoacc010.MoveNext
   Loop
   adoacc010.Close
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)
   Select Case KeyCode
      Case vbKeyF12
         Screen.MousePointer = vbHourglass
         Acc010Query
         Screen.MousePointer = vbDefault
   End Select
   KeyEnter KeyCode
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

'*************************************************
'  會計科目資料查詢
'
'*************************************************
Private Sub Acc010Query()
On Error GoTo Checking
   adoadodc1.Close
   adoadodc1.CursorLocation = adUseClient
   Select Case Combo1
      Case strCon1
         strCondition = "a0101"
      Case strCon2
         strCondition = "a0102"
      Case strCon3
         strCondition = "a0103"
      Case strCon4
         strCondition = "a0104"
      Case strCon5
         strCondition = "a0105"
     'Add by Amy 203/12/12 +專用公司別欄位
     Case strCon6
         strCondition = "a0109"
      Case MsgText(31)
         adoadodc1.Open "select * from acc010, acc070 where a0105 = a0701 (+) order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
         Adodc1.Recordset.Requery
         Exit Sub
      Case Else
         Exit Sub
   End Select
   If Combo3 = MsgText(601) Then
      adoadodc1.Open "select * from acc010, acc070 where a0105 = a0701 (+) and " & strCondition & " = '" & Combo2 & "' order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
   Else
      If Combo2 = MsgText(601) Then
         adoadodc1.Open "select * from acc010, acc070 where a0105 = a0701 (+) and " & strCondition & " <= '" & Combo3 & "' order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
      Else
         adoadodc1.Open "select * from acc010, acc070 where a0105 = a0701 (+) and " & strCondition & " >= '" & Combo2 & "' and " & strCondition & " <= '" & Combo3 & "' order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
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
   adoadodc1.Open "select * from acc010 where a0101 = '" & Combo2 & "' order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub
