VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc4110 
   AutoRedraw      =   -1  'True
   Caption         =   "會計科目基本資料"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5115
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
      Left            =   6840
      TabIndex        =   14
      Top             =   960
      Width           =   1572
   End
   Begin VB.CommandButton Command1 
      Height          =   300
      Left            =   3000
      Picture         =   "Frmacc4110.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   240
      Width           =   350
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc4110.frx":0102
      Height          =   3400
      Left            =   240
      TabIndex        =   12
      Top             =   1560
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   5980
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   18
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
         Size            =   11.25
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
         Caption         =   "專用"
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
            ColumnWidth     =   555.024
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   555.024
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   2025.071
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   312
      Left            =   240
      Top             =   1440
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
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
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
      Left            =   2040
      TabIndex        =   11
      Top             =   960
      Width           =   3370
   End
   Begin VB.TextBox Text8 
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
      Left            =   1440
      MaxLength       =   2
      TabIndex        =   5
      Top             =   960
      Width           =   612
   End
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
      Height          =   300
      Left            =   6840
      TabIndex        =   4
      Top             =   600
      Width           =   1572
   End
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
      Height          =   300
      Left            =   6840
      MaxLength       =   1
      TabIndex        =   2
      Top             =   240
      Width           =   612
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
      Height          =   300
      Left            =   1440
      MaxLength       =   6
      TabIndex        =   0
      Top             =   240
      Width           =   1572
   End
   Begin MSForms.TextBox Text3 
      Height          =   300
      Left            =   1440
      TabIndex        =   3
      Top             =   600
      Width           =   3975
      VariousPropertyBits=   679493659
      BackColor       =   16777215
      MaxLength       =   35
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "專用公司別"
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
      Left            =   5520
      TabIndex        =   13
      Top             =   960
      Width           =   1200
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "分攤類別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   360
      TabIndex        =   10
      Top             =   960
      Width           =   972
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Top             =   4920
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   1332
      Left            =   240
      Top             =   120
      Width           =   8292
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "借/貸方別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   5640
      TabIndex        =   9
      Top             =   600
      Width           =   1092
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "科目名稱"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   360
      TabIndex        =   8
      Top             =   600
      Width           =   1092
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "科目層級"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   5640
      TabIndex        =   7
      Top             =   240
      Width           =   972
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "科目編號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   360
      TabIndex        =   6
      Top             =   240
      Width           =   972
   End
End
Attribute VB_Name = "Frmacc4110"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/08/23 Form2.0已修改 Text3/DataGrid1
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/3 日期欄已修改
Option Explicit
Public adoacc010 As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

'Add by Amy 2013/12/12
Private Sub Combo1_Validate(Cancel As Boolean)
    Dim stTmp As String 'Add by Amy 2020/03/30
    
    'Modify by Amy 2020/03/30
    If Combo1 <> MsgText(601) Then
        stTmp = Combo1
        If InStr(Combo1, "--") > 0 Then
            stTmp = Mid(Combo1, 1, Val(InStr(Combo1, "--")) - 1)
        End If
        'If Left(Combo1, 1) <> Left(ComboItem(254), 1) And Left(Combo1, 1) <> Left(ComboItem(255), 1) Then
        If InStr(GetBookKeepCmp, stTmp) = 0 Then
            MsgBox Label3 & MsgText(63), , MsgText(5)
            Cancel = True
            Combo1.SetFocus
            Exit Sub
        End If
    End If
End Sub
'end 2013/12/12

Private Sub Combo3_Validate(Cancel As Boolean)
   If Mid(Combo3, 1, 1) <> Mid(ComboItem(1), 1, 1) And Mid(Combo3, 1, 1) <> Mid(ComboItem(2), 1, 1) Then
      MsgBox MsgText(44), , MsgText(5)
      Cancel = True
      Exit Sub
   End If
End Sub

Private Sub Command1_Click()
   If Adodc1.Recordset.RecordCount = 0 Or Text1 = MsgText(601) Then
      Exit Sub
   End If
   Adodc1.Recordset.Find "a0101 = '" & Text1 & "'", 0, adSearchForward, 1
   If Adodc1.Recordset.EOF = False Then
      FormShow
      RecordShow
   Else
      MsgBox MsgText(33), , MsgText(5)
      Adodc1.Recordset.MoveFirst
   End If
End Sub

Private Sub DataGrid1_SelChange(Cancel As Integer)
   FormShow
   RecordShow
End Sub

Private Sub Form_Activate()
On Error GoTo Checking
   strFormName = Name
   If strItemNo = MsgText(601) Then
      Exit Sub
   End If
   Adodc1.Recordset.Find "a0101 = '" & strItemNo & "'", 0, adSearchForward, 1
   If Adodc1.Recordset.EOF = False Then
      FormShow
      RecordShow
   End If
   strItemNo = MsgText(601)
Checking:
   Exit Sub
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 8850
   Me.Height = 5500
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath1)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   'Add By Amy 2013/12/12 +專用公司別
   Combo1.AddItem MsgText(601), 0
   'Modify by Amy 2020/03/27
   If strSrvDate(1) >= 智慧所更名日 Then
        Call Pub_SetCboCmp(Combo1, False, True, False, , 1, "--")
   Else
        Combo1.AddItem ComboItem(254)
        Combo1.AddItem ComboItem(255)
   End If
   'end 2020/03/27
   Combo1 = MsgText(601)
   'end 2013/12/12
   Combo3.AddItem ComboItem(1)
   Combo3.AddItem ComboItem(2)
   Combo3 = MsgText(601)
   OpenTable
   If Adodc1.Recordset.RecordCount <> 0 Then
      Adodc1.Recordset.MoveLast
      Adodc1.Recordset.MoveFirst
      RecordShow
   End If
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
   Set Frmacc4110 = Nothing
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoadodc1.CursorLocation = adUseClient
'   adoadodc1.MaxRecords = 10
   adoadodc1.Open "select * from acc010, acc070 where a0105 = a0701 (+) order by a0101 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
'   adoadodc1.Open "select * from (select a0101, a0102, a0103, a0104, decode(a0105, null, '" & MsgText(12) & "', a0105, a0105) as a0105 from acc010), (select decode(a0701, null, '" & MsgText(12) & "', a0701, a0701) as a0701, decode(a0702, null, '" & MsgText(601) & "', a0702, a0702) as a0702 from (select a0701, a0702 from acc070 union select a0105, '" & MsgText(601) & "' from acc010) where (a0701 is not null and a0702 is not null) or (a0701 is null and a0702 is null)) where a0105 = a0701 order by a0101 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  顯示欄位資料(會計科目資料表)
'
'*************************************************
Public Sub FormShow()
   Dim stCmpNo As String 'Add by Amy 2020/03/27
    
   If IsNull(Adodc1.Recordset.Fields("a0101").Value) Then
      Text1 = MsgText(601)
   Else
      Text1 = Adodc1.Recordset.Fields("a0101").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a0104").Value) Then
      Text2 = MsgText(601)
   Else
      Text2 = Adodc1.Recordset.Fields("a0104").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a0102").Value) Then
      Text3 = MsgText(601)
   Else
      Text3 = Adodc1.Recordset.Fields("a0102").Value
   End If
   'Add by Amy  2013/12/12 +專用公司別
    If IsNull(Adodc1.Recordset.Fields("a0109").Value) Then
      Combo1 = MsgText(601)
   Else
      'Modify by Amy 2020/03/27
      If strSrvDate(1) >= 智慧所更名日 Then
        Call SetCompVal(Adodc1.Recordset.Fields("a0109").Value)
      Else
        Combo1 = Combo1.List(IIf(Val(Adodc1.Recordset.Fields("a0109").Value) = 1, 1, 2))
      End If
   End If
   'end 2013/12/12
   If IsNull(Adodc1.Recordset.Fields("a0103").Value) Then
      Combo3 = MsgText(601)
   Else
      Combo3 = Combo3.List(Val(Adodc1.Recordset.Fields("a0103").Value) - 1)
   End If
   If IsNull(Adodc1.Recordset.Fields("a0105").Value) Then
      Text8 = MsgText(601)
   Else
      If Adodc1.Recordset.Fields("a0105").Value = MsgText(12) Then
         Text8 = MsgText(601)
      Else
         Text8 = Adodc1.Recordset.Fields("a0105").Value
      End If
   End If
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   If strSaveConfirm <> MsgText(3) And strSaveConfirm <> MsgText(4) Then
      Exit Sub
   End If
   If Text1 = MsgText(601) Then
      MsgBox Label9 & MsgText(52), , MsgText(5)
      Cancel = True
      Exit Sub
   End If
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
   If Val(Text2) < 1 Or Val(Text2) > 4 Then
      MsgBox Label1 & MsgText(53), , MsgText(5)
      Cancel = True
      Exit Sub
   End If
End Sub

Private Sub Text3_GotFocus()
   StatusView MsgText(65) & "30"
   TextInverse Text3
   'add by nickc 2007/07/13 將輸入法改成使用API
   OpenIme
End Sub

Private Sub Text3_LostFocus()
   StatusView MsgText(601)
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
   If Text3 = MsgText(601) Then
      MsgBox Label2 & MsgText(52), , MsgText(5)
      Cancel = True
      Exit Sub
   End If
   If CheckLen(Label2, Text3, 30) = MsgText(603) Then
      Cancel = True
      Exit Sub
   End If
   'add by nickc 2007/07/13 將輸入法改成使用API
   If Cancel = False Then CloseIme
End Sub

Private Sub Text8_Change()
   Text4 = Left(A0702Query(Text8), 10) 'Modify by Amy 2013/12/12 顯示10個字
End Sub

Private Sub Text8_GotFocus()
   TextInverse Text8
End Sub

'*************************************************
'  重新整理 Adodc 之資料
'
'*************************************************
Public Sub AdodcRefresh()
On Error GoTo Checking
   adoadodc1.Close
   adoadodc1.CursorLocation = adUseClient
'   adoadodc1.MaxRecords = 10
   adoadodc1.Open "select * from acc010, acc070 where a0105 = a0701 (+) order by a0101 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
'   adoadodc1.Open "select * from (select a0101, a0102, a0103, a0104, decode(a0105, null, '" & MsgText(12) & "', a0105, a0105) as a0105 from acc010), (select decode(a0701, null, '" & MsgText(12) & "', a0701, a0701) as a0701, decode(a0702, null, '" & MsgText(601) & "', a0702, a0702) as a0702 from (select a0701, a0702 from acc070 union select a0105, '" & MsgText(601) & "' from acc010) where (a0701 is not null and a0702 is not null) or (a0701 is null and a0702 is null)) where a0105 = a0701 order by a0101 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Adodc1.Recordset.Requery
   If Adodc1.Recordset.RecordCount <> 0 Then
      If Text1 <> MsgText(601) Then
         Adodc1.Recordset.Find "a0101 = '" & Text1 & "'", 0, adSearchForward, 1
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

Private Sub Text8_Validate(Cancel As Boolean)
   If Text8 = MsgText(601) Then
      Exit Sub
   End If
   If ExistCheck("acc070", "a0701", Text8, Label10) = False Then
      Cancel = True
      Exit Sub
   End If
End Sub

'Add by Amy 由.bas搬回
Public Sub Frmacc4110_Delete()
On Error GoTo Checking
      If DeleteCheck("select a0101 from acc010 where a0101 = '" & Text1 & "'") = MsgText(603) Then
         Exit Sub
      End If
      adoTaie.Execute "delete from acc010 where a0101 = '" & Text1 & "'"
      AdodcRefresh
      If Adodc1.Recordset.RecordCount <> 0 Then
         Adodc1.Recordset.MoveFirst
         RecordShow
      Else
         StatusClear
      End If
   
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

Public Sub Frmacc4110_Clear()
      Text1 = ""
      Text2 = ""
      Text3 = ""
      Combo1 = "" 'Add by Amy 2013/12/12
      Combo3 = ""
      Text8 = ""
      Text4 = ""
      Text1.SetFocus
End Sub

Public Sub Frmacc4110_First()
      If Adodc1.Recordset.RecordCount <> 0 Then
         Adodc1.Recordset.MoveFirst
         Frmacc4110_Clear
         FormShow
         RecordShow
      End If
End Sub

Public Sub Frmacc4110_Previous()
      If Adodc1.Recordset.BOF = False Then
         Adodc1.Recordset.MovePrevious
         If Adodc1.Recordset.BOF Then
            Adodc1.Recordset.MoveFirst
            MsgBox MsgText(7), , MsgText(5)
         End If
         Frmacc4110_Clear
         FormShow
         RecordShow
      End If

End Sub

Public Sub Frmacc4110_Next()
      If Adodc1.Recordset.EOF = False Then
         Adodc1.Recordset.MoveNext
         If Adodc1.Recordset.EOF Then
            Adodc1.Recordset.MoveLast
            MsgBox MsgText(8), , MsgText(5)
         End If
         Frmacc4110_Clear
         FormShow
         RecordShow
      End If
End Sub

Public Sub Frmacc4110_Last()
      If Adodc1.Recordset.RecordCount <> 0 Then
         Adodc1.Recordset.MoveLast
         Frmacc4110_Clear
         FormShow
         RecordShow
      End If
End Sub

Public Sub Frmacc4110_Save()
Dim adoacc070 As New ADODB.Recordset
Dim strSql As String
Dim bolCancel As Boolean 'Add by Amy 2020/03/30

   On Error GoTo Checking
      If strSaveConfirm = MsgText(4) Then
         strSql = "select a0107, a0108 from acc010 where a0101 = '" & Text1 & "'"
         If CheckRecord(strSql, IIf(IsNull(adoadodc1.Fields("a0107").Value), 0, adoadodc1.Fields("a0107").Value), IIf(IsNull(adoadodc1.Fields("a0108").Value), 0, adoadodc1.Fields("a0108").Value)) = False Then
            strControlButton = MsgText(602)
            Text1.SetFocus
            Exit Sub
         End If
      End If
      If Text1 = MsgText(601) Then
         MsgBox MsgText(10) & Label9, , MsgText(5)
         strControlButton = MsgText(602)
         Text1.SetFocus
         Exit Sub
      Else
         If Val(Text2) < 1 Or Val(Text2) > 4 Then
            MsgBox Label1 & MsgText(53), , MsgText(5)
            strControlButton = MsgText(602)
            Text2.SetFocus
            Exit Sub
         End If
         If Text3 = MsgText(601) Then
            MsgBox Label2 & MsgText(52), , MsgText(5)
            strControlButton = MsgText(602)
            Text3.SetFocus
            Exit Sub
         End If
         'Add by Amy 2013/12/12
         If Combo1 <> MsgText(601) Then
            'Modify by Amy 2022/03/30 改抓Combo1_Validate
'            If Left(Combo1, 1) <> Left(ComboItem(254), 1) And Left(Combo1, 1) <> Left(ComboItem(255), 1) Then
'                MsgBox MsgText(63), , MsgText(5)
'                strControlButton = MsgText(602)
'                Combo1.SetFocus
'                Exit Sub
'            End If
            Call Combo1_Validate(bolCancel)
            If bolCancel = True Then
                strControlButton = MsgText(602)
                Exit Sub
            End If
         End If
         'end 2013/12/12
         If Mid(Combo3, 1, 1) <> Mid(ComboItem(1), 1, 1) And Mid(Combo3, 1, 1) <> Mid(ComboItem(2), 1, 1) Then
            MsgBox MsgText(44), , MsgText(5)
            strControlButton = MsgText(602)
            Combo3.SetFocus
            Exit Sub
         End If
         If Text8 <> MsgText(601) Then
            If ExistCheck("acc070", "a0701", Text8, Label10) = False Then
               strControlButton = MsgText(602)
               Text8.SetFocus
               Exit Sub
            End If
         End If
         If CheckLen(Label2, Text3, 30) = MsgText(603) Then
            strControlButton = MsgText(602)
            Text3.SetFocus
            Exit Sub
         End If
      End If
      'Add by Amy 2021/08/23 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
      If PUB_ChkUniText(Me) = False Then
            strControlButton = MsgText(602)
            Exit Sub
      End If
      
      adoacc010.CursorLocation = adUseClient
      adoacc010.Open "select * from acc010 where a0101 = '" & Text1 & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
      If strSaveConfirm = MsgText(3) Then
         If adoacc010.RecordCount <> 0 Then
            MsgBox MsgText(9), , MsgText(5)
            strControlButton = MsgText(602)
            adoacc010.Close
            Text1.SetFocus
            Exit Sub
         End If
         adoacc010.AddNew
      Else
         If strSaveConfirm = MsgText(4) Then
            If adoacc010.RecordCount = 0 Then
               MsgBox MsgText(28), , MsgText(5)
               strControlButton = MsgText(602)
               adoacc010.Close
               Text1.SetFocus
               Exit Sub
            End If
         End If
      End If
      adoacc010.Fields("a0101").Value = Text1
      If Text2 <> MsgText(601) Then
         adoacc010.Fields("a0104").Value = Text2
      Else
         adoacc010.Fields("a0104").Value = Null
      End If
      If Text3 <> MsgText(601) Then
         adoacc010.Fields("a0102").Value = Text3
      Else
         adoacc010.Fields("a0102").Value = Null
      End If
      'Add by Amy 2013/12/12
      If Combo1 <> MsgText(601) Then
         adoacc010.Fields("a0109").Value = Mid(Combo1, 1, 1)
      Else
         adoacc010.Fields("a0109").Value = Null
      End If
      'end 2013/12/12
      If Combo3 <> MsgText(601) Then
         adoacc010.Fields("a0103").Value = Mid(Combo3, 1, 1)
      Else
         adoacc010.Fields("a0103").Value = Null
      End If
      If Text8 <> MsgText(601) Then
         adoacc010.Fields("a0105").Value = Text8
      Else
         adoacc010.Fields("a0105").Value = Null
      End If
      If strSaveConfirm = MsgText(4) Then
         adoacc010.Fields("a0106").Value = strUserNum
         adoacc010.Fields("a0107").Value = Val(strSrvDate(2))
         adoacc010.Fields("a0108").Value = ServerTime
      End If
      adoacc010.UpdateBatch
      adoacc010.Close
      AdodcRefresh
      RecordShow
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
 End Sub
 
 Public Function FormCheck() As Boolean
    FormCheck = False
    If Text1 = MsgText(601) Then
         MsgBox "請選擇欲刪除的資料", , MsgText(5)
         strSaveConfirm = MsgText(601)
         Text1.SetFocus
         Exit Function
      End If
      If ExistCheck("acc021", "ax205", Text1, Label9, False) = True Then
         MsgBox "傳票檔有該科目代號的資料,不可刪除", , MsgText(5)
         strSaveConfirm = MsgText(601)
         Text1.SetFocus
         Exit Function
      End If
    FormCheck = True
 End Function
'end 2015/06/11

Private Sub SetCompVal(ByVal stNo As String)
    Dim i As Integer
    Dim stOrgName As String, stName As String
    
    For i = 1 To Combo1.ListCount - 1
        stOrgName = Combo1.List(i)
        stName = Mid(stOrgName, 1, Val(InStr(stOrgName, "--")) - 1)
        If stName = stNo Then
            Combo1 = stOrgName
            Exit For
        End If
    Next i
End Sub
