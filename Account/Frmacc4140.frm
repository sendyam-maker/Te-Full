VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Frmacc4140 
   AutoRedraw      =   -1  'True
   Caption         =   "部門基本資料"
   ClientHeight    =   5120
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   8760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5120
   ScaleWidth      =   8760
   Begin VB.TextBox Text7 
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
      Left            =   2520
      MaxLength       =   3
      TabIndex        =   7
      Top             =   1950
      Width           =   1140
   End
   Begin VB.TextBox Text6 
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
      Left            =   2520
      MaxLength       =   6
      TabIndex        =   6
      Top             =   1620
      Width           =   1140
   End
   Begin VB.TextBox Text5 
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
      Left            =   2520
      MaxLength       =   6
      TabIndex        =   5
      Top             =   1290
      Width           =   1140
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
      Left            =   630
      Top             =   2790
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
      Left            =   7680
      MaxLength       =   1
      TabIndex        =   2
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Height          =   300
      Left            =   2520
      Picture         =   "Frmacc4140.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   240
      Width           =   350
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc4140.frx":0102
      Height          =   2505
      Left            =   240
      TabIndex        =   8
      Top             =   2430
      Width           =   8295
      _ExtentX        =   14623
      _ExtentY        =   4427
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
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "a0901"
         Caption         =   "部門代號"
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
         DataField       =   "a0902"
         Caption         =   "部門名稱(中)"
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
         DataField       =   "a0903"
         Caption         =   "部門名稱(英)"
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
            ColumnWidth     =   1049.953
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3360.189
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   3390.236
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text4 
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
      Left            =   1800
      MaxLength       =   50
      TabIndex        =   4
      Top             =   960
      Width           =   6612
   End
   Begin VB.TextBox Text3 
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
      Left            =   1800
      TabIndex        =   3
      Top             =   600
      Width           =   6612
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
      Left            =   1800
      MaxLength       =   3
      TabIndex        =   0
      Top             =   240
      Width           =   732
   End
   Begin VB.Label lblA0910 
      BackColor       =   &H8000000E&
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
      Left            =   3735
      TabIndex        =   18
      Top             =   1995
      Width           =   4650
   End
   Begin VB.Label lblA0909 
      BackColor       =   &H8000000E&
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
      Left            =   3735
      TabIndex        =   17
      Top             =   1665
      Width           =   2085
   End
   Begin VB.Label lblA0908 
      BackColor       =   &H8000000E&
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
      Left            =   3735
      TabIndex        =   16
      Top             =   1335
      Width           =   2085
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "做帳部門"
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
      TabIndex        =   15
      Top             =   1980
      Width           =   2085
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "部門主管(案件用)"
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
      TabIndex        =   14
      Top             =   1650
      Width           =   2085
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "部門主管(發EMail用)"
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
      Width           =   2085
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "是否為利潤中心部門(Y/N)"
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
      Left            =   4920
      TabIndex        =   12
      Top             =   240
      Width           =   2655
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
      Height          =   2205
      Left            =   240
      Top             =   120
      Width           =   8295
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "部門名稱(英)"
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
      TabIndex        =   11
      Top             =   960
      Width           =   1452
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "部門名稱(中)"
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
      Top             =   600
      Width           =   1452
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "部門別"
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
      TabIndex        =   9
      Top             =   240
      Width           =   852
   End
End
Attribute VB_Name = "Frmacc4140"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/08/23 Form2.0已修改 (無需修改)
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/3 日期欄已修改
Option Explicit
Public adoadodc1 As New ADODB.Recordset

Private Sub Command1_Click()
   If Adodc1.Recordset.RecordCount = 0 Or Text1 = MsgText(601) Then
      Exit Sub
   End If
   Adodc1.Recordset.Find "a0901 = '" & Text1 & "'", 0, adSearchForward, 1
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
   strFormName = Name
   If strItemNo = MsgText(601) Then
      Exit Sub
   End If
   Adodc1.Recordset.Find "a0901 = '" & strItemNo & "'", 0, adSearchForward, 1
   If Adodc1.Recordset.EOF = False Then
      FormShow
      RecordShow
   End If
   strItemNo = MsgText(601)
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
   OpenTable
   If Adodc1.Recordset.RecordCount <> 0 Then
      Adodc1.Recordset.MoveLast
      Adodc1.Recordset.MoveFirst
      RecordShow
   End If
   'Add by Morgan 2010/4/26
   lblA0908.BackStyle = 0
   lblA0909.BackStyle = 0
   lblA0910.BackStyle = 0
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
   Set Frmacc4140 = Nothing
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
   adoadodc1.Open "select * from acc090 order by a0901 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  顯示欄位資料(部門別資料)
'
'*************************************************
Public Sub FormShow()
   If IsNull(Adodc1.Recordset.Fields("a0901").Value) Then
      Text1 = MsgText(601)
   Else
      Text1 = Adodc1.Recordset.Fields("a0901").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a0904").Value) Then
      Text2 = MsgText(601)
   Else
      Text2 = Adodc1.Recordset.Fields("a0904").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a0902").Value) Then
      Text3 = MsgText(601)
   Else
      Text3 = Adodc1.Recordset.Fields("a0902").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a0903").Value) Then
      Text4 = MsgText(601)
   Else
      Text4 = Adodc1.Recordset.Fields("a0903").Value
   End If
   'Add by Morgan 2010/4/26
   Text5 = "" & Adodc1.Recordset.Fields("a0908")
   Text5_Validate False
   Text6 = "" & Adodc1.Recordset.Fields("a0909")
   Text6_Validate False
   Text7 = "" & Adodc1.Recordset.Fields("a0910")
   Text7_Validate False
   
   'add by sonia 2022/4/15
   Text5.Locked = True: Text6.Locked = True
   'end 2022/4/15
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   If strSaveConfirm <> MsgText(3) And strSaveConfirm <> MsgText(4) Then
      Exit Sub
   End If
   If Text1 = MsgText(601) Then
      MsgBox Label3 & MsgText(52), , MsgText(5)
      Cancel = True
      Exit Sub
   End If
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
   If Text2 <> MsgText(602) And Text2 <> MsgText(603) Then
      MsgBox Label2 & MsgText(54), , MsgText(5)
      Cancel = True
      Exit Sub
   End If
End Sub

Private Sub Text3_GotFocus()
   StatusView MsgText(65) & "20"
   TextInverse Text3
   'add by nickc 2007/07/13 將輸入法改成使用API
   OpenIme
End Sub

Private Sub Text3_LostFocus()
   StatusView MsgText(601)
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
   If CheckLen(Label1, Text3, 20) = MsgText(603) Then
      Cancel = True
      Exit Sub
   End If
   'add by nickc 2007/07/13 將輸入法改成使用API
   If Cancel = False Then CloseIme
End Sub

Private Sub Text4_GotFocus()
   TextInverse Text4
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'*************************************************
'  重新整理 Adodc 之資料
'
'*************************************************
Public Sub AdodcRefresh()
On Error GoTo Checking
   adoadodc1.Close
   adoadodc1.CursorLocation = adUseClient
   If strConTitle = MsgText(31) Or strConTitle = MsgText(601) Then
      adoadodc1.Open "select * from acc090 order by a0901 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Else
      adoadodc1.Open "select * from acc090 where " & strConTitle & " >= '" & strCondition1 & "' and " & strConTitle & " <= '" & strCondition2 & "' order by a0901 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   End If
   Adodc1.Recordset.Requery
   If Adodc1.Recordset.RecordCount <> 0 Then
      If Text1 <> MsgText(601) Then
         Adodc1.Recordset.Find "a0901 = '" & Text1 & "'", 0, adSearchForward, 1
         If Adodc1.Recordset.EOF = False Then
            FormShow
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

Private Sub Text5_GotFocus()
   TextInverse Text5
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text5_Validate(Cancel As Boolean)
   lblA0908 = ""
   If Text5 <> MsgText(601) Then
      If PUB_GetStaffState(Text5, strExc(1), True) = 0 Then
         Cancel = True
         TextInverse Text5
      Else
         lblA0908 = strExc(1)
      End If
   End If
End Sub

Private Sub Text6_GotFocus()
   TextInverse Text6
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text6_Validate(Cancel As Boolean)
   lblA0909 = ""
   If Text6 <> MsgText(601) Then
      If PUB_GetStaffState(Text6, strExc(1), True) = 0 Then
         Cancel = True
         TextInverse Text6
      Else
         lblA0909 = strExc(1)
      End If
   End If
End Sub

Private Sub Text7_GotFocus()
   TextInverse Text7
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text7_Validate(Cancel As Boolean)
   lblA0910 = ""
   If Text7 <> MsgText(601) Then
      lblA0910 = GetDepartmentName(Text7)
      If lblA0910 = "" Then
         MsgBox "作帳部門輸入錯誤！"
         Cancel = True
      End If
   End If
End Sub
'Add by Morgan 2010/4/26
Public Function TxtValidate() As Boolean
   Dim bCancel As Boolean
   If Text5.Enabled = True Then
      Text5_Validate bCancel
      If bCancel Then Exit Function
   End If
   If Text6.Enabled = True Then
      Text6_Validate bCancel
      If bCancel Then Exit Function
   End If
   If Text7.Enabled = True Then
      Text7_Validate bCancel
      If bCancel Then Exit Function
   End If
   TxtValidate = True
End Function
