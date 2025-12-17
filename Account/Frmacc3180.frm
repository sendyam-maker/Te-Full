VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc3180 
   AutoRedraw      =   -1  'True
   Caption         =   "銀行基本資料"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8835
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5400
   ScaleWidth      =   8835
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
      Height          =   315
      Left            =   7650
      MaxLength       =   3
      TabIndex        =   4
      Top             =   570
      Width           =   765
   End
   Begin VB.CommandButton Command1 
      Height          =   300
      Left            =   3360
      Picture         =   "Frmacc3180.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   240
      Width           =   350
   End
   Begin VB.TextBox Text10 
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
      Left            =   7920
      MaxLength       =   1
      TabIndex        =   2
      Top             =   240
      Width           =   492
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc3180.frx":0102
      Height          =   2535
      Left            =   225
      TabIndex        =   10
      Top             =   2490
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   4471
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
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "A0G06"
         Caption         =   "轉帳代碼"
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
         DataField       =   "a0g01"
         Caption         =   "銀行代號"
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
         DataField       =   "a0g02"
         Caption         =   "銀行簡稱"
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
         DataField       =   "a0g04"
         Caption         =   "銀行名稱(中)"
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
         DataField       =   "a0g05"
         Caption         =   "電話"
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
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   989.858
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1379.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1980.284
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   3630.047
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1709.858
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
      Left            =   0
      Top             =   3210
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
      MaxLength       =   10
      TabIndex        =   0
      Top             =   240
      Width           =   1572
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1800
      TabIndex        =   7
      Top             =   1620
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   300
      Left            =   6840
      TabIndex        =   8
      Top             =   1620
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSForms.TextBox Text3 
      Height          =   300
      Left            =   1800
      TabIndex        =   5
      Top             =   930
      Width           =   6612
      VariousPropertyBits=   679493659
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text6 
      Height          =   300
      Left            =   1800
      TabIndex        =   9
      Top             =   1980
      Width           =   6612
      VariousPropertyBits=   679493659
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text4 
      Height          =   300
      Left            =   1800
      TabIndex        =   6
      Top             =   1260
      Width           =   6612
      VariousPropertyBits=   679493659
      MaxLength       =   100
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text2 
      Height          =   315
      Left            =   1800
      TabIndex        =   3
      Top             =   570
      Width           =   4410
      VariousPropertyBits=   679493659
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "轉帳代碼"
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
      Left            =   6480
      TabIndex        =   19
      Top             =   570
      Width           =   900
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "銀行名稱(中)"
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
      Left            =   360
      TabIndex        =   18
      Top             =   930
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "是否在台北市(Y/N)"
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
      Left            =   5880
      TabIndex        =   17
      Top             =   240
      Width           =   2052
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   2265
      Left            =   240
      Top             =   120
      Width           =   8295
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "地址"
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
      Left            =   360
      TabIndex        =   16
      Top             =   1980
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "傳真"
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
      Left            =   5880
      TabIndex        =   15
      Top             =   1620
      Width           =   495
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "電話"
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
      Left            =   360
      TabIndex        =   14
      Top             =   1620
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "銀行名稱(英)"
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
      Left            =   360
      TabIndex        =   13
      Top             =   1260
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "銀行簡稱"
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
      Left            =   360
      TabIndex        =   12
      Top             =   570
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   4800
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "銀行代號"
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
      Top             =   240
      Width           =   972
   End
End
Attribute VB_Name = "Frmacc3180"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/10/19 Form2.0已修改 Text2/Text3/Text6/DataGrid1
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/3 日期欄已修改
Option Explicit
Public adoadodc1 As New ADODB.Recordset

Private Sub Command1_Click()
   If Adodc1.Recordset.RecordCount = 0 Or Text1 = MsgText(601) Then
      Exit Sub
   End If
   Adodc1.Recordset.Find "a0g01 = '" & Text1 & "'", 0, adSearchForward, 1
   If Adodc1.Recordset.EOF = False Then
      FormShow
      RecordShow
   Else
      MsgBox MsgText(33), , MsgText(5)
      Adodc1.Recordset.MoveFirst
   End If
End Sub

Private Sub Command1_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Command1_Click
         Exit Sub
   End Select
   KeyEnter KeyCode
End Sub

Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
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
   If Adodc1.Recordset.RecordCount <> 0 Then
      Adodc1.Recordset.MoveFirst
   End If
   Adodc1.Recordset.Find "a0g01 = '" & strItemNo & "'", 0, adSearchForward, 1
   If Adodc1.Recordset.EOF = False Then
      FormShow
      RecordShow
   End If
   strItemNo = MsgText(601)
End Sub

'Add by Amy 2021/10/19
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call PUB_SaveTrackMode(0, KeyCode)  'Form2.0 記錄鍵盤傳入順序
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Form_Load()
Dim intX As Integer, intY As Integer, sglWidth As Single, sglHeight As Single, intCounter As Integer
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   'Modify by Amy 2023/10/11 原W8850 H5500
   Me.Width = 8955
   Me.Height = 5870
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
   FormEnabled False 'Add by Morgan 2007/2/7
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
   strTrackMode = "" 'Add by Amy 2021/10/19 Form2.0 記錄鍵盤傳入順序(清除)
   Set Frmacc3180 = Nothing
End Sub

Private Sub MaskEdBox1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub MaskEdBox2_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 And Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then
      KeyAscii = 0
   End If
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoadodc1.CursorLocation = adUseClient
   'Modify by Morgan 2007/2/7
   'adoadodc1.Open "select * from acc0g0 order by a0g01 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoadodc1.Open "select * from acc0g0 order by a0g06 asc,a0g01 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   'End 2007/2/7
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  顯示資料表(銀行別資料)
'
'*************************************************
Public Sub FormShow()
   Text1 = Adodc1.Recordset.Fields("a0g01").Value
   If IsNull(Adodc1.Recordset.Fields("a0g09").Value) Then
      Text10 = MsgText(601)
   Else
      Text10 = Adodc1.Recordset.Fields("a0g09").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a0g02").Value) Then
      Text2 = MsgText(601)
   Else
      Text2 = Adodc1.Recordset.Fields("a0g02").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a0g03").Value) Then
      Text4 = MsgText(601)
   Else
      Text4 = Adodc1.Recordset.Fields("a0g03").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a0g05").Value) Then
      MaskEdBox1.Text = MsgText(601)
   Else
      MaskEdBox1.Text = Adodc1.Recordset.Fields("a0g05").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a0g07").Value) Then
      MaskEdBox2.Text = MsgText(601)
   Else
      MaskEdBox2.Text = Adodc1.Recordset.Fields("a0g07").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a0g08").Value) Then
      Text6 = MsgText(601)
   Else
      Text6 = Adodc1.Recordset.Fields("a0g08").Value
   End If
   'Add by Morgan 2007/2/7
   Text5 = "" & Adodc1.Recordset.Fields("a0g06").Value
   Text3 = "" & Adodc1.Recordset.Fields("a0g04").Value
   'End 2007/2/7
   
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   If Text5.Locked = False And Text5 = "" Then
      Text5 = Mid(Text1, 3, 3)
   End If
End Sub

Private Sub Text10_GotFocus()
   TextInverse Text10
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") And KeyAscii <> Asc("N") Then
      KeyAscii = 0
   End If
End Sub

Private Sub Text10_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Text2_GotFocus()
   StatusView MsgText(65) & "100"
   TextInverse Text2
   'add by nickc 2007/07/13 將輸入法改成使用API
   OpenIme
End Sub

'Modify by Amy 2021/10/19 原:KeyCode As Integer
Private Sub Text2_KeyUp(KeyCode As MSForms.ReturnInteger, Shift As Integer)
   KeyEnter Val(KeyCode)
End Sub

Private Sub Text2_LostFocus()
   StatusView MsgText(601)
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
   If CheckLen(Label1, Text2, 100) = MsgText(603) Then
      Cancel = True
      Exit Sub
   End If
   'add by nickc 2007/07/13 將輸入法改成使用API
   If Cancel = False Then CloseIme
End Sub


'add by nickc 2007/07/13 將輸入法改成使用API
Private Sub Text3_GotFocus()
OpenIme
End Sub
'add by nickc 2007/07/13 將輸入法改成使用API
Private Sub Text3_Validate(Cancel As Boolean)
CloseIme
End Sub

Private Sub Text4_GotFocus()
   TextInverse Text4
End Sub

'Moidfy by Amy 2021/10/19 原:KeyCode As Integer
Private Sub Text4_KeyUp(KeyCode As MSForms.ReturnInteger, Shift As Integer)
   KeyEnter Val(KeyCode)
End Sub


Private Sub Text5_GotFocus()
   TextInverse Text5
   'add by nickc 2007/07/13 將輸入法改成使用API
   OpenIme
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then
      KeyAscii = 0
   End If
End Sub
'add by nickc 2007/07/13 將輸入法改成使用API
Private Sub Text5_Validate(Cancel As Boolean)
CloseIme
End Sub

Private Sub Text6_GotFocus()
   StatusView MsgText(65) & "100"
   TextInverse Text6
   'add by nickc 2007/07/13 將輸入法改成使用API
   OpenIme
End Sub

'Modify by Amy 2021/10/19 原:KeyCode As Integer
Private Sub Text6_KeyUp(KeyCode As MSForms.ReturnInteger, Shift As Integer)
   KeyEnter Val(KeyCode)
End Sub

'*************************************************
'  重新整理 Adodc 之資料
'
'*************************************************
Public Sub AdodcRefresh()
On Error GoTo Checking
   adoadodc1.Close
   adoadodc1.CacheSize = 100
   adoadodc1.CursorLocation = adUseClient
   If strConTitle = MsgText(31) Or strConTitle = MsgText(601) Then
      strSql = "select * from acc0g0"
   Else
      strSql = "select * from acc0g0 where " & strConTitle & " >= '" & strCondition1 & "' and " & strConTitle & " <= '" & strCondition2
   End If
   'Modify by Morgan 2007/2/7
   'strSQL = strSQL & " order by a0g01 asc"
   strSql = strSql & " order by a0g06 asc,a0g01 asc"
   'End 2007/2/7
   adoadodc1.Open strSql, adoTaie, adOpenDynamic, adLockBatchOptimistic
   Adodc1.Recordset.Requery
   If Adodc1.Recordset.RecordCount <> 0 Then
      If Text1 <> MsgText(601) Then
         Adodc1.Recordset.Find "a0g01 = '" & Text1 & "'", 0, adSearchForward, 1
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

Private Sub Text6_LostFocus()
   StatusView MsgText(601)
End Sub

Private Sub Text6_Validate(Cancel As Boolean)
   If CheckLen(Label7, Text6, 100) = MsgText(603) Then
      Cancel = True
      Exit Sub
   End If
   'add by nickc 2007/07/13 將輸入法改成使用API
   If Cancel = False Then CloseIme
End Sub
'Add by Morgan 2007/2/7
Public Sub FormEnabled(Optional p_bolOn As Boolean = True, Optional p_bolAll As Boolean = False)
   If p_bolAll Then
      Me.Text1.Locked = Not p_bolOn
   Else
      Me.Text1.Locked = p_bolOn
   End If
   Me.Text10.Locked = Not p_bolOn
   Me.Text2.Locked = Not p_bolOn
   Me.Text3.Locked = Not p_bolOn
   Me.Text4.Locked = Not p_bolOn
   Me.Text5.Locked = Not p_bolOn
   Me.Text6.Locked = Not p_bolOn
End Sub

Public Sub Frmacc3180_Save()
   On Error GoTo Checking
   With Frmacc3180
      If .Text1 = MsgText(601) Then
         MsgBox MsgText(10) & .Label9, , MsgText(5)
         strControlButton = MsgText(602)
         .Text1.SetFocus
         Exit Sub
      Else
         If CheckLen(.Label1, .Text2, 100) = MsgText(603) Then
            strControlButton = MsgText(602)
            .Text2.SetFocus
            Exit Sub
         End If
         If CheckLen(.Label7, .Text6, 100) = MsgText(603) Then
            strControlButton = MsgText(602)
            .Text6.SetFocus
            Exit Sub
         End If
      End If
      'Add by Amy 2021/10/19 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
      If PUB_ChkUniText(Me) = False Then
         strControlButton = MsgText(602)
         Exit Sub
      End If

      If strSaveConfirm = MsgText(3) Then
         If .Adodc1.Recordset.RecordCount <> 0 Then
            .Adodc1.Recordset.Find "a0g01 = '" & .Text1 & "'", 0, adSearchForward, 1
            If .Adodc1.Recordset.EOF = False Then
               MsgBox MsgText(9), , MsgText(5)
               strControlButton = MsgText(602)
               Exit Sub
            End If
         End If
         .Adodc1.Recordset.AddNew
      End If
      .Adodc1.Recordset.Fields("a0g01").Value = .Text1
      If .Text10 <> MsgText(601) Then
         .Adodc1.Recordset.Fields("a0g09").Value = .Text10
      Else
         .Adodc1.Recordset.Fields("a0g09").Value = Null
      End If
      If .Text2 <> MsgText(601) Then
         .Adodc1.Recordset.Fields("a0g02").Value = .Text2
      Else
         .Adodc1.Recordset.Fields("a0g02").Value = Null
      End If
      If .Text4 <> MsgText(601) Then
         .Adodc1.Recordset.Fields("a0g03").Value = .Text4
      Else
         .Adodc1.Recordset.Fields("a0g03").Value = Null
      End If
      If .MaskEdBox1.Text <> MsgText(601) Then
         .Adodc1.Recordset.Fields("a0g05").Value = .MaskEdBox1.Text
      Else
         .Adodc1.Recordset.Fields("a0g05").Value = Null
      End If
      If .MaskEdBox2.Text <> MsgText(601) Then
         .Adodc1.Recordset.Fields("a0g07").Value = .MaskEdBox2.Text
      Else
         .Adodc1.Recordset.Fields("a0g07").Value = Null
      End If
      If .Text6 <> MsgText(601) Then
         .Adodc1.Recordset.Fields("a0g08").Value = .Text6
      Else
         .Adodc1.Recordset.Fields("a0g08").Value = Null
      End If
      'Add by Morgan 2007/2/7
      If .Text5 <> "" Then
         .Adodc1.Recordset.Fields("a0g06").Value = .Text5
      Else
         .Adodc1.Recordset.Fields("a0g06").Value = "999"
      End If
      If .Text3 <> "" Then
         .Adodc1.Recordset.Fields("a0g04").Value = .Text3
      Else
         .Adodc1.Recordset.Fields("a0g04").Value = Null
      End If
      'End 2007/2/7
      .Adodc1.Recordset.UpdateBatch
      .AdodcRefresh
      .FormShow
      .RecordShow
      .FormEnabled False
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
   End With
End Sub
