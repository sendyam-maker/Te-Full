VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc3190 
   AutoRedraw      =   -1  'True
   Caption         =   "銀行帳戶基本資料"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5460
   ScaleWidth      =   8820
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
      Height          =   315
      Left            =   6000
      TabIndex        =   7
      Top             =   1690
      Width           =   2420
   End
   Begin VB.CommandButton Command1 
      Height          =   300
      Left            =   5400
      Picture         =   "Frmacc3190.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   2
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
      Left            =   1320
      MaxLength       =   6
      TabIndex        =   5
      Top             =   1320
      Width           =   1572
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc3190.frx":0102
      Height          =   3000
      Left            =   240
      TabIndex        =   8
      Top             =   2160
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   5292
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
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "a0g02"
         Caption         =   "銀行名稱"
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
         DataField       =   "a0h02"
         Caption         =   "銀行帳號"
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
         DataField       =   "a0h04"
         Caption         =   "開戶日期"
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
      BeginProperty Column03 
         DataField       =   "a0h03"
         Caption         =   "帳戶名稱"
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
         DataField       =   "a0h08"
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
         DataField       =   "a0h15"
         Caption         =   "出名人"
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
      BeginProperty Column06 
         DataField       =   "a0h16"
         Caption         =   "存款類別"
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
            ColumnWidth     =   2924.788
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1635.024
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   1365.165
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   5160.189
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
      Left            =   240
      Top             =   2040
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
   Begin VB.TextBox Text5 
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
      Left            =   4200
      MaxLength       =   10
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  '靠右對齊
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
      Height          =   300
      Left            =   6840
      TabIndex        =   13
      Top             =   600
      Width           =   1572
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
      Left            =   1320
      MaxLength       =   12
      TabIndex        =   0
      Top             =   240
      Width           =   1572
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1320
      TabIndex        =   3
      Top             =   600
      Width           =   1572
      _ExtentX        =   2778
      _ExtentY        =   529
      _Version        =   393216
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
   Begin MSForms.TextBox Text7 
      Height          =   315
      Left            =   1320
      TabIndex        =   6
      Top             =   1690
      Width           =   3495
      VariousPropertyBits=   679493659
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text4 
      Height          =   300
      Left            =   2880
      TabIndex        =   17
      Top             =   1320
      Width           =   5532
      VariousPropertyBits=   679493661
      BackColor       =   14737632
      MaxLength       =   35
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text6 
      Height          =   300
      Left            =   5760
      TabIndex        =   15
      Top             =   240
      Width           =   2655
      VariousPropertyBits=   679493661
      BackColor       =   14737632
      MaxLength       =   35
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text3 
      Height          =   300
      Left            =   1320
      TabIndex        =   4
      Top             =   960
      Width           =   7092
      VariousPropertyBits=   679493659
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "存款類別"
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
      Left            =   5040
      TabIndex        =   19
      Top             =   1695
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "出 名 人"
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
      Top             =   1690
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "會計科目"
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
      TabIndex        =   16
      Top             =   1320
      Width           =   972
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "帳戶名稱"
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
      TabIndex        =   14
      Top             =   960
      Width           =   972
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   1935
      Left            =   240
      Top             =   120
      Width           =   8295
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "餘額"
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
      TabIndex        =   12
      Top             =   600
      Width           =   732
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "開戶日期"
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
      Top             =   600
      Width           =   972
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "銀行帳號"
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
      Top             =   240
      Width           =   972
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
      Left            =   3120
      TabIndex        =   9
      Top             =   240
      Width           =   972
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   4800
      Visible         =   0   'False
      Width           =   132
   End
End
Attribute VB_Name = "Frmacc3190"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/10/19 Form2.0已修改 Text3/Text4/Text6/Text7/DataGrid1
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/3 日期欄已修改
Option Explicit
Public adoacc0h0 As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Public adocheck As New ADODB.Recordset

Private Sub Command1_Click()
   If Adodc1.Recordset.RecordCount = 0 Or Text1 = MsgText(601) Or Text5 = MsgText(601) Then
      Exit Sub
   End If
   Adodc1.Recordset.Find "a0h02 = '" & Text1 & "'", 0, adSearchForward, 1
   If Adodc1.Recordset.EOF = False Then
      Adodc1.Recordset.Find "a0h01 = '" & Text5 & "'", 0, adSearchForward, Adodc1.Recordset.Bookmark
      If Adodc1.Recordset.EOF = False Then
         FormShow
         RecordShow
      Else
         MsgBox MsgText(33), , MsgText(5)
         Adodc1.Recordset.MoveFirst
      End If
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
   If strCompanyNo = MsgText(601) Then
      Exit Sub
   End If
   If Adodc1.Recordset.RecordCount <> 0 Then
      Adodc1.Recordset.MoveFirst
   End If
   Adodc1.Recordset.Find "a0h01 = '" & strCompanyNo & "'", 0, adSearchForward, 1
   If Adodc1.Recordset.EOF = False Then
      Adodc1.Recordset.Find "a0h02 = '" & strItemNo & "'", 0, adSearchForward, 1
      If Adodc1.Recordset.EOF = False Then
         FormShow
         RecordShow
      End If
   End If
   strCompanyNo = MsgText(601)
End Sub

'Add by Amy 2021/10/19
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call PUB_SaveTrackMode(0, KeyCode)  'Form2.0 記錄鍵盤傳入順序
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
   'Modify by Amy 2023/10/11 原W8850 H5500
   Me.Width = 8940
   Me.Height = 5925
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
   Set Frmacc3190 = Nothing
End Sub

Private Sub MaskEdBox1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub MaskEdBox1_Validate(Cancel As Boolean)
   If MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Then
      MsgBox Label2 & MsgText(52), , MsgText(5)
      Cancel = True
      MaskEdBox1.SetFocus
      Exit Sub
   End If
   If DateCheck(MaskEdBox1.Text) = MsgText(603) Then
      MsgBox Label2 & MsgText(63), , MsgText(5)
      Cancel = True
      MaskEdBox1.SetFocus
      Exit Sub
   End If
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Text10_Change()
   If Text10 = MsgText(601) Then
      Exit Sub
   End If
   Text4 = A0102Query(Text10)
End Sub

Private Sub Text10_GotFocus()
   TextInverse Text10
End Sub

Private Sub Text10_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Text10_Validate(Cancel As Boolean)
   If strSaveConfirm <> MsgText(3) And strSaveConfirm <> MsgText(4) Then
      Exit Sub
   End If
   If ExistCheck("acc010", "a0101", Text10, Label5) = False Then
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

'Modify by Amy 2021/10/19 原:KeyCode As Integer
Private Sub Text3_KeyUp(KeyCode As MSForms.ReturnInteger, Shift As Integer)
   KeyEnter Val(KeyCode)
End Sub

Private Sub Text3_LostFocus()
   StatusView MsgText(601)
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
   If CheckLen(Label4, Text3, 30) = MsgText(603) Then
      Cancel = True
      Exit Sub
   End If
   'add by nickc 2007/07/13 將輸入法改成使用API
   If Cancel = False Then CloseIme
End Sub

Private Sub Text5_Change()
   If Text5 = MsgText(601) Then
      Exit Sub
   End If
   Text6 = A0g02Query(Text5)
End Sub

Private Sub Text5_GotFocus()
   TextInverse Text5
End Sub

Private Sub Text5_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoacc0h0.CursorLocation = adUseClient
   adoacc0h0.Open "select * from acc0h0 order by a0h01 asc, a0h02 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "select * from acc0h0, acc0g0 where acc0h0.a0h01 = acc0g0.a0g01 order by a0h01 asc, a0h02 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  顯示資料表(銀行帳戶資料)
'
'*************************************************
Public Sub FormShow()
Dim adoacc040 As New ADODB.Recordset
Dim adoacc0b0 As New ADODB.Recordset
Dim intYear As Integer
Dim intMonth As Integer
Dim strCurrDate As String

   Text5 = Adodc1.Recordset.Fields("a0h01").Value
   Text1 = Adodc1.Recordset.Fields("a0h02").Value
   MaskEdBox1.Mask = MsgText(601)
   If IsNull(Adodc1.Recordset.Fields("a0h04").Value) Then
      MaskEdBox1.Text = MsgText(601)
   Else
      MaskEdBox1.Text = CFDate(Trim(str(Adodc1.Recordset.Fields("a0h04").Value)))
   End If
   MaskEdBox1.Mask = DFormat
   If IsNull(Adodc1.Recordset.Fields("a0h03").Value) Then
      Text3 = MsgText(601)
   Else
      Text3 = Adodc1.Recordset.Fields("a0h03").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a0h08").Value) Then
      Text10 = MsgText(601)
   Else
      Text10 = Adodc1.Recordset.Fields("a0h08").Value
   End If
   'Add by Amy 2013/08/12 +出名人及存款類別欄位
   If IsNull(Adodc1.Recordset.Fields("a0h15").Value) Then
      Text7 = MsgText(601)
   Else
      Text7 = Adodc1.Recordset.Fields("a0h15").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a0h16").Value) Then
      Text8 = MsgText(601)
   Else
      Text8 = Adodc1.Recordset.Fields("a0h16").Value
   End If
   'end 2013/08/12
   adoacc0b0.CursorLocation = adUseClient
   adoacc0b0.Open "select * from acc0b0", adoTaie, adOpenStatic, adLockReadOnly
   If adoacc0b0.RecordCount = 0 Then
      Text2 = MsgText(601)
      adoacc0b0.Close
      Exit Sub
   End If
   If IsNull(adoacc0b0.Fields("a0b02").Value) Then
      strCurrDate = ""
   Else
      strCurrDate = adoacc0b0.Fields("a0b02").Value
   End If
   adoacc040.CursorLocation = adUseClient
   adoacc040.Open "select a0408 from acc040 where a0401 = " & Val(Mid(CFDate(IIf(IsNull(adoacc0b0.Fields("a0b02").Value), "", adoacc0b0.Fields("a0b02").Value)), 1, 3)) & " and a0402 = " & Val(Mid(CFDate(IIf(IsNull(adoacc0b0.Fields("a0b02").Value), "", adoacc0b0.Fields("a0b02").Value)), 5, 2)) & " and a0403 = '1' and a0404 = '" & MsgText(55) & "' and a0405 = '" & Text10 & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoacc040.RecordCount <> 0 Then
      If IsNull(adoacc040.Fields("a0408").Value) Then
         Text2 = MsgText(601)
      Else
         Text2 = Format(adoacc040.Fields("a0408").Value, DDollar)
      End If
   Else
      Text2 = MsgText(601)
   End If
   adoacc0b0.Close
   adoacc040.Close
   
   '未收未付之票據
   adocheck.CursorLocation = adUseClient
   adocheck.Open "select sum(a0e11) from acc0e0, acc0h0 where a0e19 = a0h01 and a0e20 = a0h02 and a0h08 = '" & Text10 & "' and a0e04 = '" & MsgText(18) & "' and (a0e21 = 0 or a0e21 is null) and a0e17 = 0 and a0e15 = 0 and a0e34 = 0", adoTaie, adOpenStatic, adLockReadOnly
   If adocheck.RecordCount <> 0 Then
      If IsNull(adocheck.Fields(0).Value) = False Then
         Text2 = Format(Val(Replace(Text2, ",", "")) + adocheck.Fields(0).Value, DDollar)
      End If
   End If
   adocheck.Close
   adocheck.CursorLocation = adUseClient
   adocheck.Open "select sum(a0e11) from acc0e0, acc0h0 where a0e01 = a0h01 and a0e07 = a0h02 and a0h08 = '" & Text10 & "' and a0e04 = '" & MsgText(19) & "' and (a0e37 = 0 or a0e37 is null) and a0e25 = 0", adoTaie, adOpenStatic, adLockReadOnly
   If adocheck.RecordCount <> 0 Then
      If IsNull(adocheck.Fields(0).Value) = False Then
         Text2 = Format(Val(Replace(Text2, ",", "")) - adocheck.Fields(0).Value, DDollar)
      End If
   End If
   adocheck.Close
   '已入帳
   adocheck.CursorLocation = adUseClient
   adocheck.Open "select sum(ax206 - ax207) from acc021, acc020 where ax201 = a0201 and ax202 = a0202 and ax205 = '" & Text10 & "' and a0205 > " & Val(strCurrDate) & "", adoTaie, adOpenStatic, adLockReadOnly
   If adocheck.RecordCount <> 0 Then
      If IsNull(adocheck.Fields(0).Value) = False Then
         Text2 = Format(Val(Replace(Text2, ",", "")) + adocheck.Fields(0).Value, DDollar)
      End If
   End If
   adocheck.Close
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
      adoadodc1.Open "select * from acc0h0, acc0g0 where acc0h0.a0h01 = acc0g0.a0g01 order by a0h01 asc, a0h02 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Else
      If strConTitle <> strCon4 Then
         adoadodc1.Open "select * from acc0h0, acc0g0 where acc0h0.a0h01 = acc0g0.a0g01 and " & strConTitle & " >= '" & strCondition1 & "' and " & strConTitle & " <= '" & strCondition2 & "' order by a0h01 asc, a0h02 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
      Else
         adoadodc1.Open "select * from acc0h0, acc0g0 where acc0h0.a0h01 = acc0g0.a0g01 and " & strConTitle & " >= " & Val(strCondition1) & " and " & strConTitle & " <= " & Val(strCondition2) & " order by a0h01 asc, a0h02 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
      End If
   End If
   Adodc1.Recordset.Requery
   If Adodc1.Recordset.RecordCount <> 0 Then
      If Text1 <> MsgText(601) And Text5 <> MsgText(601) Then
         Adodc1.Recordset.Find "a0h02 = '" & Text1 & "'", 0, adSearchForward, 1
         If Adodc1.Recordset.EOF = False Then
            Adodc1.Recordset.Find "a0h01 = '" & Text5 & "'", 0, adSearchForward, Adodc1.Recordset.Bookmark
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

Private Sub Text5_Validate(Cancel As Boolean)
   If strSaveConfirm <> MsgText(3) And strSaveConfirm <> MsgText(4) Then
      Exit Sub
   End If
   If ExistCheck("acc0g0", "a0g01", Text5, Label9) = False Then
      Cancel = True
      Exit Sub
   End If
End Sub

'Add by Amy 2013/08/12 +出名人及存款類別
Private Sub Text7_GotFocus()
    TextInverse Text7
    OpenIme
End Sub

'Add by Amy 2021/10/19 從aacc_sav搬回
Public Sub Frmacc3190_Save()
   On Error GoTo Checking
   With Frmacc3190
      If .Text5 = MsgText(601) Then
         MsgBox MsgText(10) & .Label9, , MsgText(5)
         strControlButton = MsgText(602)
         .Text5.SetFocus
         Exit Sub
      Else
         If .Text1 = MsgText(601) Then
            MsgBox MsgText(10) & .Label1, , MsgText(5)
            strControlButton = MsgText(602)
            .Text1.SetFocus
            Exit Sub
         End If
         If ExistCheck("acc0g0", "a0g01", .Text5, .Label9) = False Then
            strControlButton = MsgText(602)
            .Text5.SetFocus
            Exit Sub
         End If
         If ExistCheck("acc010", "a0101", .Text10, .Label5) = False Then
            strControlButton = MsgText(602)
            .Text10.SetFocus
            Exit Sub
         End If
         If .MaskEdBox1.Text = MsgText(601) Or .MaskEdBox1.Text = MsgText(29) Then
            MsgBox .Label2 & MsgText(52), , MsgText(5)
            strControlButton = MsgText(602)
            .MaskEdBox1.SetFocus
            Exit Sub
         Else
            If DateCheck(.MaskEdBox1.Text) = MsgText(603) Then
               MsgBox .Label2 & MsgText(63), , MsgText(5)
               strControlButton = MsgText(602)
               .MaskEdBox1.SetFocus
               Exit Sub
            End If
         End If
         If CheckLen(.Label4, .Text3, 30) = MsgText(603) Then
            strControlButton = MsgText(602)
            .Text3.SetFocus
            Exit Sub
         End If
          'Add by Amy 2013/08/12 +出名人、存款類別必填
         If .Text7 = MsgText(601) Then
            MsgBox MsgText(10) & .Label6, , MsgText(5)
            strControlButton = MsgText(602)
            .Text7.SetFocus
            Exit Sub
         Else
            '判斷只能輸入15個中文
            If CheckLengthIsOK(.Text7, 30) = False Then
                strControlButton = MsgText(602)
                .Text7.SetFocus
                Exit Sub
            End If
         End If
         If .Text8 = MsgText(601) Then
            MsgBox MsgText(10) & .Label7, , MsgText(5)
            strControlButton = MsgText(602)
            .Text8.SetFocus
            Exit Sub
         Else
            '判斷只能輸入6個中文
            If CheckLengthIsOK(.Text8, 12) = False Then
                strControlButton = MsgText(602)
                .Text8.SetFocus
                Exit Sub
            End If
         End If
         'end 2013/08/12
      End If
      'Add by Amy 2021/10/19 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
      If PUB_ChkUniText(Me) = False Then
         strControlButton = MsgText(602)
         Exit Sub
      End If

      If strSaveConfirm = MsgText(3) Then
         If .adoacc0h0.RecordCount <> 0 Then
            .adoacc0h0.Find "a0h02 = '" & .Text1 & "'", 0, adSearchForward, 1
            If .adoacc0h0.EOF = False Then
               .adoacc0h0.Find "a0h01 = '" & .Text5 & "'", 0, adSearchForward, .adoacc0h0.Bookmark
               If .adoacc0h0.EOF = False Then
                  MsgBox MsgText(9), , MsgText(5)
                  strControlButton = MsgText(602)
                  Exit Sub
               End If
            End If
         End If
         .Adodc1.Recordset.AddNew
      End If
      .Adodc1.Recordset.Fields("a0h01").Value = .Text5
      .Adodc1.Recordset.Fields("a0h02").Value = .Text1
      If .MaskEdBox1.Text <> MsgText(601) And .MaskEdBox1.Text <> MsgText(29) Then
         .Adodc1.Recordset.Fields("a0h04").Value = Val(FCDate(.MaskEdBox1.Text))
      Else
         .Adodc1.Recordset.Fields("a0h04").Value = Null
      End If
      If .Text3 <> MsgText(601) Then
         .Adodc1.Recordset.Fields("a0h03").Value = .Text3
      Else
         .Adodc1.Recordset.Fields("a0h03").Value = Null
      End If
      If .Text10 <> MsgText(601) Then
         .Adodc1.Recordset.Fields("a0h08").Value = .Text10
      Else
         .Adodc1.Recordset.Fields("a0h08").Value = Null
      End If
      'Add by Amy 2013/08/12 +出名人及存款類別
      If .Text7 <> MsgText(601) Then
         .Adodc1.Recordset.Fields("a0h15").Value = .Text7
      Else
         .Adodc1.Recordset.Fields("a0h15").Value = Null
      End If
      If .Text8 <> MsgText(601) Then
         .Adodc1.Recordset.Fields("a0h16").Value = .Text8
      Else
         .Adodc1.Recordset.Fields("a0h16").Value = Null
      End If
      'end 2013/08/12
      If strSaveConfirm = MsgText(3) Then
         .Adodc1.Recordset.Fields("a0h09").Value = Val(strSrvDate(2))
         .Adodc1.Recordset.Fields("a0h10").Value = ServerTime
         .Adodc1.Recordset.Fields("a0h11").Value = strUserNum
      Else
         .Adodc1.Recordset.Fields("a0h12").Value = Val(strSrvDate(2))
         .Adodc1.Recordset.Fields("a0h13").Value = ServerTime
         .Adodc1.Recordset.Fields("a0h14").Value = strUserNum
      End If
      .Adodc1.Recordset.UpdateBatch
      .AdodcRefresh
      .FormShow
      .RecordShow
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
   End With
End Sub
