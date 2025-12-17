VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc3130 
   AutoRedraw      =   -1  'True
   Caption         =   "票據託收作業"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5115
   ScaleWidth      =   8760
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "Frmacc3130.frx":0000
      Left            =   1230
      List            =   "Frmacc3130.frx":0002
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   180
      Width           =   2375
   End
   Begin VB.TextBox Text3 
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
      Left            =   3360
      MaxLength       =   10
      TabIndex        =   16
      Top             =   4608
      Width           =   1455
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
      Left            =   840
      MaxLength       =   10
      TabIndex        =   14
      Top             =   4608
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Height          =   372
      Left            =   8040
      Picture         =   "Frmacc3130.frx":0004
      Style           =   1  '圖片外觀
      TabIndex        =   13
      Top             =   560
      Width           =   372
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc3130.frx":066E
      Height          =   3100
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   5450
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
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
      Caption         =   "票據託收資料"
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "a0e02"
         Caption         =   "票據號碼"
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
         DataField       =   "a0e10"
         Caption         =   "到期日期"
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
      BeginProperty Column02 
         DataField       =   "a0e11"
         Caption         =   "票據金額"
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
      BeginProperty Column03 
         DataField       =   "a0g02"
         Caption         =   "收票銀行"
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
         DataField       =   "a0e07"
         Caption         =   "收票帳號"
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
         DataField       =   "a0e08"
         Caption         =   "票別"
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
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         Size            =   344
         BeginProperty Column00 
            ColumnWidth     =   1349.858
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   1425.26
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   1530.142
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   3479.811
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1590.236
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnWidth     =   585.071
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   6419.906
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   312
      Left            =   240
      Top             =   1320
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   556
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
   Begin VB.TextBox Text11 
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
      Height          =   315
      Left            =   4050
      MaxLength       =   10
      TabIndex        =   4
      Top             =   1080
      Width           =   1572
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H00FFFFFF&
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
      Left            =   4680
      MaxLength       =   10
      TabIndex        =   1
      Top             =   180
      Width           =   1272
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
      Height          =   315
      Left            =   1320
      MaxLength       =   8
      TabIndex        =   3
      Top             =   1080
      Width           =   1572
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1230
      TabIndex        =   2
      Top             =   600
      Width           =   1575
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
   Begin MSForms.TextBox Text12 
      Height          =   315
      Left            =   5640
      TabIndex        =   7
      Top             =   1080
      Width           =   2775
      VariousPropertyBits=   679493661
      BackColor       =   14737632
      MaxLength       =   35
      Size            =   "4895;556"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text10 
      Height          =   315
      Left            =   5970
      TabIndex        =   6
      Top             =   180
      Width           =   2580
      VariousPropertyBits=   679493661
      BackColor       =   14737632
      MaxLength       =   35
      Size            =   "4551;556"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "金額合計"
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
      Left            =   2280
      TabIndex        =   17
      Top             =   4608
      Width           =   972
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "筆數"
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
      Left            =   240
      TabIndex        =   15
      Top             =   4608
      Width           =   492
   End
   Begin VB.Label Label3 
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
      Height          =   255
      Left            =   270
      TabIndex        =   12
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "託收銀行"
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
      Left            =   3720
      TabIndex        =   11
      Top             =   240
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   855
      Left            =   225
      Top             =   120
      Width           =   8355
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "收票銀行"
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
      TabIndex        =   10
      Top             =   1080
      Width           =   972
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "票據號碼"
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
      Top             =   1080
      Width           =   972
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "託收日期"
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
      Left            =   270
      TabIndex        =   8
      Top             =   600
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   4248
      Visible         =   0   'False
      Width           =   132
   End
End
Attribute VB_Name = "Frmacc3130"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/10/07 Form2.0已修改 Text10/Text12/DataGrid1
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/3 日期欄已修改
Option Explicit
Public adoacc0e0 As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Public adocheck As New ADODB.Recordset
Public strA0E23 As String   '2014/1/23 add by sonia

Private Sub Command1_Click()
   AdodcDelete
End Sub

Private Sub Command1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
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
   'Add by Morgan 2006/7/21
   MaskEdBox1 = CFDate(strSrvDate(2))
   MaskEdBox1.Mask = DFormat
   PUB_SetAccount Combo1
   'end 2006/7/21
   
   OpenTable
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(133)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   StatusClear
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc3130 = Nothing
End Sub

Private Sub MaskEdBox1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub MaskEdBox1_Validate(Cancel As Boolean)
   If MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Then
      MsgBox Label2 & MsgText(52), , MsgText(5)
      Cancel = True
      MaskEdBox1.SetFocus
      Exit Sub
   Else
      If DateCheck(MaskEdBox1.Text) = MsgText(603) Then
         MsgBox Label2 & MsgText(63), , MsgText(5)
         Cancel = True
         MaskEdBox1.SetFocus
         Exit Sub
      End If
      AdodcRefresh
   End If
End Sub

Private Sub Combo1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Combo1_Validate(Cancel As Boolean)
   Dim strCombo1 As String 'Add by Amy 2023/05/23
   If Combo1 = MsgText(601) Then
      MsgBox Label3 & MsgText(52), , MsgText(5)
      Cancel = True
      Exit Sub
   Else
      If adocheck.State = adStateOpen Then adocheck.Close
      adocheck.CursorLocation = adUseClient
      'Modify by Amy 2023/05/23 原:Combo1
      If Combo1 <> MsgText(601) Then strCombo1 = Left(Combo1.Text, InStr(Combo1, " ") - 1)
      adocheck.Open "select a0h01, a0h02 from acc0h0 where a0h02 = '" & strCombo1 & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adocheck.RecordCount <> 0 Then
         If IsNull(adocheck.Fields(0).Value) Then
            Text9 = MsgText(601)
         Else
            Text9 = adocheck.Fields(0).Value
         End If
      Else
         MsgBox MsgText(28) & Label3, , MsgText(5)
         Text9 = MsgText(601)
         Cancel = True
         Combo1.SetFocus
      End If
      adocheck.Close
      '2014/1/23 add by sonia
      If Cancel = True Then Exit Sub
      'modify by sonia 2015/5/12 加智權 華銀長安0236819
      If strCombo1 = "1607750" Or strCombo1 = "0236819" Then
         strA0E23 = "J"
      'add by sonia 2020/4/7 加法律所
      ElseIf strCombo1 = "1756890" Then
      'end 2023/05/23
         strA0E23 = "L"
      'end 2020/4/7
      Else
         strA0E23 = "1"
      End If
      '2014/1/23 end
   End If
End Sub

Private Sub Text11_Change()
   Text12 = A0g02Query(Text11)
End Sub

Private Sub Text11_GotFocus()
   TextInverse Text11
End Sub

Private Sub Text11_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Text11_Validate(Cancel As Boolean)
   
   If Text11.Text = "" Then Exit Sub 'Add by Morgan 2004/11/15 空白不檢查
   
   If ExistCheck("acc0g0", "a0g01", Text11, Label9) = False Then
      Cancel = True
      Exit Sub
   End If
End Sub

Private Sub Text5_GotFocus()
   TextInverse Text5
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text5_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoadodc1.CursorLocation = adUseClient
   'Modify by Morgan 2004/11/1 加 and rownum<1
   adoadodc1.Open "select * from acc0e0 where a0e19 = '" & Text9 & "' and a0e20 = '" & Combo1 & "' and a0e14 = " & Val(FCDate(MaskEdBox1.Text)) & " and a0e25 = 0 and rownum<1 order by a0e14 asc, a0e47 desc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

Private Sub Text5_Validate(Cancel As Boolean)
   If Text5 <> MsgText(601) Then
      If adocheck.State = adStateOpen Then adocheck.Close
      adocheck.CursorLocation = adUseClient
      'Modify by Morgan 2004/10/29  加應收過濾條件 and a0e04='R'
      'adocheck.Open "select a0e01, a0e02 from acc0e0 where a0e02 = '" & Text5 & "'", adoTaie, adOpenStatic, adLockReadOnly
      adocheck.Open "select a0e01 from acc0e0 where a0e02 = '" & Text5 & "' and a0e04='R'", adoTaie, adOpenStatic, adLockReadOnly
      If adocheck.RecordCount <> 0 Then
         If IsNull(adocheck.Fields(0).Value) Then
            Text11 = MsgText(601)
         Else
            Text11 = adocheck.Fields(0).Value
         End If
      Else
         Text11 = MsgText(601)
         'Add by Morgan 2004/11/16
         MsgBox MsgText(154), , MsgText(5)
         Cancel = True
      End If
      adocheck.Close
   End If
End Sub

Private Sub Text9_Change()
   Text10 = A0g02Query(Text9)
End Sub

Private Sub Text9_GotFocus()
   TextInverse Text9
End Sub

Private Sub Text9_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

'*************************************************
'  重新整理 Adodc 之資料
'
'*************************************************
Public Sub AdodcRefresh()
Dim strUnion As String

On Error GoTo Checking
   adoadodc1.Close
   adoadodc1.CursorLocation = adUseClient
'   strUnion = "select a0e01, a0e02, a0e10, a0e11, a0g02, a0e07, a0e08 from acc0e0, acc0g0 where a0e01 = a0g01 and a0e19 = '" & Text9 & "' and a0e20 = '" & combo1 & "' and a0e14 = " & Val(FCDate(MaskEdBox1.Text)) & " and a0e25 = 0"
   'Modify by Amy 2023/05/23 原:Combo1
   strExc(1) = ""
   If Combo1 <> MsgText(601) Then strExc(1) = Left(Combo1.Text, InStr(Combo1, " ") - 1)
   strUnion = "select a0e01, a0e02, a0e10, a0e11, a0g02, a0e07, a0e08, cu04 as contect, a0e14, a0e47 from acc0e0, acc0g0, customer where a0e01 = a0g01 and a0e05 = '1' and substr(a0e06, 1, 8) = cu01 (+) and substr(a0e06, 9, 1) = cu02 (+) and a0e19 = '" & Text9 & "' and a0e20 = '" & strExc(1) & "' and a0e14 = " & Val(FCDate(MaskEdBox1.Text)) & " and a0e25 = 0"
   strUnion = strUnion & " union select a0e01, a0e02, a0e10, a0e11, a0g02, a0e07, a0e08, a0i02 as contect, a0e14, a0e47 from acc0e0, acc0g0, acc0i0 where a0e01 = a0g01 and a0e05 = '2' and a0e06 = a0i01 (+) and a0e19 = '" & Text9 & "' and a0e20 = '" & strExc(1) & "' and a0e14 = " & Val(FCDate(MaskEdBox1.Text)) & " and a0e25 = 0"
   strUnion = strUnion & " union select a0e01, a0e02, a0e10, a0e11, a0g02, a0e07, a0e08, st02 as contect, a0e14, a0e47 from acc0e0, acc0g0, staff where a0e01 = a0g01 and a0e05 = '3' and a0e06 = st01 (+) and a0e19 = '" & Text9 & "' and a0e20 = '" & strExc(1) & "' and a0e14 = " & Val(FCDate(MaskEdBox1.Text)) & " and a0e25 = 0"
   strUnion = strUnion & " union select a0e01, a0e02, a0e10, a0e11, a0g02, a0e07, a0e08, '' as contect, a0e14, a0e47 from acc0e0, acc0g0 where a0e01 = a0g01 and a0e05 = '4' and a0e19 = '" & Text9 & "' and a0e20 = '" & strExc(1) & "' and a0e14 = " & Val(FCDate(MaskEdBox1.Text)) & " and a0e25 = 0 order by a0e14 asc, a0e47 desc"
   'end 2023/05/23
   adoadodc1.Open strUnion, adoTaie, adOpenDynamic, adLockBatchOptimistic
   Adodc1.Recordset.Requery
   SumShow
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)
   Select Case KeyCode
      Case vbKeyF12
         AdodcRefresh
      Case vbKeyInsert
         'Modify by Morgan2004/10/29 存檔成功才清資料
         'Acc0e0Save
         If Acc0e0Save = True Then
            Text5 = MsgText(601)
            Text11 = MsgText(601)
            Text12 = MsgText(601)
            Text5.SetFocus
         End If
   End Select
   KeyEnter KeyCode
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(133)
End Sub

'*************************************************
'  儲存資料表(票據資料)
'
'*************************************************
'Modify by Morgan 2004/10/29
'Private Sub Acc0e0Save()
Private Function Acc0e0Save() As Boolean

On Error GoTo Checking
   
   If Text11 = "" Then 'Add by Morgan 2004/10/29
      If adocheck.State = adStateOpen Then adocheck.Close
      adocheck.CursorLocation = adUseClient
      'Modify by Morgan 2004/10/29  加應收過濾條件 and a0e04='R'
      'adocheck.Open "select a0e01, a0e02 from acc0e0 where a0e02 = '" & Text5 & "'", adoTaie, adOpenStatic, adLockReadOnly
      adocheck.Open "select a0e01 from acc0e0 where a0e02 = '" & Text5 & "' and a0e04='R' and (a0e34 = 0 or a0e34 is null) and a0e21 = 0 and a0e15 = 0", adoTaie, adOpenStatic, adLockReadOnly
      
      If adocheck.RecordCount <> 0 Then
         'Add by Morgan 2004/10/29
         If adocheck.RecordCount > 1 Then
            MsgBox "收票銀行無法確定，請自行輸入！【 " & adocheck.GetString(, , , ";") & "】" 'Add by Morgan 2004/10/29
            adocheck.Close
            Text11.SetFocus
            Exit Function
         End If
         '2004/10/29
         If IsNull(adocheck.Fields(0).Value) Then
            Text11 = MsgText(601)
         Else
            Text11 = adocheck.Fields(0).Value
         End If
      Else
         Text11 = MsgText(601)
         'Add by Morgan 2004/11/16
         adocheck.Close
         MsgBox MsgText(154), , MsgText(5)
         Text5.SetFocus
         Exit Function
         '2004/11/16
      End If
      
      If adocheck.State = adStateOpen Then adocheck.Close
   End If
   
   If Text5 = MsgText(601) Then
      MsgBox MsgText(10) & Label5, , MsgText(5)
      strControlButton = MsgText(602)
      Text5.SetFocus
      Exit Function
   Else
      If Text11 = MsgText(601) Then
         MsgBox MsgText(10) & Label9, , MsgText(5)
         strControlButton = MsgText(602)
         Text11.SetFocus
         Exit Function
      End If
      If ExistCheck("acc0g0", "a0g01", Text9, Label1) = False Then
         strControlButton = MsgText(602)
         Text9.SetFocus
         Exit Function
      End If
      If ExistCheck("acc0g0", "a0g01", Text11, Label9) = False Then
         strControlButton = MsgText(602)
         Text11.SetFocus
         Exit Function
      End If
      If MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Then
         MsgBox Label2 & MsgText(52), , MsgText(5)
         strControlButton = MsgText(602)
         MaskEdBox1.SetFocus
         Exit Function
      Else
         If DateCheck(MaskEdBox1.Text) = MsgText(603) Then
            MsgBox Label2 & MsgText(63), , MsgText(5)
            strControlButton = MsgText(602)
            MaskEdBox1.SetFocus
            Exit Function
         End If
      End If
   End If
   adoacc0e0.CursorLocation = adUseClient
   'Modify by Morgan 2004/10/29  加應收過濾條件 and a0e04='R'
   adoacc0e0.Open "select * from acc0e0 where a0e01 = '" & Text11 & "' and a0e02 = '" & Text5 & "' and (a0e34 = 0 or a0e34 is null) and a0e21 = 0 and a0e15 = 0 and a0e04='R'", adoTaie, adOpenDynamic, adLockBatchOptimistic
   If adoacc0e0.RecordCount = 0 Then
      MsgBox MsgText(154), , MsgText(5)
      strControlButton = MsgText(602)
      adoacc0e0.Close
      Text11.SetFocus
      Exit Function
   Else
      '2014/1/23 modify by sonia 加判斷公司別a0e23
      If adoacc0e0.Fields("a0e23").Value <> strA0E23 Then
         MsgBox "票據公司別與託收銀行帳號不合！", , MsgText(5)
         strControlButton = MsgText(602)
         adoacc0e0.Close
         Text5.SetFocus
         Exit Function
      End If
      '2014/1/23 end
      If adoacc0e0.Fields("a0e16").Value = 0 And adoacc0e0.Fields("a0e14").Value <> 0 Then
         MsgBox MsgText(9), , MsgText(5)
         strControlButton = MsgText(602)
         adoacc0e0.Close
         Text11.SetFocus
         Exit Function
      End If
      If adoacc0e0.Fields("a0e14") <> 0 Or adoacc0e0.Fields("a0e15").Value <> 0 Or adoacc0e0.Fields("a0e17").Value <> 0 Or adoacc0e0.Fields("a0e21").Value <> 0 Or adoacc0e0.Fields("a0e25").Value <> 0 Or adoacc0e0.Fields("a0e34").Value <> 0 Then
         MsgBox MsgText(60), , MsgText(5)
         strControlButton = MsgText(602)
         Text5.SetFocus
         Exit Function
      End If
   End If
   If Text9 <> MsgText(601) Then
      adoacc0e0.Fields("a0e19").Value = Text9
   Else
      adoacc0e0.Fields("a0e19").Value = Null
   End If
   If Combo1 <> MsgText(601) Then
      adoacc0e0.Fields("a0e20").Value = Left(Combo1.Text, InStr(Combo1, " ") - 1)  'Modify by Amy 2023/05/23 原:Combo1
   Else
      adoacc0e0.Fields("a0e20").Value = Null
   End If
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      adoacc0e0.Fields("a0e14").Value = Val(FCDate(MaskEdBox1.Text))
   Else
      adoacc0e0.Fields("a0e14").Value = 0
   End If
   adoacc0e0.Fields("a0e16").Value = 0
   adoacc0e0.Fields("a0e29").Value = Val(strSrvDate(2))
   adoacc0e0.Fields("a0e30").Value = ServerTime
   adoacc0e0.Fields("a0e31").Value = strUserNum
   adoacc0e0.Fields("a0e47").Value = ServerTime
   adoacc0e0.UpdateBatch
   Acc0e0Save = True 'Add by Morgan 2004/10/29
   AdodcRefresh
   adoacc0e0.Close
Checking:
   If Err.Number = 0 Then
      Exit Function
   End If
   MsgBox Err.Description, , MsgText(5)
End Function

'*************************************************
'  刪除 Adodc 之資料
'
'*************************************************
Private Sub AdodcDelete()
On Error GoTo Checking
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   adoTaie.Execute "update acc0e0 set a0e14 = 0, a0e19 = null, a0e20 = null where a0e01 = '" & Adodc1.Recordset.Fields("a0e01").Value & "' and a0e02 = '" & Adodc1.Recordset.Fields("a0e02").Value & "' and a0e21 = 0"
   AdodcRefresh
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

Private Sub Text9_Validate(Cancel As Boolean)
   If ExistCheck("acc0g0", "a0g01", Text9, Label1) = False Then
      Cancel = True
      Exit Sub
   End If
End Sub

'*************************************************
'  筆數及金額合計
'
'*************************************************
Private Sub SumShow()
Dim adoaccsum As New ADODB.Recordset

   adoaccsum.CursorLocation = adUseClient
   '2014/1/23 modify by sonia 加判斷公司別a0e23
   'adoaccsum.Open "select count(a0e02), sum(a0e11) from acc0e0 where a0e19 = '" & Text9 & "' and a0e20 = '" & Combo1 & "' and a0e14 = " & Val(FCDate(MaskEdBox1.Text)) & " and a0e25 = 0", adoTaie, adOpenStatic, adLockReadOnly
   'Modify by Amy 2023/05/23 原:Combo1
   strExc(1) = ""
   If Combo1 <> MsgText(601) Then strExc(1) = Left(Combo1.Text, InStr(Combo1, " ") - 1)
   adoaccsum.Open "select count(a0e02), sum(a0e11) from acc0e0 where a0e23='" & strA0E23 & "' and a0e19 = '" & Text9 & "' and a0e20 = '" & strExc(1) & "' and a0e14 = " & Val(FCDate(MaskEdBox1.Text)) & " and a0e25 = 0", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0).Value) Then
         Text2 = MsgText(601)
      Else
         Text2 = Format(adoaccsum.Fields(0).Value, DDollar)
      End If
      If IsNull(adoaccsum.Fields(1).Value) Then
         Text3 = MsgText(601)
      Else
         Text3 = Format(adoaccsum.Fields(1).Value, DDollar)
      End If
   Else
      Text2 = MsgText(601)
      Text3 = MsgText(601)
   End If
   adoaccsum.Close
End Sub



