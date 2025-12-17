VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc2140 
   AutoRedraw      =   -1  'True
   Caption         =   "銷帳作業"
   ClientHeight    =   5448
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5448
   ScaleWidth      =   8760
   Begin VB.TextBox Text7 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3570
      TabIndex        =   11
      Top             =   615
      Width           =   585
   End
   Begin VB.CommandButton Command3 
      Height          =   300
      Left            =   2640
      Picture         =   "Frmacc2140.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   255
      Width           =   350
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc2140.frx":0102
      Height          =   3495
      Left            =   240
      TabIndex        =   5
      Top             =   1710
      Width           =   8295
      _ExtentX        =   14626
      _ExtentY        =   6160
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   17
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "a1401"
         Caption         =   "銷帳單號"
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
         DataField       =   "a1402"
         Caption         =   "銷帳日期"
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
         DataField       =   "a1403"
         Caption         =   "請款編號"
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
         DataField       =   "a1411"
         Caption         =   "本所案號"
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
         DataField       =   "a1k18"
         Caption         =   "幣別"
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
         DataField       =   "a1412"
         Caption         =   "請款外幣"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "a1413"
         Caption         =   "代理人"
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
         DataField       =   "a1404"
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
         BeginProperty Column00 
            ColumnWidth     =   1247.811
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   1068.094
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1344.189
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   552.189
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   1235.906
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1272.189
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   3960
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4170
      TabIndex        =   12
      Top             =   615
      Width           =   1605
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1320
      TabIndex        =   10
      Top             =   615
      Width           =   1572
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6840
      MaxLength       =   15
      TabIndex        =   3
      Top             =   240
      Width           =   1572
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1320
      MaxLength       =   15
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   330
      Left            =   4170
      TabIndex        =   2
      Top             =   240
      Width           =   1605
      _ExtentX        =   2815
      _ExtentY        =   593
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   312
      Left            =   240
      Top             =   1320
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   2117
      _ExtentY        =   572
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
   Begin MSForms.TextBox Text6 
      Height          =   495
      Left            =   1320
      TabIndex        =   4
      Top             =   990
      Width           =   7155
      VariousPropertyBits=   -1467989989
      ScrollBars      =   2
      Size            =   "12621;873"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text5 
      Height          =   330
      Left            =   6840
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   615
      Width           =   1575
      VariousPropertyBits=   671105055
      BackColor       =   14737632
      MaxLength       =   50
      Size            =   "2778;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   4920
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   1455
      Left            =   240
      Top             =   120
      Width           =   8295
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "備註"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   360
      TabIndex        =   15
      Top             =   990
      Width           =   972
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "代理人"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   5880
      TabIndex        =   14
      Top             =   654
      Width           =   972
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "請款"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   3120
      TabIndex        =   13
      Top             =   654
      Width           =   972
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "本所案號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   360
      TabIndex        =   9
      Top             =   654
      Width           =   972
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "請款編號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   5880
      TabIndex        =   8
      Top             =   279
      Width           =   972
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "銷帳日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   3120
      TabIndex        =   7
      Top             =   279
      Width           =   972
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "銷帳單號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   360
      TabIndex        =   6
      Top             =   279
      Width           =   972
   End
End
Attribute VB_Name = "Frmacc2140"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/03 改成Form2.0 ; DataGrid1改字型=新細明體-ExtB、Text5、Text6
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit
Public adoacc1k0 As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Public strDocNo As String
Public m_Old_a1k01 As String 'Add By Sindy 2010/12/7

Private Sub Command3_Click()
   If Adodc1.Recordset.RecordCount = 0 Or Text2 = MsgText(601) Then
      Exit Sub
   End If
   Adodc1.Recordset.Find "a1401 = '" & Text2 & "'", 0, adSearchForward, 1
   If Adodc1.Recordset.EOF = False Then
      FormShow
      AdodcRefresh
      RecordShow
   Else
      MsgBox MsgText(33), , MsgText(5)
      Adodc1.Recordset.MoveFirst
   End If
End Sub

Private Sub Command3_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Command3_Click
         Exit Sub
   End Select
   KeyEnter KeyCode
End Sub

Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub DataGrid1_SelChange(Cancel As Integer)
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   FormShow
End Sub

Private Sub Form_Activate()
   strFormName = Name
   If strItemNo = MsgText(601) Then
      Exit Sub
   End If
   If Adodc1.Recordset.RecordCount <> 0 Then
      Adodc1.Recordset.MoveFirst
   End If
   Adodc1.Recordset.Find "a1401 = '" & strItemNo & "'", 0, adSearchForward, 1
   If Adodc1.Recordset.EOF = False Then
      FormShow
      RecordShow
   End If
   strItemNo = MsgText(601)
End Sub

'Added by Lydia 2021/12/03
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call PUB_SaveTrackMode(0, KeyCode)  'Added by Lydia 2021/12/03 Form2.0 記錄鍵盤傳入順序
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   Call PUB_SaveTrackMode(1, KeyCode)  'Added by Lydia 2021/12/03 Form2.0 記錄鍵盤傳入順序
   
   KeyEnter KeyCode
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   'Modified by Lydia 2021/12/07 改成模組
'   Me.Icon = LoadPicture(strIcoPath)
'   strFormName = Name
'   Me.Width = 8850
'   Me.Height = 5500
'   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
'   Image1 = LoadPicture(strBackPicPath1)
'   sglWidth = Image1.Width
'   sglHeight = Image1.Height
'   For intX = 0 To Int(ScaleWidth / sglWidth)
'       For intY = 0 To Int(ScaleHeight / sglHeight)
'           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
'       Next
'   Next
   strFormName = Name
   'Modify by Amy 2023/08/18 H5700
   PUB_InitForm Me, 8850, 5890, strBackPicPath1
   'end 2021/12/07
   
   OpenTable
   MaskEdBox1.Mask = DFormat
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
   strFormName = MsgText(601)
   strTrackMode = "" 'Added by Lydia 2021/12/03 Form2.0 記錄鍵盤傳入順序(清除)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc2140 = Nothing
End Sub

Private Sub MaskEdBox1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "select * from acc140,acc1k0 where a1403=a1k01(+) order by a1401 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  重新整理 Adodc 之資料
'
'*************************************************
Public Sub AdodcRefresh()
On Error GoTo Checking
   adoadodc1.Close
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "select * from acc140,acc1k0 where a1403=a1k01(+) order by a1401 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Adodc1.Recordset.Requery
   If Adodc1.Recordset.RecordCount <> 0 Then
      Adodc1.Recordset.Find "a1401 = '" & Text2 & "'", 0, adSearchForward, 1
      If Adodc1.Recordset.EOF = False Then
         FormShow
         RecordShow
      End If
   End If
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  顯示資料表
'
'*************************************************
Public Sub FormShow()
   Text2 = Adodc1.Recordset.Fields("a1401").Value
   MaskEdBox1.Mask = MsgText(601)
   If IsNull(Adodc1.Recordset.Fields("a1402").Value) Then
      MaskEdBox1.Text = MsgText(601)
   Else
      MaskEdBox1.Text = CFDate(Adodc1.Recordset.Fields("a1402").Value)
   End If
   MaskEdBox1.Mask = DFormat
   m_Old_a1k01 = "" 'Add By Sindy 2010/12/7
   If IsNull(Adodc1.Recordset.Fields("a1403").Value) Then
      Text1 = MsgText(601)
   Else
      Text1 = Adodc1.Recordset.Fields("a1403").Value
      m_Old_a1k01 = Text1 'Add By Sindy 2010/12/7
   End If
   If IsNull(Adodc1.Recordset.Fields("a1404").Value) Then
      Text6 = MsgText(601)
   Else
      Text6 = Adodc1.Recordset.Fields("a1404").Value
   End If
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
   If strSaveConfirm <> MsgText(3) Then
      Exit Sub
   End If
   If Mid(MaskEdBox1.Text, 1, 3) <> Mid(CFDate(ACDate(ServerDate)), 1, 3) Then
      Text2 = UpdateNo("acc140", "a1401", 5, MaskEdBox1.Text, MsgText(811))
   Else
      'Text2 = AutoNo(MsgText(811), 5)
      Text2 = strDocNo
   End If
End Sub

Private Sub Text1_Change()
   If Text1 = MsgText(601) Then
      Exit Sub
   End If
   Acc1k0Query
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Text6_GotFocus()
   TextInverse Text6
   'edit by nickc 2007/06/11  切換輸入法改用API
   OpenIme
   
End Sub

'Modified by Lydia 2021/12/03 改成Form 2.0; KeyCode As Integer=>MSForms.ReturnInteger
Private Sub Text6_KeyUp(KeyCode As MSForms.ReturnInteger, Shift As Integer)
   'Modified by Lydia 2021/12/03 +val()
   KeyEnter Val(KeyCode)
End Sub

'*************************************************
'  顯示筆數
'
'*************************************************
Public Sub RecordShow()
   Frmacc0000.StatusBar1.Panels(2).Text = Adodc1.Recordset.Bookmark & MsgText(35) & Adodc1.Recordset.RecordCount
End Sub

'*************************************************
'  顯示查詢資料
'
'*************************************************
Private Sub Acc1k0Query()
   adoacc1k0.CursorLocation = adUseClient
   '2006/7/18 MODIFY BY SONIA
   'adoacc1k0.Open "select * from acc1k0 where a1k01 = '" & Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
   If strSaveConfirm = MsgText(3) Then
      '2007/11/5 modify by sonia 收款結清後不可銷帳 X09506022
      'adoacc1k0.Open "select * from acc1k0 where A1K25 IS NULL AND A1K12 IS NULL AND a1k01 = '" & Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
      adoacc1k0.Open "select * from acc1k0 where A1K25 IS NULL AND A1K12 IS NULL and a1k29 is null AND a1k01 = '" & Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
   Else
      adoacc1k0.Open "select * from acc1k0 where a1k01 = '" & Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
   End If
   '2006/7/18 END
   If adoacc1k0.RecordCount <> 0 Then
      If IsNull(adoacc1k0.Fields("a1k13").Value) Then
         Text3 = MsgText(601)
      Else
         Text3 = adoacc1k0.Fields("a1k13").Value
         If IsNull(adoacc1k0.Fields("a1k14").Value) = False Then
            Text3 = Text3 & adoacc1k0.Fields("a1k14").Value
         End If
         If IsNull(adoacc1k0.Fields("a1k15").Value) = False Then
            Text3 = Text3 & adoacc1k0.Fields("a1k15").Value
         End If
         If IsNull(adoacc1k0.Fields("a1k16").Value) = False Then
            Text3 = Text3 & adoacc1k0.Fields("a1k16").Value
         End If
      End If
      If IsNull(adoacc1k0.Fields("a1k08").Value) Then
         Text4 = MsgText(601)
      Else
         'Modify By Sindy 2013/1/14
         'Text4 = Format(adoacc1k0.Fields("a1k08").Value, FAmount)
         If IsNull(adoacc1k0.Fields("a1k31").Value) Then
            Text4 = Format(adoacc1k0.Fields("a1k08").Value, FAmount)
         Else
            Text4 = Format(adoacc1k0.Fields("a1k08").Value - adoacc1k0.Fields("a1k31").Value, FAmount)
         End If
         '2013/1/14 End
      End If
      'Add By Sindy 2012/12/7
      If IsNull(adoacc1k0.Fields("a1k18").Value) Then
         Text7 = MsgText(601)
      Else
         Text7 = adoacc1k0.Fields("a1k18").Value
      End If
      '2012/12/7 End
      If IsNull(adoacc1k0.Fields("a1k03").Value) Then
         Text5 = MsgText(601)
      Else
         Text5 = adoacc1k0.Fields("a1k03").Value
      End If
   Else
     Text3 = MsgText(601)
     Text4 = MsgText(601)
     Text5 = MsgText(601)
     Text7 = MsgText(601) 'Add By Sindy 2012/12/7
   End If
   adoacc1k0.Close
End Sub

Private Sub Text6_LostFocus()
'edit by nickc 2007/06/11  切換輸入法改用API
CloseIme
End Sub
