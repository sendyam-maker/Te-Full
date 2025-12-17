VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc21k0 
   AutoRedraw      =   -1  'True
   Caption         =   "請款單作廢作業"
   ClientHeight    =   5436
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5436
   ScaleWidth      =   8760
   Begin VB.CommandButton Command5 
      Height          =   300
      Left            =   3360
      Picture         =   "Frmacc21k0.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   255
      Width           =   350
   End
   Begin VB.TextBox Text9 
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
      Left            =   3360
      TabIndex        =   17
      Top             =   600
      Width           =   372
   End
   Begin VB.TextBox Text8 
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
      Left            =   3120
      TabIndex        =   16
      Top             =   600
      Width           =   252
   End
   Begin VB.TextBox Text7 
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
      Left            =   2280
      TabIndex        =   15
      Top             =   600
      Width           =   852
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc21k0.frx":0102
      Height          =   2190
      Left            =   240
      TabIndex        =   4
      Top             =   2970
      Width           =   8295
      _ExtentX        =   14626
      _ExtentY        =   3874
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
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "cp09"
         Caption         =   "總收文號"
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
         DataField       =   "Rdate"
         Caption         =   "收文日期"
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
         DataField       =   "cpm03"
         Caption         =   "案件性質"
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
         DataField       =   "Fname"
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
      BeginProperty Column04 
         DataField       =   "Ono"
         Caption         =   "彼所案號"
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
         DataField       =   "Sdate"
         Caption         =   "發文日期"
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
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   1476.284
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   1260.284
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   972.284
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1344.189
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1476.284
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnWidth     =   1272.189
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text6 
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
      Left            =   4920
      TabIndex        =   14
      Top             =   600
      Width           =   1572
   End
   Begin VB.TextBox Text1 
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
      Left            =   1800
      TabIndex        =   5
      Top             =   600
      Width           =   492
   End
   Begin VB.TextBox Text5 
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
      Left            =   1800
      MaxLength       =   15
      TabIndex        =   0
      Top             =   240
      Width           =   1572
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   330
      Left            =   4920
      TabIndex        =   2
      Top             =   240
      Width           =   1572
      _ExtentX        =   2773
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
      Height          =   315
      Left            =   210
      Top             =   2520
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
   Begin MSForms.TextBox Text4 
      Height          =   330
      Left            =   1800
      TabIndex        =   21
      Top             =   1680
      Width           =   6615
      VariousPropertyBits=   671105055
      BackColor       =   14737632
      Size            =   "11668;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text3 
      Height          =   330
      Left            =   1800
      TabIndex        =   20
      Top             =   1320
      Width           =   6615
      VariousPropertyBits=   671105055
      BackColor       =   14737632
      Size            =   "11668;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text2 
      Height          =   330
      Left            =   1800
      TabIndex        =   19
      Top             =   960
      Width           =   6615
      VariousPropertyBits=   671105055
      BackColor       =   14737632
      Size            =   "11668;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text10 
      Height          =   645
      Left            =   1800
      TabIndex        =   3
      Top             =   2040
      Width           =   6615
      VariousPropertyBits=   -1467989989
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "11668;1138"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "注意事項"
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
      TabIndex        =   18
      Top             =   2040
      Width           =   972
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "請款金額"
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
      Left            =   3960
      TabIndex        =   13
      Top             =   639
      Width           =   972
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "作廢日期"
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
      Left            =   3960
      TabIndex        =   12
      Top             =   279
      Width           =   972
   End
   Begin VB.Label Label1 
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
      TabIndex        =   11
      Top             =   639
      Width           =   972
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "案件名稱"
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
      TabIndex        =   10
      Top             =   999
      Width           =   972
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "(中)"
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
      Left            =   1320
      TabIndex        =   9
      Top             =   960
      Width           =   492
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "(英)"
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
      Left            =   1320
      TabIndex        =   8
      Top             =   1359
      Width           =   492
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "(日)"
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
      Left            =   1320
      TabIndex        =   7
      Top             =   1719
      Width           =   492
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
      Height          =   2715
      Left            =   240
      Top             =   120
      Width           =   8295
   End
   Begin VB.Label Label6 
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
      Left            =   360
      TabIndex        =   6
      Top             =   279
      Width           =   972
   End
End
Attribute VB_Name = "Frmacc21k0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/01 改成Form2.0 ; Text2、Text3、Text4、Text10、DataGrid1改字型=新細明體-ExtB
'Memo By Sonia 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit
Public adoacc1k0 As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset

Private Sub Command5_Click()
   If adoacc1k0.RecordCount = 0 Or Text5 = MsgText(601) Then
      Exit Sub
   End If
   adoacc1k0.Find "a1k01 = '" & Text5 & "'", 0, adSearchForward, 1
   If adoacc1k0.EOF = False Then
      FormShow
      AdodcRefresh
      RecordShow
   Else
      MsgBox MsgText(33), , MsgText(5)
      adoacc1k0.MoveFirst
   End If
End Sub

Private Sub Command5_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Command5_Click
         Exit Sub
   End Select
   KeyEnter KeyCode
End Sub

Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Form_Activate()
'edit by nickc 2007/02/08
'   '93.3.16 ADD BY SONIA
'   If IsObject(mdiMain) Then
'      mdiMain.toolshow
'   End If
'   '93.3.16 END
   Dim formCnt As Integer
   For formCnt = 0 To Forms.Count - 1
       If UCase(Forms(formCnt).Name) = "MDIMAIN" Then
             Forms(formCnt).ToolShow
             Exit For
       End If
   Next
   strFormName = Name
   If strItemNo = MsgText(601) Then
      Exit Sub
   End If
   If adoacc1k0.RecordCount <> 0 Then
      adoacc1k0.MoveFirst
   End If
   adoacc1k0.Find "a1k01 = '" & strItemNo & "'", 0, adSearchForward, 1
   If adoacc1k0.EOF = False Then
      FormShow
      AdodcRefresh
      RecordShow
   End If
   strItemNo = MsgText(601)
End Sub

'Added by Lydia 2021/12/08
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call PUB_SaveTrackMode(0, KeyCode)  'Added by Lydia 2021/12/08 Form2.0 記錄鍵盤傳入順序
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   Call PUB_SaveTrackMode(1, KeyCode)  'Added by Lydia 2021/12/08 Form2.0 記錄鍵盤傳入順序
   
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
   PUB_InitForm Me, 8850, 5880, strBackPicPath1
   'end 2021/12/07
   
   MaskEdBox2.Mask = DFormat
   OpenTable
   If adoacc1k0.RecordCount <> 0 Then
      adoacc1k0.MoveLast
      adoacc1k0.MoveFirst
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
   strTrackMode = "" 'Added by Lydia 2021/12/08 Form2.0 記錄鍵盤傳入順序(清除)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc21k0 = Nothing
End Sub

Private Sub MaskEdBox2_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub MaskEdBox2_Validate(Cancel As Boolean)
   If MaskEdBox2.Text = MsgText(601) Or MaskEdBox2.Text = MsgText(29) Then
      MsgBox Label7 & MsgText(52), , MsgText(5)
      Cancel = True
      MaskEdBox2.SetFocus
      Exit Sub
   End If
   If DateCheck(MaskEdBox2.Text) = MsgText(603) Then
      MsgBox Label7 & MsgText(63), , MsgText(5)
      Cancel = True
      MaskEdBox2.SetFocus
      Exit Sub
   End If
End Sub

Private Sub Text1_Change()
   CaseQuery
End Sub

'Add By Sindy 2009/07/15
Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

'Add By Sindy 2009/07/15
Public Sub Text1_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(Text1) = False Then
      ' 檢查系統類別
      If IsCorrectSysKind(Text1) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "本所案號中的系統別不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         Text1_GotFocus
         GoTo EXITSUB
      End If
      ' 檢查使用者權限
      If IsUserHasRightOfSystem(strUserNum, Text1) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "您沒有使用該系統類別的權限"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         Text1_GotFocus
         GoTo EXITSUB
      End If
   Else
'      Cancel = True
'      strTit = "資料檢核"
'      strMsg = "本所案號中的系統別不可空白"
'      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'      Text1_GotFocus
'      GoTo EXITSUB
   End If
EXITSUB:
End Sub

Private Sub Text10_GotFocus()
   TextInverse Text10
   'add by nickc 2007/07/13 將輸入法改成使用API
   OpenIme
End Sub

'Modified by Lydia 2021/12/01 改成Form 2.0; KeyCode As Integer=>MSForms.ReturnInteger
Private Sub Text10_KeyUp(KeyCode As MSForms.ReturnInteger, Shift As Integer)
   'Modified by Lydia 2021/12/01 + val
   KeyEnter Val(KeyCode)
End Sub
'add by nickc 2007/07/13 將輸入法改成使用API
Private Sub Text10_Validate(Cancel As Boolean)
CloseIme
End Sub

Private Sub Text5_GotFocus()
   CloseIme 'Add by Morgan 2008/7/2
   TextInverse Text5
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
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
   adoacc1k0.CursorLocation = adUseClient
   adoacc1k0.Open "select * from acc1k0 where a1k12 > 0 order by a1k01 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "select * from caseprogress where cp60 = 'a' order by cp09 asc", adoTaie, adOpenStatic, adLockReadOnly
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
   adoacc1k0.Close
   adoacc1k0.CursorLocation = adUseClient
   adoacc1k0.Open "select * from acc1k0 where a1k12 > 0 order by a1k01 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   'adoacc1k0.Open "select * from acc1k0 where a1k12 is not null", adoTaie, adOpenDynamic, adLockBatchOptimistic
   If adoacc1k0.RecordCount <> 0 Then
      If Text5 <> MsgText(601) Then
         adoacc1k0.Find "a1k01 = '" & Text5 & "'", 0, adSearchForward, 1
      End If
   End If
   adoadodc1.Close
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "select cp09, cp05 - 19110000 as Rdate, cpm03, nvl(fa04, nvl(fa05, fa06)) as Fname, pa77 as Ono, cp27 - 19110000 as Sdate from caseprogress, acc1w0, casepropertymap, patent, fagent where cp09 = a1w02 and cp01 = cpm01 and cp10 = cpm02 and cp01 = pa01 and cp02 = pa02 and cp03 = pa03 and cp04 = pa04 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and a1w01 = '" & Text5 & "' union " & _
                  "select cp09, cp05 - 19110000 as Rdate, cpm03, nvl(fa04, nvl(fa05, fa06)) as Fname, tm45 as Ono, cp27 - 19110000 as Sdate from caseprogress, acc1w0, casepropertymap, trademark, fagent where cp09 = a1w02 and cp01 = cpm01 and cp10 = cpm02 and cp01 = tm01 and cp02 = tm02 and cp03 = tm03 and cp04 = tm04 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and a1w01 = '" & Text5 & "' union " & _
                  "select cp09, cp05 - 19110000 as Rdate, cpm03, nvl(fa04, nvl(fa05, fa06)) as Fname, lc23 as Ono, cp27 - 19110000 as Sdate from caseprogress, acc1w0, casepropertymap, lawcase, fagent where cp09 = a1w02 and cp01 = cpm01 and cp10 = cpm02 and cp01 = lc01 and cp02 = lc02 and cp03 = lc03 and cp04 = lc04 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and a1w01 = '" & Text5 & "' union " & _
                  "select cp09, cp05 - 19110000 as Rdate, cpm03, nvl(fa04, nvl(fa05, fa06)) as Fname, sp27 as Ono, cp27 - 19110000 as Sdate from caseprogress, acc1w0, casepropertymap, servicepractice, fagent where cp09 = a1w02 and cp01 = cpm01 and cp10 = cpm02 and cp01 = sp01 and cp02 = sp02 and cp03 = sp03 and cp04 = sp04 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and a1w01 = '" & Text5 & "' order by cp09 asc", adoTaie, adOpenStatic, adLockReadOnly
   Adodc1.Recordset.Requery
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  顯示查詢資料(案件基本資料)
'
'*************************************************
Private Sub CaseQuery()
   If Text1 = MsgText(601) Or Text7 = MsgText(601) Or Text8 = MsgText(601) Or Text9 = MsgText(601) Then
      Exit Sub
   End If
   Text2 = CaseNameShow(Text1, Text7, Text8, Text9, 1)
   Text3 = CaseNameShow(Text1, Text7, Text8, Text9, 2)
   Text4 = CaseNameShow(Text1, Text7, Text8, Text9, 3)
End Sub

'*************************************************
'  顯示資料表
'
'*************************************************
Public Sub FormShow()
   Text5 = adoacc1k0.Fields("a1k01").Value
   MaskEdBox2.Mask = MsgText(601)
   If IsNull(adoacc1k0.Fields("a1k12").Value) Then
      MaskEdBox2.Text = MsgText(601)
   Else
      MaskEdBox2.Text = CFDate(adoacc1k0.Fields("a1k12").Value)
   End If
   MaskEdBox2.Mask = DFormat
   If IsNull(adoacc1k0.Fields("a1k13").Value) Then
      Text1 = MsgText(601)
   Else
      Text1 = adoacc1k0.Fields("a1k13").Value
   End If
   If IsNull(adoacc1k0.Fields("a1k14").Value) Then
      Text7 = MsgText(601)
   Else
      Text7 = adoacc1k0.Fields("a1k14").Value
   End If
   If IsNull(adoacc1k0.Fields("a1k15").Value) Then
      Text8 = MsgText(601)
   Else
      Text8 = adoacc1k0.Fields("a1k15").Value
   End If
   If IsNull(adoacc1k0.Fields("a1k16").Value) Then
      Text9 = MsgText(601)
   Else
      Text9 = adoacc1k0.Fields("a1k16").Value
   End If
   If IsNull(adoacc1k0.Fields("a1k11").Value) Then
      Text6 = MsgText(601)
   Else
      Text6 = adoacc1k0.Fields("a1k11").Value
   End If
   '2013/10/22 modify by sonia a1k05改為a1k34
   If IsNull(adoacc1k0.Fields("a1k34").Value) Then
      Text10 = MsgText(601)
   Else
      Text10 = adoacc1k0.Fields("a1k34").Value
   End If
End Sub

'*************************************************
'  顯示筆數
'
'*************************************************
Public Sub RecordShow()
   If adoacc1k0.RecordCount = 0 Then
      Exit Sub
   End If
   CountShow adoacc1k0.Bookmark, adoacc1k0.RecordCount
End Sub

Private Sub Text5_Validate(Cancel As Boolean)
   If strSaveConfirm <> MsgText(3) Then
      Exit Sub
   End If
   If Text5 = MsgText(601) Then
      Exit Sub
   End If
   Acc1k0Query
   AdodcRefresh
End Sub

Private Sub Text7_Change()
   CaseQuery
End Sub

Private Sub Text8_Change()
   CaseQuery
End Sub

Private Sub Text9_Change()
   CaseQuery
End Sub

'*************************************************
'  顯示查詢資料(國外請款單資料(主檔))
'
'*************************************************
Private Sub Acc1k0Query()
   adoquery.CursorLocation = adUseClient
   adoquery.Open "select * from acc1k0 where a1k01 = '" & Text5 & "' and (a1k30 is null or a1k30 = 0) and a1k12 is null", adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
'      MaskEdBox2.Mask = MsgText(601)
'      If IsNull(adoquery.Fields("a1k12").Value) Then
'         MaskEdBox2.Text = MsgText(601)
'      Else
'         MaskEdBox2.Text = adoquery.Fields("a1k12").Value
'      End If
'      MaskEdBox2.Mask = DFormat
      If IsNull(adoquery.Fields("a1k13").Value) Then
         Text1 = MsgText(601)
      Else
         Text1 = adoquery.Fields("a1k13").Value
      End If
      If IsNull(adoquery.Fields("a1k14").Value) Then
         Text7 = MsgText(601)
      Else
         Text7 = adoquery.Fields("a1k14").Value
      End If
      If IsNull(adoquery.Fields("a1k15").Value) Then
         Text8 = MsgText(601)
      Else
         Text8 = adoquery.Fields("a1k15").Value
      End If
      If IsNull(adoquery.Fields("a1k16").Value) Then
         Text9 = MsgText(601)
      Else
         Text9 = adoquery.Fields("a1k16").Value
      End If
      If IsNull(adoquery.Fields("a1k11").Value) Then
         Text6 = MsgText(601)
      Else
         Text6 = adoquery.Fields("a1k11").Value
      End If
      '2013/10/22 modify by sonia a1k05改為a1k34
      If IsNull(adoquery.Fields("a1k34").Value) Then
         Text10 = MsgText(601)
      Else
         Text10 = adoquery.Fields("a1k34").Value
      End If
   Else
      MsgBox MsgText(28), , MsgText(5)
      Text5 = MsgText(601)
      MaskEdBox2.Mask = MsgText(601)
      MaskEdBox2.Text = MsgText(601)
      MaskEdBox2.Mask = DFormat
      Text1 = MsgText(601)
      Text7 = MsgText(601)
      Text8 = MsgText(601)
      Text9 = MsgText(601)
      Text6 = MsgText(601)
      Text2 = MsgText(601)
      Text3 = MsgText(601)
      Text4 = MsgText(601)
      Text10 = MsgText(601)
   End If
   adoquery.Close
End Sub
