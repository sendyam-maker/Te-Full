VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc21j0 
   AutoRedraw      =   -1  'True
   Caption         =   "帳單作廢作業"
   ClientHeight    =   5292
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5292
   ScaleWidth      =   8760
   Begin VB.CommandButton Command5 
      Height          =   300
      Left            =   2640
      Picture         =   "Frmacc21j0.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   240
      Width           =   350
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc21j0.frx":0102
      Height          =   3105
      Left            =   240
      TabIndex        =   3
      Top             =   1950
      Width           =   8295
      _ExtentX        =   14626
      _ExtentY        =   5440
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
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "axf02"
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
         DataField       =   "axf03"
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
      BeginProperty Column02 
         DataField       =   "axf12"
         Caption         =   "案件名稱"
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
         DataField       =   "axf04"
         Caption         =   "帳單金額"
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
      BeginProperty Column04 
         DataField       =   "axf13"
         Caption         =   "收據抬頭"
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
            ColumnWidth     =   1344.189
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1391.811
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   3420.284
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2675.906
         EndProperty
      EndProperty
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
      Left            =   1320
      TabIndex        =   7
      Top             =   600
      Width           =   1572
   End
   Begin VB.TextBox Text4 
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
      Left            =   6840
      TabIndex        =   6
      Top             =   600
      Width           =   1572
   End
   Begin VB.TextBox Text5 
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
      Left            =   4080
      TabIndex        =   5
      Top             =   960
      Width           =   1572
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
      Left            =   6840
      TabIndex        =   4
      Top             =   960
      Width           =   1572
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   330
      Left            =   1320
      TabIndex        =   8
      Top             =   960
      Width           =   1572
      _ExtentX        =   2773
      _ExtentY        =   593
      _Version        =   393216
      BackColor       =   14737632
      Enabled         =   0   'False
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
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   300
      Left            =   4080
      TabIndex        =   2
      Top             =   240
      Width           =   1572
      _ExtentX        =   2773
      _ExtentY        =   529
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
      Left            =   240
      Top             =   1680
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
   Begin MSForms.TextBox Text3 
      Height          =   330
      Left            =   2910
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   600
      Width           =   2190
      VariousPropertyBits=   671105055
      BackColor       =   14737632
      MaxLength       =   50
      Size            =   "3863;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text8 
      Height          =   405
      Left            =   1320
      TabIndex        =   17
      Top             =   1320
      Width           =   7095
      VariousPropertyBits=   -1467989985
      BackColor       =   14737632
      ScrollBars      =   2
      Size            =   "12515;714"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   1725
      Left            =   240
      Top             =   120
      Width           =   8295
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
      Left            =   3120
      TabIndex        =   16
      Top             =   240
      Width           =   972
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   4680
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "帳單編號"
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
      Top             =   240
      Width           =   972
   End
   Begin VB.Label Label2 
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
      Left            =   360
      TabIndex        =   14
      Top             =   600
      Width           =   972
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "代理人D/N No."
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
      Left            =   5280
      TabIndex        =   13
      Top             =   600
      Width           =   1572
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "帳單日期"
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
      TabIndex        =   12
      Top             =   960
      Width           =   972
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "幣別"
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
      TabIndex        =   11
      Top             =   960
      Width           =   972
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "帳單金額"
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
      TabIndex        =   10
      Top             =   960
      Width           =   972
   End
   Begin VB.Label Label9 
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
      TabIndex        =   9
      Top             =   1320
      Width           =   852
   End
End
Attribute VB_Name = "Frmacc21j0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/08 改成Form2.0 ; DataGrid1改字型=新細明體-ExtB、Text3、Text8
'Memo By Sonia 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit

Public adoacc150 As New ADODB.Recordset
Public adoacc150q As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Dim RQstr As String  'Add by Lydia 2014/10/31
'Add By Sindy 2018/2/22
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Public m_strCP01 As String, m_strCP02 As String, m_strCP03 As String, m_strCP04 As String
Public m_RDate As String
Dim m_Done As Boolean
Public m_PrevForm As Form  '前一畫面
'2018/2/22 END

Private Sub Command5_Click()
'   If adoacc150.RecordCount = 0 Or Text2 = MsgText(601) Then
'      Exit Sub
'   End If
'   adoacc150.Find "a1501 = '" & Text2 & "'", 0, adSearchForward, 1
   'Add by Lydia 2014/10/31 改為frmacc2150的方式

   Acc150Refresh
  ' AdodcClear
   If adoacc150.RecordCount <> 0 And adoacc150.EOF = False Then
      FormShow
      AdodcRefresh
      RecordShow
      
      'Added by Sindy 2018/2/22
      If Val(m_RDate) > 0 Then
         MaskEdBox2.Text = CFDate(Val(Me.m_RDate))
         MaskEdBox2.Mask = DFormat
      End If
      '2018/2/22 END
   Else
      If FMP2open = True Then
        MsgBox "權限不足或查無符合資料 !", vbInformation
      Else
        MsgBox MsgText(33), , MsgText(5)
      End If
      If adoacc150.BOF <> adoacc150.EOF Then adoacc150.MoveFirst
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
   
   'Added by Sindy 2018/2/22
   If m_strIR01 <> "" And m_Done = False Then
      m_Done = True
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "＜" & m_strCP01 & m_strCP02 & m_strCP03 & m_strCP04 & "＞）"
      KeyEnter vbKeyF2 'Set新增狀態
   End If
   '2018/2/22 END
   
   If strItemNo = MsgText(601) Then
      Exit Sub
   End If
   If adoacc150.RecordCount <> 0 Then
      adoacc150.MoveFirst
   End If
   adoacc150.Find "a1501 = '" & strItemNo & "'", 0, adSearchForward, 1
   If adoacc150.EOF = False Then
      FormShow
      AdodcRefresh
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
   'Modify by Amy 2023/08/18 H5500
   PUB_InitForm Me, 8850, 5730, strBackPicPath1
   'end 2021/12/07
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat
   'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
   FMP2open = PUB_FMPtoCheck(1, 0, Pub_strUserST05)
    
   OpenTable
   If adoacc150.RecordCount <> 0 Then
      adoacc150.MoveLast
      adoacc150.MoveFirst
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
   KeyEnter vbKeyEscape
   MenuEnabled
   
   'Add By Sindy 2018/2/23
   If m_strIR01 <> "" Then
      If Not m_PrevForm Is Nothing Then
         Call m_PrevForm.GoNext
      End If
   End If
   If Not m_PrevForm Is Nothing Then
      Set m_PrevForm = Nothing
   End If
   '2018/2/23 END
   
   Set Frmacc21j0 = Nothing
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
   If Text1 = MsgText(601) Then
      Exit Sub
   End If
   'Text3 = FagentQuery(Text1, 2)
'Add by Lydia 2014/11/13 改變讀取代理人名稱的方式
   If ClsPDGetAgent(Text1, strExc(0)) = True Then
      Text3 = strExc(0)
   Else
      Text3 = ""
   End If
   
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

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoacc150.CursorLocation = adUseClient
   'adoacc150.Open "select * from acc150 where a1507 > 0 order by a1501 asc", adoTaie, adOpenStatic, adLockReadOnly
   'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
   Dim midSql As String
   midSql = " select m0.* from acc150 m0 where a1507 > 0 and a1501>='" & Text2 & "' "
   If FMP2open = True Then
      RQstr = " select m1.axf01 from acc151 m1,caseprogress f0 where m0.a1501=m1.axf01(+) and m1.axf02=f0.cp09(+) " & FMP2openSQL
      midSql = midSql & " and a1501 in (" & RQstr & ") "
   End If
   midSql = midSql & " order by 1 asc "
   adoacc150.Open midSql, adoTaie, adOpenStatic, adLockReadOnly
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "select * from acc151 where axf01 = '" & Text2 & "' order by axf02 asc", adoTaie, adOpenStatic, adLockReadOnly
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  重新整理國外帳單資料
'
'*************************************************
Public Sub Acc150Refresh()
On Error GoTo Checking
   If adoacc150.State = adStateOpen Then
      adoacc150.Close
   End If
   adoacc150.CursorLocation = adUseClient
   adoacc150.MaxRecords = intMax
   'adoacc150.Open "select * from acc150 where a1501 >= '" & Text2 & "' order by a1501 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。=>RQstr
   Dim midSql As String
   midSql = " select m0.* from acc150 m0 where a1501 >= '" & Text2 & "' "
   If FMP2open = True Then
      midSql = midSql & " and a1501 in (" & RQstr & ") "
   End If
   midSql = midSql & " order by 1 asc "
   adoacc150.Open midSql, adoTaie, adOpenDynamic, adLockBatchOptimistic
   
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

   If adoadodc1.State = adStateOpen Then
      adoadodc1.Close
   End If
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "select * from acc151 where axf01 = '" & Text2 & "' order by axf02 asc", adoTaie, adOpenStatic, adLockReadOnly
   Adodc1.Recordset.Requery
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  顯示查詢資料(國外帳單資料(主檔))
'
'*************************************************
Private Sub Acc150Query()
   Dim MQuery As Boolean 'Add by Lydia 2014/10/31 判斷主檔查詢是否有符合的資料
   adoacc150q.CursorLocation = adUseClient
  ' adoacc150q.Open "select * from acc150 where a1501 = '" & Text2 & "'", adoTaie, adOpenStatic, adLockReadOnly
     'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
   Dim midSql As String
   midSql = " select m0.* from acc150 m0 where a1501 = '" & Text2 & "'"
   If FMP2open = True Then
      midSql = midSql & " and a1501 in (" & RQstr & ") "
   End If
   midSql = midSql & " order by 1 asc "
   MQuery = False
   
   adoacc150q.Open midSql, adoTaie, adOpenStatic, adLockReadOnly
   If adoacc150q.RecordCount <> 0 Then
      MQuery = True
      MaskEdBox2.Mask = MsgText(601)
      If IsNull(adoacc150q.Fields("a1507").Value) Then
         MaskEdBox2.Text = MsgText(601)
      Else
         MaskEdBox2.Text = CFDate(adoacc150q.Fields("a1507").Value)
      End If
      MaskEdBox2.Mask = DFormat
      If IsNull(adoacc150q.Fields("a1503").Value) Then
         Text1 = MsgText(601)
      Else
         Text1 = adoacc150q.Fields("a1503").Value
      End If
      If IsNull(adoacc150q.Fields("a1504").Value) Then
         Text4 = MsgText(601)
      Else
         Text4 = adoacc150q.Fields("a1504").Value
      End If
      MaskEdBox1.Mask = MsgText(601)
      If IsNull(adoacc150q.Fields("a1502").Value) Then
         MaskEdBox1.Text = MsgText(601)
      Else
         MaskEdBox1.Text = CFDate(adoacc150q.Fields("a1502").Value)
      End If
      MaskEdBox1.Mask = DFormat
      If IsNull(adoacc150q.Fields("a1505").Value) Then
         Text5 = MsgText(601)
      Else
         Text5 = adoacc150q.Fields("a1505").Value
      End If
      If IsNull(adoacc150q.Fields("a1506").Value) Then
         Text6 = MsgText(601)
      Else
         Text6 = adoacc150q.Fields("a1506").Value
      End If
      If IsNull(adoacc150q.Fields("a1509").Value) Then
         Text8 = MsgText(601)
      Else
         Text8 = adoacc150q.Fields("a1509").Value
      End If
   Else
      MaskEdBox2.Mask = ""
      MaskEdBox2.Text = ""
      MaskEdBox2.Mask = DFormat
      Text1 = ""
      Text3 = ""
      Text4 = ""
      MaskEdBox1.Mask = ""
      MaskEdBox1.Text = ""
      MaskEdBox1.Mask = DFormat
      Text5 = ""
      Text6 = ""
      Text8 = ""
   End If
   adoacc150q.Close
   If adoadodc1.State = adStateOpen Then
      adoadodc1.Close
   End If
    'Add by Lydia 2014/10/31 判斷主檔查詢是否有符合的資料
   If MQuery = True Then
        adoadodc1.CursorLocation = adUseClient
        adoadodc1.Open "select * from acc151 where axf01 = '" & Text2 & "' order by axf02 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
        Adodc1.Recordset.Requery
   End If
End Sub

'*************************************************
'  顯示資料表
'
'*************************************************
Public Sub FormShow()
   Text2 = adoacc150.Fields("a1501").Value
   MaskEdBox2.Mask = MsgText(601)
   If IsNull(adoacc150.Fields("a1507").Value) Then
      MaskEdBox2.Text = MsgText(601)
   Else
      MaskEdBox2.Text = CFDate(adoacc150.Fields("a1507").Value)
   End If
   MaskEdBox2.Mask = DFormat
   If IsNull(adoacc150.Fields("a1503").Value) Then
      Text1 = MsgText(601)
   Else
      Text1 = adoacc150.Fields("a1503").Value
   End If
   If IsNull(adoacc150.Fields("a1504").Value) Then
      Text4 = MsgText(601)
   Else
      Text4 = adoacc150.Fields("a1504").Value
   End If
   MaskEdBox1.Mask = MsgText(601)
   If IsNull(adoacc150.Fields("a1502").Value) Then
      MaskEdBox1.Text = MsgText(601)
   Else
      MaskEdBox1.Text = CFDate(adoacc150.Fields("a1502").Value)
   End If
   MaskEdBox1.Mask = DFormat
   If IsNull(adoacc150.Fields("a1505").Value) Then
      Text5 = MsgText(601)
   Else
      Text5 = adoacc150.Fields("a1505").Value
   End If
   If IsNull(adoacc150.Fields("a1506").Value) Then
      Text6 = MsgText(601)
   Else
      Text6 = adoacc150.Fields("a1506").Value
   End If
   If IsNull(adoacc150.Fields("a1509").Value) Then
      Text8 = MsgText(601)
   Else
      Text8 = adoacc150.Fields("a1509").Value
   End If
End Sub

'*************************************************
'  顯示筆數
'
'*************************************************
Public Sub RecordShow()
   If adoacc150.RecordCount = 0 Then
      Exit Sub
   End If
   CountShow adoacc150.Bookmark, adoacc150.RecordCount
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
   If Text2 = MsgText(601) Then
      Exit Sub
   End If
   Acc150Query
End Sub
