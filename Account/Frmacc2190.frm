VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Frmacc2190 
   AutoRedraw      =   -1  'True
   Caption         =   "結匯資料輸入 --> FC暫收款退費"
   ClientHeight    =   5120
   ClientLeft      =   50
   ClientTop       =   270
   ClientWidth     =   8760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5120
   ScaleWidth      =   8760
   Begin VB.CommandButton Command3 
      Caption         =   "其他"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2184
      TabIndex        =   6
      Top             =   4656
      Width           =   1092
   End
   Begin VB.CommandButton Command5 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   8040
      Picture         =   "Frmacc2190.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   2
      ToolTipText     =   "取消"
      Top             =   120
      Width           =   450
   End
   Begin VB.CommandButton Command4 
      Caption         =   "抵帳單"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1104
      TabIndex        =   5
      Top             =   4656
      Width           =   972
   End
   Begin VB.CommandButton Command2 
      Caption         =   "帳單"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   264
      TabIndex        =   4
      Top             =   4656
      Width           =   732
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc2190.frx":066A
      Height          =   4000
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   8295
      _ExtentX        =   14623
      _ExtentY        =   7056
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "a1702"
         Caption         =   "FC暫收款退費單號"
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
         DataField       =   "a1703"
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
      BeginProperty Column02 
         DataField       =   "a1704"
         Caption         =   "金額"
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
      BeginProperty Column03 
         DataField       =   "FagentName"
         Caption         =   "代理人名稱"
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
         DataField       =   "a1705"
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
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   2099.906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   840.189
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   1480.252
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   3300.095
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1709.858
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text1 
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
      Height          =   300
      Left            =   2280
      MaxLength       =   15
      TabIndex        =   0
      Top             =   240
      Width           =   1572
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   5040
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2787
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   312
      Left            =   240
      Top             =   480
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
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "作業日期"
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
      Left            =   4080
      TabIndex        =   8
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   24
      Top             =   4776
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "FC暫收款退費單號"
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
      TabIndex        =   7
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "Frmacc2190"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/06 改成Form2.0 ; DataGrid1改字型=新細明體-ExtB
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit
Public adoacc130 As New ADODB.Recordset
Public adoacc170 As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Dim adoaccrpt216 As New ADODB.Recordset
Dim adoaccrpt217 As New ADODB.Recordset
Dim adoaccrpt218 As New ADODB.Recordset
Dim dllaccrpt As Object

Private Sub Command2_Click()
   strExitControl = MsgText(601)
   tool3_enabled
   Frmacc2170.Show
   Unload Me
End Sub

Private Sub Command3_Click()
   strExitControl = MsgText(601)
   tool3_enabled
   Frmacc2191.Show
   Unload Me
End Sub

Private Sub Command4_Click()
   strExitControl = MsgText(601)
   tool3_enabled
   Frmacc2180.Show
   Unload Me
End Sub

Private Sub Command5_Click()
   AdodcDelete
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
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
   PUB_InitForm Me, 8880, 5550, strBackPicPath1
   'end 2021/12/07
   MaskEdBox1.Mask = MsgText(601)
   MaskEdBox1.Text = CFDate(ACDate(ServerDate))
   MaskEdBox1.Mask = DFormat
   OpenTable
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(133) & ", " & MsgText(134)
   Set dllaccrpt = CreateObject("AccReport.ReportSelect")
   
   Text1.Text = "O" '2015/5/18 add by sonia
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(601)
   Screen.MousePointer = vbDefault
   If strExitControl = MsgText(602) Then
      strFormName = MsgText(601)
      KeyEnter vbKeyEscape
      MenuEnabled
      Set dllaccrpt = Nothing
      Set Frmacc2190 = Nothing
   End If
   strExitControl = MsgText(602)
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
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(133) & ", " & MsgText(134)
End Sub

Private Sub Text1_GotFocus()
   'MODIFY BY SONIA 2015/5/18
   'TextInverse Text1
   If Len(Text1) > 0 Then
      Text1.SelStart = 1
      Text1.SelLength = Len(Text1) - 1
   End If
   '2015/5/18
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   If adoadodc1.State = adStateOpen Then
      adoadodc1.Close
   End If
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "select a1702, a1703, a1704, a1705, a1706, a1707, nvl(fa05||fa63||fa64||fa65,nvl(fa04, fa06)) FagentName from acc170, fagent where substr(a1705, 1, 8) = fa01 (+) and substr(a1705, 9, 1) = fa02 (+) and a1701 = '3' and (a1709 is null or a1709 = '') order by a1702 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
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
   adoadodc1.Open "select a1702, a1703, a1704, a1705, a1706, a1707, nvl(fa05||fa63||fa64||fa65,nvl(fa04, fa06)) FagentName from acc170, fagent where substr(a1705, 1, 8) = fa01 (+) and substr(a1705, 9, 1) = fa02 (+) and a1701 = '3' and (a1709 is null or a1709 = '')  order by a1702 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Adodc1.Recordset.Requery
   If Adodc1.Recordset.RecordCount <> 0 Then
      Adodc1.Recordset.Find "a1702 = '" & Text1 & "'", 0, adSearchForward, 1
      If Adodc1.Recordset.EOF Then
         Exit Sub
      Else
         DataGrid1.SelBookmarks.add Adodc1.Recordset.Bookmark
      End If
   End If
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  儲存資料表(國外結匯資料)
'
'*************************************************
Private Sub Acc170Save()
Dim stra1703 As String
Dim stra1704 As String
Dim stra1705 As String
Dim stra1706 As String
Dim stra1707 As String
Dim stra1708 As String

On Error GoTo Checking
   adoacc130.CursorLocation = adUseClient
   adoacc130.Open "select * from acc130 where a1301 = '" & Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoacc130.RecordCount <> 0 Then
      adoacc170.CursorLocation = adUseClient
      adoacc170.Open "select * from acc170 where a1701 = '3' order by a1702 asc", adoTaie, adOpenStatic, adLockReadOnly
      If adoacc170.RecordCount <> 0 Then
         adoacc170.Find "a1702 = '" & Text1 & "'", 0, adSearchForward, 1
         If adoacc170.EOF = False Then
            adoacc130.Close
            adoacc170.Close
            MsgBox MsgText(9), , MsgText(5)
            Exit Sub
         End If
      End If
      adoacc170.Close
      If IsNull(adoacc130.Fields("a1306").Value) Then
         stra1703 = ""
      Else
         stra1703 = adoacc130.Fields("a1306").Value
      End If
      If IsNull(adoacc130.Fields("a1307").Value) Then
         stra1704 = "0"
      Else
         stra1704 = adoacc130.Fields("a1307").Value
      End If
      If IsNull(adoacc130.Fields("a1304").Value) Then
         stra1705 = ""
      Else
         stra1705 = adoacc130.Fields("a1304").Value
      End If
      '2010/2/6 add by sonia 付款對象固定改為另一家,財務說不限CFT案,CFP也要
      Select Case Mid(stra1705, 1, 6)
         '2012/7/9 modify by sonia 婧瑄說加Y20908
         '2014/11/24 modify by sonia 婉莘再加Y54053
         'modify by sonia 2016/4/28 陳經理再加Y54052
         'modify by sonia 2017/8/1  陳經理再加Y54715
         'modify by sonia 2020/1/16 陳經理再加Y55351
         'modify by sonia 2024/8/12 再加Y56014
         'modify by sonia 2025/3/4  再加Y56137
         Case "Y20908", "Y20915", "Y20919", "Y20929", "Y20934", "Y34282", "Y22247", "Y30249", "Y51368", "Y51523", "Y52243", "Y20076", "Y20339", "Y54053", "Y54052", "Y54715", "Y55351", "Y56014", "Y56137"
            stra1705 = "Y20076000"          '2009/9/23 add by sonia 外商陳經理提出中東地區Abu-Ghazaleh代理人之付款對象改為Y20076,財務說不限CFT案,CFP也要
         '2010/11/19 add by sonia 陳經理說再加一組
         '2010/12/6 modify by sonia 加Y53117--Y53119,Y53122-3
         '2011/4/11 modify by sonia 婧瑄說加Y53121
         '2012/5/4  modify by sonia 加Y51352
         Case "Y45778", "Y53120", "Y20917", "Y53117", "Y53118", "Y53119", "Y53122", "Y53123", "Y53121", "Y51352"
            stra1705 = "Y20917000"
         '2010/11/22 add by sonia 陳經理說再加一組
         Case "Y49419", "Y53188"
            stra1705 = "Y53188000"
         'add by sonia 2018/1/26
         Case Else
            If stra1705 = "Y20026020" Then stra1705 = "Y20026000"
            If stra1705 = "Y45878000" Then stra1705 = "Y55253000"   'add by sonia 2019/5/28 婉莘
            If stra1705 = "Y51333010" Then stra1705 = "Y51333000"   'add by sonia 2019/6/24 婉莘
            If stra1705 = "Y52754020" Then stra1705 = "Y52754010"   'add by sonia 2022/1/20 婉莘
         'end 2018/1/26
      End Select
      '2010/2/6 end
      'Add by Amy 2013/11/18 +帳款處理訊息
      strExc(0) = GetDizhang("" & adoacc130.Fields("a1304").Value, , True)
      'end 2013/11/18
      If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
         stra1708 = Val(FCDate(MaskEdBox1.Text))
      Else
         stra1708 = ""
      End If
      adoTaie.Execute "insert into acc170 (a1701, a1702, a1703, a1704, a1705, a1708, a1710, a1711, a1712) values ('3', '" & Text1 & "', " & CNULL(stra1703) & ", " & CNULL(stra1704) & ", " & CNULL(stra1705) & ", " & CNULL(stra1708) & ", " & strSrvDate(2) & ", " & ServerTime & ", '" & strUserNum & "')"
      AdodcRefresh
   Else
      MsgBox MsgText(28), , MsgText(5)
   End If
   adoacc130.Close
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
On Error GoTo Checking
   Select Case KeyCode
      Case vbKeyF12
         AdodcRefresh
      Case vbKeyInsert
         Acc170Save
         Text1 = ""
         Text1.SetFocus
   End Select
   KeyEnter KeyCode
Checking:
   Exit Sub
End Sub

'*************************************************
'  刪除資料表
'
'*************************************************
Private Sub AdodcDelete()
On Error GoTo Checking
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   adoTaie.Execute "delete from acc170 where a1701 = '3' and a1702 = '" & Adodc1.Recordset.Fields("a1702").Value & "'"
   AdodcRefresh
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  產生結匯前核對表
'
'*************************************************
'2005/9/9 CANCEL BY SONIA 改抓Frmacc2170.ProcessData
'Private Sub ProcessData()

