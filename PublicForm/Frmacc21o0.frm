VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Frmacc21o0 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "預估結匯匯率資料維護"
   ClientHeight    =   5100
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   6165
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   6165
   Begin VB.CommandButton cmdTWBank 
      Caption         =   "台銀匯入"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2910
      TabIndex        =   12
      Top             =   1410
      Visible         =   0   'False
      Width           =   1065
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   0
      Top             =   3450
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C0FFC0&
      Caption         =   "列印(&P)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4005
      Style           =   1  '圖片外觀
      TabIndex        =   11
      Top             =   1410
      Width           =   1590
   End
   Begin VB.TextBox txtBase 
      Alignment       =   1  '靠右對齊
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
      Left            =   4005
      MaxLength       =   4
      TabIndex        =   3
      Top             =   1020
      Width           =   1572
   End
   Begin VB.TextBox txtQRate 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
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
      Left            =   1230
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1410
      Width           =   1572
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  '靠右對齊
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
      Left            =   1230
      MaxLength       =   15
      TabIndex        =   2
      Top             =   1020
      Width           =   1572
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0FFFF&
      Height          =   300
      Left            =   1230
      TabIndex        =   0
      Top             =   645
      Width           =   1584
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc21o0.frx":0000
      Height          =   3135
      Left            =   255
      TabIndex        =   5
      Top             =   1830
      Width           =   5400
      _ExtentX        =   9525
      _ExtentY        =   5530
      _Version        =   393216
      AllowUpdate     =   0   'False
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
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "a2101"
         Caption         =   "預估日期"
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
      BeginProperty Column01 
         DataField       =   "a2102"
         Caption         =   "幣別"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "###/##/##"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "a2103"
         Caption         =   "預估匯率"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.000000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "a2111"
         Caption         =   "報價匯率"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.000000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
            Alignment       =   2
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   1260.284
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1260.284
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   45
      Top             =   1950
      Visible         =   0   'False
      Width           =   1200
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
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   4020
      TabIndex        =   1
      Top             =   645
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   529
      _Version        =   393216
      BackColor       =   12648447
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
   Begin VB.Label Label6 
      Caption         =   "每日 16:30 系統自動匯入台銀當日收盤賣出即期匯率為預估匯率，已輸入資料將被取代 ，預設基數為 1.03，若有調整請通知電腦中心!!!"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   465
      Left            =   90
      TabIndex        =   13
      Top             =   90
      Width           =   5970
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "預設基數 "
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
      Left            =   3015
      TabIndex        =   10
      Top             =   1020
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "報價匯率 "
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
      Left            =   225
      TabIndex        =   9
      Top             =   1410
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "預估日期"
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
      Left            =   3015
      TabIndex        =   8
      Top             =   645
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "預估匯率 "
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
      Left            =   225
      TabIndex        =   7
      Top             =   1020
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "幣別"
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
      TabIndex        =   6
      Top             =   660
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Top             =   4320
      Visible         =   0   'False
      Width           =   135
   End
End
Attribute VB_Name = "Frmacc21o0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2021/12/01 Form2.0已檢查 (無需修改的物件)
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo By Sindy 2010/8/12 日期欄已修改
Option Explicit

Public adoacc210 As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
'Add by Morgan 2005/6/23 列印用
Dim PLeft(0 To 12) As Integer, iPrint As Integer, iPage As Integer, strTemp(1 To 4) As String
Dim m_iTitleFontSize As Single, m_iFontSize As Single
Dim m_iStartX As Integer, m_iStartY As Integer
Dim m_iPageHeight As Integer, m_iLineHeight As Integer, m_stTmp As String
Dim m_iMargin As Integer

Private Sub cmdTWBank_Click()
   Dim strDate As String, strText As String
   Me.Enabled = False
   Screen.MousePointer = vbHourglass
   
   strDate = InputBox("請輸入匯入日期!!!", , strSrvDate(2))
   If strDate <> "" Then
      strText = PUB_GetTwBankRate(Me.Inet1, strDate)
      If strText <> "" Then
         If PUB_ImportRate(strText, strDate) = True Then
            AdodcRefresh
            MsgBox "匯入完成！", vbInformation
         End If
      End If
   End If
   Screen.MousePointer = vbDefault
   Me.Enabled = True
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Combo1_Validate(Cancel As Boolean)
   If Combo1 = MsgText(601) Then
      Exit Sub
   End If
   If ExistCheck("acc1y0", "a1y01", Combo1, Label1) = False Then
      Cancel = True
      Combo1.SetFocus
   End If
End Sub

Private Sub DataGrid1_SelChange(Cancel As Integer)
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   FormShow
   RecordShow
End Sub

Private Sub Form_Activate()
   '93.3.16 ADD BY SONIA
   'Modified by Morgan 2015/12/21
   'If IsObject(mdiMain) Then
   '   ToolShow
   'End If
   If UCase(Forms(0).Name) = "MDIMAIN" Then
      Forms(0).ToolShow
   End If
   'end 2015/12/21
   '93.3.16 END
   
   strFormName = Name
   If strItemNo = MsgText(601) Then
      Exit Sub
   End If
   If adoadodc1.RecordCount <> 0 Then
      adoadodc1.MoveFirst
   End If
   adoadodc1.Find "a2101 = '" & strItemNo & "'", 0, adSearchForward, 1
   If adoadodc1.EOF = False Then
      adoadodc1.Find "a2102 = '" & strCustNo & "'", 0, adSearchForward, adoadodc1.Bookmark
      If adoadodc1.EOF = False Then
         FormShow
         RecordShow
      End If
   End If
   strItemNo = MsgText(601)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Form_Load()

'Modified by Morgan 2019/7/10
'Dim intX As Integer
'Dim intY As Integer
'Dim sglWidth As Single
'Dim sglHeight As Single
'
'   Me.Icon = LoadPicture(strIcoPath)
'   strFormName = Name
'   Me.Width = 6000
'   Me.Height = 5000
'   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
'   Image1 = LoadPicture(strBackPicPath1)
'   sglWidth = Image1.Width
'   sglHeight = Image1.Height
'   For intX = 0 To Int(ScaleWidth / sglWidth)
'       For intY = 0 To Int(ScaleHeight / sglHeight)
'           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
'       Next
'   Next
   PUB_InitForm Me, Me.Width, Me.Height
   If Pub_StrUserSt03 = "M51" Then cmdTWBank.Visible = True
'end 2019/7/10
   OpenTable
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
   Set Frmacc21o0 = Nothing
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoadodc1.CursorLocation = adUseClient
   'Modify by Morgan 2005/6/14 加報價匯率
   'adoadodc1.Open "select * from acc210 order by a2101 DESC, a2102 asc", adoTaie, adOpenStatic, adLockReadOnly
   adoadodc1.Open "select a2101,a2102,a2103,a2110,nvl(a2110,1.03)*a2103 a2111 from acc210 order by a2101 DESC, a2102 asc", adoTaie, adOpenStatic, adLockReadOnly
   Set Adodc1.Recordset = adoadodc1
   
   If adoquery.State = adStateOpen Then adoquery.Close
   adoquery.CursorLocation = adUseClient
   adoquery.Open "select * from acc1y0 order by a1y01 asc", adoTaie, adOpenStatic, adLockReadOnly
   Do While adoquery.EOF = False
      Combo1.AddItem adoquery.Fields("a1y01").Value
      adoquery.MoveNext
   Loop
   adoquery.Close
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  顯示欄位資料(美金匯率資料表)
'
'*************************************************
Public Sub FormShow()
   If IsNull(Adodc1.Recordset.Fields("a2102").Value) Then
      Combo1 = MsgText(601)
   Else
      Combo1 = Adodc1.Recordset.Fields("a2102").Value
   End If
   MaskEdBox1.Mask = MsgText(601)
   If IsNull(Adodc1.Recordset.Fields("a2101").Value) Then
      MaskEdBox1.Text = MsgText(601)
   Else
      MaskEdBox1.Text = CFDate(Adodc1.Recordset.Fields("a2101").Value)
   End If
   MaskEdBox1.Mask = DFormat
   If IsNull(Adodc1.Recordset.Fields("a2103").Value) Then
      Text5 = MsgText(601)
   Else
      Text5 = Adodc1.Recordset.Fields("a2103").Value
   End If
   'Add by Morgan 2005/6/14
   txtBase = "" & Adodc1.Recordset.Fields("a2110").Value
   If txtBase = "" Then txtBase = "1.03"
End Sub

'*************************************************
'  重新整理 Adodc 之資料
'
'*************************************************
Public Sub AdodcRefresh()
On Error GoTo Checking
   'Removed by Morgan 2019/7/10
   'adoadodc1.Close
   'adoadodc1.CursorLocation = adUseClient
   ''Modify by Morgan 2005/6/14 加報價匯率
   ''adoadodc1.Open "select * from acc210 order by a2101 asc, a2102 asc", adoTaie, adOpenStatic, adLockReadOnly
   'adoadodc1.Open "select a2101,a2102,a2103,a2110,nvl(a2110,1.03)*a2103 a2111 from acc210 order by a2101 asc, a2102 asc", adoTaie, adOpenStatic, adLockReadOnly
   'end 2019/7/10
   Adodc1.Recordset.Requery
   If Adodc1.Recordset.RecordCount <> 0 Then
      If Combo1 <> MsgText(601) Then
         Adodc1.Recordset.Find "a2101 = " & Val(FCDate(MaskEdBox1.Text)) & "", 0, adSearchForward, 1
         If Adodc1.Recordset.EOF = False Then
            Adodc1.Recordset.Find "a2102 = '" & Combo1 & "'", 0, adSearchForward, adoadodc1.Bookmark
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
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   CountShow Adodc1.Recordset.Bookmark, Adodc1.Recordset.RecordCount
End Sub


Private Sub Text5_Change()
   SetQRate
End Sub

Private Sub Text5_GotFocus()
   TextInverse Text5
End Sub

Private Sub txtBase_Change()
   SetQRate
End Sub

Private Sub txtBase_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then
      Beep
      KeyAscii = 0
   End If
End Sub

Private Sub SetQRate()
   txtQRate = Val(txtBase) * Val(Text5)
End Sub

Public Function FormCheck(ByRef p_Update As Boolean, ByRef p_AddNew As Boolean) As Boolean
   With Me
      p_Update = False
      p_AddNew = False
      If .MaskEdBox1.Text = MsgText(601) Or .MaskEdBox1.Text = MsgText(29) Then
         MsgBox "請輸入預估日期", vbExclamation
         .MaskEdBox1.SetFocus
          Exit Function
      End If
      If .Combo1 = MsgText(601) Then
         MsgBox MsgText(10) & .Label1, , MsgText(5)
         strControlButton = MsgText(602)
         .Combo1.SetFocus
         Exit Function
      Else
         If .Text5 = MsgText(601) Then
            MsgBox MsgText(10) & .Label2, , MsgText(5)
            strControlButton = MsgText(602)
            .Text5.SetFocus
            Exit Function
         End If
      End If
      
      'Add by Morgan 2010/6/29
      '若與去年同期比較超過10%時彈訊息提醒
      strExc(1) = Val(FCDate(.MaskEdBox1.Text) - 10000)
      strExc(0) = "select SQLDATET(a2101),a2103 from acc210 a where a2101 =(select max(b.a2101) from acc210 b where b.a2101<= " & strExc(1) & " and b.a2102 = a.a2102) and a2102 = '" & .Combo1 & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strExc(2) = Val(Text5) - Val("" & RsTemp(1))
         If Abs(strExc(2)) > 0.1 * Val("" & RsTemp(1)) Then
            If MsgBox("本次預估匯率 " & Text5 & " 較去年同期匯率 " & RsTemp(1) & " ( " & RsTemp(0) & " ) 差異超過 10%，是否要繼續？", vbYesNo + vbDefaultButton2) = vbNo Then
               strControlButton = MsgText(602)
               .Text5.SetFocus
               Exit Function
            End If
         End If
      End If
      'end 2010/6/29
      
      If .adoacc210.State = adStateOpen Then .adoacc210.Close
      .adoacc210.CursorLocation = adUseClient
      .adoacc210.Open "select * from acc210 where a2101 = " & Val(FCDate(.MaskEdBox1.Text)) & " and a2102 = '" & .Combo1 & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
      If strSaveConfirm = MsgText(3) Then
         If .adoacc210.RecordCount <> 0 Then
            MsgBox MsgText(9), , MsgText(5)
            strControlButton = MsgText(602)
            .adoacc210.Close
            .Combo1.SetFocus
            Exit Function
         End If
         p_AddNew = True
      Else
         If strSaveConfirm = MsgText(4) Then
            If .adoacc210.RecordCount = 0 Then
               MsgBox MsgText(28), , MsgText(5)
               strControlButton = MsgText(602)
               .adoacc210.Close
               .Combo1.SetFocus
               Exit Function
            End If
         End If
      End If
   End With
   
   
   
   If UpdateCheck(p_Update) = False Then Exit Function
   FormCheck = True
End Function

Private Function UpdateCheck(ByRef p_Update As Boolean) As Boolean

   strSql = "select a2110 from acc210 where a2101 = " & Val(FCDate(Me.MaskEdBox1.Text)) & " and a2102 <> '" & Me.Combo1 & "' and nvl(a2110,1.03)<>" & Val("" & txtBase)
   p_Update = False
On Error GoTo ErrHnd

   CheckOC
   With adoRecordset
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount > 0 Then
         If MsgBox("預設基數與該日其他筆不同，是否一並更新？", vbYesNo + vbDefaultButton2) = vbYes Then
            p_Update = True
            UpdateCheck = True
         Else
            txtBase.SetFocus
         End If
      Else
         UpdateCheck = True
      End If
   End With
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical

End Function
'Add by Morgan 2005/6/23
Private Sub cmdPrint_Click()
   Dim rstClone As New ADODB.Recordset
   Set rstClone = Adodc1.Recordset.Clone
   Dim stDate As String, iRecs As Integer
   
   With rstClone
      GetPleft
      If .RecordCount > 0 Then
         .Sort = "a2101 desc,a2101 asc"
         .MoveFirst
         stDate = "" & .Fields("a2101")
         iPage = 1
         PrintPageHeader
         PrintPageHeader1
         Do While Not .EOF
            If stDate = "" & .Fields("a2101") Then
               iRecs = iRecs + 1
               strTemp(1) = Format(rstClone.Fields("a2101"), "###/##/##")
               strTemp(2) = rstClone.Fields("a2102")
               strTemp(3) = Format(rstClone.Fields("a2103"), "#.000000")
               strTemp(4) = Format(rstClone.Fields("a2111"), "#.000000")
               PrintDetail
               .MoveNext
            Else
               Exit Do
            End If
         Loop
         Call PrintReportFooter(iRecs)
      End If
   End With
   Set rstClone = Nothing
End Sub

Sub GetPleft()

   Printer.Orientation = 1
   m_iTitleFontSize = 22
   m_iFontSize = 12
   m_iStartX = 500
   m_iStartY = 500
   m_iPageHeight = Printer.Height
   m_iLineHeight = 300
   m_iMargin = 500
   
   Erase PLeft
   PLeft(0) = 500
   '預估日期(1200)
   PLeft(1) = 500
   '幣別(800)
   PLeft(2) = PLeft(1) + 1200
   '預估匯率(1500)
   PLeft(3) = PLeft(2) + 800
   '報價匯率(1500)
   PLeft(4) = PLeft(3) + 1500
   PLeft(5) = PLeft(4) + 1500
   
    
End Sub

Sub PrintPageHeader()
    iPrint = m_iStartY
    Printer.FontName = "細明體"
    Printer.Font.Size = m_iTitleFontSize
    Printer.Font.Bold = True
    Printer.Font.Underline = True
    m_stTmp = "最新預估結匯匯率資料表"
    Printer.CurrentX = (Printer.Width - Printer.TextWidth(m_stTmp)) / 2
    Printer.CurrentY = iPrint
    Printer.Print m_stTmp
    iPrint = iPrint + 500
    Printer.Font.Size = m_iFontSize
    Printer.Font.Bold = False
    Printer.Font.Underline = False
    PrintNewLine
    Printer.CurrentX = m_iStartX
    Printer.CurrentY = iPrint
    Printer.Print "列印人：" & strUserName
    Printer.CurrentX = Printer.Width - m_iMargin - 2500
    Printer.CurrentY = iPrint
    Printer.Print "列印日期：" & Format(strSrvDate(2), "##/##/##")
    PrintNewLine
    Printer.CurrentX = Printer.Width - m_iMargin - 2500
    Printer.CurrentY = iPrint
    Printer.Print "頁    次：" & str(iPage)
    PrintNewLine
End Sub

Sub PrintPageHeader1()

    Call PrintNewLine(False, 1)
    
    Printer.CurrentX = PLeft(1)
    Printer.CurrentY = iPrint
    Printer.Print "預估日期"
    Printer.CurrentX = PLeft(2)
    Printer.CurrentY = iPrint
    Printer.Print "幣別"
    m_stTmp = "預估匯率"
    Printer.CurrentX = PLeft(4) - Printer.TextWidth(m_stTmp)
    Printer.CurrentY = iPrint
    Printer.Print m_stTmp
    m_stTmp = "報價匯率"
    Printer.CurrentX = PLeft(5) - Printer.TextWidth(m_stTmp)
    Printer.CurrentY = iPrint
    Printer.Print "報價匯率"
    PrintNewLine
    Printer.CurrentX = m_iStartX
    Printer.CurrentY = iPrint
    Printer.Print String(200, "-")
    
End Sub
'列印表尾
Private Sub PrintReportFooter(Optional ByVal iRecCount As Integer = 0)

    Call PrintNewLine(True, 2)
    Printer.CurrentX = m_iStartX
    Printer.CurrentY = iPrint
    Printer.Print String(200, "-")
    PrintNewLine
    Printer.CurrentX = m_iStartX
    Printer.CurrentY = iPrint
    Printer.Print "合計： " & iRecCount & " 筆"
    Printer.EndDoc
End Sub

Sub PrintDetail()

    Dim iCol As Integer

   PrintNewLine
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iPrint
   Printer.Print strTemp(1)
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iPrint
   Printer.Print strTemp(2)
   Printer.CurrentX = PLeft(4) - Printer.TextWidth(strTemp(3))
   Printer.CurrentY = iPrint
   Printer.Print strTemp(3)
   Printer.CurrentX = PLeft(5) - Printer.TextWidth(strTemp(4))
   Printer.CurrentY = iPrint
   Printer.Print strTemp(4)
   
        
    
End Sub

Private Sub PrintNewLine(Optional ByVal p_bolHeader1 As Boolean = True, Optional ByVal p_iExtraLines As Integer = 1)

   iPrint = iPrint + m_iLineHeight
   If iPrint >= (m_iPageHeight - m_iLineHeight - p_iExtraLines * m_iLineHeight) Then
      Printer.CurrentX = m_iStartX
      Printer.CurrentY = iPrint
      Printer.Print String(200, "-")
      iPage = iPage + 1
      Printer.NewPage
      PrintPageHeader
      If p_bolHeader1 Then
         PrintPageHeader1
      End If
      iPrint = iPrint + m_iLineHeight
    End If
    
End Sub
