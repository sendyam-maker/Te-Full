VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Frmacc2180 
   AutoRedraw      =   -1  'True
   Caption         =   "結匯資料輸入 --> 抵帳單"
   ClientHeight    =   5120
   ClientLeft      =   50
   ClientTop       =   270
   ClientWidth     =   8760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
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
      Left            =   3384
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
      Picture         =   "Frmacc2180.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   2
      ToolTipText     =   "取消"
      Top             =   120
      Width           =   450
   End
   Begin VB.CommandButton Command4 
      Caption         =   "FC暫收款退費"
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
      Left            =   1584
      TabIndex        =   5
      Top             =   4656
      Width           =   1692
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
      Width           =   1092
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc2180.frx":066A
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
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "a1702"
         Caption         =   "抵帳單編號"
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
         Caption         =   "抵帳單金額"
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
         DataField       =   "a1706"
         Caption         =   "代理人C/N No."
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
         DataField       =   "a1707"
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
      BeginProperty Column05 
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
      BeginProperty Column06 
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
            ColumnWidth     =   1470.047
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   569.764
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   1340.221
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1539.78
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   3339.78
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1379.906
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
      Left            =   1560
      MaxLength       =   15
      TabIndex        =   0
      Top             =   240
      Width           =   1572
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   4320
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   1572
      _ExtentX        =   2769
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
      Height          =   252
      Left            =   3360
      TabIndex        =   8
      Top             =   240
      Visible         =   0   'False
      Width           =   972
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
      Caption         =   "抵帳單編號"
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
      TabIndex        =   7
      Top             =   240
      Width           =   1212
   End
End
Attribute VB_Name = "Frmacc2180"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/06 改成Form2.0 ; DataGrid1改字型=新細明體-ExtB
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit
Public adoacc160 As New ADODB.Recordset
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
   Frmacc2190.Show
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
   
   Text1.Text = "V" 'Add by Morgan 2004/11/25
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(601)
   If strExitControl = MsgText(602) Then
      strFormName = MsgText(601)
      KeyEnter vbKeyEscape
      MenuEnabled
      Set dllaccrpt = Nothing
      Set Frmacc2180 = Nothing
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
   TextInverse Text1
   'TextInverse Text1
   If Len(Text1) > 0 Then
      Text1.SelStart = 1
      Text1.SelLength = Len(Text1) - 1
   End If
   '2014/3/19 END
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
   adoadodc1.Open "select a1702, a1703, axg04 as a1704, a1705, a1706, axg03 as a1707, nvl(fa05||fa63||fa64||fa65,nvl(fa04, fa06)) FagentName from acc161, acc170, fagent where axg01 = a1702 and substr(a1705, 1, 8) = fa01 (+) and substr(a1705, 9, 1) = fa02 (+) and a1701 = '2' and (a1709 is null or a1709 = '') order by a1702 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
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
   adoadodc1.Open "select a1702, a1703, axg04 as a1704, a1705, a1706, axg03 as a1707, nvl(fa05||fa63||fa64||fa65,nvl(fa04, fa06)) FagentName from acc161, acc170, fagent where axg01 = a1702 and substr(a1705, 1, 8) = fa01 (+) and substr(a1705, 9, 1) = fa02 (+) and a1701 = '2' and (a1709 is null or a1709 = '')  order by a1702 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
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

'Added by Morgan 2017/2/21 改 Acc170Save 為此共用函數(Frmacc2171也要用)
Public Function Acc170SaveNew(pBillNo As String, Optional pRefreshGrid As Boolean, Optional pEBill As Boolean = False) As Boolean
   Dim stSQL As String, intR As Integer
   Dim rsQuery As ADODB.Recordset
   
   Dim stra1703 As String
   Dim stra1704 As String
   Dim stra1705 As String
   Dim stra1706 As String
   Dim stra1707 As String
   Dim stra1708 As String

   '檢查是否已有結匯資料
   stSQL = "select * from acc170 where a1701 = '2' and a1702='" & pBillNo & "'"
   intR = 1
   Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      MsgBox MsgText(9), , MsgText(5)
      pRefreshGrid = True
      GoTo ExitFunction
   End If
      
   stSQL = "select * from acc161, acc160 where axg01 = a1601 and axg01 = '" & pBillNo & "'"
   intR = 1
   Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
   If intR <> 1 Then
      MsgBox MsgText(28), , MsgText(5)
   Else
      With rsQuery
      stra1703 = "" & .Fields("a1605").Value
      stra1704 = Val("" & .Fields("a1606").Value)
      stra1705 = "" & .Fields("a1603").Value
      
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
      strExc(0) = GetDizhang("" & .Fields("a1603").Value, , True)
      'end 2013/11/18
      
      stra1706 = "" & .Fields("a1604").Value
      stra1707 = "" & .Fields("axg03").Value
      
      'Modified by Morgan 2017/2/21 畫面欄位已隱藏,比照帳單直接設系統日
      'If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      '   stra1708 = Val(FCDate(MaskEdBox1.Text))
      'Else
      '   stra1708 = ""
      'End If
      stra1708 = strSrvDate(2)
      'end 2017/2/21
      'Modified by Morgan 2018/1/17 +a1719
      adoTaie.Execute "insert into acc170 (a1701, a1702, a1703, a1704, a1705, a1706, a1707, a1708, a1710, a1711, a1712, a1719) values ('2', '" & pBillNo & "', " & CNULL(stra1703) & ", " & CNULL(stra1704) & ", " & CNULL(stra1705) & ", " & CNULL(stra1706) & ", " & CNULL(stra1707) & ", " & CNULL(stra1708) & ", " & strSrvDate(2) & ", to_char(sysdate, 'HH24MISS'), '" & strUserNum & "','" & IIf(pEBill, "Y", "") & "')"
      pRefreshGrid = True
      Acc170SaveNew = True
      
        'Added by Lydia 2017/11/13 J公司或獨立水單都要彈訊息提醒
        Select Case Left(stra1707, Len(stra1707) - 9)
           Case "CFP", "P"
              stSQL = "select nvl(p1.pa26,p2.pa26) pa26,nvl(p1.pa161,p2.pa161) pa161 from acc170, acc151, patent p1, acc161, patent p2 where a1702='" & pBillNo & "' and a1702=axf01(+) and a1702=axg01(+) " & _
                        "and substr(axf03, 1, length(axf03) - 9)=p1.pa01(+) and substr(axf03, length(axf03) - 8, 6)=p1.pa02(+) and substr(axf03, length(axf03) - 2, 1)=p1.pa03(+) and substr(axf03, length(axf03) - 1, 2)=p1.pa04(+) " & _
                        "and substr(axg03, 1, length(axg03) - 9)=p2.pa01(+) and substr(axg03, length(axg03) - 8, 6)=p2.pa02(+) and substr(axg03, length(axg03) - 2, 1)=p2.pa03(+) and substr(axg03, length(axg03) - 1, 2)=p2.pa04(+) "
           Case "TF", "T", "CFT"
              stSQL = "select nvl(t1.tm23,t2.tm23) pa26,nvl(t1.tm130,t2.tm130) pa161 from acc170, acc151, trademark t1, acc161, trademark t2 where a1702='" & pBillNo & "' and a1702=axf01(+) and a1702=axg01(+) " & _
                        "and substr(axf03, 1, length(axf03) - 9)=t1.tm01(+) and substr(axf03, length(axf03) - 8, 6)=t1.tm02(+) and substr(axf03, length(axf03) - 2, 1)=t1.tm03(+) and substr(axf03, length(axf03) - 1, 2)=t1.tm04(+) " & _
                        "and substr(axg03, 1, length(axg03) - 9)=t2.tm01(+) and substr(axg03, length(axg03) - 8, 6)=t2.tm02(+) and substr(axg03, length(axg03) - 2, 1)=t2.tm03(+) and substr(axg03, length(axg03) - 1, 2)=t2.tm04(+) "
           Case "L", "CFL"
              stSQL = "select nvl(L1.LC11,L2.LC11) pa26,nvl(L1.lc48,L2.lc48) pa161 from acc170, acc151, LAWCASE L1, acc161, LAWCASE L2 where a1702='" & pBillNo & "' and a1702=axf01(+) and a1702=axg01(+) " & _
                        "and substr(axf03, 1, length(axf03) - 9)=L1.LC01(+) and substr(axf03, length(axf03) - 8, 6)=L1.LC02(+) and substr(axf03, length(axf03) - 2, 1)=L1.LC03(+) and substr(axf03, length(axf03) - 1, 2)=L1.LC04(+) " & _
                        "and substr(axg03, 1, length(axg03) - 9)=L2.LC01(+) and substr(axg03, length(axg03) - 8, 6)=L2.LC02(+) and substr(axg03, length(axg03) - 2, 1)=L2.LC03(+) and substr(axg03, length(axg03) - 1, 2)=L2.LC04(+) "
           Case Else
              stSQL = "select nvl(s1.sp08,s2.sp08) pa26,nvl(s1.sp85,s2.sp85) pa161 from acc170, acc151, servicepractice s1, acc161, servicepractice s2 where a1702='" & pBillNo & "' and a1702=axf01(+) and a1702=axg01(+) " & _
                        "and substr(axf03, 1, length(axf03) - 9)=s1.sp01(+) and substr(axf03, length(axf03) - 8, 6)=s1.sp02(+) and substr(axf03, length(axf03) - 2, 1)=s1.sp03(+) and substr(axf03, length(axf03) - 1, 2)=s1.sp04(+) " & _
                        "and substr(axg03, 1, length(axg03) - 9)=s2.sp01(+) and substr(axg03, length(axg03) - 8, 6)=s2.sp02(+) and substr(axg03, length(axg03) - 2, 1)=s2.sp03(+) and substr(axg03, length(axg03) - 1, 2)=s2.sp04(+) "
        End Select
        
        intR = 1
        Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
        If intR = 1 Then
           With rsQuery
           If IsNull(.Fields("pa26").Value) = False Then
'modify by sonia 2020/7/20 改用函數判斷
'              '高國碩要求X60149改為只要母號,關係企業不要
'              If Left(.Fields("pa26").Value, 6) = "X44551" Or Left(.Fields("pa26").Value, 6) = "X62079" _
'              Or Left(.Fields("pa26").Value, 6) = "X43988" Or Left(.Fields("pa26").Value, 6) = "X63219" Or Left(.Fields("pa26").Value, 6) = "X62319" _
'              Or Left(.Fields("pa26").Value, 6) = "X60498" Or Left(.Fields("pa26").Value, 6) = "X62702" Or Left(.Fields("pa26").Value, 6) = "X63838" _
'               Then
'                 MsgBox "此為特定客戶, 水單不能合併！", , MsgText(5)
'              End If
'              '單獨編號而關係企業不要的,例:王副總的華碩客戶(X69011010)
'              '高國碩要求X60149改為只要母號,關係企業不要
'              '張詠翔要求取消X6014900
'              '郭雅娟要求加X60738010國立清華大學
'              If Left(.Fields("pa26").Value, 8) = "X6901101" Or Left(.Fields("pa26").Value, 8) = "X6073801" Then
'                 MsgBox "此為特定客戶, 水單不能合併！", , MsgText(5)
'              End If
'
              If PUB_ChkNoMergePayCust("", .Fields("pa26")) = True Then
                 MsgBox "此為特定客戶, 水單不能合併！", , MsgText(5)
              End If
'end2020/7/20
           End If
           'Modified by Morgan 2018/2/6 電子結匯畫面已有註記不用再提醒 -- 婉莘
           If "" & .Fields("pa161").Value = "J" And pEBill = False Then
              MsgBox "此為智權公司出名案件！", , MsgText(5)
           End If
           End With
        End If
        'end 2017/11/13
   
      End With
   End If
   
ExitFunction:
   
   If Not rsQuery Is Nothing Then
      If rsQuery.State = adStateOpen Then rsQuery.Close
      Set rsQuery = Nothing
   End If
   
End Function

'*************************************************
'  儲存資料表(國外結匯資料)
'
'*************************************************
'2017/2/21 已改用Acc170SaveNew
'Private Sub Acc170Save()
'Dim stra1703 As String
'Dim stra1704 As String
'Dim stra1705 As String
'Dim stra1706 As String
'Dim stra1707 As String
'Dim stra1708 As String
'
'On Error GoTo Checking
'   If adoacc160.State = adStateOpen Then
'      adoacc160.Close
'   End If
'   adoacc160.CursorLocation = adUseClient
'   adoacc160.Open "select * from acc161, acc160 where acc161.axg01 = acc160.a1601 and axg01 = '" & Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
'   If adoacc160.RecordCount <> 0 Then
'      adoacc160.MoveFirst
'      adoacc170.CursorLocation = adUseClient
'      adoacc170.Open "select * from acc170 where a1701 = '2' order by a1702 asc", adoTaie, adOpenStatic, adLockReadOnly
'      If adoacc170.RecordCount <> 0 Then
'         adoacc170.Find "a1702 = '" & Text1 & "'", 0, adSearchForward, 1
'         If adoacc170.EOF = False Then
'            adoacc160.Close
'            adoacc170.Close
'            MsgBox MsgText(9), , MsgText(5)
'            AdodcRefresh
'            Exit Sub
'         End If
'      End If
'      adoacc170.Close
'      If IsNull(adoacc160.Fields("a1605").Value) Then
'         stra1703 = ""
'      Else
'         stra1703 = adoacc160.Fields("a1605").Value
'      End If
'      If IsNull(adoacc160.Fields("a1606").Value) Then
'         stra1704 = "0"
'      Else
'         stra1704 = adoacc160.Fields("a1606").Value
'      End If
'      If IsNull(adoacc160.Fields("a1603").Value) Then
'         stra1705 = ""
'      Else
'         stra1705 = adoacc160.Fields("a1603").Value
'      End If
'      '2010/2/6 add by sonia 付款對象固定改為另一家,財務說不限CFT案,CFP也要
'      Select Case Mid(stra1705, 1, 6)
'         '2012/7/9 modify by sonia 婧瑄說加Y20908
'         '2014/11/24 modify by sonia 婉莘再加Y54053
'         'modify by sonia 2016/4/28 陳經理再加Y54052
'         'modify by sonia 2017/8/1  陳經理再加Y54715
'         'modify by sonia 2020/1/16 陳經理再加Y55351
'         'modify by sonia 2024/8/12 再加Y56014
'         Case "Y20908", "Y20915", "Y20919", "Y20929", "Y20934", "Y34282", "Y22247", "Y30249", "Y51368", "Y51523", "Y52243", "Y20076", "Y20339", "Y54053", "Y54052", "Y54715", "Y55351", "Y56014"
'            stra1705 = "Y20076000"          '2009/9/23 add by sonia 外商陳經理提出中東地區Abu-Ghazaleh代理人之付款對象改為Y20076,財務說不限CFT案,CFP也要
'         '2010/11/19 add by sonia 陳經理說再加一組
'         '2010/12/6 modify by sonia 加Y53117--Y53119,Y53122-3
'         '2011/4/11 modify by sonia 婧瑄說加Y53121
'         '2012/5/4  modify by sonia 加Y51352
'         Case "Y45778", "Y53120", "Y20917", "Y53117", "Y53118", "Y53119", "Y53122", "Y53123", "Y53121", "Y51352"
'            stra1705 = "Y20917000"
'         '2010/11/22 add by sonia 陳經理說再加一組
'         Case "Y49419", "Y53188"
'            stra1705 = "Y53188000"
'         'add by sonia 2018/1/26
'         Case Else
'            If stra1705 = "Y20026020" Then stra1705 = "Y20026000"
'            If stra1705 = "Y45878000" Then stra1705 = "Y55253000"   'add by sonia 2019/5/28 婉莘
'            If stra1705 = "Y51333010" Then stra1705 = "Y51333000"   'add by sonia 2019/6/24 婉莘
'         'end 2018/1/26
'      End Select
'      '2010/2/6 end
'      'Add by Amy 2013/11/18 +帳款處理訊息
'      strExc(0) = GetDizhang("" & adoacc160.Fields("a1603").Value, , True)
'      'end 2013/11/18
'      If IsNull(adoacc160.Fields("a1604").Value) Then
'         stra1706 = ""
'      Else
'         stra1706 = adoacc160.Fields("a1604").Value
'      End If
'      If IsNull(adoacc160.Fields("axg03").Value) Then
'         stra1707 = ""
'      Else
'         stra1707 = adoacc160.Fields("axg03").Value
'      End If
'      If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
'         stra1708 = Val(FCDate(MaskEdBox1.Text))
'      Else
'         stra1708 = ""
'      End If
'      adoTaie.Execute "insert into acc170 (a1701, a1702, a1703, a1704, a1705, a1706, a1707, a1708, a1710, a1711, a1712) values ('2', '" & Text1 & "', " & CNULL(stra1703) & ", " & CNULL(stra1704) & ", " & CNULL(stra1705) & ", " & CNULL(stra1706) & ", " & CNULL(stra1707) & ", " & CNULL(stra1708) & ", " & strSrvDate(2) & ", " & ServerTime & ", '" & strUserNum & "')"
'      AdodcRefresh
'   Else
'      MsgBox MsgText(28), , MsgText(5)
'   End If
'   adoacc160.Close
'Checking:
'   If Err.Number = 0 Then
'      Exit Sub
'   End If
'   MsgBox Err.Description, , MsgText(5)
'End Sub
'2017/2/21

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)
   Dim bolRefreshGrid As Boolean 'Added by Morgan 2017/2/21
   
On Error GoTo Checking
   Select Case KeyCode
      Case vbKeyInsert
         'Modified by Morgan 2017/2/21
         'Acc170Save
         Acc170SaveNew Me.Text1, bolRefreshGrid
         If bolRefreshGrid Then AdodcRefresh
         'end 2017/2/21
         Text1.Text = "V"
         Text1.SetFocus
         Text1_GotFocus  'ADD BY SONIA 2014/3/19
      Case vbKeyF12
         AdodcRefresh
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
   adoTaie.Execute "delete from acc170 where a1701 = '2' and a1702 = '" & Adodc1.Recordset.Fields("a1702").Value & "'"
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
'

