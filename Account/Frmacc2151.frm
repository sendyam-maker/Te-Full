VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Frmacc2151 
   AutoRedraw      =   -1  'True
   Caption         =   "帳單資料查詢"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5025
   ScaleWidth      =   8760
   Begin VB.ComboBox Combo3 
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
      Left            =   5760
      TabIndex        =   2
      Top             =   240
      Width           =   2772
   End
   Begin VB.ComboBox Combo2 
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
      Left            =   2640
      TabIndex        =   1
      Top             =   240
      Width           =   2772
   End
   Begin VB.ComboBox Combo1 
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
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2052
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc2151.frx":0000
      Height          =   4092
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   8292
      _ExtentX        =   14631
      _ExtentY        =   7223
      _Version        =   393216
      AllowUpdate     =   0   'False
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
      Caption         =   "帳單資料"
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "a1501"
         Caption         =   "帳單編號"
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
         DataField       =   "a1502"
         Caption         =   "帳單日期"
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
         DataField       =   "a1503"
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
      BeginProperty Column03 
         DataField       =   "a1504"
         Caption         =   "代理人D/N編號"
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
         DataField       =   "a1505"
         Caption         =   "幣別"
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
      BeginProperty Column05 
         DataField       =   "a1506"
         Caption         =   "帳單金額"
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
         DataField       =   "a1509"
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
      BeginProperty Column07 
         DataField       =   "a1507"
         Caption         =   "作廢日期"
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
            ColumnWidth     =   1335.118
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   1260.284
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1349.858
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1679.811
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   599.811
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   1530.142
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   4529.764
         EndProperty
         BeginProperty Column07 
            Alignment       =   2
            ColumnWidth     =   1230.236
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   312
      Left            =   240
      Top             =   600
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
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   4800
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "~"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   5520
      TabIndex        =   5
      Top             =   240
      Width           =   132
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "="
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   2400
      TabIndex        =   4
      Top             =   240
      Width           =   132
   End
End
Attribute VB_Name = "Frmacc2151"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/07 改成Form2.0 ; DataGrid1改字型=新細明體-ExtB
'Memo By Sonia 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit
Public adoacc150 As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Dim RQstr As String 'Add by Lydia 2014/10/31
Private Sub Combo1_Change()
'   SelectScope
End Sub

Private Sub Combo1_Click()
'   SelectScope
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Combo2_Validate(Cancel As Boolean)
   Combo3 = Combo2
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
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
'   Me.Height = 5400
'   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
'   Image1 = LoadPicture(strBackPicPath2)
'   sglWidth = Image1.Width
'   sglHeight = Image1.Height
'   For intX = 0 To Int(ScaleWidth / sglWidth)
'       For intY = 0 To Int(ScaleHeight / sglHeight)
'           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
'       Next
'   Next
   strFormName = Name
   PUB_InitForm Me, 8850, 5400, strBackPicPath2
   'end 2021/12/07
   
   strCon1 = "帳單編號"
   strCon2 = "帳單日期"
   strCon3 = "代理人"
   strCon4 = "代理人D/N編號"
   strCon5 = "幣別"
   Combo1.AddItem MsgText(31)
   Combo1.AddItem strCon1
   Combo1.AddItem strCon2
   Combo1.AddItem strCon3
   Combo1.AddItem strCon4
   Combo1.AddItem strCon5
   Combo1 = MsgText(31)
   'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
   'Added by Lydia 2019/09/10 寰華案控制輸入帳單、已提申、發證書輸入，P的程序不能輸入
   If Pub_StrUserSt03 <> "M51" And Pub_StrUserSt03 <> "M31" Then
        FMP2open = PUB_FMPtoCheck(1, 0, Pub_strUserST05, "INVERSE_SQL")
        '內專才控制回傳
        If UCase(App.EXEName) <> "PATPRO" And UCase(App.EXEName) <> "TEPATPRO" Then
            FMP2openSQL = ""
        End If
   Else
   'end 2019/09/10
        FMP2open = PUB_FMPtoCheck(1, 0, Pub_strUserST05)
   End If
   
   SelectScope
   OpenTable
   StatusView MsgText(98)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Adodc1.Recordset.RecordCount <> 0 Then
      strItemNo = Adodc1.Recordset.Fields("a1501").Value
   Else
      strItemNo = MsgText(601)
   End If
   StatusClear
   tool1_enabled
   Select Case strFormLink
      Case "Frmacc2150"
         Frmacc2150.Enabled = True
         Frmacc2150.Show
      Case "Frmacc21j0"
         Frmacc21j0.Enabled = True
         Frmacc21j0.Show
   End Select
   Set Frmacc2151 = Nothing
End Sub

'*************************************************
'  搜尋條件範圍值，並代入 Combo2、Combo3 之中
'
'*************************************************
Private Sub SelectScope()
   'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
   '設別名m0
   strCondition = MsgText(601)
   If Combo1 = MsgText(31) Then
      Exit Sub
   End If
   Select Case Combo1
      Case strCon1
         strCondition = "a1501"
      Case strCon2
         strCondition = "a1502"
      Case strCon3
         strCondition = "a1503"
      Case strCon4
         strCondition = "a1504"
      Case strCon5
         strCondition = "a1505"
   End Select
   If strCondition = MsgText(601) Then
      Exit Sub
   End If
   Combo2.Clear
   Combo3.Clear
   adoacc150.CursorLocation = adUseClient
   adoacc150.Open "select distinct " & strCondition & " from acc150 order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
   Do While adoacc150.EOF = False
      If IsNull(adoacc150.Fields(0).Value) = False Then
         Combo2.AddItem adoacc150.Fields(0).Value
         Combo3.AddItem adoacc150.Fields(0).Value
      End If
      adoacc150.MoveNext
   Loop
   adoacc150.Close
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)
   Select Case KeyCode
      Case vbKeyF12
         'Add by Lydia 2014/11/14 防止未輸入查詢條件
         If Combo1 = "全部" Then
            MsgBox "請輸入有效查詢條件!!", vbCritical
            Exit Sub
         ElseIf Combo2 = "" Or Combo3 = "" Then
            MsgBox "請輸入有效查詢條件!!", vbCritical
            Exit Sub
         End If
         Acc150Query
   End Select
   KeyEnter KeyCode
   StatusView MsgText(98)
End Sub

'*************************************************
'  票據資料查詢
'
'*************************************************
Private Sub Acc150Query()
On Error GoTo Checking
   'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
   '設別名m0
   adoadodc1.Close
   adoadodc1.CursorLocation = adUseClient
   Select Case Combo1
      Case strCon1
         strCondition = "a1501"
      Case strCon2
         strCondition = "a1502"
      Case strCon3
         strCondition = "a1503"
      Case strCon4
         strCondition = "a1504"
      Case strCon5
         strCondition = "a1505"
      Case MsgText(31)
         'Modified by Lydia 2019/09/10  寰華案控制輸入帳單、已提申、發證書輸入，P的程序不能輸入
         'If FMP2open = True Then
         If FMP2openSQL <> "" And Pub_StrUserSt03 <> "M31" And Pub_StrUserSt03 <> "M51" Then
           adoadodc1.Open "select m0.* from acc150 m0 where " & Mid(RQstr, 5, Len(RQstr) - 4) & " order by a1501 asc", adoTaie, adOpenStatic, adLockReadOnly
         Else
           adoadodc1.Open "select m0.* from acc150 m0 order by a1501 asc", adoTaie, adOpenStatic, adLockReadOnly
         End If
         Adodc1.Recordset.Requery
         Exit Sub
      Case Else
         Exit Sub
   End Select
   If Combo3 = MsgText(601) Then
      If Combo1 = strCon2 Then
         adoadodc1.Open "select m0.* from acc150 m0 where " & strCondition & " = " & Val(Combo2) & " " & RQstr & " order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
      Else
         adoadodc1.Open "select m0.* from acc150 m0 where " & strCondition & " = '" & Combo2 & "' " & RQstr & " order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
      End If
   Else
      If Combo2 = MsgText(601) Then
         If Combo1 = strCon2 Then
            adoadodc1.Open "select m0.* from acc150 m0 where " & strCondition & " <= " & Val(Combo3) & " " & RQstr & " order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
         Else
            adoadodc1.Open "select m0.* from acc150 m0 where " & strCondition & " <= '" & Combo3 & "' " & RQstr & " order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
         End If
      Else
         If Combo1 = strCon2 Then
            adoadodc1.Open "select m0.* from acc150 m0 where " & strCondition & " >= " & Val(Combo2) & " and " & strCondition & " <= " & Val(Combo3) & " " & RQstr & " order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
         Else
            adoadodc1.Open "select m0.* from acc150 m0 where " & strCondition & " >= '" & Combo2 & "' and " & strCondition & " <= '" & Combo3 & "' " & RQstr & " order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
         End If
      End If
   End If
   Adodc1.Recordset.Requery
   If Adodc1.Recordset.RecordCount = 0 Then
      MsgBox MsgText(33), , MsgText(5)
   End If
   Exit Sub
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
   adoadodc1.CursorLocation = adUseClient
   'adoadodc1.Open "select m0.* from acc150 m0 where a1501 = '" & Combo2 & "' order by a1501 asc", adoTaie, adOpenStatic, adLockReadOnly
   'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
   '設別名m0
   Dim midSql As String
   midSql = " select m0.* from acc150 m0 where a1501 = '" & Combo2 & "' "
   'Modified by Lydia 2019/09/10  寰華案控制輸入帳單、已提申、發證書輸入，P的程序不能輸入
   'If FMP2open = True Then
   If FMP2openSQL <> "" And Pub_StrUserSt03 <> "M31" And Pub_StrUserSt03 <> "M51" Then
      RQstr = " select m1.axf01 from acc151 m1,caseprogress f0 where m0.a1501=m1.axf01(+) and m1.axf02=f0.cp09(+) " & FMP2openSQL
      midSql = midSql & " and a1501 in (" & RQstr & ") "
      RQstr = " and a1501 in (" & RQstr & ") " 'acc150Query用
   End If
   midSql = midSql & " order by 1 asc "
   adoadodc1.Open midSql, adoTaie, adOpenStatic, adLockReadOnly
   
   Set Adodc1.Recordset = adoadodc1
End Sub

