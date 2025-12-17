VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Frmacc11d1 
   AutoRedraw      =   -1  'True
   Caption         =   "案件進度資料查詢"
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
      ItemData        =   "Frmacc11d1.frx":0000
      Left            =   240
      List            =   "Frmacc11d1.frx":0002
      TabIndex        =   0
      Top             =   240
      Width           =   2052
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc11d1.frx":0004
      Height          =   4215
      Left            =   150
      TabIndex        =   3
      Top             =   720
      Width           =   8445
      _ExtentX        =   14896
      _ExtentY        =   7435
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   18
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
         Name            =   "新細明體-ExtB"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "案件進度資料"
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "CaseNo"
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
      BeginProperty Column01 
         DataField       =   "RDate"
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
         DataField       =   "cp09"
         Caption         =   "收文號"
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
      BeginProperty Column03 
         DataField       =   "CaseProperty"
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
      BeginProperty Column04 
         DataField       =   "a0k20"
         Caption         =   "收據智權人員"
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
      BeginProperty Column05 
         DataField       =   "cp14"
         Caption         =   "承辦人"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "cp16"
         Caption         =   "費用"
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
      BeginProperty Column07 
         DataField       =   "cp17"
         Caption         =   "規費"
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
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         Size            =   344
         BeginProperty Column00 
            ColumnWidth     =   1425.26
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   1049.953
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1289.764
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1244.976
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            ColumnWidth     =   1049.953
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            ColumnWidth     =   989.858
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
Attribute VB_Name = "Frmacc11d1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/14 Form2.0已修改
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/11/26 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit

Public adocase As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset


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
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 8850
   Me.Height = 5400
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath2)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   strCon1 = "收文號"
   strCon2 = "收文日期"
   '2012/8/22 MODIFY BY SONIA
   'strCon3 = "智權人員"
   strCon3 = "收據智權人員"
   strCon4 = "承辦人"
   strCon5 = "本所案號"
   Combo1.AddItem MsgText(31)
   Combo1.AddItem strCon1
   Combo1.AddItem strCon2
   Combo1.AddItem strCon3
   Combo1.AddItem strCon4
   Combo1.AddItem strCon5
   Combo1 = MsgText(31)
   SelectScope
   OpenTable
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Adodc1.Recordset.RecordCount <> 0 Then
      strItemNo = Adodc1.Recordset.Fields("cp09").Value
   Else
      strItemNo = MsgText(601)
   End If
   StatusClear
   'Modify by Morgan 2004/1/12
   'tool8_enabled
   tool14_enabled
   'Modify end------------------
   Frmacc11d0.Enabled = True
   Frmacc11d0.Show
   Set Frmacc11d1 = Nothing
End Sub

'*************************************************
'  搜尋條件範圍值，並代入 Combo2、Combo3 之中
'
'*************************************************
Private Sub SelectScope()
   strCondition = MsgText(601)
   If Combo1 = MsgText(31) Then
      Exit Sub
   End If
   Select Case Combo1
      Case strCon1
         strCondition = "cp09"
      Case strCon2
         strCondition = "cp05"
      Case strCon3
         '2012/8/23 MODIFY BY SONIA
         'strCondition = "cp13"
         strCondition = "a0k20"
      Case strCon4
         strCondition = "cp14"
      Case strCon5
         strCondition = "cp01||cp02||cp03||cp04"
   End Select
   If strCondition = MsgText(601) Then
      Exit Sub
   End If
   Combo2.Clear
   Combo3.Clear
   adocase.CursorLocation = adUseClient
   adocase.Open "select distinct " & strCondition & " from caseprogress order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
   Do While adocase.EOF = False
      If IsNull(adocase.Fields(0).Value) = False Then
         Combo2.AddItem adocase.Fields(0).Value
         Combo3.AddItem adocase.Fields(0).Value
      End If
      adocase.MoveNext
   Loop
   adocase.Close
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)
   Select Case KeyCode
      Case vbKeyF12
         CaseQuery
   End Select
   KeyEnter KeyCode
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

'*************************************************
'  票據資料查詢
'
'*************************************************
Private Sub CaseQuery()
On Error GoTo Checking
   adoadodc1.Close
   adoadodc1.CursorLocation = adUseClient
   Select Case Combo1
      Case strCon1
         strCondition = "cp09"
      Case strCon2
         strCondition = "cp05"
      Case strCon3
         '2012/8/23 MODIFY BY SONIA
         'strCondition = "cp13"
         strCondition = "a0k20"
      Case strCon4
         strCondition = "cp14"
      Case strCon5
         strCondition = "cp01||cp02||cp03||cp04"
      Case MsgText(31)
         '2012/8/23 MODIFY BY SONIA 智權人員改先抓收據智權人員並加第二,三排序條件收文日,本所案號
         'adoadodc1.Open "select cp01||cp02||cp03||cp04 as CaseNo, cp09, cp05 - 19110000 as RDate, cp13, cp14, cp16, cp17, nvl(cpm03, cpm04) as CaseProperty from caseprogress, casepropertymap where cp01 = cpm01 (+) and cp10 = cpm02 (+) and substr(cp01, 1, 2) <> 'FC' and (cp16 <> 0 and cp16 is not null) and cp16 <> cp17 order by cp09 asc", adoTaie, adOpenStatic, adLockReadOnly
         adoadodc1.Open "select cp01||cp02||cp03||cp04 as CaseNo, cp09, cp05 - 19110000 as RDate, a0k20, cp14, cp16, cp17, nvl(cpm03, cpm04) as CaseProperty from caseprogress, casepropertymap, acc0k0 where cp01 = cpm01 (+) and cp10 = cpm02 (+) and substr(cp01, 1, 2) <> 'FC' and (cp16 <> 0 and cp16 is not null) and cp16 <> cp17 and cp60=a0k01(+) order by cp09 asc, cp05, cp01||cp02||cp03||cp04", adoTaie, adOpenStatic, adLockReadOnly
         Adodc1.Recordset.Requery
         Exit Sub
      Case Else
         Exit Sub
   End Select
   If Combo3 = MsgText(601) Then
      If Combo1 = strCon2 Then
         '2012/8/23 MODIFY BY SONIA 智權人員改先抓收據智權人員
         'adoadodc1.Open "select cp01||cp02||cp03||cp04 as CaseNo, cp09, cp05 - 19110000 as RDate, cp13, cp14, cp16, cp17, nvl(cpm03, cpm04) as CaseProperty from caseprogress, casepropertymap where cp01 = cpm01 (+) and cp10 = cpm02 (+) and substr(cp01, 1, 2) <> 'FC' and (cp16 <> 0 and cp16 is not null) and cp16 <> cp17 and " & strCondition & " = " & Val(Combo2) & " order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
         adoadodc1.Open "select cp01||cp02||cp03||cp04 as CaseNo, cp09, cp05 - 19110000 as RDate, a0k20, cp14, cp16, cp17, nvl(cpm03, cpm04) as CaseProperty from caseprogress, casepropertymap, acc0k0 where cp01 = cpm01 (+) and cp10 = cpm02 (+) and substr(cp01, 1, 2) <> 'FC' and (cp16 <> 0 and cp16 is not null) and cp16 <> cp17 and cp60=a0k01(+) and " & strCondition & " = " & Val(Combo2) + 19110000 & " order by " & strCondition & " asc, cp01||cp02||cp03||cp04", adoTaie, adOpenStatic, adLockReadOnly
      Else
         '2012/8/23 MODIFY BY SONIA 智權人員改先抓收據智權人員
         'adoadodc1.Open "select cp01||cp02||cp03||cp04 as CaseNo, cp09, cp05 - 19110000 as RDate, cp13, cp14, cp16, cp17, nvl(cpm03, cpm04) as CaseProperty from caseprogress, casepropertymap where cp01 = cpm01 (+) and cp10 = cpm02 (+) and substr(cp01, 1, 2) <> 'FC' and (cp16 <> 0 and cp16 is not null) and cp16 <> cp17 and " & strCondition & " = '" & Combo2 & "' order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
         adoadodc1.Open "select cp01||cp02||cp03||cp04 as CaseNo, cp09, cp05 - 19110000 as RDate, a0k20, cp14, cp16, cp17, nvl(cpm03, cpm04) as CaseProperty from caseprogress, casepropertymap, acc0k0 where cp01 = cpm01 (+) and cp10 = cpm02 (+) and substr(cp01, 1, 2) <> 'FC' and (cp16 <> 0 and cp16 is not null) and cp16 <> cp17 and cp60=a0k01(+) and " & strCondition & " = '" & Combo2 & "' order by " & strCondition & " asc, cp05, cp01||cp02||cp03||cp04", adoTaie, adOpenStatic, adLockReadOnly
      End If
   Else
      If Combo2 = MsgText(601) Then
         If Combo1 = strCon2 Then
            '2012/8/23 MODIFY BY SONIA 智權人員改先抓收據智權人員
            'adoadodc1.Open "select cp01||cp02||cp03||cp04 as CaseNo, cp09, cp05 - 19110000 as RDate, cp13, cp14, cp16, cp17, nvl(cpm03, cpm04) as CaseProperty from caseprogress, casepropertymap where cp01 = cpm01 (+) and cp10 = cpm02 (+) and substr(cp01, 1, 2) <> 'FC' and (cp16 <> 0 and cp16 is not null) and cp16 <> cp17 and " & strCondition & " <= " & Val(Combo3) & " order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
            adoadodc1.Open "select cp01||cp02||cp03||cp04 as CaseNo, cp09, cp05 - 19110000 as RDate, a0k20, cp14, cp16, cp17, nvl(cpm03, cpm04) as CaseProperty from caseprogress, casepropertymap, acc0k0 where cp01 = cpm01 (+) and cp10 = cpm02 (+) and substr(cp01, 1, 2) <> 'FC' and (cp16 <> 0 and cp16 is not null) and cp16 <> cp17 and cp60=a0k01(+) and " & strCondition & " <= " & Val(Combo3) + 19110000 & " order by " & strCondition & " asc, cp01||cp02||cp03||cp04", adoTaie, adOpenStatic, adLockReadOnly
         Else
            '2012/8/23 MODIFY BY SONIA 智權人員改先抓收據智權人員
            'adoadodc1.Open "select cp01||cp02||cp03||cp04 as CaseNo, cp09, cp05 - 19110000 as RDate, cp13, cp14, cp16, cp17, nvl(cpm03, cpm04) as CaseProperty from caseprogress, casepropertymap where cp01 = cpm01 (+) and cp10 = cpm02 (+) and substr(cp01, 1, 2) <> 'FC' and (cp16 <> 0 and cp16 is not null) and cp16 <> cp17 and " & strCondition & " <= '" & Combo3 & "' order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
            adoadodc1.Open "select cp01||cp02||cp03||cp04 as CaseNo, cp09, cp05 - 19110000 as RDate, a0k20, cp14, cp16, cp17, nvl(cpm03, cpm04) as CaseProperty from caseprogress, casepropertymap, acc0k0 where cp01 = cpm01 (+) and cp10 = cpm02 (+) and substr(cp01, 1, 2) <> 'FC' and (cp16 <> 0 and cp16 is not null) and cp16 <> cp17 and cp60=a0k01(+) and " & strCondition & " <= '" & Combo3 & "' order by " & strCondition & " asc, cp05, cp01||cp02||cp03||cp04", adoTaie, adOpenStatic, adLockReadOnly
         End If
      Else
         If Combo1 = strCon2 Then
            '2012/8/23 MODIFY BY SONIA 智權人員改先抓收據智權人員
            'adoadodc1.Open "select cp01||cp02||cp03||cp04 as CaseNo, cp09, cp05 - 19110000 as RDate, cp13, cp14, cp16, cp17, nvl(cpm03, cpm04) as CaseProperty from caseprogress, casepropertymap where cp01 = cpm01 (+) and cp10 = cpm02 (+) and substr(cp01, 1, 2) <> 'FC' and (cp16 <> 0 and cp16 is not null) and cp16 <> cp17 and " & strCondition & " >= " & Val(Combo2) & " and " & strCondition & " <= " & Val(Combo3) & " order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
            adoadodc1.Open "select cp01||cp02||cp03||cp04 as CaseNo, cp09, cp05 - 19110000 as RDate, a0k20, cp14, cp16, cp17, nvl(cpm03, cpm04) as CaseProperty from caseprogress, casepropertymap, acc0k0 where cp01 = cpm01 (+) and cp10 = cpm02 (+) and substr(cp01, 1, 2) <> 'FC' and (cp16 <> 0 and cp16 is not null) and cp16 <> cp17 and cp60=a0k01(+) and " & strCondition & " >= " & Val(Combo2) + 19110000 & " and " & strCondition & " <= " & Val(Combo3) + 19110000 & " order by " & strCondition & " asc, cp01||cp02||cp03||cp04", adoTaie, adOpenStatic, adLockReadOnly
         Else
            '2012/8/23 MODIFY BY SONIA 智權人員改先抓收據智權人員
            'adoadodc1.Open "select cp01||cp02||cp03||cp04 as CaseNo, cp09, cp05 - 19110000 as RDate, cp13, cp14, cp16, cp17, nvl(cpm03, cpm04) as CaseProperty from caseprogress, casepropertymap where cp01 = cpm01 (+) and cp10 = cpm02 (+) and substr(cp01, 1, 2) <> 'FC' and (cp16 <> 0 and cp16 is not null) and cp16 <> cp17 and " & strCondition & " >= '" & Combo2 & "' and " & strCondition & " <= '" & Combo3 & "' order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
            adoadodc1.Open "select cp01||cp02||cp03||cp04 as CaseNo, cp09, cp05 - 19110000 as RDate, a0k20, cp14, cp16, cp17, nvl(cpm03, cpm04) as CaseProperty from caseprogress, casepropertymap, acc0k0 where cp01 = cpm01 (+) and cp10 = cpm02 (+) and substr(cp01, 1, 2) <> 'FC' and (cp16 <> 0 and cp16 is not null) and cp16 <> cp17 and cp60=a0k01(+) and " & strCondition & " >= '" & Combo2 & "' and " & strCondition & " <= '" & Combo3 & "' order by " & strCondition & " asc, cp05, cp01||cp02||cp03||cp04", adoTaie, adOpenStatic, adLockReadOnly
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
On Error GoTo Checking
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "select * from caseprogress where cp09 = '" & Combo2 & "' order by cp09 asc", adoTaie, adOpenStatic, adLockReadOnly
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

