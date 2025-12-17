VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Frmacc11a1 
   AutoRedraw      =   -1  'True
   Caption         =   "暫收款資料查詢"
   ClientHeight    =   5030
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   8760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5030
   ScaleWidth      =   8760
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
      Height          =   312
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2052
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
      Height          =   312
      Left            =   2640
      TabIndex        =   1
      Top             =   240
      Width           =   2772
   End
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
      Height          =   312
      Left            =   5760
      TabIndex        =   2
      Top             =   240
      Width           =   2772
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc11a1.frx":0000
      Height          =   4215
      Left            =   150
      TabIndex        =   3
      Top             =   720
      Width           =   8445
      _ExtentX        =   14887
      _ExtentY        =   7444
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   18
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "暫收款資料"
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "a0t01"
         Caption         =   "暫收款單號"
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
      BeginProperty Column01 
         DataField       =   "a0t03"
         Caption         =   "輸入日期"
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
         DataField       =   "a0t04"
         Caption         =   "欲處理日期"
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
         DataField       =   "a0t08"
         Caption         =   "暫收款金額"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "st02"
         Caption         =   "智權人員"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "cu04"
         Caption         =   "客戶名稱"
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
         DataField       =   "a0t07"
         Caption         =   "單據編號"
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
            ColumnWidth     =   1409.953
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   1230.236
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   1280.126
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1429.795
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   3309.732
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1289.764
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
      TabIndex        =   5
      Top             =   240
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
      TabIndex        =   4
      Top             =   240
      Width           =   132
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   4800
      Visible         =   0   'False
      Width           =   132
   End
End
Attribute VB_Name = "Frmacc11a1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/14 Form2.0已修改
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/11/26 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit

Public adoacc0t0 As New ADODB.Recordset
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
   If Combo1 = "客戶編號" Then
      Select Case Len(Combo2)
         Case 6
            Combo2 = Combo2 & "000"
         Case 8
            Combo2 = Combo2 & "0"
      End Select
   End If
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
   strCon1 = "暫收款單號"
   strCon2 = "輸入日期"
   strCon3 = "欲處理日期"
   strCon4 = "智權人員"
   strCon5 = "客戶編號"
   strCon6 = "單據編號"
   Combo1.AddItem MsgText(31)
   Combo1.AddItem strCon1
   Combo1.AddItem strCon2
   Combo1.AddItem strCon3
   Combo1.AddItem strCon4
   Combo1.AddItem strCon5
   Combo1.AddItem strCon6
   Combo1 = MsgText(31)
   SelectScope
   OpenTable
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Adodc1.Recordset.RecordCount <> 0 Then
      strItemNo = Adodc1.Recordset.Fields("a0t01").Value
   Else
      strItemNo = MsgText(601)
   End If
   StatusClear
   tool1_enabled
   Frmacc11a0.Enabled = True
   Frmacc11a0.Show
   Set Frmacc11a1 = Nothing
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
         strCondition = "a0t01"
      Case strCon2
         strCondition = "a0t03"
      Case strCon3
         strCondition = "a0t04"
      Case strCon4
         strCondition = "a0t05"
      Case strCon5
         strCondition = "a0t06"
      Case strCon6
         strCondition = "a0t07"
   End Select
   If strCondition = MsgText(601) Then
      Exit Sub
   End If
   Combo2.Clear
   Combo3.Clear
   adoacc0t0.CursorLocation = adUseClient
   adoacc0t0.Open "select distinct " & strCondition & " from acc0t0 order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
   Do While adoacc0t0.EOF = False
      If IsNull(adoacc0t0.Fields(0).Value) = False Then
         Combo2.AddItem adoacc0t0.Fields(0).Value
         Combo3.AddItem adoacc0t0.Fields(0).Value
      End If
      adoacc0t0.MoveNext
   Loop
   adoacc0t0.Close
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)
   Select Case KeyCode
      Case vbKeyF12
         Screen.MousePointer = vbHourglass
         Acc0t0Query
         Screen.MousePointer = vbDefault
   End Select
   KeyEnter KeyCode
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

'*************************************************
'  票據資料查詢
'
'*************************************************
Private Sub Acc0t0Query()
On Error GoTo Checking
   adoadodc1.Close
   adoadodc1.CursorLocation = adUseClient
   Select Case Combo1
      Case strCon1
         strCondition = "a0t01"
      Case strCon2
         strCondition = "a0t03"
      Case strCon3
         strCondition = "a0t04"
      Case strCon4
         strCondition = "a0t05"
      Case strCon5
         strCondition = "a0t06"
      Case strCon6
         strCondition = "a0t07"
      Case MsgText(31)
         If Frmacc11a0.Option1.Value Then
            '2007/10/30 modify by sonia 因J09400660轉國外收款沖到,故A1P02再加'F'
            'modify by sonia 2021/8/23 很多暫收款財務處自行以總帳傳票沖銷,故人工上A0T10以區別,故加入A0T10為判斷條件
            'modify by sonia 2024/5/14 a1p02加入'E'，因為a1p04='聯米企業股份有限公司971'有沖銷J09700301
            adoadodc1.Open "select * from acc0t0, staff, customer where a0t05 = st01 (+) and substr(a0t06, 1, 8) = cu01 (+)" & _
               " and substr(a0t06, 9, 1) = cu02 (+) and a0t10 is null and a0t01 not in (select a1p23 from acc1p0 where a1p02 in ('A', 'Z', 'W', 'F', 'E')" & _
               " and a1p05 = '2401' and a1p07 <> 0 and a1p23 is not null) and a0t01 not in (select a0s02 from acc0s0" & _
               " where substr(a0s02, 1, 1) = 'J') order by a0t01 asc", adoTaie, adOpenStatic, adLockReadOnly
         Else
            adoadodc1.Open "select * from acc0t0, staff, customer where a0t05 = st01 (+) and substr(a0t06, 1, 8) = cu01 (+)" & _
               " and substr(a0t06, 9, 1) = cu02 (+) order by a0t01 asc", adoTaie, adOpenStatic, adLockReadOnly
         End If
         Adodc1.Recordset.Requery
         Exit Sub
      Case Else
         Exit Sub
   End Select
   If Frmacc11a0.Option1.Value Then
      If Combo3 = MsgText(601) Then
         If Combo1 = strCon2 Or Combo1 = strCon3 Then
            'modify by sonia 2021/8/23 很多暫收款財務處自行以總帳傳票沖銷,故人工上A0T10以區別,故加入A0T10為判斷條件
            'modify by sonia 2024/5/14 a1p02加入'E'，因為a1p04='聯米企業股份有限公司971'有沖銷J09700301
            adoadodc1.Open "select * from acc0t0, staff, customer where a0t05 = st01 (+) and substr(a0t06, 1, 8) = cu01 (+)" & _
               " and substr(a0t06, 9, 1) = cu02 (+) and a0t10 is null and a0t01 not in (select a1p23 from acc1p0 where a1p02 in ('A', 'Z', 'W', 'F', 'E')" & _
               " and a1p05 = '2401' and a1p07 <> 0 and a1p23 is not null) and a0t01 not in (select a0s02 from acc0s0" & _
               " where substr(a0s02, 1, 1) = 'J') and " & strCondition & " = " & Val(Combo2) & " order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
         Else
            'modify by sonia 2021/8/23 很多暫收款財務處自行以總帳傳票沖銷,故人工上A0T10以區別,故加入A0T10為判斷條件
            'modify by sonia 2024/5/14 a1p02加入'E'，因為a1p04='聯米企業股份有限公司971'有沖銷J09700301
            adoadodc1.Open "select * from acc0t0, staff, customer where a0t05 = st01 (+) and substr(a0t06, 1, 8) = cu01 (+)" & _
               " and substr(a0t06, 9, 1) = cu02 (+) and a0t10 is null and a0t01 not in (select a1p23 from acc1p0 where a1p02 in ('A', 'Z', 'W', 'F', 'E')" & _
               " and a1p05 = '2401' and a1p07 <> 0 and a1p23 is not null) and a0t01 not in (select a0s02 from acc0s0" & _
               " where substr(a0s02, 1, 1) = 'J') and " & strCondition & " = '" & Combo2 & "' order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
         End If
      Else
         If Combo2 = MsgText(601) Then
            If Combo1 = strCon2 Or Combo1 = strCon3 Then
               'modify by sonia 2021/8/23 很多暫收款財務處自行以總帳傳票沖銷,故人工上A0T10以區別,故加入A0T10為判斷條件
               'modify by sonia 2024/5/14 a1p02加入'E'，因為a1p04='聯米企業股份有限公司971'有沖銷J09700301
               adoadodc1.Open "select * from acc0t0, staff, customer where a0t05 = st01 (+) and substr(a0t06, 1, 8) = cu01 (+)" & _
                  " and substr(a0t06, 9, 1) = cu02 (+) and a0t10 is null and a0t01 not in (select a1p23 from acc1p0 where a1p02 in ('A', 'Z', 'W', 'F', 'E')" & _
                  " and a1p05 = '2401' and a1p07 <> 0 and a1p23 is not null) and a0t01 not in (select a0s02 from acc0s0" & _
                  " where substr(a0s02, 1, 1) = 'J') and " & strCondition & " <= " & Val(Combo3) & " order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
            Else
               'modify by sonia 2021/8/23 很多暫收款財務處自行以總帳傳票沖銷,故人工上A0T10以區別,故加入A0T10為判斷條件
               'modify by sonia 2024/5/14 a1p02加入'E'，因為a1p04='聯米企業股份有限公司971'有沖銷J09700301
               adoadodc1.Open "select * from acc0t0, staff, customer where a0t05 = st01 (+) and substr(a0t06, 1, 8) = cu01 (+)" & _
                  " and substr(a0t06, 9, 1) = cu02 (+) and a0t10 is null and a0t01 not in (select a1p23 from acc1p0 where a1p02 in ('A', 'Z', 'W', 'F', 'E')" & _
                  " and a1p05 = '2401' and a1p07 <> 0 and a1p23 is not null) and a0t01 not in (select a0s02 from acc0s0" & _
                  " where substr(a0s02, 1, 1) = 'J') and " & strCondition & " <= '" & Combo3 & "' order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
            End If
         Else
            If Combo1 = strCon2 Or Combo1 = strCon3 Then
               'modify by sonia 2021/8/23 很多暫收款財務處自行以總帳傳票沖銷,故人工上A0T10以區別,故加入A0T10為判斷條件
               'modify by sonia 2024/5/14 a1p02加入'E'，因為a1p04='聯米企業股份有限公司971'有沖銷J09700301
               adoadodc1.Open "select * from acc0t0, staff, customer where a0t05 = st01 (+) and substr(a0t06, 1, 8) = cu01 (+)" & _
                  " and substr(a0t06, 9, 1) = cu02 (+) and a0t10 is null and a0t01 not in (select a1p23 from acc1p0 where a1p02 in ('A', 'Z', 'W', 'F', 'E')" & _
                  " and a1p05 = '2401' and a1p07 <> 0 and a1p23 is not null) and a0t01 not in (select a0s02 from acc0s0" & _
                  " where substr(a0s02, 1, 1) = 'J') and " & strCondition & " >= " & Val(Combo2) & " and " & _
                  strCondition & " <= " & Val(Combo3) & " order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
            Else
               'modify by sonia 2021/8/23 很多暫收款財務處自行以總帳傳票沖銷,故人工上A0T10以區別,故加入A0T10為判斷條件,並加入,a0t01 desc 排序
               'modify by sonia 2024/5/14 a1p02加入'E'，因為a1p04='聯米企業股份有限公司971'有沖銷J09700301
               adoadodc1.Open "select * from acc0t0, staff, customer where a0t05 = st01 (+) and substr(a0t06, 1, 8) = cu01 (+)" & _
                  " and substr(a0t06, 9, 1) = cu02 (+) and a0t10 is null and a0t01 not in (select a1p23 from acc1p0 where a1p02 in ('A', 'Z', 'W', 'F', 'E')" & _
                  " and a1p05 = '2401' and a1p07 <> 0 and a1p23 is not null) and a0t01 not in (select a0s02 from acc0s0" & _
                  " where substr(a0s02, 1, 1) = 'J') and " & strCondition & " >= '" & Combo2 & "' and " & strCondition & _
                  " <= '" & Combo3 & "' order by " & strCondition & " asc,a0t01 desc", adoTaie, adOpenStatic, adLockReadOnly
            End If
         End If
      End If
   Else
      If Combo3 = MsgText(601) Then
         If Combo1 = strCon2 Or Combo1 = strCon3 Then
            adoadodc1.Open "select * from acc0t0, staff, customer where a0t05 = st01 (+) and substr(a0t06, 1, 8) = cu01 (+)" & _
               " and substr(a0t06, 9, 1) = cu02 (+) and " & strCondition & " = " & Val(Combo2) & " order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
         Else
            adoadodc1.Open "select * from acc0t0, staff, customer where a0t05 = st01 (+) and substr(a0t06, 1, 8) = cu01 (+)" & _
               " and substr(a0t06, 9, 1) = cu02 (+) and " & strCondition & " = '" & Combo2 & "' order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
         End If
      Else
         If Combo2 = MsgText(601) Then
            If Combo1 = strCon2 Or Combo1 = strCon3 Then
               adoadodc1.Open "select * from acc0t0, staff, customer where a0t05 = st01 (+) and substr(a0t06, 1, 8) = cu01 (+)" & _
                  " and substr(a0t06, 9, 1) = cu02 (+) and " & strCondition & " <= " & Val(Combo3) & " order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
            Else
               adoadodc1.Open "select * from acc0t0, staff, customer where a0t05 = st01 (+) and substr(a0t06, 1, 8) = cu01 (+)" & _
                  " and substr(a0t06, 9, 1) = cu02 (+) and " & strCondition & " <= '" & Combo3 & "' order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
            End If
         Else
            If Combo1 = strCon2 Or Combo1 = strCon3 Then
               adoadodc1.Open "select * from acc0t0, staff, customer where a0t05 = st01 (+) and substr(a0t06, 1, 8) = cu01 (+)" & _
                  " and substr(a0t06, 9, 1) = cu02 (+) and " & strCondition & " >= " & Val(Combo2) & " and " & strCondition & _
                  " <= " & Val(Combo3) & " order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
            Else
               adoadodc1.Open "select * from acc0t0, staff, customer where a0t05 = st01 (+) and substr(a0t06, 1, 8) = cu01 (+)" & _
                  " and substr(a0t06, 9, 1) = cu02 (+) and " & strCondition & " >= '" & Combo2 & "' and " & _
                  strCondition & " <= '" & Combo3 & "' order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
            End If
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
   adoadodc1.Open "select * from acc0t0 where a0t01 = '" & Combo2 & "' order by a0t01 asc", adoTaie, adOpenStatic, adLockReadOnly
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub


