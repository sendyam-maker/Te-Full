VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Frmacc41d1 
   AutoRedraw      =   -1  'True
   Caption         =   "應收付分錄調整查詢"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
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
      Height          =   312
      Left            =   5772
      TabIndex        =   2
      Top             =   204
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
      Height          =   312
      Left            =   2652
      TabIndex        =   1
      Top             =   204
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
      Height          =   312
      Left            =   252
      TabIndex        =   0
      Top             =   204
      Width           =   2052
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc41d1.frx":0000
      Height          =   4092
      Left            =   252
      TabIndex        =   3
      Top             =   684
      Width           =   8292
      _ExtentX        =   14631
      _ExtentY        =   7223
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   18
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
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "a1p01"
         Caption         =   "公司別"
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
      BeginProperty Column01 
         DataField       =   "a1p04"
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
      BeginProperty Column02 
         DataField       =   "a1p18"
         Caption         =   "入帳日期"
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
         DataField       =   "a0102"
         Caption         =   "會計科目"
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
         DataField       =   "a1p07"
         Caption         =   "借方金額"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "a1p08"
         Caption         =   "貸方金額"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "a0902"
         Caption         =   "部門別"
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
         Size            =   275
         BeginProperty Column00 
            ColumnWidth     =   780.095
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1289.764
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   989.858
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1679.811
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column06 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   312
      Left            =   252
      Top             =   564
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
   Begin VB.Image Image1 
      Height          =   132
      Left            =   12
      Top             =   4764
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
      Left            =   5532
      TabIndex        =   5
      Top             =   204
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
      Left            =   2412
      TabIndex        =   4
      Top             =   204
      Width           =   132
   End
End
Attribute VB_Name = "Frmacc41d1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/10/25 Form2.0已修改 DataGrid1
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/3 日期欄已修改
Option Explicit
Public adoacc020 As New ADODB.Recordset
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
   strCon1 = "公司別"
   strCon2 = "單據編號(含客戶抬頭)"
   strCon3 = "入帳日期"
   strCon4 = "部門別"
   strCon5 = "會計科目"
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
      strCompanyNo = Adodc1.Recordset.Fields("a1p01").Value
      strItemNo = Adodc1.Recordset.Fields("a1p04").Value
      'Modified by Lydia 2019/11/07 瑞婷反應,有時候點到"應收付分錄調整Frmacc41d0"的程式會出錯
      'strExc(0) = Adodc1.Recordset.Fields("a1p02").Value 'Add by Amy 2014/02/07
      Frmacc41d0.stra1p02 = "" & Adodc1.Recordset.Fields("a1p02").Value
   Else
      strCompanyNo = MsgText(601)
      strItemNo = MsgText(601)
      'Modified by Lydia 2019/11/07
      'strExc(0) = MsgText(601) 'Add by Amy 2014/02/07
      Frmacc41d0.stra1p02 = MsgText(601)
   End If
   StatusClear
   tool1_enabled
   Frmacc41d0.Enabled = True
   Frmacc41d0.Show
   Set Frmacc41d1 = Nothing
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
         strCondition = "a1p01"
      Case strCon2
         strCondition = "a1p04"
      Case strCon3
         strCondition = "a1p18"
      Case strCon4
         strCondition = "a1p06"
      Case strCon5
         strCondition = "a1p05"
   End Select
   If strCondition = MsgText(601) Then
      Exit Sub
   End If
   Combo2.Clear
   Combo3.Clear
   adoacc020.CursorLocation = adUseClient
   adoacc020.Open "select distinct " & strCondition & " from acc1p0 order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
   Do While adoacc020.EOF = False
      If IsNull(adoacc020.Fields(0).Value) = False Then
         Combo2.AddItem adoacc020.Fields(0).Value
         Combo3.AddItem adoacc020.Fields(0).Value
      End If
      adoacc020.MoveNext
   Loop
   adoacc020.Close
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)
   Select Case KeyCode
      Case vbKeyF12
         Screen.MousePointer = vbHourglass
         Acc020Query
         Screen.MousePointer = vbDefault
   End Select
   KeyEnter KeyCode
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

'*************************************************
'  傳票資料查詢
'
'*************************************************
Private Sub Acc020Query()
On Error GoTo Checking
   If adoadodc1.State = adStateOpen Then
      adoadodc1.Close
   End If
   adoadodc1.CursorLocation = adUseClient
   Select Case Combo1
      Case strCon1
         strCondition = "a1p01"
      Case strCon2
         strCondition = "a1p04"
      Case strCon3
         strCondition = "a1p18"
      Case strCon4
         strCondition = "a1p06"
      Case strCon5
         strCondition = "a1p05"
      Case MsgText(31)
         adoadodc1.Open "select * from acc1p0, acc090, acc010 where a1p06 = a0901 (+) and a1p05 = a0101 and a1p02 in ('Z', 'E', 'W', 'L', 'G') order by a1p01 asc, a1p18 asc, a1p04 asc", adoTaie, adOpenStatic, adLockReadOnly
         Adodc1.Recordset.Requery
         Exit Sub
      Case Else
         Exit Sub
   End Select
   If Combo3 = MsgText(601) Then
      If Combo1 = strCon3 Then
         adoadodc1.Open "select * from acc1p0, acc090, acc010 where a1p06 = a0901 (+) and a1p05 = a0101 and a1p02 in ('Z', 'E', 'W', 'L', 'G') and " & strCondition & " = " & Val(Combo2) & " order by a1p01 asc, a1p18 asc, a1p04 asc", adoTaie, adOpenStatic, adLockReadOnly
      Else
         adoadodc1.Open "select * from acc1p0, acc090, acc010 where a1p06 = a0901 (+) and a1p05 = a0101 and a1p02 in ('Z', 'E', 'W', 'L', 'G') and instr(" & strCondition & ", '" & Combo2 & "') = 1 order by a1p01 asc, a1p18 asc, a1p04 asc", adoTaie, adOpenStatic, adLockReadOnly
      End If
   Else
      If Combo2 = MsgText(601) Then
         If Combo1 = strCon3 Then
            adoadodc1.Open "select * from acc1p0, acc090, acc010 where a1p06 = a0901 (+) and a1p05 = a0101 and a1p02 in ('Z', 'E', 'W', 'L', 'G') and " & strCondition & " <= " & Val(Combo3) & " order by a1p01 asc, a1p18 asc, a1p04 asc", adoTaie, adOpenStatic, adLockReadOnly
         Else
            adoadodc1.Open "select * from acc1p0, acc090, acc010 where a1p06 = a0901 (+) and a1p05 = a0101 and a1p02 in ('Z', 'E', 'W', 'L', 'G') and " & strCondition & " <= '" & Combo3 & "' order by a1p01 asc, a1p18 asc, a1p04 asc", adoTaie, adOpenStatic, adLockReadOnly
         End If
      Else
         If Combo1 = strCon3 Then
            adoadodc1.Open "select * from acc1p0, acc090, acc010 where a1p06 = a0901 (+) and a1p05 = a0101 and a1p02 in ('Z', 'E', 'W', 'L', 'G') and " & strCondition & " >= " & Val(Combo2) & " and " & strCondition & " <= " & Val(Combo3) & " order by a1p01 asc, a1p18 asc, a1p04 asc", adoTaie, adOpenStatic, adLockReadOnly
         Else
            adoadodc1.Open "select * from acc1p0, acc090, acc010 where a1p06 = a0901 (+) and a1p05 = a0101 and a1p02 in ('Z', 'E', 'W', 'L', 'G') and " & strCondition & " >= '" & Combo2 & "' and " & strCondition & " <= '" & Combo3 & "' order by a1p01 asc, a1p18 asc, a1p04 asc", adoTaie, adOpenStatic, adLockReadOnly
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
   adoadodc1.Open "select * from acc1p0 where a1p04 = 'ZZ' order by a1p01 asc, a1p04 asc, a1p18 asc", adoTaie, adOpenStatic, adLockReadOnly
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

