VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Frmacc2112 
   AutoRedraw      =   -1  'True
   Caption         =   "收款資料查詢"
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
      Height          =   330
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
      Height          =   330
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
      Height          =   330
      ItemData        =   "Frmacc2112.frx":0000
      Left            =   240
      List            =   "Frmacc2112.frx":0002
      TabIndex        =   0
      Top             =   240
      Width           =   2052
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc2112.frx":0004
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
      Caption         =   "收款資料"
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "a0y01"
         Caption         =   "收款單號"
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
         DataField       =   "a0y02"
         Caption         =   "收款日期"
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
         DataField       =   "a0y03"
         Caption         =   "幣別"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "####-####"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Amount"
         Caption         =   "收款金額"
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
      BeginProperty Column04 
         DataField       =   "a0y04"
         Caption         =   "匯率"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.0000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "a0y07"
         Caption         =   "代理人1"
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
         DataField       =   "a0y08"
         Caption         =   "代理人2"
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
         DataField       =   "a0y09"
         Caption         =   "代理人3"
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
      BeginProperty Column08 
         DataField       =   "A1P23"
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
         BeginProperty Column00 
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   599.811
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   975.118
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1289.764
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1230.236
         EndProperty
         BeginProperty Column08 
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
Attribute VB_Name = "Frmacc2112"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/03 改成Form2.0 ; DataGrid1改字型=新細明體-ExtB
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit
Public adoacc0y0 As New ADODB.Recordset
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
   '2009/6/26 ADD BY SONIA
   If (Combo1 = "代理人1" Or Combo1 = "代理人2" Or Combo1 = "代理人3") And Len(Combo2) = 6 Then
      Combo2 = AfterZero(Combo2)
   End If
   '2009/6/26 END
   Combo3 = Combo2
   '2009/6/2 ADD BY SONIA 預設尾碼999
   'Modify By Sindy 2014/8/11 999=>ZZZ
   'If Combo2 <> "" And (Combo1 = strCon3 Or Combo1 = strCon4 Or Combo1 = strCon5) Then Combo3 = Left(Combo2, 6) & "999"
   If Combo2 <> "" And (Combo1 = strCon3 Or Combo1 = strCon4 Or Combo1 = strCon5) Then Combo3 = Left(Combo2, 6) & "ZZZ"
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
   strCon1 = "收款單號"
   strCon2 = "收款日期"
   strCon3 = "代理人1"
   strCon4 = "代理人2"
   strCon5 = "代理人3"
   strCon6 = "單據編號"      '2006/3/8 ADD BY SONIA
   Combo1.AddItem MsgText(31)
   Combo1.AddItem strCon1
   Combo1.AddItem strCon2
   Combo1.AddItem strCon3
   Combo1.AddItem strCon4
   Combo1.AddItem strCon5
   Combo1.AddItem strCon6    '2006/3/8 ADD BY SONIA
   Combo1 = MsgText(31)
   SelectScope
   OpenTable
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Adodc1.Recordset.RecordCount <> 0 Then
      strItemNo = Adodc1.Recordset.Fields("a0y01").Value
   Else
      strItemNo = MsgText(601)
   End If
   StatusClear
   tool1_enabled
   Frmacc2110.Enabled = True
   Frmacc2110.Show
   Set Frmacc2112 = Nothing
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
         strCondition = "a0y01"
      Case strCon2
         strCondition = "a0y02"
      Case strCon3
         strCondition = "a0y07"
      Case strCon4
         strCondition = "a0y08"
      Case strCon5
         strCondition = "a0y09"
      '2006/3/8 ADD BY SONIA
      Case strCon6
         strCondition = "A1P23"
      '2006/3/8 END
   End Select
   If strCondition = MsgText(601) Then
      Exit Sub
   End If
   Combo2.Clear
   Combo3.Clear
   adoacc0y0.CursorLocation = adUseClient
   '2006/3/8 MODIFY BY SONIA
   'adoacc0y0.Open "select distinct " & strCondition & " from acc0y0 order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
   If Combo1 = strCon6 Then
      'Modified by Lydia 2018/10/08 + a0y01 asc
      adoacc0y0.Open "select distinct " & strCondition & " from acc0y0,ACC1P0 WHERE A0Y01=A1P04 AND A1P23 IS NOT NULL order by " & strCondition & " asc, a0y01 asc", adoTaie, adOpenStatic, adLockReadOnly
   Else
      'Modified by Lydia 2018/10/08 + a0y01 asc
      adoacc0y0.Open "select distinct " & strCondition & " from acc0y0 order by " & strCondition & " asc, a0y01 asc", adoTaie, adOpenStatic, adLockReadOnly
   End If
   '2006/3/8 END
   Do While adoacc0y0.EOF = False
      If IsNull(adoacc0y0.Fields(0).Value) = False Then
         Combo2.AddItem adoacc0y0.Fields(0).Value
         Combo3.AddItem adoacc0y0.Fields(0).Value
      End If
      adoacc0y0.MoveNext
   Loop
   adoacc0y0.Close
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)
   Select Case KeyCode
      Case vbKeyF12
         Acc0y0Query
   End Select
   KeyEnter KeyCode
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

'*************************************************
'  票據資料查詢
'
'*************************************************
Private Sub Acc0y0Query()
On Error GoTo Checking
   adoadodc1.Close
   adoadodc1.CursorLocation = adUseClient
   Select Case Combo1
      Case strCon1
         strCondition = "a0y01"
      Case strCon2
         strCondition = "a0y02"
      Case strCon3
         strCondition = "a0y07"
      Case strCon4
         strCondition = "a0y08"
      Case strCon5
         strCondition = "a0y09"
      '2006/3/8 ADD BY SONIA
      Case strCon6
         strCondition = "A1P23"
      '2006/3/8 END
      Case MsgText(31)
         'Modify by Morgan 2004/11/2 加收款金額
         'adoadodc1.Open "select * from acc0y0 order by a0y01 asc", adoTaie, adOpenStatic, adLockReadOnly
         'Modify by Morgan 2005/5/27 減溢收金額
         'adoadodc1.Open "select a0y01,a0y02,a0y03,a0y04,a0y07,a0y08,a0y09,sum(a0z04) as Amount from acc0y0, acc0z0 where a0z01(+)=a0y01 group by a0y01,a0y02,a0y03,a0y04,a0y07,a0y08,a0y09 order by a0y01 asc", adoTaie, adOpenStatic, adLockReadOnly
         adoadodc1.Open "select a0y01,a0y02,a0y03,a0y04,a0y07,a0y08,a0y09,sum(a0z04)+nvl(max(a0y06),0) as Amount,'' AS A1P23 from acc0y0, acc0z0 where a0z01(+)=a0y01 group by a0y01,a0y02,a0y03,a0y04,a0y07,a0y08,a0y09 order by a0y01 asc", adoTaie, adOpenStatic, adLockReadOnly
         Adodc1.Recordset.Requery
         Exit Sub
      Case Else
         Exit Sub
   End Select
   If Combo3 = MsgText(601) Then
      Select Case Combo1
      Case strCon2
         'Modify by Morgan 2004/11/2 加收款金額
         'adoadodc1.Open "select * from acc0y0 where " & strCondition & " = " & Val(Combo2) & " order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
         'Modify by Morgan 2005/5/27 減溢收金額
         'adoadodc1.Open "select a0y01,a0y02,a0y03,a0y04,a0y07,a0y08,a0y09,sum(a0z04) as Amount from acc0y0, acc0z0 where a0z01(+)=a0y01 and " & strCondition & " = " & Val(Combo2) & " group by a0y01,a0y02,a0y03,a0y04,a0y07,a0y08,a0y09 order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
         'Modified by Lydia 2018/10/08 + a0y01 asc
         adoadodc1.Open "select a0y01,a0y02,a0y03,a0y04,a0y07,a0y08,a0y09,sum(a0z04)+nvl(max(a0y06),0) as Amount,'' AS A1P23 from acc0y0, acc0z0 where a0z01(+)=a0y01 and " & strCondition & " = " & Val(Combo2) & " group by a0y01,a0y02,a0y03,a0y04,a0y07,a0y08,a0y09 order by " & strCondition & " asc, a0y01 asc", adoTaie, adOpenStatic, adLockReadOnly
      '2006/3/8 ADD BY SONIA
      Case strCon6
         'Modified by Lydia 2018/10/08 + a0y01 asc
         adoadodc1.Open "select a0y01,a0y02,a0y03,a0y04,a0y07,a0y08,a0y09,sum(a0z04)+nvl(max(a0y06),0) as Amount,A1P23 from acc0y0, acc0z0, ACC1P0 where a0z01(+)=a0y01 and " & strCondition & " = '" & Combo2 & "' AND A0Z01=A1P04 AND A1P23 IS NOT NULL group by a0y01,a0y02,a0y03,a0y04,a0y07,a0y08,a0y09,A1P23 order by " & strCondition & " asc, a0y01 asc", adoTaie, adOpenStatic, adLockReadOnly
      '2006/3/8 END
      Case Else
         'Modify by Morgan 2004/11/2 加收款金額
         'adoadodc1.Open "select * from acc0y0 where " & strCondition & " = '" & Combo2 & "' order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
         'Modify by Morgan 2005/5/27 減溢收金額
         'adoadodc1.Open "select a0y01,a0y02,a0y03,a0y04,a0y07,a0y08,a0y09,sum(a0z04) as Amount from acc0y0, acc0z0 where a0z01(+)=a0y01 and " & strCondition & " = '" & Combo2 & "' group by a0y01,a0y02,a0y03,a0y04,a0y07,a0y08,a0y09 order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
         'Modified by Lydia 2018/10/08 + a0y01 asc
         adoadodc1.Open "select a0y01,a0y02,a0y03,a0y04,a0y07,a0y08,a0y09,sum(a0z04)+nvl(max(a0y06),0) as Amount,'' AS A1P23 from acc0y0, acc0z0 where a0z01(+)=a0y01 and " & strCondition & " = '" & Combo2 & "' group by a0y01,a0y02,a0y03,a0y04,a0y07,a0y08,a0y09 order by " & strCondition & " asc, a0y01 asc", adoTaie, adOpenStatic, adLockReadOnly
      End Select
   Else
      If Combo2 = MsgText(601) Then
         Select Case Combo1
         Case strCon2
            'Modify by Morgan 2004/11/2 加收款金額
            'adoadodc1.Open "select * from acc0y0 where " & strCondition & " <= " & Val(Combo3) & " order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
            'Modify by Morgan 2005/5/27 減溢收金額
            'adoadodc1.Open "select a0y01,a0y02,a0y03,a0y04,a0y07,a0y08,a0y09,sum(a0z04) as Amount from acc0y0, acc0z0 where a0z01(+)=a0y01 and " & strCondition & " <= " & Val(Combo3) & " group by a0y01,a0y02,a0y03,a0y04,a0y07,a0y08,a0y09 order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
            'Modified by Lydia 2018/10/08 + a0y01 asc
            adoadodc1.Open "select a0y01,a0y02,a0y03,a0y04,a0y07,a0y08,a0y09,sum(a0z04)+nvl(max(a0y06),0) as Amount,'' AS A1P23 from acc0y0, acc0z0 where a0z01(+)=a0y01 and " & strCondition & " <= " & Val(Combo3) & " group by a0y01,a0y02,a0y03,a0y04,a0y07,a0y08,a0y09 order by " & strCondition & " asc, a0y01 asc", adoTaie, adOpenStatic, adLockReadOnly
         '2006/3/8 ADD BY SONIA
         Case strCon6
            'Modified by Lydia 2018/10/08 + a0y01 asc
            adoadodc1.Open "select a0y01,a0y02,a0y03,a0y04,a0y07,a0y08,a0y09,sum(a0z04)+nvl(max(a0y06),0) as Amount,A1P23 from acc0y0, acc0z0, ACC1P0 where a0z01(+)=a0y01 and " & strCondition & " <= '" & Combo3 & "' AND A0Z01=A1P04 AND A1P23 IS NOT NULL group by a0y01,a0y02,a0y03,a0y04,a0y07,a0y08,a0y09,A1P23 order by " & strCondition & " asc, a0y01 asc", adoTaie, adOpenStatic, adLockReadOnly
         '2006/3/8 END
         Case Else
            'Modify by Morgan 2004/11/2 加收款金額
            'adoadodc1.Open "select * from acc0y0 where " & strCondition & " <= '" & Combo3 & "' order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
            'Modify by Morgan 2005/5/27 減溢收金額
            'adoadodc1.Open "select a0y01,a0y02,a0y03,a0y04,a0y07,a0y08,a0y09,sum(a0z04) as Amount from acc0y0, acc0z0 where a0z01(+)=a0y01 and " & strCondition & " <= '" & Combo3 & "' group by a0y01,a0y02,a0y03,a0y04,a0y07,a0y08,a0y09 order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
            'Modified by Lydia 2018/10/08 + a0y01 asc
            adoadodc1.Open "select a0y01,a0y02,a0y03,a0y04,a0y07,a0y08,a0y09,sum(a0z04)+nvl(max(a0y06),0) as Amount,'' AS A1P23 from acc0y0, acc0z0 where a0z01(+)=a0y01 and " & strCondition & " <= '" & Combo3 & "' group by a0y01,a0y02,a0y03,a0y04,a0y07,a0y08,a0y09 order by " & strCondition & " asc, a0y01 asc", adoTaie, adOpenStatic, adLockReadOnly
         End Select
      Else
         Select Case Combo1
         Case strCon2
            'Modify by Morgan 2004/11/2 加收款金額
            'adoadodc1.Open "select * from acc0y0 where " & strCondition & " >= " & Val(Combo2) & " and " & strCondition & " <= " & Val(Combo3) & " order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
            'Modify by Morgan 2005/5/27 減溢收金額
            'adoadodc1.Open "select a0y01,a0y02,a0y03,a0y04,a0y07,a0y08,a0y09,sum(a0z04) as Amount from acc0y0, acc0z0 where a0z01(+)=a0y01 and " & strCondition & " >= " & Val(Combo2) & " and " & strCondition & " <= " & Val(Combo3) & " group by a0y01,a0y02,a0y03,a0y04,a0y07,a0y08,a0y09 order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
            'Modified by Lydia 2018/10/08 + a0y01 asc
            adoadodc1.Open "select a0y01,a0y02,a0y03,a0y04,a0y07,a0y08,a0y09,sum(a0z04)+nvl(max(a0y06),0) as Amount,'' AS A1P23 from acc0y0, acc0z0 where a0z01(+)=a0y01 and " & strCondition & " >= " & Val(Combo2) & " and " & strCondition & " <= " & Val(Combo3) & " group by a0y01,a0y02,a0y03,a0y04,a0y07,a0y08,a0y09 order by " & strCondition & " asc, a0y01 asc", adoTaie, adOpenStatic, adLockReadOnly
         '2006/3/8 ADD BY SONIA
         Case strCon6
            'Modified by Lydia 2018/10/08 + a0y01 asc
            adoadodc1.Open "select a0y01,a0y02,a0y03,a0y04,a0y07,a0y08,a0y09,sum(a0z04)+nvl(max(a0y06),0) as Amount,A1P23 from acc0y0, acc0z0, ACC1P0 where a0z01(+)=a0y01 and " & strCondition & " >= '" & Combo2 & "' and " & strCondition & " <= '" & Combo3 & "' AND A0Z01=A1P04 AND A1P23 IS NOT NULL group by a0y01,a0y02,a0y03,a0y04,a0y07,a0y08,a0y09,A1P23 order by " & strCondition & " asc, a0y01 asc", adoTaie, adOpenStatic, adLockReadOnly
         '2006/3/8 END
         Case Else
            'Modify by Morgan 2004/11/2 加收款金額
            'adoadodc1.Open "select * from acc0y0 where " & strCondition & " >= '" & Combo2 & "' and " & strCondition & " <= '" & Combo3 & "' order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
            'Modify by Morgan 2005/5/27 減溢收金額
            'adoadodc1.Open "select a0y01,a0y02,a0y03,a0y04,a0y07,a0y08,a0y09,sum(a0z04) as Amount from acc0y0, acc0z0 where a0z01(+)=a0y01 and " & strCondition & " >= '" & Combo2 & "' and " & strCondition & " <= '" & Combo3 & "' group by a0y01,a0y02,a0y03,a0y04,a0y07,a0y08,a0y09 order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
            'Modified by Lydia 2018/10/08 + a0y01 asc
            adoadodc1.Open "select a0y01,a0y02,a0y03,a0y04,a0y07,a0y08,a0y09,sum(a0z04)+nvl(max(a0y06),0) as Amount,'' AS A1P23 from acc0y0, acc0z0 where a0z01(+)=a0y01 and " & strCondition & " >= '" & Combo2 & "' and " & strCondition & " <= '" & Combo3 & "' group by a0y01,a0y02,a0y03,a0y04,a0y07,a0y08,a0y09 order by " & strCondition & " asc, a0y01 asc", adoTaie, adOpenStatic, adLockReadOnly
         End Select
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
   adoadodc1.Open "select * from acc0y0 where a0y01 = '" & Combo2 & "' order by a0y01 asc", adoTaie, adOpenStatic, adLockReadOnly
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

