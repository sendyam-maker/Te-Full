VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Frmacc2121 
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
      Height          =   300
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
      Height          =   330
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
      Height          =   330
      Left            =   5760
      TabIndex        =   2
      Top             =   240
      Width           =   2772
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc2121.frx":0000
      Height          =   4092
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   8292
      _ExtentX        =   14623
      _ExtentY        =   7214
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
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "a1201"
         Caption         =   "暫收款單號"
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
         DataField       =   "a1202"
         Caption         =   "暫收款日期"
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
         DataField       =   "a1203"
         Caption         =   "代理人"
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
         DataField       =   "a1204"
         Caption         =   "幣別"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "#,##0.0000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "a1205"
         Caption         =   "匯率"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.0000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "a1207"
         Caption         =   "外幣金額"
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
         DataField       =   "a1211"
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
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   1340.221
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   1230.236
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   569.764
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   849.827
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   4160.126
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
Attribute VB_Name = "Frmacc2121"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/03 改成Form2.0 ; DataGrid1改字型=新細明體-ExtB
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit
Public adoacc120 As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset

Private Sub Combo1_Change()
'   SelectScope
End Sub

Private Sub Combo1_Click()
'   SelectScope
End Sub

Private Sub Combo2_GotFocus()
   CloseIme
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Combo2_Validate(Cancel As Boolean)
   '2009/6/26 ADD BY SONIA
   If Combo1 = "代理人" And Len(Combo2) = 6 Then
      Combo2 = AfterZero(Combo2)
   End If
   '2009/6/26 END
   Combo3 = Combo2
   '2009/6/2 MODIFY BY SONIA 預設尾碼999
   'Modify By Sindy 2014/8/11 999=>ZZZ
   'If Combo2 <> "" And Combo1 = strCon3 Then Combo3 = Left(Combo2, 6) & "999"
   If Combo2 <> "" And Combo1 = strCon3 Then Combo3 = Left(Combo2, 6) & "ZZZ"
End Sub

Private Sub Combo3_GotFocus()
   CloseIme
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
   strCon2 = "暫收款日期"
   strCon3 = "代理人"
   strCon4 = "幣別"
   strCon5 = "外幣金額"
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
      strItemNo = Adodc1.Recordset.Fields("a1201").Value
   Else
      strItemNo = MsgText(601)
   End If
   StatusClear
   tool1_enabled
   Frmacc2120.Enabled = True
   Frmacc2120.Show
   Set Frmacc2121 = Nothing
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
         strCondition = "a1201"
      Case strCon2
         strCondition = "a1202"
      Case strCon3
         strCondition = "a1203"
      Case strCon4
         strCondition = "a1204"
      Case strCon5
         strCondition = "a1207"
   End Select
   If strCondition = MsgText(601) Then
      Exit Sub
   End If
   Combo2.Clear
   Combo3.Clear
   adoacc120.CursorLocation = adUseClient
   'Modified by Lydia 2018/10/08 + a1201 asc
   adoacc120.Open "select distinct " & strCondition & " from acc120 order by " & strCondition & " asc, a1201 asc", adoTaie, adOpenStatic, adLockReadOnly
   Do While adoacc120.EOF = False
      If IsNull(adoacc120.Fields(0).Value) = False Then
         Combo2.AddItem adoacc120.Fields(0).Value
         Combo3.AddItem adoacc120.Fields(0).Value
      End If
      adoacc120.MoveNext
   Loop
   adoacc120.Close
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)
   Select Case KeyCode
      Case vbKeyF12
         Acc120Query
   End Select
   KeyEnter KeyCode
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

'*************************************************
'  票據資料查詢
'
'*************************************************
Private Sub Acc120Query()
On Error GoTo Checking
   adoadodc1.Close
   adoadodc1.CursorLocation = adUseClient
   Select Case Combo1
      Case strCon1
         strCondition = "a1201"
      Case strCon2
         strCondition = "a1202"
      Case strCon3
         strCondition = "a1203"
      Case strCon4
         strCondition = "a1204"
      Case strCon5
         strCondition = "a1207"
      Case MsgText(31)
         If Frmacc2120.Option1 Then
            '92.6.16 MODIFY BY SONIA
            'adoadodc1.Open "select * from acc120 where a1201 not in (select a1303 from acc130) and a1201 not in (select a1p30 from acc1p0 where a1p02 = 'F' and a1p05 = '2401' and a1p07 <> 0 and a1p30 is not null) order by a1201 asc", adoTaie, adOpenStatic, adLockReadOnly
            'modify by sonia 2024/7/5 a1p02加入'I',Frmacc2120早已加入
            adoadodc1.Open "select * from acc120 where a1201 not in (select a1303 from acc130) and a1201 not in (select a1p30 from acc1p0 where a1p02 IN ('I','F','K') and a1p05 = '2401' and a1p07 <> 0 and a1p30 is not null) order by a1201 asc", adoTaie, adOpenStatic, adLockReadOnly
         Else
            adoadodc1.Open "select * from acc120 order by a1201 asc", adoTaie, adOpenStatic, adLockReadOnly
         End If
         Adodc1.Recordset.Requery
         Exit Sub
      Case Else
         Exit Sub
   End Select
   If Frmacc2120.Option1.Value Then
      If Combo3 = MsgText(601) Then
         If Combo1 = strCon2 Or Combo1 = strCon5 Then
            '92.6.16 MODIFY BY SONIA
            'adoadodc1.Open "select * from acc120 where " & strCondition & " = " & Val(Combo2) & " and a1201 not in (select a1303 from acc130) and a1201 not in (select a1p30 from acc1p0 where a1p02 = 'F' and a1p05 = '2401' and a1p07 <> 0 and a1p30 is not null) order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
            'Modified by Lydia 2018/12/08 + a1201 asc
            'modify by sonia 2024/7/5 a1p02加入'I',Frmacc2120早已加入
            adoadodc1.Open "select * from acc120 where " & strCondition & " = " & Val(Combo2) & " and a1201 not in (select a1303 from acc130) and a1201 not in (select a1p30 from acc1p0 where a1p02 IN ('I','F','K') and a1p05 = '2401' and a1p07 <> 0 and a1p30 is not null) order by " & strCondition & " asc, a1201 asc ", adoTaie, adOpenStatic, adLockReadOnly
         Else
            '92.6.16 MODIFY BY SONIA
            'adoadodc1.Open "select * from acc120 where " & strCondition & " = '" & Combo2 & "' and a1201 not in (select a1303 from acc130) and a1201 not in (select a1p30 from acc1p0 where a1p02 = 'F' and a1p05 = '2401' and a1p07 <> 0 and a1p30 is not null) order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
            'Modified by Lydia 2018/12/08 + a1201 asc
            'modify by sonia 2024/7/5 a1p02加入'I',Frmacc2120早已加入
            adoadodc1.Open "select * from acc120 where " & strCondition & " = '" & Combo2 & "' and a1201 not in (select a1303 from acc130) and a1201 not in (select a1p30 from acc1p0 where a1p02 IN ('I','F','K') and a1p05 = '2401' and a1p07 <> 0 and a1p30 is not null) order by " & strCondition & " asc, a1201 asc", adoTaie, adOpenStatic, adLockReadOnly
         End If
      Else
         If Combo2 = MsgText(601) Then
            If Combo1 = strCon2 Or Combo1 = strCon5 Then
               '92.6.16 MODIFY BY SONIA
               'adoadodc1.Open "select * from acc120 where " & strCondition & " <= " & Val(Combo3) & " and a1201 not in (select a1303 from acc130) and a1201 not in (select a1p30 from acc1p0 where a1p02 = 'F' and a1p05 = '2401' and a1p07 <> 0 and a1p30 is not null) order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
               'Modified by Lydia 2018/12/08 + a1201 asc
               'modify by sonia 2024/7/5 a1p02加入'I',Frmacc2120早已加入
               adoadodc1.Open "select * from acc120 where " & strCondition & " <= " & Val(Combo3) & " and a1201 not in (select a1303 from acc130) and a1201 not in (select a1p30 from acc1p0 where a1p02 IN ('I','F','K') and a1p05 = '2401' and a1p07 <> 0 and a1p30 is not null) order by " & strCondition & " asc, a1201 asc", adoTaie, adOpenStatic, adLockReadOnly
            Else
               '92.6.16 MODIFY BY SONIA
               'adoadodc1.Open "select * from acc120 where " & strCondition & " <= '" & Combo3 & "' and a1201 not in (select a1303 from acc130) and a1201 not in (select a1p30 from acc1p0 where a1p02 = 'F' and a1p05 = '2401' and a1p07 <> 0 and a1p30 is not null) order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
               'Modified by Lydia 2018/12/08 + a1201 asc
               'modify by sonia 2024/7/5 a1p02加入'I',Frmacc2120早已加入
               adoadodc1.Open "select * from acc120 where " & strCondition & " <= '" & Combo3 & "' and a1201 not in (select a1303 from acc130) and a1201 not in (select a1p30 from acc1p0 where a1p02 IN ('I','F','K') and a1p05 = '2401' and a1p07 <> 0 and a1p30 is not null) order by " & strCondition & " asc, a1201 asc", adoTaie, adOpenStatic, adLockReadOnly
            End If
         Else
            If Combo1 = strCon2 Or Combo1 = strCon5 Then
               '92.6.16 MODIFY BY SONIA
               'adoadodc1.Open "select * from acc120 where " & strCondition & " >= " & Val(Combo2) & " and " & strCondition & " <= " & Val(Combo3) & " and a1201 not in (select a1303 from acc130) and a1201 not in (select a1p30 from acc1p0 where a1p02 = 'F' and a1p05 = '2401' and a1p07 <> 0 and a1p30 is not null) order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
               'Modified by Lydia 2018/12/08 + a1201 asc
               'modify by sonia 2024/7/5 a1p02加入'I',Frmacc2120早已加入
               adoadodc1.Open "select * from acc120 where " & strCondition & " >= " & Val(Combo2) & " and " & strCondition & " <= " & Val(Combo3) & " and a1201 not in (select a1303 from acc130) and a1201 not in (select a1p30 from acc1p0 where a1p02 IN ('I','F','K') and a1p05 = '2401' and a1p07 <> 0 and a1p30 is not null) order by " & strCondition & " asc, a1201 asc", adoTaie, adOpenStatic, adLockReadOnly
            Else
               '92.6.16 MODIFY BY SONIA
               'adoadodc1.Open "select * from acc120 where " & strCondition & " >= '" & Combo2 & "' and " & strCondition & " <= '" & Combo3 & "' and a1201 not in (select a1303 from acc130) and a1201 not in (select a1p30 from acc1p0 where a1p02 = 'F' and a1p05 = '2401' and a1p07 <> 0 and a1p30 is not null) order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
               'Modified by Lydia 2018/12/08 + a1201 asc
               'modify by sonia 2024/7/5 a1p02加入'I',Frmacc2120早已加入
               adoadodc1.Open "select * from acc120 where " & strCondition & " >= '" & Combo2 & "' and " & strCondition & " <= '" & Combo3 & "' and a1201 not in (select a1303 from acc130) and a1201 not in (select a1p30 from acc1p0 where a1p02 IN ('I','F','K') and a1p05 = '2401' and a1p07 <> 0 and a1p30 is not null) order by " & strCondition & " asc, a1201 asc", adoTaie, adOpenStatic, adLockReadOnly
            End If
         End If
      End If
   Else
      If Combo3 = MsgText(601) Then
         If Combo1 = strCon2 Or Combo1 = strCon5 Then
            'Modified by Lydia 2018/12/08 + a1201 asc
            adoadodc1.Open "select * from acc120 where " & strCondition & " = " & Val(Combo2) & " order by " & strCondition & " asc,  a1201 asc", adoTaie, adOpenStatic, adLockReadOnly
         Else
            'Modified by Lydia 2018/12/08 + a1201 asc
            adoadodc1.Open "select * from acc120 where " & strCondition & " = '" & Combo2 & "' order by " & strCondition & " asc,  a1201 asc", adoTaie, adOpenStatic, adLockReadOnly
         End If
      Else
         If Combo2 = MsgText(601) Then
            If Combo1 = strCon2 Or Combo1 = strCon5 Then
               'Modified by Lydia 2018/12/08 + a1201 asc
               adoadodc1.Open "select * from acc120 where " & strCondition & " <= " & Val(Combo3) & " order by " & strCondition & " asc, a1201 asc", adoTaie, adOpenStatic, adLockReadOnly
            Else
               'Modified by Lydia 2018/12/08 + a1201 asc
               adoadodc1.Open "select * from acc120 where " & strCondition & " <= '" & Combo3 & "' order by " & strCondition & " asc, a1201 asc", adoTaie, adOpenStatic, adLockReadOnly
            End If
         Else
            If Combo1 = strCon2 Or Combo1 = strCon5 Then
               'Modified by Lydia 2018/12/08 + a1201 asc
               adoadodc1.Open "select * from acc120 where " & strCondition & " >= " & Val(Combo2) & " and " & strCondition & " <= " & Val(Combo3) & " order by " & strCondition & " asc, a1201 asc", adoTaie, adOpenStatic, adLockReadOnly
            Else
               'Modified by Lydia 2018/12/08 + a1201 asc
               adoadodc1.Open "select * from acc120 where " & strCondition & " >= '" & Combo2 & "' and " & strCondition & " <= '" & Combo3 & "' order by " & strCondition & " asc, a1201 asc", adoTaie, adOpenStatic, adLockReadOnly
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
   adoadodc1.Open "select * from acc120 where a1201 = '" & Combo2 & "' order by a1201 asc", adoTaie, adOpenStatic, adLockReadOnly
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub
