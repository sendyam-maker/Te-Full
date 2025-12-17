VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Frmacc4280 
   AutoRedraw      =   -1  'True
   Caption         =   "日記帳查詢"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9405
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5115
   ScaleWidth      =   9405
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1110
      TabIndex        =   0
      Top             =   240
      Width           =   3500
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc4280.frx":0000
      Height          =   4200
      Left            =   240
      TabIndex        =   6
      Top             =   810
      Width           =   9000
      _ExtentX        =   15875
      _ExtentY        =   7408
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
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
      Caption         =   "日記帳資料"
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "公司別"
         Caption         =   "公司"
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
         DataField       =   "傳票日期"
         Caption         =   "傳票日期"
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
         DataField       =   "科目代號"
         Caption         =   "科目代號"
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
         DataField       =   "科目名稱"
         Caption         =   "科目名稱"
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
      BeginProperty Column04 
         DataField       =   "借方金額"
         Caption         =   "借方金額"
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
      BeginProperty Column05 
         DataField       =   "貸方金額"
         Caption         =   "貸方金額"
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
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   569.764
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1035.213
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2700.284
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1484.787
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   1470.047
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
      Left            =   240
      Top             =   690
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
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   5850
      TabIndex        =   1
      Top             =   240
      Width           =   1335
      _ExtentX        =   2355
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
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   300
      Left            =   7530
      TabIndex        =   2
      Top             =   240
      Width           =   1335
      _ExtentX        =   2355
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
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   585
      Left            =   240
      Top             =   120
      Width           =   9000
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7290
      TabIndex        =   5
      Top             =   240
      Width           =   135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "傳票日期"
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
      Left            =   4890
      TabIndex        =   4
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "公司別"
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
      TabIndex        =   3
      Top             =   240
      Width           =   732
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   4632
      Visible         =   0   'False
      Width           =   132
   End
End
Attribute VB_Name = "Frmacc4280"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/13 Form2.0已修改 (無需修改)
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/3 日期欄已修改
Option Explicit
Public adoadodc1 As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Dim strCond1, strCond2, strCond3, strCond4 As String
Dim strField1, strField2, strField3 As String
Dim strSql As String
Dim strSum As String
'Add by Amy 2020/04/14
Private Sub Combo1_GotFocus()
    TextInverse Combo1
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Combo1_Validate(Cancel As Boolean)
    Dim strCmp As String
    
    If Trim(Combo1) = MsgText(601) Then Exit Sub
    
    strCmp = Combo1
    If InStr(strCmp, "　") > 0 Then
        strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
    End If
    If InStr(GetBookKeepCmp, strCmp) = 0 Then
        MsgBox Label3 & MsgText(63), , MsgText(5)
        Cancel = True
        Combo1.SetFocus
        Exit Sub
    ElseIf Len(Trim(Combo1)) = 1 Then
        Combo1 = Trim(strCmp) & "　" & A0802Query(strCmp)
    End If
End Sub
'end 2020/04/14

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
   Me.Width = 9500
   Me.Height = 5500
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath2)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   'Add by Amy 2020/04/14
   Combo1.AddItem "", 0
   Call Pub_SetCboCmp(Combo1, False, False, False, , 1)
   'end 2020/04/14
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat
   OpenTable
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
   '20140124START Modify By eric
   'Me.Text4.Text = "1"
   'Text5 = A0802Query(Text4)
   '20140124END
End Sub

Private Sub Form_Unload(Cancel As Integer)
   StatusClear
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc4280 = Nothing
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
Dim StrSQLa As String
Dim strCmp As String 'Add by Amy 2020/04/14

On Error GoTo Checking
   'Modify by Amy 2020/03/31 改下拉 原:Text4
   If Trim(Combo1) <> MsgText(601) Then
      strCmp = Combo1
      If InStr(strCmp, "　") > 0 Then
            strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
      End If
    End If
    adoadodc1.CursorLocation = adUseClient
    '20140224START Modify By eric
    StrSQLa = "select ax201 as 公司別,a0205 as 傳票日期 , ax205 as 科目代號, a0102 as 科目名稱, sum(ax206) as 借方金額, sum(ax207) as 貸方金額 from acc021, acc020, acc010 where acc021.ax201 = acc020.a0201 and acc021.ax202 = acc020.a0202 and acc021.ax205 = acc010.a0101 " & _
                        " and ax201 = '" & strCmp & "' and a0205 >= " & Val(FCDate(MaskEdBox1.Text)) & " and a0205 <= " & Val(FCDate(MaskEdBox2.Text)) & _
                        " group by a0205, ax205, a0102, ax201 "
    'end 2020/04/14
    'StrSQLa = "select a0205 as 傳票日期 , ax205 as 科目代號, a0102 as 科目名稱, sum(ax206) as 借方金額, sum(ax207) as 貸方金額 from acc021, acc020, acc010 where acc021.ax201 = acc020.a0201 and acc021.ax202 = acc020.a0202 and acc021.ax205 = acc010.a0101 " & _
    '                    " and ax201 = '" & Text4.Text & "' and a0205 >= " & Val(FCDate(MaskEdBox1.Text)) & " and a0205 <= " & Val(FCDate(MaskEdBox2.Text)) & _
    '                    " group by a0205, ax205, a0102 "
    '20140124END
    adoadodc1.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
    Set Adodc1.Recordset = adoadodc1
Checking:
    If Err.Number = 0 Then
        Exit Sub
    End If
    MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  查詢資料表(傳票資料)
'
'*************************************************
Public Sub QueryTable()
Dim rsA As New ADODB.Recordset
Dim ii As Double
Dim strNoteDate As String '傳票日期
Dim dblDayDebTotal As Double  '日期借方小計
Dim dblDayCrtTotal As Double  '日期借方小計
Dim dblDebTotal As Double '借方合計
Dim dblCrtTotal As Double '貸方合計
Dim strCmp As String 'Add by Amy 2020/04/14

'add by nickc 2007/02/08
Dim StrSQLa As String

On Error GoTo Checking
    adoTaie.Execute "Delete From ACCRPT426 Where ID='" & strUserNum & "' "
    strSql = ""
    strSum = ""
    If adoadodc1.State = adStateOpen Then
        adoadodc1.Close
    End If
    adoadodc1.CursorLocation = adUseClient
    '公司別
   '20140124START Modify By eric
   'Modify by Amy 2020/04/14 公司別改下拉 原:Text4
   If Trim(Combo1) <> MsgText(601) Then
      strCmp = Combo1
      If InStr(strCmp, "　") > 0 Then
            strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
      End If
      strSql = " and ax201 = '" & strCmp & "'"
   End If
   'end 2020/04/14
   'If Text4 <> MsgText(601) Then
   '   strSql = " and ax201 = '" & Text4 & "'"
   'End If
   '20140124END
   '傳票日期
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      strSql = strSql & " and a0205 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
   End If
   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      strSql = strSql & " and a0205 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
   End If
    '20140124START Modify By eric
    StrSQLa = "select ax201 as 公司別 , a0205 as 傳票日期 , ax205 as 科目代號, a0102 as 科目名稱, sum(ax206) as 借方金額, sum(ax207) as 貸方金額 from acc021, acc020, acc010 where acc021.ax201 = acc020.a0201 and acc021.ax202 = acc020.a0202 and acc021.ax205 = acc010.a0101 " & _
                     strSql & " group by a0205, ax205, ax201, a0102 Order By 2, 1, 3 "
    'StrSQLa = "select a0205 as 傳票日期 , ax205 as 科目代號, a0102 as 科目名稱, sum(ax206) as 借方金額, sum(ax207) as 貸方金額 from acc021, acc020, acc010 where acc021.ax201 = acc020.a0201 and acc021.ax202 = acc020.a0202 and acc021.ax205 = acc010.a0101 " & _
    '                 strSql & " group by a0205, ax205, a0102 Order By 1, 2 , 3 "
    '20140124END
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
    If rsA.RecordCount <= 0 Then
       Set Adodc1.Recordset = rsA
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
        MsgBox MsgText(28), , MsgText(5)
        Exit Sub
    End If
    ii = 0
    '20140124START Modify By eric
    strNoteDate = "" & rsA.Fields(1).Value
    'strNoteDate = "" & rsA.Fields(0).Value
    '20140124END
    
    dblDayDebTotal = 0: dblDayCrtTotal = 0: dblDebTotal = 0: dblCrtTotal = 0
    While Not rsA.EOF
        '20140124START Modify By eric
        '若日期不同
        'If strNoteDate <> "" & rsA.Fields(0).Value Then
        '    ii = ii + 1
        '    adoTaie.Execute " Insert Into ACCRPT426 Values(" & ii & ",Null, Null,'日期小計:'," & dblDayDebTotal & "," & dblDayCrtTotal & ",'" & strUserNum & "' ) "
        '    dblDayDebTotal = 0: dblDayCrtTotal = 0
        '    strNoteDate = "" & rsA.Fields(0).Value
        'End If
        'ii = ii + 1
        'adoTaie.Execute " Insert Into ACCRPT426 Values(" & ii & ",'" & rsA.Fields(1).Value & "','" & rsA.Fields(2).Value & "'," & rsA.Fields(3).Value & "," & rsA.Fields(4).Value & ",'" & strUserNum & "' ) "
        'dblDayDebTotal = dblDayDebTotal + CDbl(rsA.Fields(3).Value)
        'dblDayCrtTotal = dblDayCrtTotal + CDbl(rsA.Fields(4).Value)
        'dblDebTotal = dblDebTotal + CDbl(rsA.Fields(3).Value)
        'dblCrtTotal = dblCrtTotal + CDbl(rsA.Fields(4).Value)
        
        '若日期不同
        If strNoteDate <> "" & rsA.Fields(1).Value Then
            ii = ii + 1
            adoTaie.Execute " Insert Into ACCRPT426 Values(" & ii & ",Null, Null,'日期小計:'," & dblDayDebTotal & "," & dblDayCrtTotal & ",'" & strUserNum & "' , Null) "
            dblDayDebTotal = 0: dblDayCrtTotal = 0
            strNoteDate = "" & rsA.Fields(1).Value
        End If
        ii = ii + 1
        adoTaie.Execute " Insert Into ACCRPT426 Values(" & ii & "," & rsA.Fields(1).Value & ",'" & rsA.Fields(2).Value & "','" & rsA.Fields(3).Value & "'," & rsA.Fields(4).Value & "," & rsA.Fields(5).Value & ",'" & strUserNum & "','" & rsA.Fields(0).Value & "' ) "
        dblDayDebTotal = dblDayDebTotal + CDbl(rsA.Fields(4).Value)
        dblDayCrtTotal = dblDayCrtTotal + CDbl(rsA.Fields(5).Value)
        dblDebTotal = dblDebTotal + CDbl(rsA.Fields(4).Value)
        dblCrtTotal = dblCrtTotal + CDbl(rsA.Fields(5).Value)
        '20140124END
        
        rsA.MoveNext
        
    Wend
    
    '20140124START Modisy By eric
    'ii = ii + 1
    'adoTaie.Execute " Insert Into ACCRPT426 Values(" & ii & ",Null, Null,'日期小計:'," & dblDayDebTotal & "," & dblDayCrtTotal & ",'" & strUserNum & "' ) "
    'ii = ii + 1
    'adoTaie.Execute " Insert Into ACCRPT426 Values(" & ii & ",Null, Null,'合　　計:'," & dblDebTotal & "," & dblCrtTotal & ",'" & strUserNum & "' ) "
    'StrSQLa = "select R42602 as 傳票日期 ,R42603 as 科目代號, R42604 as 科目名稱, R42605 as 借方金額, R42606 as 貸方金額 from ACCRPT426 where id='" & strUserNum & "' Order By R42601 "
    ii = ii + 1
    adoTaie.Execute " Insert Into ACCRPT426 Values(" & ii & ",Null, Null,'日期小計:'," & dblDayDebTotal & "," & dblDayCrtTotal & ",'" & strUserNum & "', Null ) "
    ii = ii + 1
    adoTaie.Execute " Insert Into ACCRPT426 Values(" & ii & ",Null, Null,'合　　計:'," & dblDebTotal & "," & dblCrtTotal & ",'" & strUserNum & "',Null ) "
    StrSQLa = "select R42602 as 傳票日期 ,R42603 as 科目代號, R42604 as 科目名稱, R42605 as 借方金額, R42606 as 貸方金額, R42607 as 公司別 from ACCRPT426 where id='" & strUserNum & "' Order By R42601 "
    '20140124END
    
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
   Set Adodc1.Recordset = rsA
   If Adodc1.Recordset.RecordCount = 0 Then
      Adodc1.Recordset.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   End If
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

''Mark by Amy 2020/04/14 公司別改下拉
'Private Sub Text4_Change()
'   '20140124START Modify By eric
'   'Text5 = A0802Query(Text4)
'   Select Case Text4
'      Case "1"
'         Text5 = A0802Query(Text4)
'      Case "2"
'         Text5 = A0802Query("J")
'   End Select
'   '20140124END
'End Sub
'
'Private Sub Text4_GotFocus()
'   TextInverse Text4
'   '20140124START Add By eric
'   CloseIme
'   '20140124END
'End Sub
'20140124START By eric
'Private Sub Text4_LostFocus()
'   If Text4.Text <> "1" And Text4.Text <> "2" And Text4.Text <> "" Then
'      MsgBox "公司別僅可為 1 或 2 或空白  ! (1.台一 2.智權 空白.全部)"
'      Text4.Text = ""
'      Text4.SetFocus
'      Exit Sub
'   End If
'End Sub
'end 2020/04/14

'*************************************************
'  功能鍵定義
'
'*************************************************
Public Sub KeyDefine(KeyCode As Integer)
   Dim HasShowMsg As Boolean 'Add by Amy 2020/04/14
   Select Case KeyCode
      Case vbKeyF12
         'Modify by Amy 2020/04/14
         If FormCheck(HasShowMsg) Then
            Screen.MousePointer = vbHourglass
            QueryTable
            Screen.MousePointer = vbDefault
            Exit Sub
         ElseIf HasShowMsg = False Then
            MsgBox MsgText(181), , MsgText(5)
         End If
         'end 2020/04/14
   End Select
   KeyEnter KeyCode
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck(HasShowMsg As Boolean) As Boolean
   Dim bCancel As Boolean 'Add by Amy 2020/04/14
   
   'Modif by Amy 2020/04/14 公司別改下拉 原:Text4
   If Trim(Combo1) <> MsgText(601) Then
      Call Combo1_Validate(bCancel)
      If bCancel = True Then
        HasShowMsg = True
      Else
        FormCheck = True
      End If
      Exit Function
   End If
   'end 2020/04/14
   If MaskEdBox1.Text <> MsgText(29) Then
      FormCheck = True
      Exit Function
   End If
   If MaskEdBox2.Text <> MsgText(29) Then
      FormCheck = True
      Exit Function
   End If
   FormCheck = False
End Function




