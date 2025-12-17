VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Frmacc4210 
   AutoRedraw      =   -1  'True
   Caption         =   "傳票資料查詢"
   ClientHeight    =   5112
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9408
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5112
   ScaleWidth      =   9408
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1320
      TabIndex        =   0
      Top             =   210
      Width           =   3520
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5970
      TabIndex        =   12
      Top             =   4560
      Width           =   1500
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7488
      TabIndex        =   11
      Top             =   4560
      Width           =   1500
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc4210.frx":0000
      Height          =   3300
      Left            =   240
      TabIndex        =   10
      Top             =   1200
      Width           =   9000
      _ExtentX        =   15875
      _ExtentY        =   5821
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   18
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "傳票資料"
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "ax201"
         Caption         =   "公司別"
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
         DataField       =   "ax202"
         Caption         =   "傳票編號"
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
         DataField       =   "a0205"
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
      BeginProperty Column03 
         DataField       =   "ax203"
         Caption         =   "項次"
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
         DataField       =   "a0102"
         Caption         =   "科目名稱"
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
         DataField       =   "ax206"
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
      BeginProperty Column06 
         DataField       =   "ax207"
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
      BeginProperty Column07 
         DataField       =   "ax204"
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
      BeginProperty Column08 
         DataField       =   "ax212"
         Caption         =   "摘要"
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
            Alignment       =   2
            ColumnWidth     =   720
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1272.189
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   1128.189
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   492.095
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1800
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column07 
            Alignment       =   2
            ColumnWidth     =   684.284
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   4680
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   312
      Left            =   240
      Top             =   1080
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   550
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
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7590
      TabIndex        =   4
      Top             =   600
      Width           =   1332
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5910
      TabIndex        =   3
      Top             =   600
      Width           =   1332
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1320
      TabIndex        =   1
      Top             =   570
      Width           =   1532
      _ExtentX        =   2709
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.8
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
      Left            =   3330
      TabIndex        =   2
      Top             =   600
      Width           =   1530
      _ExtentX        =   2688
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "合計"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   5400
      TabIndex        =   13
      Top             =   4560
      Width           =   612
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   975
      Left            =   240
      Top             =   120
      Width           =   9000
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7350
      TabIndex        =   9
      Top             =   600
      Width           =   135
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3030
      TabIndex        =   8
      Top             =   600
      Width           =   135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "傳票號碼"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4950
      TabIndex        =   7
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "傳票日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   360
      TabIndex        =   6
      Top             =   600
      Width           =   972
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "公司別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   210
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   4632
      Visible         =   0   'False
      Width           =   132
   End
End
Attribute VB_Name = "Frmacc4210"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/13 Form2.0已修改 DataGrid1
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
   Me.Width = 9700 'Modify by Amy 2023/07/19 原:9500
   Me.Height = 5600 'Modify by Amy 2023/07/19 原:5500
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
   '20140122REMARK By eric (公司別欄位可選擇 1本所 或 2智權)
   'Text4 = "1"
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   StatusClear
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc4210 = Nothing
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoadodc1.CursorLocation = adUseClient
   'Modify by Amy 2020/04/14 公司別改下拉 原:Text4
   Select Case strAccount
      Case "2"
         adoadodc1.Open "select ax301 as ax201, ax302 as ax202, a0305 as a0205, ax303 as ax203, ax306 as ax206, ax307 as ax207, ax304 as ax204, ax312 as ax212, a0102 from acc031, acc010, acc030 where ax305 = a0101 and ax301 = a0301 and ax302 = a0302 and ax301 = '" & Combo1 & "' and a0305 >= " & Val(FCDate(MaskEdBox1.Text)) & " and a0305 <= " & Val(FCDate(MaskEdBox2.Text)) & " and ax302 >= '" & Text1 & "' and ax302 <= '" & Text3 & "' order by a0305 desc, ax302 asc, ax303 asc", adoTaie, adOpenStatic, adLockReadOnly
      Case Else
         adoadodc1.Open "select * from acc021, acc010, acc020 where acc021.ax205 = acc010.a0101 and acc021.ax201 = acc020.a0201 and acc021.ax202 = acc020.a0202 and ax201 = '" & Combo1 & "' and a0205 >= " & Val(FCDate(MaskEdBox1.Text)) & " and a0205 <= " & Val(FCDate(MaskEdBox2.Text)) & " and ax202 >= '" & Text1 & "' and ax202 <= '" & Text3 & "' order by a0205 desc, ax202 asc, ax203 asc", adoTaie, adOpenStatic, adLockReadOnly
   End Select
   'end 2020/04/14
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
Dim strCmp As String 'Add by Amy 2020/04/14

On Error GoTo Checking
   strSql = ""
   strSum = ""
   'Modify by Amy 2020/04/14 改下拉 原:Text4
   If Trim(Combo1) <> MsgText(601) Then
      strCmp = Combo1
      If InStr(strCmp, "　") > 0 Then
            strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
      End If
   End If
   'end 2020/04/14
   If adoadodc1.State = adStateOpen Then
      adoadodc1.Close
   End If
   adoadodc1.CursorLocation = adUseClient
   Select Case strAccount
      Case "2"
         If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
            strSql = strSql & " and a0305 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
         End If
         If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
            strSql = strSql & " and a0305 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
         End If
         If Text1 <> MsgText(601) Then
            strSql = strSql & " and ax302 >= '" & Text1 & "'"
         End If
         If Text3 <> MsgText(601) Then
            strSql = strSql & " and ax302 <= '" & Text3 & "'"
         End If
         'Modify by Amy 2020/04/14 公司別改下拉 原:Text4
         If strCmp <> MsgText(601) Then
            strSql = strSql & " and ax301 = '" & strCmp & "'"
         End If
         strSum = strSql
         strSql = "select ax301 as ax201, ax302 as ax202, a0305 as a0205, ax303 as ax203, ax306 as ax206, ax307 as ax207, ax304 as ax204, ax312 as ax212, a0102 from acc031, acc010, acc030 where ax305 = a0101 (+) and ax301 = a0301 and ax302 = a0302" & strSql & " order by ax301 asc, ax302 asc, ax303 asc"
         adoadodc1.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
      Case Else
         If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
            strSql = strSql & " and a0205 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
         End If
         If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
            strSql = strSql & " and a0205 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
         End If
         If Text1 <> MsgText(601) Then
            strSql = strSql & " and ax202 >= '" & Text1 & "'"
         End If
         If Text3 <> MsgText(601) Then
            strSql = strSql & " and ax202 <= '" & Text3 & "'"
         End If
         'Modify by Amy 2020/04/14 公司別改下拉 原:Text4
         If strCmp <> MsgText(601) Then
            strSql = strSql & " and ax201 = '" & strCmp & "'"
         End If
         strSum = strSql
         strSql = "select * from acc021, acc010, acc020 where ax205 = a0101 (+) and ax201 = a0201 and ax202 = a0202" & strSql & " order by ax201 asc, ax202 asc, ax203 asc"
         adoadodc1.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
   End Select
   Adodc1.Recordset.Requery
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

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text3_GotFocus()
   TextInverse Text3
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Mark by Amy 2020/04/14 公司別改下拉
'Private Sub Text4_Change()
'   If Text4 = MsgText(601) Then
'      Exit Sub
'   End If
'   '20140122START Add By eric
'   If Text4 <> "1" And Text4 <> "J" Then
'      MsgBox "公司別僅能為 1 或 J ! (1:台一/J:智權)"
'      Text4.Text = ""
'      Text4.SetFocus
'   End If
'   '20140122END
'   Text5 = A0802Query(Text4)
'End Sub
'
'Private Sub Text4_GotFocus()
'   TextInverse Text4
'   '20140122START Add By eric
'   CloseIme
'   '20140122END
'End Sub

''20140122START Add By eric
'Private Sub Text4_KeyPress(KeyAscii As Integer)
'   KeyAscii = UpperCase(KeyAscii)
'End Sub
'
''20140120START By eric
'Private Sub Text4_LostFocus()
'   If Text4.Text = "" Then
'      MsgBox "公司別不可空白! (1:台一/J:智權)"
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
         'Modify by Amy 2020/04/14 +HasShowMsg
         If FormCheck(HasShowMsg) Then
            Screen.MousePointer = vbHourglass
            QueryTable
            SumShow
            Screen.MousePointer = vbDefault
            Exit Sub
         ElseIf HasShowMsg = False Then
         'end 2020/04/14
            MsgBox MsgText(181), , MsgText(5)
         End If
   End Select
   KeyEnter KeyCode
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

'*************************************************
'  計算並顯示合計
'
'*************************************************
Public Sub SumShow()
   adoaccsum.CursorLocation = adUseClient
   Select Case strAccount
      Case "2"
         strSum = "select sum(ax306), sum(ax307) from acc031, acc010, acc030 where acc031.ax305 = acc010.a0101 (+) and acc031.ax301 = acc030.a0301 and acc031.ax302 = acc030.a0302" & strSum
         'If Text4 <> MsgText(601) Then
         '   strSum = "select sum(ax306), sum(ax307) from (" & strSum & ") where ax301 = '" & Text4 & "'"
         'End If
         adoaccsum.Open strSum, adoTaie, adOpenStatic, adLockReadOnly
      Case Else
         strSum = "select sum(ax206), sum(ax207) from acc021, acc010, acc020 where acc021.ax205 = acc010.a0101 (+) and acc021.ax201 = acc020.a0201 and acc021.ax202 = acc020.a0202" & strSum
         'If Text4 <> MsgText(601) Then
         '   strSum = "select sum(ax206), sum(ax207) from (" & strSum & ") where ax201 = '" & Text4 & "'"
         'End If
         adoaccsum.Open strSum, adoTaie, adOpenStatic, adLockReadOnly
   End Select
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0).Value) Then
         Text6 = MsgText(601)
      Else
         Text6 = Format(adoaccsum.Fields(0).Value, FDollar)
      End If
      If IsNull(adoaccsum.Fields(1).Value) Then
         Text2 = MsgText(601)
      Else
         Text2 = Format(adoaccsum.Fields(1).Value, FDollar)
      End If
   Else
      Text6 = MsgText(601)
      Text2 = MsgText(601)
   End If
   strSql = MsgText(601)
   adoaccsum.Close
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck(HasShowMsg As Boolean) As Boolean
   'Add by Amy 2020/04/14
   Dim bCancel As Boolean
   
   If Trim(Combo1) <> MsgText(601) Then
      Call Combo1_Validate(bCancel)
      If bCancel = True Then
        HasShowMsg = True
        FormCheck = False
        Exit Function
      End If
   End If
   'end 2020/04/14
   If Text1 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Text3 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
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




