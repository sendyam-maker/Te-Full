VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Frmacc4230 
   AutoRedraw      =   -1  'True
   Caption         =   "科目餘額查詢"
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
      Left            =   1320
      TabIndex        =   0
      Top             =   210
      Width           =   3500
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc4230.frx":0000
      Height          =   3372
      Left            =   240
      TabIndex        =   13
      Top             =   1080
      Width           =   8292
      _ExtentX        =   14631
      _ExtentY        =   5953
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
      Caption         =   "科目餘額資料"
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "a0405"
         Caption         =   "會計科目代號"
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
         DataField       =   "a0102"
         Caption         =   "會計科目名稱"
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
         DataField       =   "a0408"
         Caption         =   "目前餘額"
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
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   1470.047
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3809.764
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   2505.26
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   312
      Left            =   240
      Top             =   960
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
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   6240
      TabIndex        =   12
      Top             =   600
      Width           =   852
      _ExtentX        =   1508
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
   Begin VB.TextBox Text2 
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
      Left            =   3240
      TabIndex        =   3
      Top             =   600
      Width           =   1572
   End
   Begin VB.TextBox Text10 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
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
      Left            =   6840
      TabIndex        =   10
      Top             =   210
      Visible         =   0   'False
      Width           =   1572
   End
   Begin VB.TextBox Text9 
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
      Left            =   6240
      TabIndex        =   1
      Top             =   210
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.TextBox Text7 
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
      Left            =   1320
      TabIndex        =   2
      Top             =   600
      Width           =   1572
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
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
      Left            =   5880
      TabIndex        =   9
      Top             =   4560
      Width           =   2415
   End
   Begin VB.Label Label4 
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
      Height          =   252
      Left            =   3000
      TabIndex        =   11
      Top             =   600
      Width           =   132
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   4800
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "合計"
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
      Left            =   5040
      TabIndex        =   8
      Top             =   4560
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   852
      Left            =   240
      Top             =   120
      Width           =   8292
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "年月"
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
      Left            =   5400
      TabIndex        =   7
      Top             =   600
      Width           =   732
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "會計科目"
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
      TabIndex        =   6
      Top             =   600
      Width           =   972
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "部門別"
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
      Left            =   5400
      TabIndex        =   5
      Top             =   210
      Visible         =   0   'False
      Width           =   735
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
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   210
      Width           =   735
   End
End
Attribute VB_Name = "Frmacc4230"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/13 Form2.0已修改 (無需修改)
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/3 日期欄已修改
Option Explicit
Public adoaccsum As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset

'Add by Amy 2020/04/08
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
'end 2020/04/08

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
   MaskEdBox1.Mask = Mid(DFormat, 1, 6)
   '20140124REMARK By eric (公司別欄位可選擇 1本所 或 2智權)
   'Text5 = "1"
   'Add by Amy 2020/04/08
   Combo1.AddItem "", 0
   Call Pub_SetCboCmp(Combo1, False, False, False, , 1)
   'end 2020/04/08
   OpenTable
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   Set Frmacc4230 = Nothing
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

'Mark by Amy 2020/04/08 公司別改下拉
'Private Sub Text5_Change()
'   If Text5 = MsgText(601) Then
'      Exit Sub
'   End If
'   '20140124START Add By eric
'   If Text5 <> "1" And Text5 <> "J" Then
'      MsgBox "公司別僅能為 1 或 J ! (1:台一/J:智權)"
'      Text5.Text = ""
'      Text5.SetFocus
'   End If
'   '20140124END
'   Text6 = A0802Query(Text5)
'End Sub
'
'Private Sub Text5_GotFocus()
'   TextInverse Text5
'   '20140124START Add By eric
'   CloseIme
'   '20140124END
'End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
Dim strCmp As String 'Add by Amy 2020/04/08

On Error GoTo Checking
   'Modify by Amy 2020/04/08 改下拉 原:Text5
   strCmp = Combo1
   If InStr(strCmp, "　") > 0 Then
        strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
   End If
   adoadodc1.CursorLocation = adUseClient
   Select Case strAccount
      Case "2"
         adoadodc1.Open "select a0101, a0102, a0501 as a0401, a0502 as a0402, a0503 as a0403, a0504 as a0404, a0505 as a0405, a0508 as a0408 from acc050, acc010 where acc050.a0505 = acc010.a0101 and a0503 = '" & strCmp & "' and a0505 >= '" & Text7 & "' and a0505 <= '" & Text2 & "' and a0501 = " & Val(Mid(MaskEdBox1.Text, 1, 3)) & " and a0502 = " & Val(Mid(MaskEdBox1.Text, 5, 2)) & " order by a0501 desc, a0502 desc, a0505 asc", adoTaie, adOpenStatic, adLockReadOnly
      Case Else
         adoadodc1.Open "select a0101, a0102, a0401, a0402, a0403, a0404, a0405, a0408 from acc040, acc010 where acc040.a0405 = acc010.a0101 and a0403 = '" & strCmp & "' and a0405 >= '" & Text7 & "' and a0405 <= '" & Text2 & "' and a0401 = " & Val(Mid(MaskEdBox1.Text, 1, 3)) & " and a0402 = " & Val(Mid(MaskEdBox1.Text, 5, 2)) & " order by a0401 desc, a0402 desc, a0405 asc", adoTaie, adOpenStatic, adLockReadOnly
   End Select
   'end 2020/04/08
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  查詢資料表(會計科目餘額資料)
'
'*************************************************
Public Sub QueryTable()
Dim strSql As String
Dim strCmp As String 'Add by Amy 2020/04/08

On Error GoTo Checking
   'Add by Amy 2020/04/08 原:Text5
   If Trim(Combo1) <> MsgText(601) Then
        strCmp = Combo1
        If InStr(strCmp, "　") > 0 Then
            strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
        End If
   End If
   If adoadodc1.State = adStateOpen Then
      adoadodc1.Close
   End If
   If adoaccsum.State = adStateOpen Then
      adoaccsum.Close
   End If
   adoadodc1.CursorLocation = adUseClient
   adoaccsum.CursorLocation = adUseClient
   Select Case strAccount
      Case "2"
         '20140124START Add By eric
         'Modify by Amy 2020/04/08 原:Text5
         If strCmp <> MsgText(601) Then
            strSql = strSql & " and a0503 = '" & strCmp & "'"
         End If
         '20140124END
         'end 2020/04/08
         If Text7 <> MsgText(601) Then
            strSql = strSql & " and a0505 >= '" & Text7 & "'"
         End If
         If Text2 <> MsgText(601) Then
            strSql = strSql & " and a0505 <= '" & Text2 & "'"
         End If
         If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> Mid(MsgText(29), 1, 6) Then
            strSql = strSql & " and a0501 = " & Val(Mid(MaskEdBox1.Text, 1, 3)) & ""
            strSql = strSql & " and a0502 = " & Val(Mid(MaskEdBox1.Text, 5, 2)) & ""
         End If
         adoadodc1.Open "select a0102, a0505 as a0405, sum(a0508) as a0408 from acc050, acc010 where acc050.a0505 = acc010.a0101 (+) and a0504 = '" & MsgText(55) & "'" & strSql & " group by a0102, a0505 order by a0505", adoTaie, adOpenStatic, adLockReadOnly
         adoaccsum.Open "select sum(a0508) from acc050 where A0504 = '" & MsgText(55) & "'" & strSql, adoTaie, adOpenStatic, adLockReadOnly
      Case Else
         '20140124START Add By eric
         'Modify by Amy 2020/04/08 原:Text5
         If strCmp <> MsgText(601) Then
            strSql = " and a0403 = '" & strCmp & "'"
         End If
         '20140124END
         If Text7 <> MsgText(601) Then
            strSql = strSql & " and a0405 >= '" & Text7 & "'"
         End If
         If Text2 <> MsgText(601) Then
            strSql = strSql & " and a0405 <= '" & Text2 & "'"
         End If
         If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> Mid(MsgText(29), 1, 6) Then
            strSql = strSql & " and a0401 = " & Val(Mid(MaskEdBox1.Text, 1, 3)) & ""
            strSql = strSql & " and a0402 = " & Val(Mid(MaskEdBox1.Text, 5, 2)) & ""
         End If
         adoadodc1.Open "select a0102, a0405, sum(a0408) as a0408 from acc040, acc010 where acc040.a0405 = acc010.a0101 (+) and a0404 = '" & MsgText(55) & "'" & strSql & " group by a0102, a0405 order by a0405", adoTaie, adOpenStatic, adLockReadOnly
         adoaccsum.Open "select sum(a0408) from acc040 where A0404 = '" & MsgText(55) & "'" & strSql, adoTaie, adOpenStatic, adLockReadOnly
   End Select
   Adodc1.Recordset.Requery
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0).Value) Then
         Text4 = MsgText(601)
      Else
         Text4 = Format(adoaccsum.Fields(0).Value, FDollar)
      End If
   End If
   adoaccsum.Close
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

Private Sub Text7_GotFocus()
   TextInverse Text7
End Sub

Private Sub Text9_Change()
   If Text9 = MsgText(601) Then
      Exit Sub
   End If
   Text10 = A0902Query(Text9)
End Sub

Private Sub Text9_GotFocus()
   TextInverse Text9
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Public Sub KeyDefine(KeyCode As Integer)
Dim bolShowMsg As Boolean 'Add by Amy 2020/04/08

   Select Case KeyCode
      Case vbKeyF12
         If FormCheck(bolShowMsg) Then
            Screen.MousePointer = vbHourglass
            QueryTable
            Screen.MousePointer = vbDefault
            Exit Sub
         Else
            If bolShowMsg = False Then MsgBox MsgText(181), , MsgText(5)
         End If
   End Select
   KeyEnter KeyCode
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck(ByRef bolMsg As Boolean) As Boolean
   Dim bolCancel As Boolean 'Add by Amy 2020/04/08
  
   'Modify by Amy 2020/04/08 公司別改下拉 原:Text5
   If Trim(Combo1) <> MsgText(601) Then
      Call Combo1_Validate(bolCancel)
      If bolCancel = False Then
        FormCheck = True
        Exit Function
      Else
        bolMsg = True '已show過訊息
        FormCheck = False
        Exit Function
      End If
   End If
   'end 2020/04/08
   If Text9 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Text7 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Text2 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   
   FormCheck = False
End Function

'Mark by Amy 2020/04/08 公司別改下拉
'20140124START Add By eric
'Private Sub Text5_KeyPress(KeyAscii As Integer)
'   KeyAscii = UpperCase(KeyAscii)
'End Sub
'
''20140124START Add By eric
'Private Sub Text5_LostFocus()
'   If Text5.Text = "" Then
'      MsgBox "公司別不可空白! (1:台一/J:智權)"
'      Text5.Text = ""
'      Text5.SetFocus
'      Exit Sub
'   End If
'End Sub
'end 2020/04/08
