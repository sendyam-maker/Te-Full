VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Frmacc4190 
   AutoRedraw      =   -1  'True
   Caption         =   "分攤類別比率資料"
   ClientHeight    =   4584
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   6012
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4584
   ScaleWidth      =   6012
   Begin VB.CommandButton cmdCopy 
      Caption         =   "複製前一年資料"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   14
      Top             =   1680
      Width           =   1675
   End
   Begin VB.CommandButton Command3 
      Height          =   300
      Left            =   5160
      Picture         =   "Frmacc4190.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   600
      Width           =   350
   End
   Begin VB.CommandButton Command1 
      Caption         =   "產生部門分攤"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   4
      Top             =   1080
      Width           =   1675
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0FFFF&
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
      Left            =   1200
      MaxLength       =   3
      TabIndex        =   1
      Top             =   600
      Width           =   612
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc4190.frx":0102
      Height          =   2895
      Left            =   1200
      TabIndex        =   5
      Top             =   1080
      Width           =   2775
      _ExtentX        =   4890
      _ExtentY        =   5101
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   20
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
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "a0604"
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
      BeginProperty Column01 
         DataField       =   "a0605"
         Caption         =   "分攤比率(%)"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.00"
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
            ColumnWidth     =   768.189
         EndProperty
         BeginProperty Column01 
            Alignment       =   1
            ColumnWidth     =   1500.095
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00C0FFFF&
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
      Left            =   3000
      MaxLength       =   2
      TabIndex        =   2
      Top             =   600
      Width           =   612
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
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
      Left            =   3600
      TabIndex        =   12
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
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
      Left            =   1200
      MaxLength       =   1
      TabIndex        =   0
      Top             =   240
      Width           =   612
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
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
      Left            =   1800
      TabIndex        =   9
      Top             =   240
      Width           =   3732
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
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
      Left            =   2400
      TabIndex        =   7
      Top             =   4080
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
      Left            =   1200
      Top             =   960
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   2117
      _ExtentY        =   572
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
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   $"Frmacc4190.frx":0117
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   2130
      Left            =   4200
      TabIndex        =   15
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "年度"
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
      TabIndex        =   13
      Top             =   600
      Width           =   732
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   852
      Left            =   240
      Top             =   120
      Width           =   5412
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "分攤類別"
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
      Left            =   2040
      TabIndex        =   11
      Top             =   600
      Width           =   972
   End
   Begin VB.Label Label4 
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
      Height          =   252
      Left            =   360
      TabIndex        =   10
      Top             =   240
      Width           =   732
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "%"
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
      Left            =   3840
      TabIndex        =   8
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "Total"
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
      Left            =   1200
      TabIndex        =   6
      Top             =   4080
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   3840
      Visible         =   0   'False
      Width           =   132
   End
End
Attribute VB_Name = "Frmacc4190"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/10/25 Form2.0已修改 (無需修改)
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/3 日期欄已修改
Option Explicit
Public adoacc061 As New ADODB.Recordset
'Public adoacc060 As New ADODB.Recordset
Public adoacc090 As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Public adoins061 As New ADODB.Recordset                                 '20140108 Add by eric

'產生部門分攤
Private Sub Command1_Click()
   'Add by Amy 2024/08/15
   Dim strQ As String, strWhere As String
   
   If FormCheck("Cmd1") = False Then
      Exit Sub
   End If
   strWhere = "a0601='" & Text6 & "' And a0602='" & Text5 & "' And a0603='" & Text1 & "' "
   If Pub_GetField("Acc060", strWhere, "Count(*)") <> "0" Then
      MsgBox "資料已存在！", vbCritical
      Exit Sub
   End If
   'end 2024/08/15
'   adoacc060.CursorLocation = adUseClient
'   adoacc060.Open "select * from acc060 where a0601 = " & Val(Text6) & " and a0602 = '" & Text5 & "' and a0603 = '" & Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
'   If adoacc060.RecordCount = 0 Then
'      MsgBox MsgText(9), , MsgText(5)
'      adoacc060.Close
'      Exit Sub
'   End If
'   adoacc060.Close
   Frmacc4190_Save
   If strControlButton = MsgText(602) Then
      Exit Sub
   End If
   'Modify by Amy 2024/08/13 非L公司,Grid之部門不能有L部門的分攤比率；不管哪個公司,Grid之部門都不能有TOT
   strWhere = " And a0901<>'TOT' "
   If Text1 <> "L" And Val(Text6) >= 110 Then
      strWhere = strWhere & " And a0901<>'L' "
   End If
   strQ = "select * from acc090 where a0904 = '" & MsgText(602) & "' " & strWhere & " order by a0901 asc"
   adoacc090.CursorLocation = adUseClient
   adoacc090.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
   'end 2024/08/13
   Do While adoacc090.EOF = False
      adoTaie.Execute "insert into acc060 values (" & Val(Text6) & ", '" & Text5 & "', '" & Text1 & "', '" & adoacc090.Fields("a0901").Value & "', 0)"
      adoacc090.MoveNext
   Loop
   adoacc090.Close
   AdodcRefresh
   SumShow
End Sub

Private Sub Command1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

'查詢
Private Sub Command3_Click()
   'Modify by Amy 2024/08/15 檢查改至FormChcek
   If FormCheck("Cmd3") = False Then
      Exit Sub
   End If
   'If adoacc061.RecordCount = 0 Or Text1 = MsgText(601) Or Text5 = MsgText(601) Or Text6 = MsgText(601) Then
   If adoacc061.RecordCount = 0 Then
   'end 2024/08/15
      Exit Sub
   End If
   adoacc061.Find "ax601 = " & Val(Text6) & "", 0, adSearchForward, 1
   If adoacc061.EOF = False Then
      adoacc061.Find "ax602 = '" & Text5 & "'", 0, adSearchForward, adoacc061.Bookmark
      If adoacc061.EOF = False Then
         adoacc061.Find "ax603 = '" & Text1 & "'", 0, adSearchForward, adoacc061.Bookmark
         'Modify by Amy 2024/08/13 避免第一筆查有資料,第二筆沒資料,畫面又停留在第二筆後操作其他按鈕,導致錯誤
         If adoacc061.EOF = False Then
'            AdodcRefresh
'            SumShow
'            RecordShow
         Else
            MsgBox MsgText(33), , MsgText(5)
            adoacc061.MoveFirst
         End If
      Else
         MsgBox MsgText(33), , MsgText(5)
         adoacc061.MoveFirst
      End If
   Else
      MsgBox MsgText(33), , MsgText(5)
      adoacc061.MoveFirst
   End If
   AdodcRefresh
   SumShow
   RecordShow
   'end 2024/08/13
End Sub

Private Sub Command3_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Command3_Click
         Exit Sub
   End Select
   KeyEnter KeyCode
End Sub

Private Sub DataGrid1_AfterColUpdate(ByVal ColIndex As Integer)
   'Add by Amy 2024/08/13 輸完資料直接按存檔不會更新,再按查詢鈕才更新
   If DataGrid1.AllowUpdate = False Then Exit Sub
   
   Select Case ColIndex
      Case 1
         If DataGrid1.Columns(1).Text = MsgText(601) Then
            DataGrid1.Columns(1).Value = 0
         End If
         Adodc1.Recordset.UpdateBatch
         SumShow
         'Mark by Amy 2020/11/17 改至存檔前檢查
'         If Val(Text3) > 100 Then
'            MsgBox MsgText(49), , MsgText(5)
'            DataGrid1.Columns(1).Value = 0
'            Adodc1.Recordset.UpdateBatch
'            SumShow
'            Exit Sub
'         End If
   End Select
End Sub

'Add by Amy 2024/08/15
Private Sub DataGrid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
   Dim strWhere As String
   
   Select Case ColIndex
      Case 0 '部門
         If IsNumeric(DataGrid1.Columns(ColIndex).Value) Then
            MsgBox "不可輸入數字！", vbCritical
            Cancel = True
         End If
         '已有部門不可改
         strWhere = "a0904='Y' And a0901<>'TOT' And a0901='" & DataGrid1.Columns(ColIndex).Value & "'"
         If Text1 <> "L" And Val(Text6) >= 110 Then strWhere = strWhere & "And a0901<>'L' "
         '判斷是否有此部門
         If Pub_GetField("Acc090", strWhere, "A0901") = DataGrid1.Columns(ColIndex).Value Then
            strWhere = "a0601='" & Text6 & "' And a0602='" & Text5 & "' And a0603='" & Text1 & "' " & _
                                 "And a0604='" & DataGrid1.Columns(ColIndex).Value & "'"
            If Pub_GetField("Acc060", strWhere, "A0604") = DataGrid1.Columns(ColIndex).Value Then
               MsgBox "資料已存在！", vbCritical
               Cancel = True
            End If
         Else
            MsgBox "無此部門別！", vbCritical
            Cancel = True
         End If
      Case 1 '比例
         If Not IsNumeric(DataGrid1.Columns(ColIndex).Value) Then
            MsgBox "只可輸入數字！", vbCritical
            Cancel = True
         End If
   End Select
End Sub

'Add by Amy 2024/08/15
Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Form_Activate()
   strFormName = Name
   If strCon1 = MsgText(601) Then
      Exit Sub
   End If
   adoacc061.Find "ax601 = " & Val(strCon1) & "", 0, adSearchForward, 1
   If adoacc061.EOF = False Then
      adoacc061.Find "ax602 = '" & strCon2 & "'", 0, adSearchForward, adoacc061.Bookmark
      If adoacc061.EOF = False Then
         adoacc061.Find "ax603 = '" & strCon3 & "'", 0, adSearchForward, adoacc061.Bookmark
         If adoacc061.EOF = False Then
            FormShow
            AdodcRefresh
            SumShow
            RecordShow
         End If
      End If
   End If
   strCon1 = MsgText(601)
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 6000
   Me.Height = 5000
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath1)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   OpenTable
   If adoacc061.RecordCount <> 0 Then
      adoacc061.MoveLast
      adoacc061.MoveFirst
      RecordShow
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
      Cancel = 1
      Exit Sub
   End If
   StatusClear
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc4190 = Nothing
End Sub

Private Sub Text1_Change()
   If Text1 = MsgText(601) Then
      Exit Sub
   End If
   Text2 = A0802Query(Text1)
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii) 'Add by Amy 2020/04/07
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Text6.SetFocus
         Exit Sub
   End Select
   KeyEnter KeyCode
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
   Dim strQ As String, strWhere As String 'Add by Amy 2024/08/13
   
On Error GoTo Checking
   If adoacc061.State <> adStateClosed Then adoacc061.Close 'Add by Amy 2020/04/07
   adoacc061.CursorLocation = adUseClient
   'Modify by Amy 2024/08/30 排序改為 公司別+年度+類別-秀玲 原:order by ax601 asc, ax602 asc, ax603 asc
   adoacc061.Open "select * from acc061 order by ax603, ax601, ax602", adoTaie, adOpenDynamic, adLockBatchOptimistic
   
  'Modify by Amy 2024/08/13 非L公司,Grid之部門不能有L部門的分攤比率；不管哪個公司,Grid之部門都不能有TOT
   strWhere = " And a0604<>'TOT' "
   If Text1 <> "L" And Val(Text6) >= 110 Then
      strWhere = strWhere & " And a0604<>'L' "
   End If
   strQ = "select * from acc060 where a0603 = '" & Text1 & "' and a0601 = " & Val(Text6) & " and a0602 = '" & Text5 & "' " & strWhere & _
               "order by a0604 asc"
  If adoadodc1.State <> adStateClosed Then adoadodc1.Close 'Add by Amy 2020/04/07
  adoadodc1.CursorLocation = adUseClient
  adoadodc1.Open strQ, adoTaie, adOpenDynamic, adLockBatchOptimistic
  Set Adodc1.Recordset = adoadodc1
  
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  顯示欄位資料(費用分攤比率資料)
'
'*************************************************
Public Sub FormShow()
   If IsNull(adoacc061.Fields("ax603").Value) Then
      Text1 = MsgText(601)
   Else
      Text1 = adoacc061.Fields("ax603").Value
   End If
   If IsNull(adoacc061.Fields("ax601").Value) Then
      Text6 = MsgText(601)
   Else
      Text6 = adoacc061.Fields("ax601").Value
   End If
   If IsNull(adoacc061.Fields("ax602").Value) Then
      Text5 = MsgText(601)
   Else
      Text5 = adoacc061.Fields("ax602").Value
   End If
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   If strSaveConfirm <> MsgText(3) And strSaveConfirm <> MsgText(4) Then
      Exit Sub
   End If
   If ExistCheck("acc080", "a0801", Text1, Label4) = False Then
      Cancel = True
      Exit Sub
   End If
End Sub

Private Sub Text5_Change()
   If Text5 = MsgText(601) Then
      Exit Sub
   End If
   Text4 = A0702Query(Text5)
End Sub

Private Sub Text5_GotFocus()
   TextInverse Text5
End Sub

Private Sub Text5_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Text1.SetFocus
         Exit Sub
   End Select
   KeyEnter KeyCode
End Sub

'*************************************************
'  計算並顯示合計
'
'*************************************************
Public Sub SumShow()
   Dim strQ As String, strWhere As String 'Add by Amy 2024/08/13
   'Modify by Amy 2024/08/13 非L公司,Grid之部門不能有L部門的分攤比率；不管哪個公司,Grid之部門都不能有TOT
   strWhere = " And a0604<>'TOT' "
   If Text1 <> "L" And Val(Text6) >= 110 Then
      strWhere = strWhere & " And a0604<>'L' "
   End If
   strQ = "select sum(a0605) from acc060 where a0603 = '" & Text1 & "' and a0601 = " & Val(Text6) & " and a0602 = '" & Text5 & "' " & _
                  strWhere
   adoaccsum.CursorLocation = adUseClient
   adoaccsum.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
   'end 2024/08/13
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0).Value) Then
         Text3 = MsgText(601)
      Else
         Text3 = Format(adoaccsum.Fields(0).Value, strPercent)
      End If
   End If
   adoaccsum.Close
End Sub

'*************************************************
'  重新整理 Adodc 之資料
'
'*************************************************
Public Sub AdodcRefresh()
   Dim strQ As String, strWhere As String 'Add by Amy 2024/08/13
   
On Error GoTo Checking
   'Modify by Amy 2024/08/13 非L公司,Grid之部門不能有L部門的分攤比率；不管哪個公司,Grid之部門都不能有TOT
   strWhere = " And a0604<>'TOT' "
   If Text1 <> "L" And Val(Text6) >= 110 Then
      strWhere = strWhere & " And a0604<>'L' "
   End If
   strQ = "select * from acc060 where a0603 = '" & Text1 & "' and a0601 = " & Val(Text6) & " and a0602 = '" & Text5 & "' " & strWhere & _
               "order by a0604 asc"
   adoadodc1.Close
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open strQ, adoTaie, adOpenDynamic, adLockBatchOptimistic
   'end 2024/08/13
   Adodc1.Recordset.Requery
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  顯示筆數
'
'*************************************************
Public Sub RecordShow()
   'Modify by Amy 2024/08/13 +if
   If adoacc061.EOF = False Then
      Frmacc0000.StatusBar1.Panels(2).Text = adoacc061.Bookmark & MsgText(35) & adoacc061.RecordCount
   Else
      Frmacc0000.StatusBar1.Panels(2).Text = ""
   End If
End Sub

Private Sub Text5_Validate(Cancel As Boolean)
   If Text5 = MsgText(601) Then Exit Sub 'Add by Amy 2024/08/15
   
   If ExistCheck("acc070", "a0701", Text5, Label1) = False Then
      Cancel = True
      Exit Sub
   End If
End Sub

Private Sub Text6_GotFocus()
   TextInverse Text6
End Sub

Private Sub Text6_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Text5.SetFocus
         Exit Sub
   End Select
   KeyEnter KeyCode
End Sub

'20140108START Add by eric
'2014/1/27 modify by sonia 再修改為複製前一年資料(原為複製1公司至J公司)
Private Sub cmdCopy_Click()
   
   Dim stTitle As String
   Dim iYear As Integer
   Dim stYear As String, stPeriod As String
 
   
   stTitle = "複製前一年資料"
   iYear = Val(Left(strSrvDate(1), 4)) - 1911
  
   Do
      stYear = InputBox("請輸入欲複製到的(新)年度:", stTitle, iYear)
      If stYear = "" Then Exit Sub
    
      If Not IsNumeric(stYear) Then
         If MsgBox("年度輸入錯誤！", vbRetryCancel, stTitle) = vbCancel Then
            Exit Sub
         End If
      Else
         Exit Do
      End If
   Loop
   
   
   If doCopy(stYear, intI) = True Then
      If intI > 0 Then
         Call OpenTable 'Add by Amy 2020/04/07 解決無法即時查詢
         Call AdodcRefresh 'Add by Amy 2020/08/08
         MsgBox "已複製 " & stYear - 1 & " 年度資料至 " & stYear & " 年度！"
      Else
         MsgBox "無資料可供複製！"
      End If
   End If
   
End Sub
'20140108END Add by eric

'20140108START Add by eric
Private Function doCopy(p_stToYear As String, Optional p_iRec As Integer) As Boolean

   Dim strACC060 As String
   Dim acc061(1 To 6) As String
   Dim strWhere As String 'Add by Amy 2024/08/13
         
On Error GoTo ErrHnd

   'Modified by Lydia 2018/02/07 aacc_var已有adoTaie.begintrans包住
   'cnnConnection.BeginTrans
   If strSaveConfirm = MsgText(601) Then adoTaie.BeginTrans
   
           'Modified by Lydia 2018/02/07 cnnConnection=>adoTaie
          adoins061.CursorLocation = adUseClient
          adoins061.Open "select ax601,ax602,ax603,ax604,ax605,ax606,ax607,ax608,ax609 from ACC061 a where AX601= '" & p_stToYear - 1 & "'" & _
                        " and not exists(select * from ACC061 b where b.AX601='" & p_stToYear & "' and b.AX602=a.AX602 and b.AX603=a.AX603) ORDER BY a.ax601,a.ax602,a.ax603 ", adoTaie, adOpenStatic, adLockReadOnly
        
          If adoins061.RecordCount <> 0 Then
             adoins061.MoveFirst
          
             Do While Not adoins061.EOF
                 p_iRec = 0
                      
                 acc061(1) = Val("" & p_stToYear & "")
                 acc061(2) = "'" & adoins061.Fields("ax602") & "'"
                 acc061(3) = "'" & adoins061.Fields("ax603") & "'"
                 acc061(4) = "'" & strUserNum & "'"
                 acc061(5) = "'" & strSrvDate(2) & "'"
                 acc061(6) = "'" & ServerTime & "'"
                
                 'INSERT ACC061費用分攤比率資料(主檔)
                 strSql = "insert into acc061(ax601,ax602,ax603,ax604,ax605,ax606)" & _
                         " values(" & acc061(1) & "," & acc061(2) & "," & acc061(3) & "," & acc061(4) & "," & acc061(5) & "," & acc061(6) & ")"
                  'Modified by Lydia 2018/02/07 cnnConnection=>adoTaie
                 adoTaie.Execute strSql, p_iRec
   
                 If p_iRec = 0 Then
                    'Modify by Amy 2020/04/07 +L公司文字
                    MsgBox "主檔資料異常或已存在J或L公司資料，請確認後重新執行！", vbCritical
                    adoins061.Close
                    GoTo ErrHnd
                 End If
                
                 'INSERT ACC060費用分攤比率資料(交易檔) J/L公司資料
                 'Modify by Amy 2024/08/13 非L公司,Grid之部門不能有L部門的分攤比率；不管哪個公司,Grid之部門都不能有TOT
                  strWhere = " And a0604<>'TOT' "
                  If Text1 <> "L" And Val(Text6) >= 110 Then
                     strWhere = strWhere & " And a0604<>'L' "
                  End If
                 strACC060 = "insert into acc060(a0601,a0602,a0603,a0604,a0605)" & _
                            " select  '" & acc061(1) & "'," & acc061(2) & "," & acc061(3) & ", a0604, a0605 " & _
                            " from ACC060 c " & _
                            " where c.A0601 = '" & p_stToYear - 1 & "' and c.A0602 = " & acc061(2) & " and c.A0603 = " & acc061(3) & "" & strWhere & _
                            " and not exists(select * from ACC060 d where d.A0601 = '" & p_stToYear & "' and d.A0602=" & acc061(2) & " and d.A0603 =" & acc061(3) & strWhere & " and d.A0604=c.A0604) "
                 'end 2024/08/13
                 p_iRec = 0
                  'Modified by Lydia 2018/02/07 cnnConnection=>adoTaie
                 adoTaie.Execute strACC060, p_iRec
  
                 If p_iRec = 0 Then
                    MsgBox "交易檔資料異常，請確認後重新執行！", vbCritical
                    adoins061.Close
                    GoTo ErrHnd
                 End If
                 adoins061.MoveNext
                      
             Loop
             adoins061.Close            '物件須先CLOSE
                 
          Else
             MsgBox "無符合條件資料可複製，請確認後重新執行！", vbCritical
             adoins061.Close
             'Modified by Lydia 2018/02/07
             'adoTaie.RollbackTrans
             'GoTo ErrHnd
             If strSaveConfirm = MsgText(601) Then adoTaie.RollbackTrans
             Exit Function
             'end 2018/02/07
          End If
          'end 2018/02/07
          
   'Modified by Lydia 2018/02/07 aacc_var已有adoTaie.begintrans包住
   'cnnConnection.CommitTrans
   If strSaveConfirm = MsgText(601) Then adoTaie.CommitTrans

      doCopy = True
   
ErrHnd:
      If Err.Number <> 0 Then
          'Modified by Lydia 2018/02/07 cnnConnection=>adoTaie
         'cnnConnection.RollbackTrans
         adoTaie.RollbackTrans
         MsgBox Err.Description, vbCritical
         Err.Clear
      Else
         Exit Function
      End If
      Screen.MousePointer = vbDefault
   
End Function
'20140108END Add by eric

'Add by Amy 2020/04/06 從aacc_sav搬過來
Public Sub Frmacc4190_Save()
Dim strSave As String

   On Error GoTo Checking
   'Memo by Amy 2024/08/15 原檢查程式搬至FormCheckm,避免有未檢查到的
   
   'Modify by Amy 2024/08/15 原修改存檔不會Run 此(更新修改人日時間)
   '                                                     拿掉 if 判斷,若修改又按[複製前一年資料]鈕,資料可能跳至第一筆,導致更新資料出現  違反唯一限制條件...
   'If strSaveConfirm = MsgText(3) Then
      If adoacc061.RecordCount <> 0 Then
         adoacc061.MoveFirst
         adoacc061.Find "ax601 = " & Val(Text6) & "", 0, adSearchForward, 1
         If adoacc061.EOF = False Then
            adoacc061.Find "ax602 = '" & Text5 & "'", 0, adSearchForward, adoacc061.Bookmark
            If adoacc061.EOF = False Then
               'Add by Morgan 2008/2/13
               If adoacc061.Fields("ax601") <> Val(Text6) Then
                  adoacc061.AddNew
               Else
                  adoacc061.Find "ax603 = '" & Text1 & "'", 0, adSearchForward, adoacc061.Bookmark
                  If adoacc061.EOF = False Then
                     'Add by Morgan 2008/2/13
                     If adoacc061.Fields("ax601") <> Val(Text6) Or adoacc061.Fields("ax602") <> Text5 Then
                        adoacc061.AddNew
                     End If
                  Else
                     adoacc061.AddNew
                  End If
               End If
            Else
               adoacc061.AddNew
            End If
         Else
            adoacc061.AddNew
         End If
      Else
         adoacc061.AddNew
      End If
   'End If
      
   adoacc061.Fields("ax601").Value = Val(Text6)
   adoacc061.Fields("ax602").Value = Text5
   adoacc061.Fields("ax603").Value = Text1
   If strSaveConfirm = MsgText(3) Then
      adoacc061.Fields("ax605").Value = Val(strSrvDate(2))
      adoacc061.Fields("ax606").Value = ServerTime
      adoacc061.Fields("ax604").Value = strUserNum
   Else
      adoacc061.Fields("ax608").Value = Val(strSrvDate(2))
      adoacc061.Fields("ax609").Value = ServerTime
      adoacc061.Fields("ax607").Value = strUserNum
   End If
   adoacc061.UpdateBatch
   AdodcRefresh
   SumShow
   RecordShow
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'Add by Amy 2020/11/17 +存檔前檢查,由aacc_var搬回
'Modify by Amy 2024/08/15 +stState
Public Function FormCheck(stState As String) As Boolean
   Dim bCancel As Boolean 'Add by Amy 2024/08/15
   
   FormCheck = False
   '公司別
   If Text1 = MsgText(601) Then
      MsgBox Label4 & "不可為空", , MsgText(5)
      Text1.SetFocus
      Exit Function
   End If
   Call Text1_Validate(bCancel)
   If bCancel = True Then
      Exit Function
   End If
   '年度
   If Text6 = MsgText(601) Then
      MsgBox Label2 & "不可為空", , MsgText(5)
      Text6.SetFocus
      Exit Function
   End If
   '分攤類別
   If Text5 = MsgText(601) Then
      MsgBox Label1 & "不可為空", , MsgText(5)
      Text5.SetFocus
      Exit Function
   End If
   Call Text5_Validate(bCancel)
   If bCancel = True Then
      Exit Function
   End If
   
   '修改
   If stState = "F3" Then
      If adoadodc1.RecordCount = 0 Then
         MsgBox "無資料可修改", , MsgText(5)
         Exit Function
      End If
   '存檔
   ElseIf stState = "F9" Then
      'Add by Amy 2024/08/13 輸完資料直接按存檔不會更新,因Focus 還在DataGrid1,不會觸發DataGrid1_AfterColUpdate
      Command3.SetFocus
      SumShow
      
      If Val(Text3) <> 100 Then
         MsgBox MsgText(51), , MsgText(5)
         Exit Function
      End If
   End If
 
    FormCheck = True
End Function

'Add by Amy 2024/08/13 將aacc_var程式搬回
Public Sub SetData(ByVal strKeyCode As String)
   Select Case strKeyCode
      Case "F2" '新增
         Frmacc4190_Clear
         AdodcRefresh
         Command1.Enabled = True
         DataGrid1.AllowUpdate = True
      Case "F3" '修改
         DataGrid1.AllowUpdate = True
      Case "F5" '刪除
         Frmacc4190_Delete
         Frmacc4190_Clear
      Case "F9" '存檔
         Frmacc4190_Save
         Command1.Enabled = False
         DataGrid1.AllowUpdate = False
      Case "F10" '取消
         '避免Grid 輸資料按取消資料仍為輸的資料,故重抓
         If Text1 <> MsgText(601) And Text6 <> MsgText(601) And Text5 <> MsgText(601) Then
            Command3_Click
         End If
   End Select
End Sub
