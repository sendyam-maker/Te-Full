VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Frmacc3210 
   AutoRedraw      =   -1  'True
   Caption         =   "銀行帳號流動資金查詢"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9270
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5175
   ScaleWidth      =   9270
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc3210.frx":0000
      Height          =   4095
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   8800
      _ExtentX        =   15531
      _ExtentY        =   7223
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
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
      Caption         =   "流動資金資料"
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "t0101"
         Caption         =   "發票銀行名稱"
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
         DataField       =   "t0102"
         Caption         =   "銀行帳號"
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
         DataField       =   "t0107"
         Caption         =   "預測日期"
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
         DataField       =   "t0103"
         Caption         =   "期初餘額"
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
      BeginProperty Column04 
         DataField       =   "t0104"
         Caption         =   "收入"
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
         DataField       =   "t0105"
         Caption         =   "支出"
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
      BeginProperty Column06 
         DataField       =   "t0106"
         Caption         =   "帳號餘額"
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
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1289.764
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   959.811
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1049.953
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            ColumnWidth     =   929.764
         EndProperty
      EndProperty
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
      Left            =   2880
      TabIndex        =   1
      Top             =   240
      Width           =   1572
   End
   Begin VB.TextBox Text1 
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
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   1572
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   6720
      TabIndex        =   2
      Top             =   240
      Width           =   1575
      _ExtentX        =   2778
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
      Left            =   120
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
   Begin VB.Label Label2 
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
      Left            =   2640
      TabIndex        =   5
      Top             =   240
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   495
      Left            =   120
      Top             =   120
      Width           =   8800
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   -120
      Top             =   4800
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "最後傳票日期"
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
      Left            =   5160
      TabIndex        =   4
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "銀行別"
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
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "Frmacc3210"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/07 Form2.0已修改 DataGrid1
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/3 日期欄已修改
Option Explicit
Public adoacc0h0 As New ADODB.Recordset
Public adosubrsum As New ADODB.Recordset
Public adosubpsum As New ADODB.Recordset
Public adotmp01 As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset

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
   'Modify by Amy 2023/10/11 避免切畫面仍要調整,故調大小 原W9500 H5400/(lngWidth - Me.Width) / 2-瑞婷
   Me.Width = 9300
   Me.Height = 5640
   Me.Move 0, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath2)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   MaskEdBox1.Mask = DFormat
   OpenTable
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98) & " " & MsgText(150)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   AdodcDelete
   StatusClear
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc3210 = Nothing
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adotmp01.CursorLocation = adUseClient
   adotmp01.Open "select * from acctmp01 order by t0101 asc, t0102 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoTaie.Execute "delete from acctmp01"
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "select * from acctmp01 order by t0101 asc, t0102 asc", adoTaie, adOpenStatic, adLockReadOnly
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  計算票據預計收入、支出及預計餘額
'
'*************************************************
Public Sub Calculate()
Dim lngBeginAmt As Long, strSql As String
Dim intYear, intMonth As Integer
Dim adoquery As New ADODB.Recordset
Dim adoacc0b0 As New ADODB.Recordset
Dim strProDate As String
Dim adoacc0e0 As New ADODB.Recordset, Amt As Long 'Add by Amy 2013/08/22


   AdodcDelete
   adoacc0h0.CursorLocation = adUseClient
   If Text1 <> MsgText(601) Then
      strSql = " and a0h01 >= '" & Text1 & "'"
   End If
   If Text2 <> MsgText(601) Then
      strSql = strSql & " and a0h01 <= '" & Text2 & "'"
   End If
   adoacc0b0.CursorLocation = adUseClient
   adoacc0b0.Open "select * from acc0b0", adoTaie, adOpenStatic, adLockReadOnly
   If adoacc0b0.RecordCount = 0 Then
      If Mid(ServerDate, 5, 2) = 1 Then
         intMonth = 12
         intYear = Val(Mid(CFDate(ACDate(ServerDate)), 1, 3)) - 1
      Else
         intMonth = Val(Mid(ServerDate, 5, 2)) - 1
         intYear = Val(Mid(CFDate(ACDate(ServerDate)), 1, 3))
      End If
      strProDate = intYear & IIf(intMonth > 9, intMonth, "0" & intMonth) & "00"
   Else
      If IsNull(adoacc0b0.Fields("a0b02").Value) Then
         If Mid(ServerDate, 5, 2) = 1 Then
            intMonth = 12
            intYear = Val(Mid(CFDate(ACDate(ServerDate)), 1, 3)) - 1
         Else
            intMonth = Val(Mid(ServerDate, 5, 2)) - 1
            intYear = Val(Mid(CFDate(ACDate(ServerDate)), 1, 3))
         End If
         strProDate = intYear & IIf(intMonth > 9, intMonth, "0" & intMonth) & "00"
      Else
        intMonth = Val(Mid(CFDate(adoacc0b0.Fields("a0b02").Value), 5, 2))
        intYear = Val(Mid(CFDate(adoacc0b0.Fields("a0b02").Value), 1, 3))
        strProDate = adoacc0b0.Fields("a0b02").Value
      End If
   End If
   adoacc0b0.Close
   adoacc0h0.Open "select * from acc0h0, acc0g0 where a0h01 = a0g01" & strSql & " order by a0h01 asc, a0h02 asc", adoTaie, adOpenStatic, adLockReadOnly
   Do While adoacc0h0.EOF = False
      lngBeginAmt = 0
      adotmp01.AddNew
      adotmp01.Fields("t0101").Value = adoacc0h0.Fields("a0g02").Value
      adotmp01.Fields("t0102").Value = adoacc0h0.Fields("a0h02").Value
      adotmp01.Fields("t0103").Value = 0
      adotmp01.Fields("t0104").Value = 0
      adotmp01.Fields("t0105").Value = 0
      adotmp01.Fields("t0106").Value = 0
      If MaskEdBox1.Mask <> MsgText(601) Then
         adotmp01.Fields("t0107").Value = Val(FCDate(MaskEdBox1.Text))
      Else
         adotmp01.Fields("t0107").Value = Null
      End If
      adotmp01.UpdateBatch
      adoquery.CursorLocation = adUseClient
      'adoquery.Open "select a0408 from acc040 where a0403 = '1' and a0401 = " & intYear & "  and a0404 = '" & MsgText(55) & "' and a0405 = '" & adoacc0h0.Fields("a0h08").Value & "' and a0402 in (select max(a0402) from acc040 where a0403 = '1' and a0401 = " & intYear & " and a0405 = '" & adoacc0h0.Fields("a0h08").Value & "')", adoTaie, adOpenStatic, adLockReadOnly
      adoquery.Open "select a0408 from acc040 where a0403 = '1' and a0401 = " & intYear & "  and a0404 = '" & MsgText(55) & "' and a0405 = '" & adoacc0h0.Fields("a0h08").Value & "' and a0402 = " & intMonth & "", adoTaie, adOpenStatic, adLockReadOnly
      If adoquery.RecordCount <> 0 Then
         If IsNull(adoquery.Fields(0).Value) Then
            lngBeginAmt = 0
         Else
            lngBeginAmt = Val(adoquery.Fields(0).Value)
         End If
      Else
         lngBeginAmt = 0
      End If
      adoquery.Close
      adoTaie.Execute "update acctmp01 set t0103 = " & lngBeginAmt & " where t0101 = '" & adoacc0h0.Fields("a0g02").Value & "' and t0102 = '" & adoacc0h0.Fields("a0h02").Value & "'"
      '92.11.10 cancel by sonia
      ''未兌現
      'adosubrsum.CursorLocation = adUseClient
      ''adosubrsum.Open "select sum(ax206) from acc021, acc020 where ax201 = a0201 and ax202 = a0202 and ax205 = '" & adoacc0h0.Fields("a0h08").Value & "' and a0205 > " & Val(strProDate) & " and a0205 <= " & Val(FCDate(MaskEdBox1.Text)) & "", adoTaie, adOpenStatic, adLockReadOnly
      ''Ken 91/03/26 -- Start
      ''adosubrsum.Open "select sum(a0e11) from acc0e0, acc0h0 where a0e19 = a0h01 and a0e20 = a0h02 and a0h08 = '" & adoacc0h0.Fields("a0h08").Value & "' and a0e04 = '" & MsgText(18) & "' and (a0e10 <= " & Val(FCDate(MaskEdBox1.Text)) & " and a0e17 = 0 and a0e15 = 0 and a0e34 = 0 and a0e21 = 0 and a0e16 = 0)", adoTaie, adOpenStatic, adLockReadOnly
      'adosubrsum.Open "select sum(a0e11) from acc0e0, acc0h0 where a0e19 = a0h01 and a0e20 = a0h02 and a0h08 = '" & adoacc0h0.Fields("a0h08").Value & "' and a0e04 = '" & MsgText(18) & "' and (a0e10 > " & Val(FCDate(MaskEdBox1.Text)) & " and a0e17 = 0 and a0e15 = 0 and a0e34 = 0 and a0e21 = 0 and a0e16 = 0)", adoTaie, adOpenStatic, adLockReadOnly
      ''Ken 91/03/26 -- End
      'If adosubrsum.RecordCount <> 0 Then
      '   If IsNull(adosubrsum.Fields(0).Value) = False Then
      '      adoTaie.Execute "update acctmp01 set t0104 = " & Val(adosubrsum.Fields(0).Value) & " where t0101 = '" & adoacc0h0.Fields("a0g02").Value & "' and t0102 = '" & adoacc0h0.Fields("a0h02").Value & "'"
      '   End If
      'End If
      'adosubrsum.Close
      '92.11.10 end
      
      'Add by Amy 2013/08/22 +未收票據
      adoacc0e0.CursorLocation = adUseClient
      adoacc0e0.Open "select sum(a0e11) from acc0e0, acc0h0 where a0e19 = a0h01 and a0e20 = a0h02 and a0h08 = '" & adoacc0h0.Fields("a0h08").Value & "' and a0e04 = '" & MsgText(18) & "' and (a0e10 > " & Val(FCDate(MaskEdBox1.Text)) & " and a0e17 = 0 and a0e15 = 0 and a0e34 = 0 and a0e21 = 0)", adoTaie, adOpenStatic, adLockReadOnly
     
      If adoacc0e0.RecordCount <> 0 Then
         If IsNull(adoacc0e0.Fields(0).Value) Then
            Amt = 0
         Else
             Amt = Val(adoacc0e0.Fields(0).Value)
         End If
      Else
         Amt = 0
      End If
      adoacc0e0.Close
      'end 2013/08/22
      
      '已入帳
      adosubrsum.CursorLocation = adUseClient
      'Ken 91/03/26 -- Start
      'adosubrsum.Open "select sum(ax206) from acc021, acc020 where ax201 = a0201 and ax202 = a0202 and ax205 = '" & adoacc0h0.Fields("a0h08").Value & "' and a0205 > " & Val(strProDate) & "", adoTaie, adOpenStatic, adLockReadOnly
      adosubrsum.Open "select sum(ax206) from acc021, acc020 where ax201 = a0201 and ax202 = a0202 and ax205 = '" & adoacc0h0.Fields("a0h08").Value & "' and a0205 > " & Val(strProDate) & " and a0205 <= " & Val(FCDate(MaskEdBox1.Text)) & "", adoTaie, adOpenStatic, adLockReadOnly
      'Ken 91/03/26 -- End
      If adosubrsum.RecordCount <> 0 Then
         If IsNull(adosubrsum.Fields(0).Value) = False Then
            'Modify by Amy 2013/08/22 搬至下方
            'adoTaie.Execute "update acctmp01 set t0104 = t0104 + " & Val(adosubrsum.Fields(0).Value) & " +" & Amt & " where t0101 = '" & adoacc0h0.Fields("a0g02").Value & "' and t0102 = '" & adoacc0h0.Fields("a0h02").Value & "'"
            Amt = Amt + Val(adosubrsum.Fields(0).Value)
         End If
      End If
      adoTaie.Execute "update acctmp01 set t0104 = t0104 + " & Amt & " where t0101 = '" & adoacc0h0.Fields("a0g02").Value & "' and t0102 = '" & adoacc0h0.Fields("a0h02").Value & "'"
      adosubrsum.Close
      '92.11.10 cancel by sonia
      ''未兌領
      'adosubpsum.CursorLocation = adUseClient
      ''adosubpsum.Open "select sum(ax207) from acc021, acc020 where ax201 = a0201 and ax202 = a0202 and ax205 = '" & adoacc0h0.Fields("a0h08").Value & "' and a0205 > " & Val(strProDate) & " and a0205 <= " & Val(FCDate(MaskEdBox1.Text)) & "", adoTaie, adOpenStatic, adLockReadOnly
      ''Ken 91/03/26 -- Start
      ''adosubpsum.Open "select sum(a0e11) from acc0e0, acc0h0 where a0e01 = a0h01 and a0e07 = a0h02 and a0h08 = '" & adoacc0h0.Fields("a0h08").Value & "' and a0e04 = '" & MsgText(19) & "' and (a0e10 <= " & Val(FCDate(MaskEdBox1.Text)) & "  and a0e25 = 0 and (a0e37 = 0 or a0e37 is null))", adoTaie, adOpenStatic, adLockReadOnly
      'adosubpsum.Open "select sum(a0e11) from acc0e0, acc0h0 where a0e01 = a0h01 and a0e07 = a0h02 and a0h08 = '" & adoacc0h0.Fields("a0h08").Value & "' and a0e04 = '" & MsgText(19) & "' and (a0e10 > " & Val(FCDate(MaskEdBox1.Text)) & "  and a0e25 = 0 and (a0e37 = 0 or a0e37 is null))", adoTaie, adOpenStatic, adLockReadOnly
      ''Ken 91/03/26 -- End
      'If adosubpsum.RecordCount <> 0 Then
      '   If IsNull(adosubpsum.Fields(0).Value) = False Then
      '      adoTaie.Execute "update acctmp01 set t0105 = " & Val(adosubpsum.Fields(0).Value) & " where t0101 = '" & adoacc0h0.Fields("a0g02").Value & "' and t0102 = '" & adoacc0h0.Fields("a0h02").Value & "'"
      '   End If
      'End If
      'adosubpsum.Close
      '92.11.10 end
      
      'Add by Amy 2013/08/22 +未付票據
      adoacc0e0.CursorLocation = adUseClient
      adoacc0e0.Open "select sum(a0e11) from acc0e0, acc0h0 where a0e01 = a0h01 and a0e07 = a0h02 and a0h08 = '" & adoacc0h0.Fields("a0h08").Value & "' and a0e04 = '" & MsgText(19) & "' and (a0e10 > " & Val(FCDate(MaskEdBox1.Text)) & "  and a0e25 = 0 and (a0e37 = 0 or a0e37 is null))", adoTaie, adOpenStatic, adLockReadOnly
      If adoacc0e0.RecordCount <> 0 Then
         If IsNull(adoacc0e0.Fields(0).Value) Then
            Amt = 0
         Else
            Amt = Val(adoacc0e0.Fields(0).Value)
         End If
      Else
         Amt = 0
      End If
      adoacc0e0.Close
      'end 2013/08/22
      
      '已入帳
      adosubpsum.CursorLocation = adUseClient
      'Ken 91/03/26 -- Start
      'adosubpsum.Open "select sum(ax207) from acc021, acc020 where ax201 = a0201 and ax202 = a0202 and ax205 = '" & adoacc0h0.Fields("a0h08").Value & "' and a0205 > " & Val(strProDate) & "", adoTaie, adOpenStatic, adLockReadOnly
      adosubpsum.Open "select sum(ax207) from acc021, acc020 where ax201 = a0201 and ax202 = a0202 and ax205 = '" & adoacc0h0.Fields("a0h08").Value & "' and a0205 > " & Val(strProDate) & " and a0205 <= " & Val(FCDate(MaskEdBox1.Text)) & "", adoTaie, adOpenStatic, adLockReadOnly
      'Ken 91/03/26 -- End
      If adosubpsum.RecordCount <> 0 Then
         If IsNull(adosubpsum.Fields(0).Value) = False Then
            'Modify by Amy 2013/08/22 搬至下方
            'adoTaie.Execute "update acctmp01 set t0105 = t0105 + " & Val(adosubpsum.Fields(0).Value) & " where t0101 = '" & adoacc0h0.Fields("a0g02").Value & "' and t0102 = '" & adoacc0h0.Fields("a0h02").Value & "'"
            Amt = Amt + Val(adosubpsum.Fields(0).Value)
         End If
      End If
      adoTaie.Execute "update acctmp01 set t0105 = t0105 + " & Amt & " where t0101 = '" & adoacc0h0.Fields("a0g02").Value & "' and t0102 = '" & adoacc0h0.Fields("a0h02").Value & "'"
      adosubpsum.Close
      
      adotmp01.Close
      adotmp01.CursorLocation = adUseClient
      adotmp01.Open "select * from acctmp01 where t0101 = '" & adoacc0h0.Fields("a0g02").Value & "' and t0102 = '" & adoacc0h0.Fields("a0h02").Value & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
      If adotmp01.RecordCount <> 0 Then
         adotmp01.Fields("t0106").Value = Val(adotmp01.Fields("t0103").Value) + Val(adotmp01.Fields("t0104").Value) - Val(adotmp01.Fields("t0105").Value)
         adotmp01.UpdateBatch
      End If
      adoacc0h0.MoveNext
   Loop
   adoacc0h0.Close
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)
   Select Case KeyCode
      Case vbKeyF12
         If FormCheck Then
            Screen.MousePointer = vbHourglass
            Calculate
            AdodcRefresh
            Screen.MousePointer = vbDefault
            Exit Sub
         Else
            MsgBox MsgText(181), , MsgText(5)
         End If
   End Select
   KeyEnter KeyCode
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98) & MsgText(150)
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

'*************************************************
'  重新整理 Adodc 之資料
'
'*************************************************
Public Sub AdodcRefresh()
On Error GoTo Checking
   If adoadodc1.State = adStateOpen Then
      adoadodc1.Close
   End If
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "select * from acctmp01 order by t0101 asc, t0102 asc", adoTaie, adOpenStatic, adLockReadOnly
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

'*************************************************
'  清除資料表(銀行帳號流動資金預測查詢暫存檔)
'
'*************************************************
Private Sub AdodcDelete()
   adoTaie.Execute "delete from acctmp01"
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
   If Text1 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Text2 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If MaskEdBox1.Text <> MsgText(29) Then
      FormCheck = True
      Exit Function
   End If
   FormCheck = False
End Function

