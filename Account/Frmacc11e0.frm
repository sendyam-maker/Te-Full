VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc11e0 
   AutoRedraw      =   -1  'True
   Caption         =   "繳款書寄出明細"
   ClientHeight    =   5115
   ClientLeft      =   5640
   ClientTop       =   1830
   ClientWidth     =   9405
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5115
   ScaleWidth      =   9405
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc11e0.frx":0000
      Height          =   4125
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   9000
      _ExtentX        =   15875
      _ExtentY        =   7276
      _Version        =   393216
      AllowArrows     =   -1  'True
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   18
      TabAction       =   1
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
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "T0702"
         Caption         =   "名條收件人"
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
         DataField       =   "T0709"
         Caption         =   "回執客戶名稱"
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
         DataField       =   "T0707"
         Caption         =   "名條"
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
         DataField       =   "T0708"
         Caption         =   "分所"
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
         DataField       =   "T0703"
         Caption         =   "客戶地址"
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
         DataField       =   "T0705"
         Caption         =   "金額"
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
         DataField       =   "T0706"
         Caption         =   "份數"
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
         DataField       =   "T0704"
         Caption         =   "客戶電話"
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
            ColumnWidth     =   2129.953
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2594.835
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   540.284
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   510.236
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   3000.189
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            ColumnWidth     =   555.024
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1170.142
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   7656
      Picture         =   "Frmacc11e0.frx":0015
      Style           =   1  '圖片外觀
      TabIndex        =   2
      ToolTipText     =   "取消"
      Top             =   204
      Width           =   350
   End
   Begin VB.CommandButton Command1 
      Caption         =   "列印"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   8280
      TabIndex        =   3
      Top             =   204
      Width           =   975
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
      Height          =   315
      Left            =   1200
      MaxLength       =   9
      TabIndex        =   0
      Top             =   240
      Width           =   1572
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   240
      Top             =   600
      Visible         =   0   'False
      Width           =   1200
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
   Begin MSForms.TextBox Text2 
      Height          =   285
      Left            =   2760
      TabIndex        =   5
      Top             =   240
      Width           =   4845
      VariousPropertyBits=   671105051
      BackColor       =   14737632
      Size            =   "8546;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "PS : 不印名條於名條欄輸入 N, 寄分所於分所欄輸入 Y"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   1560
      TabIndex        =   6
      Top             =   600
      Width           =   6015
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Top             =   6120
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "客戶編號"
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
      TabIndex        =   4
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "Frmacc11e0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2022/2/17 Form2.0已修改 (Printer列印未改)
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/11/26 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit

Public adoacc0e0 As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
Public adoaccrpt111 As New ADODB.Recordset
Dim dllaccrpt111 As Object
Dim iCount As Integer  '2005/11/28 ADD BY SONIA
'預設印表機
Dim m_DefaultPrinter As String


Private Sub Command1_Click()
   Screen.MousePointer = vbHourglass
   If ProduceData Then
      PrintAddress    '地址條
   End If
   Screen.MousePointer = vbDefault
   MsgBox MsgText(212), , MsgText(21)
   Screen.MousePointer = vbHourglass
   '繳款書
   Set dllaccrpt111 = CreateObject("AccReport.ReportSelect")
   dllaccrpt111.Acc11e0 ReportTitle(1151), StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
   Me.SetFocus 'Add by Morgan 2004/11/15 將 Focus 設回
   Screen.MousePointer = vbDefault
   'Remove by Morgan 2008/7/18 不必再寄--瑞婷
   'Screen.MousePointer = vbHourglass
   'PrintPayNotice      '回執單
   'Screen.MousePointer = vbDefault
   'end 2008/7/18
   adoTaie.Execute "delete from acctmp07"
   FormRefresh
   Text1 = "X"
   Text2 = ""
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(133)
End Sub

Private Sub Command2_Click()
   FormDelete
   FormRefresh
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(133) & " " & MsgText(151)
End Sub

Private Sub DataGrid1_AfterColUpdate(ByVal ColIndex As Integer)
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   Adodc1.Recordset.UpdateBatch
End Sub

Private Sub DataGrid1_GotFocus()
   DataGrid1.col = 0
End Sub

Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
Dim intCounter As Integer

   Select Case KeyCode
      Case vbKeyReturn
         'Modify by Morgan 2006/10/30
         Select Case DataGrid1.col
            Case 0, 1, 2, 3, 4, 5, 6
               SendKeys "{TAB}"
            Case 7
               SendKeys "{DOWN}"
               For intCounter = 1 To 7
                  SendKeys "{LEFT}"
               Next intCounter
         End Select
   End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
   If KeyCode <> vbKeyEscape Then
      Frmacc0000.StatusBar1.Panels(1).Text = MsgText(133) & " " & MsgText(151)
   End If
End Sub

Private Sub Form_Load()
   '表單初始化
   PUB_InitForm Me, 9500, 5500
   OpenTable
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(133)
   Text1 = "X"   '2005/11/25 ADD BY SONIA
   iCount = 0    '2005/11/28 ADD BY SONIA
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   Set Frmacc11e0 = Nothing
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)
   Select Case KeyCode
      Case vbKeyInsert
         FormSave
         FormRefresh
   End Select
   KeyEnter KeyCode
End Sub

'*************************************************
'  儲存資料表
'
'*************************************************
Private Sub FormSave()
   '2005/11/25 ADD BY SONIA
   Select Case Len(Text1)
      Case 6
         Text1 = Text1 & "000"
      Case 8
         Text1 = Text1 & "0"
   End Select
   '2005/11/25 END
   adoquery.CursorLocation = adUseClient
   'edit by nick 2004/07/23 寄客戶應該要使用聯絡地址
   'adoquery.Open "select cu04, cu23, cu16 from customer where cu01 = '" & Mid(Text1, 1, 8) & "' and cu02 = '" & IIf(Mid(Text1, 9, 1) = "", "0", Mid(Text1, 9, 1)) & "'", adoTaie, adOpenStatic, adLockReadOnly
   'Modify by Morgan 2007/1/22 客戶地址若客戶狀態有資料時優先抓
   adoquery.Open "select cu04, NVL(CU80,cu30||' '||cu31), cu16 from customer where cu01 = '" & Mid(Text1, 1, 8) & "' and cu02 = '" & IIf(Mid(Text1, 9, 1) = "", "0", Mid(Text1, 9, 1)) & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      If IsNull(adoquery.Fields("cu04").Value) Then
         Text2 = ""
      Else
         Text2 = adoquery.Fields("cu04").Value
      End If
      adoTaie.Execute "delete from acctmp07 where t0701 = '" & Text1 & "'"
      '2005/11/25 MODIFY BY SONIA 加客戶名稱T0709
      'adoTaie.Execute "insert into acctmp07 (t0701, t0702, t0703, t0704) values ('" & Text1 & "', '" & IIf(IsNull(adoquery.Fields("cu04").Value), "", adoquery.Fields("cu04").Value) & "', '" & IIf(IsNull(adoquery.Fields(1).Value), "", adoquery.Fields(1).Value) & "', '" & IIf(IsNull(adoquery.Fields("cu16").Value), "", adoquery.Fields("cu16").Value) & "')"
      iCount = iCount + 1
      adoTaie.Execute "insert into acctmp07 (t0701, t0702, t0703, t0704,t0709,T0710) values ('" & Text1 & "', '" & IIf(IsNull(adoquery.Fields("cu04").Value), "", adoquery.Fields("cu04").Value) & "', '" & IIf(IsNull(adoquery.Fields(1).Value), "", adoquery.Fields(1).Value) & "', '" & IIf(IsNull(adoquery.Fields("cu16").Value), "", adoquery.Fields("cu16").Value) & "','" & IIf(IsNull(adoquery.Fields("cu04").Value), "", adoquery.Fields("cu04").Value) & "'," & iCount & ")"
      '2005/11/25 END
   Else
      Text2 = ""
   End If
   adoquery.Close
   Text1_GotFocus
End Sub

'*************************************************
'  重新整理資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoadodc1.CursorLocation = adUseClient
   '2005/11/28 MODIFY BY SONIA
   'adoadodc1.Open "select * from acctmp07 order by t0701 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoadodc1.Open "select * from acctmp07 order by t0710 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Set Adodc1.Recordset = adoadodc1
   iCount = adoadodc1.RecordCount '2005/11/28 ADD BY SONIA
Checking:
   Exit Sub
End Sub

'*************************************************
'  重新整理資料表
'
'*************************************************
Private Sub FormRefresh()
   Adodc1.Recordset.Requery
End Sub

'*************************************************
'  刪除資料表
'
'*************************************************
Private Sub FormDelete()
On Error GoTo Checking
   If Adodc1.Recordset.RecordCount <> 0 Then
      adoTaie.Execute "delete from acctmp07 where t0701 = '" & Adodc1.Recordset.Fields("t0701").Value & "'"
   End If
Checking:
   Exit Sub
End Sub

Private Sub Text1_GotFocus()
Dim intPos As Integer  '2005/11/25 ADD BY SONIA
      
   TextInverse Text1
   '2005/11/25 ADD BY SONIA
   If Len("" & Text1) > 0 Then
      intPos = InStr("" & Text1, "X")
      If intPos - 1 = 0 Then
         If Len("" & Text1) > 1 Then
            Text1.SelStart = 1
         Else
            Text1.SelStart = 2
         End If
      End If
      Text1.SelLength = Len("" & Text1) - 1
   End If
   '2005/11/25 END
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   '2005/11/25 MODIFY BY SONIA
   'If Text1 = MsgText(601) Then
   If Text1 = MsgText(601) Or Text1 = "X" Then
   '2005/11/25 END
      Exit Sub
   End If
   adoquery.CursorLocation = adUseClient
   adoquery.Open "select cu04 from customer where cu01 = '" & Mid(Text1, 1, 8) & "' and cu02 = '" & IIf(Mid(Text1, 9, 1) = "", "0", Mid(Text1, 9, 1)) & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      If IsNull(adoquery.Fields(0).Value) Then
         Text2 = ""
      Else
         Text2 = adoquery.Fields(0).Value
      End If
   Else
      Text2 = ""
      MsgBox MsgText(28), , MsgText(5)
      Cancel = True
      Text1.SetFocus
   End If
   adoquery.Close
End Sub

'*************************************************
'  列印付款通知單
'
'*************************************************
Private Sub PrintPayNotice()

   Dim strAmount As String, intLength As Integer
   
   Printer.Font = "新細明體"
   Printer.FontSize = 12
   adoquery.CursorLocation = adUseClient
   '2005/11/28 MODIFY BY SONIA
   'adoquery.Open "select * from acctmp07 order by t0701 asc", adoTaie, adOpenStatic, adLockReadOnly
   adoquery.Open "select * from acctmp07 order by t0710 asc", adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      Screen.MousePointer = vbDefault
      MsgBox MsgText(100) & ReportTitle(115), , MsgText(5)
      Screen.MousePointer = vbHourglass
   Else
      adoquery.Close
      Exit Sub
   End If
   Do While adoquery.EOF = False
      PrintNoticeHead
      Printer.CurrentX = 200
      Printer.CurrentY = 1500
      If IsNull(adoquery.Fields("t0709").Value) Then
         Printer.Print ""
      Else
         Printer.Print adoquery.Fields("t0709").Value & ReportSum(43)
      End If
      Printer.CurrentX = 7000
      Printer.CurrentY = 2500
      Printer.Print IIf(IsNull(adoquery.Fields("t0706").Value), 0, adoquery.Fields("t0706").Value)
      strAmount = "$" & Format(IIf(IsNull(adoquery.Fields("t0705").Value), 0, adoquery.Fields("t0705").Value), DDollar) & "**"
      intLength = Printer.TextWidth(strAmount)
      Printer.CurrentX = 11000 - intLength
      Printer.CurrentY = 2500
      Printer.Print strAmount
      
      PUB_PrintReceipt2 adoquery, 0, 1
   
      adoquery.MoveNext
      If adoquery.EOF = False Then
         Printer.NewPage
      End If
   Loop
   Printer.EndDoc
   adoquery.Close
End Sub

'*************************************************
'  列印付款通知單 (抬頭及報表格式)
'
'*************************************************
Private Sub PrintNoticeHead()

   Printer.FontSize = 14
   Printer.CurrentX = 3650
   Printer.CurrentY = 300
   Printer.Print A0802Query("1")
   Printer.FontSize = 12
   Printer.CurrentX = 200
   Printer.CurrentY = 1000
   Printer.Print ReportSum(35) & CFDate(ACDate(ServerDate))
   Printer.Line (200, 1250)-(11500, 1250)
   Printer.CurrentX = 700
   Printer.CurrentY = 2500
   Printer.Print ReportSum(52)
   Printer.CurrentX = 200
   Printer.CurrentY = 3500
   Printer.Print ReportSum(53)
   Printer.CurrentX = 6000
   Printer.CurrentY = 6000
   Printer.Print A0802Query("1")
   
End Sub

'*************************************************
'  產生報表資料
'
'*************************************************
Private Function ProduceData() As Boolean

On Error GoTo Checking
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
   adoTaie.Execute "delete from accrpt111"
   adoaccrpt111.CursorLocation = adUseClient
   adoaccrpt111.Open "select * from accrpt111", adoTaie, adOpenDynamic, adLockBatchOptimistic
   
' 繳款書寄出明細
   adoacc0e0.CursorLocation = adUseClient
   adoacc0e0.Open "select * from acctmp07 order by t0710 asc", adoTaie, adOpenStatic, adLockReadOnly
   If adoacc0e0.RecordCount = 0 Then
      ProduceData = False
      adoacc0e0.Close
      Exit Function
   Else
      ProduceData = True
   End If
   Do While adoacc0e0.EOF = False
      adoaccrpt111.AddNew
      adoaccrpt111.Fields("r11101").Value = strUserNum
      adoaccrpt111.Fields("r11104").Value = adoacc0e0.Fields("t0710").Value  '2005/11/28 ADD BY SONIA
      If IsNull(adoacc0e0.Fields("t0701").Value) Then
         adoaccrpt111.Fields("r11102").Value = Null
      Else
         adoaccrpt111.Fields("r11102").Value = adoacc0e0.Fields("t0701").Value
      End If
      If IsNull(adoacc0e0.Fields("t0701").Value) Then
         adoaccrpt111.Fields("r11103").Value = Null
      Else
         adoaccrpt111.Fields("r11103").Value = adoacc0e0.Fields("t0702").Value
      End If
      If IsNull(adoacc0e0.Fields("t0705").Value) Then
         adoaccrpt111.Fields("r11106").Value = Null
      Else
         adoaccrpt111.Fields("r11106").Value = adoacc0e0.Fields("t0705").Value
      End If
      If IsNull(adoacc0e0.Fields("t0703").Value) Then
         adoaccrpt111.Fields("r11107").Value = ""
      Else
         adoaccrpt111.Fields("r11107").Value = adoacc0e0.Fields("t0703").Value
      End If
      '2005/11/25 ADD BY SONIA
      If IsNull(adoacc0e0.Fields("t0707").Value) Then
         adoaccrpt111.Fields("r11105").Value = ""
      Else
         adoaccrpt111.Fields("r11105").Value = adoacc0e0.Fields("t0707").Value
      End If
      If IsNull(adoacc0e0.Fields("t0709").Value) Then
         adoaccrpt111.Fields("r11109").Value = ""
      Else
         adoaccrpt111.Fields("r11109").Value = adoacc0e0.Fields("t0709").Value
      End If
      '2005/11/25 END
      If IsNull(adoacc0e0.Fields("t0704").Value) = False Then
         adoaccrpt111.Fields("r11107").Value = adoaccrpt111.Fields("r11107").Value & "  " & adoacc0e0.Fields("t0704").Value
      End If
      '2005/11/25 MODIFY BY SONIA
      'adoaccrpt111.Fields("R11110").Value = ComboItem(82)
      If IsNull(adoacc0e0.Fields("t0708").Value) = False Then
         adoaccrpt111.Fields("R11110").Value = ComboItem(83)
      Else
         adoaccrpt111.Fields("R11110").Value = ComboItem(82)
      End If
      '2005/11/25 END
      adoaccrpt111.Fields("R11108").Value = ComboItem(116)
      adoaccrpt111.UpdateBatch
      adoacc0e0.MoveNext
   Loop
   adoacc0e0.Close
   adoaccrpt111.Close
   StatusClear
Checking:
   If Err.Number = 0 Then
      Exit Function
   End If
   MsgBox Err.Description, , MsgText(5)
End Function

'*************************************************
'  列印地址條
'
'*************************************************
Private Sub PrintAddress()

   Dim intCounter As Integer
   Dim intLine As Integer
   
   adoacc0e0.CursorLocation = adUseClient
   '2005/11/25 MODIFY BY SONIA
   'adoacc0e0.Open "select * from acctmp07 order by t0701 asc, t0702 asc", adoTaie, adOpenStatic, adLockReadOnly
   adoacc0e0.Open "select * from acctmp07 WHERE T0707 IS NULL order by t0710 asc, t0702 asc", adoTaie, adOpenStatic, adLockReadOnly
   '2005/11/25 END
   If adoacc0e0.RecordCount <> 0 Then
      Screen.MousePointer = vbDefault
      MsgBox MsgText(100) & ReportTitle(1113), , MsgText(5)
      Screen.MousePointer = vbHourglass
   Else
      adoacc0e0.Close
      Exit Sub
   End If
   
   intCounter = 0
   
   'Modify by Morgan 2008/3/25 XP自定紙張需手動設定並將印表機預設為該紙張
   '9x
   If pub_OS = "1" Then
      Printer.Height = 2880
      Printer.Width = 13000
   Else
      Printer.PaperSize = PUB_GetPaperSize(2)
   End If
   'end 2008/3/25
   Printer.Font = "@新細明體"
   Printer.FontSize = 12
   Do While adoacc0e0.EOF = False
      intLine = 0
      If Not IsNull(adoacc0e0.Fields("t0703").Value) Then
         'Modify by Morgan 2005/9/28 若長度超過時跳行
         'Printer.CurrentX = 100
         'Printer.CurrentY = 300 + 2200 * intCounter
         'Printer.Print adoacc0e0.Fields("t0703").Value
         PUB_PrintAddress adoacc0e0.Fields("t0703").Value, intCounter, intLine
      End If
      'Modify by Morgan 2007/10/19 下移兩格--瑞婷
      'Printer.CurrentX = 100
      Printer.CurrentX = 600
      'end 2007/10/19
      Printer.CurrentY = 1000 + 2200 * intCounter
      If IsNull(adoacc0e0.Fields("t0702").Value) Then
         Printer.Print ""
      Else
         Printer.Print adoacc0e0.Fields("t0702").Value & MsgText(104)
      End If
      
      'Modify by Morgan 2006/1/18 改用大張地址條
'      intCounter = intCounter + 1
'      If intCounter = 3 Then
'         intCounter = 0
'         Printer.NewPage
'      End If
      Printer.NewPage
      
      adoacc0e0.MoveNext
   Loop
   adoacc0e0.Close
   Printer.Font = "新細明體"
   Printer.EndDoc
End Sub


