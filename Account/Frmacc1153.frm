VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Frmacc1153 
   AutoRedraw      =   -1  'True
   Caption         =   "本所案號收款資料輸入"
   ClientHeight    =   4695
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9405
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4695
   ScaleWidth      =   9405
   Begin VB.CommandButton Command1 
      Caption         =   "取消"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7785
      TabIndex        =   3
      Top             =   4200
      Width           =   1500
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1028
         SubFormatType   =   1
      EndProperty
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
      Left            =   1056
      TabIndex        =   1
      Top             =   4176
      Width           =   1572
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc1153.frx":0000
      Height          =   3852
      Left            =   72
      TabIndex        =   0
      Top             =   240
      Width           =   9252
      _ExtentX        =   16325
      _ExtentY        =   6800
      _Version        =   393216
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   14
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "case"
         Caption         =   "本所案號"
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
         DataField       =   "a1u03"
         Caption         =   "總收文號"
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
         DataField       =   "a0j09"
         Caption         =   "應收服務費"
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
      BeginProperty Column03 
         DataField       =   "a0j10"
         Caption         =   "應收規費"
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
         DataField       =   "T03"
         Caption         =   "已收服務費"
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
         DataField       =   "T04"
         Caption         =   "已收規費"
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
         DataField       =   "a1u04"
         Caption         =   "本次服務費"
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
         DataField       =   "a1u05"
         Caption         =   "本次規費"
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
      BeginProperty Column08 
         DataField       =   "a1u06"
         Caption         =   "本次扣繳額"
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
         Size            =   315
         BeginProperty Column00 
            Locked          =   -1  'True
            ColumnWidth     =   1305.071
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   945.071
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   929.764
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   884.976
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   929.764
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            ColumnWidth     =   1124.787
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            ColumnWidth     =   840.189
         EndProperty
         BeginProperty Column08 
            Alignment       =   1
            ColumnWidth     =   1184.882
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   336
      Left            =   72
      Top             =   120
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
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "收款金額"
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
      Left            =   96
      TabIndex        =   2
      Top             =   4176
      Width           =   972
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Top             =   3960
      Visible         =   0   'False
      Width           =   135
   End
End
Attribute VB_Name = "Frmacc1153"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/14 Form2.0已修改
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/11/26 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/28 日期欄已修改
Option Explicit

Public adocheck As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Dim douLAmount As Double
Dim douTax As Double
'Add by Morgan 2011/10/20
Dim bolNoUnloadCheck As Boolean
Dim douSAmount As Double
Dim douFAmount As Double


Private Sub Command1_Click()
   If MsgBox("程式將忽略金額檢查,直接回前畫面!!是否確定??", vbYesNo + vbDefaultButton2) = vbYes Then
      bolNoUnloadCheck = True
      Unload Me
   End If
End Sub

Private Sub DataGrid1_AfterColUpdate(ByVal ColIndex As Integer)
   Dim douTAmount(2) As Double
On Error GoTo Checking
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   With DataGrid1
      adocheck.CursorLocation = adUseClient
      'Modify by Morgan 2011/10/21 考慮拆收據情形
      adocheck.Open "select sum(a1u04), sum(a1u05) from acc1u0 where a1u03 = '" & Adodc1.Recordset.Fields("a1u03").Value & "' and a1u02='" & Adodc1.Recordset.Fields("a1u02").Value & "' and a1u01 <> '" & strCon1 & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adocheck.RecordCount <> 0 Then
         If IsNull(adocheck.Fields(0).Value) Then
            douTAmount(0) = 0
         Else
            douTAmount(0) = adocheck.Fields(0).Value
         End If
         If IsNull(adocheck.Fields(1).Value) Then
            douTAmount(1) = 0
         Else
            douTAmount(1) = adocheck.Fields(1).Value
         End If
      Else
         douTAmount(0) = 0
         douTAmount(1) = 0
      End If
      adocheck.Close
      'Modify by Morgan 2011/10/5
      'If (douTAmount(0) + douTAmount(1) + Val(.Columns(6).Value) + Val(.Columns(7).Value)) > Val(.Columns(2).Value) Then
      '   MsgBox MsgText(81), , MsgText(5)
      If (douTAmount(0) + Val(.Columns(6).Value)) > Val(.Columns(2).Value) Then
         MsgBox "收款服務費不可大於應收服務費...", , MsgText(5)
         Exit Sub
      End If
      
      If (douTAmount(1) + Val(.Columns(7).Value)) > (Val(.Columns(3).Value)) Then
         MsgBox MsgText(83), , MsgText(5)
         Exit Sub
      End If
      'Ken 91/11/01 小計改為銷帳規費
      'Adodc1.Recordset.Fields("a1u09").Value = Val(.Columns(6).Value) + Val(.Columns(7).Value)
      Adodc1.Recordset.Fields("a1u09").Value = 0
   End With
   'Added by Morgan 2022/4/27
   With Adodc1.Recordset
   strSql = "update acc1u0 set a1u04=" & CNULL("" & .Fields("a1u04"), True) & ", a1u05=" & CNULL("" & .Fields("a1u05"), True) & _
               " ,a1u06=" & CNULL("" & .Fields("a1u06"), True) & ", a1u09=" & CNULL("" & .Fields("a1u09"), True) & _
               " where a1u02='" & .Fields("a1u02") & "' and a1u01='" & .Fields("a1u01") & "' and a1u03='" & .Fields("a1u03") & "' "
   adoTaie.Execute strSql, intI
   End With
   'end 2022/4/27
   Adodc1.Recordset.UpdateBatch
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
   Err.Clear
   Adodc1.Recordset.Requery
End Sub

Private Sub DataGrid1_GotFocus()
Dim intCounter As Integer

   DataGrid1.col = 0
   For intCounter = 1 To 6
      SendKeys "{RIGHT}"
   Next intCounter
End Sub

Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Select Case DataGrid1.col
            Case 6
               SendKeys "{RIGHT}"
            Case 7
               SendKeys "{RIGHT}"
            Case 8
               SendKeys "{DOWN}"
               SendKeys "{LEFT}"
               SendKeys "{LEFT}"
         End Select
   End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 9500
   Me.Height = 5100
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath2)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   OpenTable
   StatusClear
End Sub

Private Function UnloadCheck() As Boolean
   Dim douTAmt As Double, douSumTax As Double
   Dim douSAmt As Double, douFAmt As Double

On Error GoTo ErrHnd

   adocheck.CursorLocation = adUseClient
   adocheck.Open "select sum(a1u04), sum(a1u05), sum(a1u06) from acc1u0 where a1u01 = '" & strCon1 & "' and a1u02 = '" & strCon2 & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adocheck.RecordCount <> 0 Then
      If IsNull(adocheck.Fields(0).Value) Then
         douTAmt = 0
      Else
         douTAmt = adocheck.Fields(0).Value
         douSAmt = adocheck.Fields(0).Value 'Add by Morgan 2011/10/20
      End If
      If IsNull(adocheck.Fields(1).Value) = False Then
         douTAmt = douTAmt + adocheck.Fields(1).Value
         douFAmt = adocheck.Fields(1).Value 'Add by Morgan 2011/10/20
      End If
      If IsNull(adocheck.Fields(2).Value) Then
         douSumTax = 0
      Else
         douSumTax = adocheck.Fields(2).Value
      End If
   Else
      douTAmt = 0
   End If
   adocheck.Close
   
   If Adodc1.Recordset.RecordCount <> 0 Then
      If douTAmt <> douLAmount Then
         MsgBox MsgText(84), , MsgText(5)
         DataGrid1.SetFocus
         strExitControl = MsgText(602)
         Exit Function
      End If
      If douSumTax <> douTax Then
         MsgBox MsgText(116), , MsgText(5)
         DataGrid1.SetFocus
         strExitControl = MsgText(602)
         Exit Function
      End If
      
      'Add by Morgan 2011/10/20
      If douSAmt <> douSAmount Then
         MsgBox "收款服務費實際不符!!"
         DataGrid1.SetFocus
         Exit Function
      End If
      If douFAmt <> douFAmount Then
         MsgBox "收款規費實際不符!!"
         DataGrid1.SetFocus
         Exit Function
      End If
      'end 2011/10/20
   End If
   'Modified by Morgan 2015/9/14 +a1u06=0 Ex.E10420653 分兩次收款,第2次扣繳全部
   adoTaie.Execute "delete from acc1u0 where a1u01 = '" & strCon1 & "' and a1u02 = '" & strCon2 & "' and a1u04 = 0 and a1u05 = 0 and a1u06=0 and a1u07 = 0 and a1u08 = 0 and a1u09 = 0 and a1u10 = 0"
   UnloadCheck = True
   Exit Function
ErrHnd:
   MsgBox Err.Description
   
End Function
Private Sub Form_Unload(Cancel As Integer)
   If bolNoUnloadCheck = False Then
      If UnloadCheck = False Then
         Cancel = 1
         Exit Sub
      End If
   Else
      bolNoUnloadCheck = False
   End If
   strCon1 = ""
   tool3_enabled
   Frmacc1151.Enabled = True
   Frmacc1151.DataGrid1.SetFocus
   Set Frmacc1153 = Nothing
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
   douSAmount = 0
   douFAmount = 0
   
On Error GoTo Checking
   'Add by Morgan 2011/10/5 考慮拆收據已收金要改抓1U0資料,又配合 DATAGRID 更新故要先寫暫存
   adoTaie.Execute "delete ACCTMP08 where T05='" & Me.Name & "' and T14='" & strUserNum & "'"
   adoTaie.Execute "insert into ACCTMP08(T01,T02,T03,T04,T05,T14) SELECT A1U02 T01,A1U03 T02,sum(a1u04) T03,sum(a1u05) T04,'" & Me.Name & "' T05,'" & strUserNum & "' T14 from acc1u0 b where a1u01 <>'" & strCon1 & "' and a1u02 = '" & strCon2 & "' group by A1U02,a1u03"
   
   adoadodc1.CursorLocation = adUseClient
   'Modify by Morgan 2011/10/4 考慮拆收據改抓 acc0j0
   'adoadodc1.Open "select cp01||cp02||cp03||cp04 as case, a1u03, nvl(cp16, 0) as cp16, nvl(cp17, 0) as cp17, cp73, cp74, a1u04, a1u05, a1u06, a1u09, a1u01, a1u02 from caseprogress, acc1u0 where cp09 = a1u03 and a1u01 = '" & strCon1 & "' and a1u02 = '" & strCon2 & "' order by cp01||cp02||cp03||cp04 asc, a1u03 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoadodc1.Open "select cp01||cp02||cp03||cp04 as case, a1u03, a0j09, a0j10,nvl(T03,0) T03,NVL(T04,0) T04, a1u04, a1u05, a1u06, a1u09, a1u01, a1u02 from acc1u0 a,acc0j0,caseprogress,ACCTMP08 where a1u01 = '" & strCon1 & "' and a1u02 = '" & strCon2 & "' and a0j01(+)=a1u03 and a0j13(+)=a1u02 and cp09(+) = a1u03 and t01(+)=a1u02 and t02(+)=a1u03 and t05(+)='" & Me.Name & "' and T14(+)='" & strUserNum & "' order by cp01||cp02||cp03||cp04 asc, a1u03 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   'Modified by Morgan 2022/4/27
   'Set Adodc1.Recordset = adoadodc1
   Set Adodc1.Recordset = PUB_CreateRecordset(adoadodc1, , , , Me.Name)
   'end 2022/4/27
   adocheck.CursorLocation = adUseClient
   adocheck.Open "select t0204, t0205, t0206 from acctmp02 where t0201 = '" & strCon1 & "' and t0202 = '" & strCon2 & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adocheck.RecordCount <> 0 Then
      If IsNull(adocheck.Fields(0).Value) Then
         douLAmount = 0
      Else
         douLAmount = adocheck.Fields(0).Value
         douSAmount = adocheck.Fields(0).Value '本次服務費 Add by Morgan 2011/10/20
      End If
      
      If IsNull(adocheck.Fields(1).Value) = False Then
         douLAmount = douLAmount + adocheck.Fields(1).Value
         douFAmount = adocheck.Fields(1).Value '本次規費 Add by Morgan 2011/10/20
      End If
      If IsNull(adocheck.Fields(2).Value) Then
         douTax = 0
      Else
         douTax = adocheck.Fields(2).Value
      End If
   Else
      douLAmount = 0
      douTax = 0
   End If
   Text3 = douLAmount
   adocheck.Close
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub
