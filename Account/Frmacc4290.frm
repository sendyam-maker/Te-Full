VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Frmacc4290 
   AutoRedraw      =   -1  'True
   Caption         =   "過帳前綜合損益餘額查詢及列印"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8925
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5295
   ScaleWidth      =   8925
   Begin VB.ComboBox CboCmp 
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
      Left            =   1080
      TabIndex        =   0
      Top             =   180
      Width           =   4700
   End
   Begin VB.TextBox Text3 
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
      Left            =   1080
      MaxLength       =   6
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox Text4 
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
      Left            =   2790
      MaxLength       =   6
      TabIndex        =   4
      Top             =   593
      Width           =   1335
   End
   Begin VB.ComboBox cmbPrinter 
      Height          =   300
      Left            =   1080
      Style           =   2  '單純下拉式
      TabIndex        =   8
      Top             =   960
      Width           =   3060
   End
   Begin VB.CommandButton cmdQuery 
      BackColor       =   &H00C0FFC0&
      Caption         =   "查詢(&Q)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7545
      Style           =   1  '圖片外觀
      TabIndex        =   5
      Top             =   240
      Width           =   1275
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C0FFC0&
      Caption         =   "列印(&P)"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7545
      Style           =   1  '圖片外觀
      TabIndex        =   6
      Top             =   600
      Width           =   1275
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc4290.frx":0000
      Height          =   3750
      Left            =   90
      TabIndex        =   10
      Top             =   1380
      Width           =   8760
      _ExtentX        =   15452
      _ExtentY        =   6615
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   14
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
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
      ColumnCount     =   6
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
         DataField       =   "ax205"
         Caption         =   "科目代號"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "###/##/##"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "A0102"
         Caption         =   "會計科目"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "###/##/##"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "sumcredit"
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
      BeginProperty Column04 
         DataField       =   "sumdebit"
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
      BeginProperty Column05 
         DataField       =   "amount"
         Caption         =   "餘額"
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
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   689.953
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   1769.953
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1440
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1440
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   1844.787
         EndProperty
      EndProperty
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   4920
      TabIndex        =   1
      Top             =   600
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   6
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
      Left            =   6120
      TabIndex        =   2
      Top             =   600
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   6
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
      Left            =   5355
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
      Caption         =   "Adodc2"
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
      Height          =   135
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label6 
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
      Left            =   5880
      TabIndex        =   15
      Top             =   623
      Width           =   135
   End
   Begin VB.Label Label1 
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
      Left            =   150
      TabIndex        =   14
      Top             =   180
      Width           =   735
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
      Height          =   255
      Left            =   150
      TabIndex        =   13
      Top             =   623
      Width           =   975
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
      Height          =   255
      Left            =   4320
      TabIndex        =   12
      Top             =   623
      Width           =   735
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
      Height          =   255
      Left            =   2520
      TabIndex        =   11
      Top             =   600
      Width           =   135
   End
   Begin VB.Label Label15 
      BackStyle       =   0  '透明
      Caption         =   "印表機"
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
      Left            =   150
      TabIndex        =   9
      Top             =   990
      Width           =   750
   End
   Begin VB.Label lblSales 
      BackStyle       =   0  '透明
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
      Left            =   7155
      TabIndex        =   7
      Top             =   450
      Width           =   1290
   End
End
Attribute VB_Name = "Frmacc4290"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2022/01/05 Form2.0已修改 DataGrid1/Printer列印未改
'Added by Lydia 2014/12/11 總帳-過帳前綜合損益查詢及列印
'2010/12/1 memo by sonia 員工編號欄已修改
'2010/8/18 sonia 日期欄已修改
Option Explicit
'公用
Dim rs4290 As New ADODB.Recordset
'頁次
Dim iPage As Integer
'欄位座標
Dim PLeft(0 To 20) As Integer
'Y座標
Dim intY As Integer
'暫存金額
Dim stMoneyTemp As String
'表頭欄位
Dim stSalesName As String, stTitle As String, stCustNo As String, stCustName As String
'預設印表機
Dim m_DefaultPrinter As String, m_Prn As Printer
'列高
Dim iRowPix As Integer
'邊界調整
Private Const intDefault As Integer = 500
'一個字元的點數
Dim iBytePix As Integer
'金錢格式
Private Const DDollar As String = "###,###,###,##0"
'群組
Dim stLstGroup1 As String, stCurGroup1 As String
Dim stCompName As String
'Dim arrSubtot(1 To 4) As Long '科目小計 'Modified by Lydia 2014/12/12 為了小數點,改成double
Dim arrSubtot(1 To 4) As Double '科目小計
Dim idx As Integer
Dim strRsQ As String 'Add by Amy 2017/08/17
Dim strCmp As String 'Add by Amy 2020/05/26

Private Function Process() As Boolean
   Dim stSQL As String, stAcc020Con As String, stAcc0w0Con As String, stAcc0l0Con As String, stCustNoCol As String
   Dim bolOK2Pring As Boolean
   Dim bCancel As Boolean 'Add by Amy 2020/05/26
   
On Error GoTo flgErr

   '公司別
   'Modify by Amy 2020/05/26 公司別改下拉
'   If Not (Text1 = "1" Or Text1 = "2" Or Text1 = "") Then
   If Trim(CboCmp) <> MsgText(601) Then
      Call CboCmp_Validate(bCancel)
      If bCancel = True Then
        CboCmp.SetFocus
        Exit Function
      End If
   End If
    'end 2020/05/26
   
   '年月
   If MaskEdBox1 = "" Or MaskEdBox2 = "" Then
      MsgBox "年月條件不可空白!!!", vbExclamation
      MaskEdBox1.SetFocus
      Exit Function
   ElseIf MaskEdBox1 > MaskEdBox2 Then
      MsgBox "年月條件範圍錯誤!!!", vbExclamation
      MaskEdBox1.SetFocus
      Exit Function
   End If
   
   '會計科目
   If Text3 = "" Or Text4 = "" Then
      MsgBox "會計科目不可空白!!!", vbExclamation
      Text3.SetFocus
      Exit Function
   ElseIf Text3 > Text4 Then
      MsgBox "會計科目範圍錯誤!!!", vbExclamation
      Text3.SetFocus
      Exit Function
   ElseIf Text3 < "4" Or Text3 >= "8" Then
      MsgBox "會計科目只可為4~7字頭!!!", vbExclamation
      Text3.SetFocus
      Exit Function
   ElseIf Text4 < "4" Or Text3 > "7" Then
      MsgBox "會計科目只可為4~7字頭!!!", vbExclamation
      Text4.SetFocus
      Exit Function
   End If
   
   stAcc020Con = ""
   '公司別
   'Modify by Amy 2020/05/26 公司別改下拉
'   If Text1 = "1" Then
'      stAcc020Con = stAcc020Con & " and a0201='" & Text1 & "'"
'      stCompName = A0802Query("1")
'   ElseIf Text1 = "2" Then
'      stAcc020Con = stAcc020Con & " and a0201='J'"
'      stCompName = A0802Query("J")
'   Else
'      stCompName = "台一　專利商標/智權"
'   End If
   strCmp = "": stCompName = ""
   If Trim(CboCmp) <> MsgText(601) Then
       strCmp = CboCmp
       If InStr(strCmp, "　") > 0 Then
             strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
       End If
        If InStr(strCmp, "+") > 0 Then
            stAcc020Con = stAcc020Con & " and a0201 In ('" & Replace(strCmp, "+", "','") & "') "
        Else
            stAcc020Con = stAcc020Con & " and a0201='" & strCmp & "'"
        End If
        stCompName = GetAccReportCmpN(strCmp, , True)
   End If
   'end 2020/05/26
   
   '收款日期
   If MaskEdBox1 <> MsgText(601) Then
      '2008/11/7 MODIFY BY SONIA
      'stAcc020Con = stAcc020Con & " and a0205 >= " & Val(FCDate(MaskEdBox1)) & "01"
      stAcc020Con = stAcc020Con & " and a0205 >= " & Val(FCDate(MaskEdBox1 & "/01"))
   End If
   If MaskEdBox2 <> MsgText(601) Then
      '2008/11/7 MODIFY BY SONIA
      'stAcc020Con = stAcc020Con & " and a0205 <= " & Val(FCDate(MaskEdBox2)) & "31"
      stAcc020Con = stAcc020Con & " and a0205 <= " & Val(FCDate(MaskEdBox2 & "/31"))
   End If
   
   '會計科目
   If Text3 <> "" Then
      stAcc020Con = stAcc020Con & " and aX205>='" & Text3 & "'"
   End If
   If Text4 <> "" Then
      stAcc020Con = stAcc020Con & " and aX205<='" & Text4 & "'"
      
   End If
   
   '會計科目4~7
   stAcc020Con = stAcc020Con & " and aX205>='4' and aX205<'8'"
   
   If Left(stAcc020Con, 4) = " and" Then stAcc020Con = Mid(stAcc020Con, 5, Len(stAcc020Con) - 4)
   
   'Modify by Amy 2020/05/26 公司別抓變數 原:Text1
   If strCmp = MsgText(601) Or InStr(strCmp, "+") > 0 Then '全部-公司別=空白
      stSQL = "SELECT aX201,aX205,A0102,A sumcredit,B sumdebit,DECODE(A0103,'1',A-B,B-A) amount,A0103,substr(aX205,1,1) Srt" & _
         " FROM ( SELECT DISTINCT ' ' aX201,aX205,A0102,A0103,SUM(aX206) A,SUM(aX207) B" & _
         " FROM ACC020,ACC021,ACC010 WHERE " & stAcc020Con & _
         " AND A0201=aX201(+) AND A0202=aX202(+) AND aX205=A0101(+) " & _
         "GROUP BY aX205,A0102,A0103 ) " & _
         "ORDER BY aX205,A0103 desc"
   Else
      stSQL = "SELECT aX201,aX205,A0102,A sumcredit,B sumdebit,DECODE(A0103,'1',A-B,B-A) amount,A0103,substr(aX205,1,1) Srt" & _
         " FROM ( SELECT DISTINCT aX201,aX205,A0102,A0103,SUM(aX206) A,SUM(aX207) B" & _
         " FROM ACC020,ACC021,ACC010 WHERE " & stAcc020Con & _
         " AND A0201=aX201(+) AND A0202=aX202(+) AND aX205=A0101(+) " & _
         "GROUP BY aX201,aX205,A0102,A0103 ) " & _
         "ORDER BY aX205,A0103 desc"
   End If
      '公司別為空白,但要有項目加總
      '   stSQL = "SELECT AX301,AX305,A0102,A sumcredit,B sumdebit,DECODE(A0103,'1',A-B,B-A) amount,A0103,substr(ax305,1,1) Srt" & _
      " FROM ( SELECT DISTINCT AX301,AX305,A0102,A0103,SUM(AX306) A,SUM(AX307) B" & _
      " FROM ACC030,ACC031,ACC011 WHERE " & stAcc030Con & _
      " AND A0301=AX301(+) AND A0302=AX302(+) AND AX305=A0101(+) " & _
      "GROUP BY AX301,AX305,A0102,A0103 ) " & _
      "ORDER BY AX301,AX305,A0103 desc"
   With rs4290
      If .State = adStateOpen Then .Close
      .CursorLocation = adUseClient
      .Open stSQL, adoTaie, adOpenForwardOnly, adLockReadOnly
      If .RecordCount > 0 Then
         Process = True
      Else
         MsgBox MsgText(28), , MsgText(5)
      End If
   End With
   
flgErr:
   
   If Err.Number <> 0 Then MsgBox Err.Description
   'Resume
End Function

Private Sub cmdPrint_Click()
   
   Screen.MousePointer = vbHourglass
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
   If rs4290.RecordCount > 0 Then
      If FormPrint = True Then
         MsgBox "列印完成！"
      End If
   End If
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
   Screen.MousePointer = vbDefault
   
End Sub

Private Function FormPrint() As Boolean

On Error GoTo flgErr

   '設定使用者所選擇的印表機成預設印表機
   For Each m_Prn In Printers
      If m_Prn.DeviceName = cmbPrinter.Text Then
         Set Printer = m_Prn
         Exit For
      End If
   Next

   GetPleft

   iPage = 0
   With Adodc1.Recordset
      .MoveFirst
      Do While Not .EOF
         stCurGroup1 = .Fields("aX201")
         If iPage = 0 Then
          '  stCompName = A0802Query(stCurGroup1)
            PrintHead
         ElseIf (stCurGroup1 <> stLstGroup1) Then
            Printer.NewPage
            PrintHead
           ' stCompName = A0802Query(stCurGroup1)
         Else
            NewLine
         End If
         Select Case Right(.Fields("Srt"), 1)
            Case "X"
               PrintLine 2
               PrintData
               PrintLine 3
            Case "Z"
               PrintData
               PrintLine 3
            Case Else
               PrintData
         End Select
         stLstGroup1 = stCurGroup1
         .MoveNext
      Loop
   End With
   NewLine
   '表尾
   Printer.FontSize = 14
   Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth("*** 結束 ***")) / 2
   Printer.CurrentY = intY
   Printer.FontBold = True
   Printer.Print "*** 結束 ***"
   Printer.EndDoc
   FormPrint = True

flgErr:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical

End Function

Private Sub AddExpRows(ByRef p_Rst As ADODB.Recordset)
   With p_Rst
      .AddNew
      .Fields("aX201") = stLstGroup1
      .Fields("A0102") = "營業收入："
      .Fields("AMOUNT") = arrSubtot(1)
      .Fields("Srt") = "4X"
      
      .AddNew
      .Fields("aX201") = stLstGroup1
      .Fields("A0102") = "營業支出："
      .Fields("AMOUNT") = arrSubtot(2)
      .Fields("Srt") = "6X"
      
      .AddNew
      .Fields("aX201") = stLstGroup1
      .Fields("A0102") = "營業損益："
      .Fields("AMOUNT") = arrSubtot(1) - arrSubtot(2)
      .Fields("Srt") = "6Z"
      
      .AddNew
      .Fields("aX201") = stLstGroup1
      .Fields("A0102") = "營業外收入："
      .Fields("AMOUNT") = arrSubtot(3)
      .Fields("Srt") = "71X"
      
      .AddNew
      .Fields("aX201") = stLstGroup1
      .Fields("A0102") = "營業外支出："
      .Fields("AMOUNT") = arrSubtot(4)
      .Fields("Srt") = "72X"
      
      .AddNew
      .Fields("aX201") = stLstGroup1
      .Fields("A0102") = "稅前淨損益："
      .Fields("AMOUNT") = arrSubtot(1) + arrSubtot(3) - arrSubtot(2) - arrSubtot(4)
      .Fields("Srt") = "8Z"
   End With
End Sub

Private Function SetGrid() As Boolean
   Dim iField As Integer
   'Dim ColInfo()
   Dim intDeduction As Integer 'Add by Amy 2014/11/13
   'Add by Amy 2017/08/17
   Dim RsQ As New ADODB.Recordset
   Dim intQ As Integer

On Error GoTo flgErr

   strRsQ = "SELECT '' aX201,'' aX205,'' A0102,'' sumcredit,'' sumdebit,'' amount,'' A0103,'' Srt From Dual Where Rownum<1"
   intQ = 1
   Set RsQ = ClsLawReadRstMsg(intQ, strRsQ)
   With rs4290
      .MoveFirst
      stLstGroup1 = "": stCurGroup1 = ""
    '  stCompName = ""
      Erase arrSubtot

      Set RsTemp = PUB_CreateRecordset(RsQ, , , , Me.Name)
      Do While Not .EOF
         '公司別 1234
         stCurGroup1 = "" & .Fields("aX201")
         If (stCurGroup1 <> stLstGroup1) And stLstGroup1 <> "" Then
            AddExpRows RsTemp
            Erase arrSubtot
         End If
         '複製原來資料
         RsTemp.AddNew
         For iField = 0 To RsQ.Fields.Count - 1 'Modify by Amy 2017/08/17
            RsTemp.Fields(iField) = .Fields(iField)
         Next
         
         '會計科目第一碼
         'Modify by Amy 2014/11/13 +減項(若為減項 *-1) ex:銷貨成本4301為銷貨的減項
         If GetClassDebitCredit(.Fields("Srt")) = "" & .Fields("A0103") Then
            intDeduction = 1
         Else
            intDeduction = -1
         End If
        
         Select Case .Fields("Srt")
            Case "4"
               arrSubtot(1) = arrSubtot(1) + Val("" & .Fields("AMOUNT")) * intDeduction
            Case "6"
               arrSubtot(2) = arrSubtot(2) + Val("" & .Fields("AMOUNT")) * intDeduction
            Case "7"
               If "" & .Fields("A0103") = "2" Then
                  RsTemp.Fields("Srt") = "71"
                  arrSubtot(3) = arrSubtot(3) + Val("" & .Fields("AMOUNT"))
               Else
                  RsTemp.Fields("Srt") = "72"
                  arrSubtot(4) = arrSubtot(4) + Val("" & .Fields("AMOUNT"))
               End If
         End Select
         'end 2014/11/13
         stLstGroup1 = stCurGroup1
         .MoveNext
      Loop
      AddExpRows RsTemp
      RsTemp.UPDATE
   End With
   
   Set Adodc1.Recordset = RsTemp.Clone
   Adodc1.Recordset.Sort = "aX201 Asc, Srt Asc"
   Set DataGrid1.DataSource = Adodc1

   SetGrid = True

flgErr:
   Set RsQ = Nothing
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   
End Function

Private Sub cmdQuery_Click()

   Screen.MousePointer = vbHourglass
   Set DataGrid1.DataSource = Nothing
   DataGrid1.Refresh
   CmdPrint.Enabled = False
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
   If Process = True Then
      If SetGrid = True Then CmdPrint.Enabled = True
   End If
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
   Screen.MousePointer = vbDefault
   
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
   If KeyCode <> vbKeyEscape Then
      Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
   End If
End Sub

Private Sub Form_Load()
    
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 9045
   Me.Height = 5700
   MoveFormToCenter Me
   Image1 = LoadPicture(strBackPicPath4)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   
   FormClear
   m_DefaultPrinter = Printer.DeviceName
   For Each m_Prn In Printers
      If m_Prn.DeviceName <> m_DefaultPrinter Then
         cmbPrinter.AddItem m_Prn.DeviceName
      End If
   Next
   cmbPrinter.ListIndex = 0
   'Add by Amy 2020/04/26 公司別下拉
   CboCmp.Clear
   CboCmp.AddItem "", 0
   Call Pub_SetCboCmp(CboCmp, True, False, False, , 1)
   'end 2020/05/26
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   For Each m_Prn In Printers
      If m_Prn.DeviceName = m_DefaultPrinter Then
         Set Printer = m_Prn
         Exit For
      End If
   Next
   Set rs4290 = Nothing
   Set Frmacc4290 = Nothing
End Sub


'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
   'Modify by Amy 2020/05/26 公司別改下拉
   'Text1 = ""
   CboCmp = ""
   'end 2020/05/26
   Text3 = "4"
   Text4 = "799999"
   MaskEdBox1.Mask = Mid(DFormat, 1, 6)
   MaskEdBox2.Mask = Mid(DFormat, 1, 6)
   
End Sub

Sub GetPleft()
   
    Erase PLeft
    Printer.Orientation = 1
    Printer.PaperSize = vbPRPSA4
    Printer.FontSize = 12
    iBytePix = Printer.TextWidth("會")
    iRowPix = 350
    
    PLeft(0) = 300
    PLeft(9) = Printer.ScaleWidth - 6 * iBytePix
    
    '會計科目
    PLeft(1) = PLeft(0)
    '科目名稱
    PLeft(2) = PLeft(1) + 7 * iBytePix
    '金額
    PLeft(3) = PLeft(9) - 6 * iBytePix
    PLeft(4) = PLeft(9)
    
End Sub


Private Sub PrintHead()
   Dim strTitle As String
   strTitle = "*** 過帳前綜合損益餘額表 ***"
   '起始列印列
   intY = 800 - intDefault
   iPage = iPage + 1
   With Printer
      '表頭
      .FontSize = 14
      .FontBold = True
      .CurrentX = (.ScaleWidth - .TextWidth(strTitle)) / 2 - 500
      .CurrentY = intY
      Printer.Print strTitle
      '跳列
      intY = intY + 500

      .FontSize = 12
      '條件
      .CurrentX = 4200 - Printer.TextWidth("公司別: ")
      .CurrentY = intY
      .FontBold = True
      Printer.Print "公司別: "
      .CurrentX = 4200
      .CurrentY = intY
      .FontBold = False
      Printer.Print stCurGroup1 & "  " & stCompName
      
      intY = intY + iRowPix
      
      .CurrentX = 4200 - Printer.TextWidth("年月: ")
      .CurrentY = intY
      .FontBold = True
      Printer.Print "年月: "
      .CurrentX = 4200
      .CurrentY = intY
      .FontBold = False
      Printer.Print MaskEdBox1 & " － " & MaskEdBox2
      
      intY = intY + iRowPix
      
      .CurrentX = 300
      .CurrentY = intY
      .FontBold = True
      Printer.Print "列印人員: "
      .CurrentX = 300 + Printer.TextWidth("列印人員: ")
      .CurrentY = intY
      .FontBold = False
      Printer.Print StaffQuery(strUserNum)

      
      .CurrentX = .ScaleWidth - 14 * iBytePix
      .CurrentY = intY
      .FontBold = True
      Printer.Print "列印日期: "
      .CurrentX = .ScaleWidth - 14 * iBytePix + Printer.TextWidth("列印日期: ")
      .CurrentY = intY
      .FontBold = False
      Printer.Print ChangeTStringToTDateString(strSrvDate(2))
      
      intY = intY + iRowPix
      
      .CurrentX = .ScaleWidth - 14 * iBytePix
      .CurrentY = intY
      .FontBold = True
      Printer.Print "頁次: "
      .CurrentX = .ScaleWidth - 14 * iBytePix + Printer.TextWidth("頁次: ")
      .CurrentY = intY
      .FontBold = False
      Printer.Print Right(Space(4) + Format(iPage), 4)
            
      intY = intY + iRowPix
      
      .FontBold = True
      .CurrentX = PLeft(1)
      .CurrentY = intY
      Printer.Print "會計科目"
      .CurrentX = PLeft(2)
      .CurrentY = intY
      Printer.Print "科目名稱"
      .CurrentX = PLeft(3)
      .CurrentY = intY
      Printer.Print "金額"

      intY = intY + iRowPix
      PrintLine 1
      .FontBold = False
   End With
   
End Sub
'劃線
Private Sub PrintLine(ByVal p_Type As Integer)
   Select Case p_Type
      Case 1
         Printer.DrawStyle = vbSolid
         Printer.Line (PLeft(0), intY)-(PLeft(9), intY)
         NewLine 150
      Case 2
         Printer.CurrentX = PLeft(3)
         Printer.CurrentY = intY
         Printer.Print ReportSum(4)
         NewLine
      Case 3
         NewLine
         Printer.CurrentX = PLeft(3)
         Printer.CurrentY = intY
         Printer.Print ReportSum(8)
   End Select
   
End Sub

Private Sub PrintData()
   Dim strDeduction As String 'Add by Amy 2014/11/13
   With Printer
      '會計科目
      .CurrentX = PLeft(1)
      .CurrentY = intY
      Printer.Print "" & Adodc1.Recordset("aX205")
      
      '科目名稱
      .CurrentX = PLeft(2)
      .CurrentY = intY
      Printer.Print "" & Adodc1.Recordset("A0102")
      
      '金額
      'Modify by Amy 2014/11/13 +減項(若為減項 *-1, 排除7科目)
      strDeduction = ""
      If Val(Adodc1.Recordset("Srt")) <= 6 Then
        If GetClassDebitCredit(Adodc1.Recordset("Srt")) <> "" & Adodc1.Recordset("A0103") Then
           If Left(LTrim(Adodc1.Recordset("AMOUNT")), 1) <> "-" Then strDeduction = "-"
        End If
      End If
      strExc(0) = strDeduction & Format(Adodc1.Recordset("AMOUNT"), FDollar)
      'end 2014/11/13
      .CurrentX = PLeft(4) - Printer.TextWidth(strExc(0))
      .CurrentY = intY
      Printer.Print strExc(0)
   End With
End Sub

Private Function NewLine(Optional ByVal iLine As Integer = 0) As Boolean
   If iLine = 0 Then iLine = iRowPix
   intY = intY + iLine
   If intY > Printer.ScaleHeight - 400 Then
      Printer.NewPage
      PrintHead
      NewLine = True
   End If
End Function

'Mark by Amy 2020/05/26 公司別改下拉
'Private Sub Text1_GotFocus()
'  TextInverse Text1
'End Sub
'
'Private Sub Text1_KeyPress(KeyAscii As Integer)
'   If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") Then
'      KeyAscii = 0
'   End If
'End Sub
'end 2020/05/26

'Add by Amy 2020/05/26
Private Sub CboCmp_GotFocus()
    TextInverse CboCmp
End Sub

Private Sub CboCmp_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub CboCmp_Validate(Cancel As Boolean)
    Dim strCmp As String
    
    If Trim(CboCmp) = MsgText(601) Then Exit Sub
    
    strCmp = CboCmp
    If InStr(strCmp, "　") > 0 Then
        strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
    End If

    If InStr(GetBookKeepCmp & 組合作帳公司 & ",", strCmp) = 0 Then
        MsgBox Label1 & MsgText(63), , MsgText(5)
        Cancel = True
        CboCmp.SetFocus
        Exit Sub
    ElseIf Len(Trim(CboCmp)) = 1 Then
        CboCmp = Trim(strCmp) & "　" & A0802Query(strCmp)
    End If
End Sub
'end 2020/05/26
