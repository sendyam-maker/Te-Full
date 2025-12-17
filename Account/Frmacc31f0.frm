VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Frmacc31f0 
   AutoRedraw      =   -1  'True
   Caption         =   "分所智慧局送件付款作業"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4785
   ScaleWidth      =   8880
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3285
      MaxLength       =   7
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   180
      Width           =   1185
   End
   Begin VB.CommandButton Command1 
      Caption         =   "查詢"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4590
      TabIndex        =   2
      Top             =   150
      Width           =   1050
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1125
      MaxLength       =   7
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   180
      Width           =   1185
   End
   Begin VB.ComboBox cboNo 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6975
      Style           =   2  '單純下拉式
      TabIndex        =   3
      Top             =   150
      Visible         =   0   'False
      Width           =   1620
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   3855
      Left            =   225
      TabIndex        =   4
      Top             =   570
      Width           =   8400
      _ExtentX        =   14817
      _ExtentY        =   6800
      _Version        =   393216
      BackColor       =   -2147483624
      Cols            =   6
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   6
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "開票日："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2430
      TabIndex        =   6
      Top             =   210
      Width           =   900
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "發文日："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   270
      TabIndex        =   5
      Top             =   210
      Width           =   900
   End
End
Attribute VB_Name = "Frmacc31f0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/10/20 Form2.0已修改 (無需修改)
'Memo By Sonia 2012/12/4 智權人員欄已修改
'Modify by Morgan 2011/6/29 畫面加開票日
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/3 日期欄已修改
'Create by Morgan 2008/9/23
Option Explicit

Public sNums As String '支票號碼字串,用逗點區隔
Public sAmts As String '支票金額字串,用逗點區隔
Public sPays As String '支付規費字串,用逗點區隔

Const iCol As Integer = 5 '票據行數
Dim iNowRow As Integer '本次點選列數
Dim iLstRow As Integer '前次點選列數
Dim sData() As String '支票號碼,支票金額, 設定金額

Private Sub cboNo_Click()
   grdDataList.TextMatrix(iNowRow, iCol) = cboNo.Text
End Sub

Private Sub Command1_Click()
   Dim bCancel  As Boolean
   Text1_Validate bCancel
   If bCancel Then Exit Sub
   Text2_Validate bCancel
   If bCancel = False Then
      SetDataListWidth
      cboNo.Visible = False
      SetGrid
   End If
End Sub

Private Sub Form_Load()
   'Modify by Amy 2023/10/11 原5000
   PUB_InitForm Me, 9000, 5250
   Text1.Text = TransDate(CompWorkDay(1, CompDate(2, -1, strSrvDate(1)), 1), 1)
   Text2.Text = Text1.Text 'Add by Morgan 2011/6/29
   Me.grdDataList.RowHeightMin = Me.cboNo.Height
   SetDataListWidth
   Me.grdDataList.Enabled = False
   SetGrid
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(133)
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   StatusClear
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc31f0 = Nothing
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)
   KeyEnter KeyCode
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(133)
End Sub

Private Sub SetDataListWidth()
   With grdDataList
      .Clear
      .Cols = 10
      .Rows = 2
      .row = 0:
      .col = 0: .ColWidth(.col) = 600: .Text = "時段"
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .col = 1: .ColWidth(.col) = 600: .Text = "所別"
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .col = 2: .ColWidth(.col) = 1800: .Text = "本所案號"
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .col = 3: .ColWidth(.col) = 1600: .Text = "案件性質"
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .col = 4: .ColWidth(.col) = 1200: .Text = "規費"
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .col = 5: .ColWidth(.col) = 1600: .Text = "票據號碼"
      cboNo.Width = .ColWidth(.col)
      .ColAlignment(.col) = flexAlignLeftCenter
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      For intI = 6 To .Cols
         .ColWidth(intI) = 0
      Next
      iNowRow = 0
      iLstRow = 0
   End With
End Sub

Private Sub SetGrid()
   
   Dim strVTblX As String
   
On Error GoTo ErrHnd
   
   strSql = "SELECT AL02 C01,AL03 C02,LPAD(A.CP01||'-'||A.CP02||'-'||A.CP03||'-'||A.CP04,15,' ') C03,RPAD(CPM03,12,' ') C04,A.CP84 C05,'' C06,ALD09 C07" & _
      ",substr(cu04,1,10)||'/'||CPM03 C08,A.CP01||A.CP02||A.CP03||A.CP04 C09" & _
      " From APPLIST,APPLISTDETAIL,CASEPROGRESS A,CASEPROGRESS B, NEXTPROGRESS,CASEPROPERTYMAP,PATENT,CUSTOMER" & _
      " WHERE AL01=" & DBDATE(Text1) & " AND AL06 IS NULL AND AL04<>'7' AND AL04<>'8' AND AL02<'5'" & _
      " AND ALD01(+)=AL01 AND ALD02(+)=AL02 AND ALD03(+)=AL03 AND ALD04(+)=AL04 AND ALD10 IS NULL" & _
      " AND A.CP01(+)=ALD05 AND A.CP02(+)=ALD06 AND A.CP03(+)=ALD07 AND A.CP04(+)=ALD08 " & _
      " AND A.CP27(+)=ALD01 AND A.CP84>0 AND (A.CP82<DECODE(ALD03,'0',AL05) OR A.CP82>=DECODE(ALD03,'1',AL05))" & _
      " AND B.CP09(+)=DECODE(A.CP10,'404',A.CP43,NULL) AND NP01(+)=DECODE(A.CP10,'404',A.CP43,NULL)" & _
      " AND CPM01=A.CP01 AND CPM02=DECODE(A.CP10,'404',NVL(NP07,B.CP10),A.CP10)" & _
      " and pa01(+)=A.cp01 and pa02(+)=A.cp02 and pa03(+)=A.cp03 and pa04(+)=A.cp04 AND PA01 IS NOT NULL" & _
      " AND cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9,1)" & _
      " ORDER BY 1,2,C07 DESC,C03,C05 DESC,C04"

   CheckOC
   With adoRecordset
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount > 0 Then
         If SetRow = True Then
            If SetData = True Then MatchData
            '2012/8/9 modify by sonia
            'tool15_enabled
            tool17_enabled
            Me.grdDataList.Enabled = True
         Else
            MsgBox "無待作業資料！", vbInformation, "注意"
         End If
      Else
         MsgBox "無待作業資料！", vbInformation, "注意"
      End If
   End With
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical

End Sub

Private Function SetRow() As Boolean
   
   Dim iCol As Integer, iRow As Integer
   Dim stDept As String '部門
   Dim lngAmount As String
   Dim strTmp As String
   
   With grdDataList
            
      '上午
      iRow = 0
      adoRecordset.MoveFirst
      While Not adoRecordset.EOF
         iRow = iRow + 1
         .Rows = iRow + 1
         '時段
         iCol = 0
         strTmp = "" & adoRecordset.Fields("C02")
         Select Case strTmp
            Case "0": .TextMatrix(iRow, iCol) = "上午"
            Case "1": .TextMatrix(iRow, iCol) = "下午"
         End Select
         '所別
         iCol = 1
         strTmp = "" & adoRecordset.Fields("C01")
         Select Case strTmp
            Case "2": .TextMatrix(iRow, iCol) = "中所"
            Case "3": .TextMatrix(iRow, iCol) = "南所"
            Case "4": .TextMatrix(iRow, iCol) = "高所"
         End Select
         iCol = 2
         .TextMatrix(iRow, iCol) = "" & adoRecordset.Fields("C03")
         iCol = 3
         .TextMatrix(iRow, iCol) = "" & adoRecordset.Fields("C04")
         '規費
         lngAmount = Val("" & adoRecordset.Fields("C05"))
         iCol = 4: .TextMatrix(iRow, iCol) = Format(lngAmount, DDollar)
         .TextMatrix(iRow, 6) = "" & adoRecordset.Fields("C07")
         .TextMatrix(iRow, 7) = "" & adoRecordset.Fields("C08")
         .TextMatrix(iRow, 8) = "" & adoRecordset.Fields("C09")
         adoRecordset.MoveNext
      Wend
      
      If .TextMatrix(1, 0) <> "" Then
         .row = 0
         SetRow = True
      End If
      
   End With
      

End Function

Private Sub GrdDataList_Click()
   cboNo.Visible = False
   If grdDataList.MouseRow > 0 Then
      iLstRow = iNowRow: iNowRow = grdDataList.MouseRow
      SetBox
   End If
   
End Sub

Private Sub SetBox(Optional iDirect As Integer = 0)
   
   Dim i As Integer
   Dim lngLeft As Long, lngTop As Long
   
   iNowRow = iNowRow + iDirect
   
   '若超出可選範圍則不作用
   If iNowRow = 0 Or iNowRow = grdDataList.Rows Then
      iNowRow = iLstRow
      
   Else
   
      '還原
      If iLstRow > 0 Then
         grdDataList.row = iLstRow
         For i = 0 To 5
            grdDataList.col = i
            grdDataList.CellForeColor = grdDataList.ForeColor
            grdDataList.CellBackColor = grdDataList.BackColor
         Next i
      End If
      
      '反白
      grdDataList.row = iNowRow
      For i = 0 To 5
         grdDataList.col = i
         grdDataList.CellForeColor = grdDataList.BackColor
         grdDataList.CellBackColor = grdDataList.BackColorSel
      Next i
      
      iLstRow = iNowRow
      
      If grdDataList.TextMatrix(iNowRow, iCol) <> "" Then
         cboNo.Text = grdDataList.TextMatrix(iNowRow, iCol)
      Else
         cboNo.ListIndex = -1
      End If
      cboNo.Visible = True
      cboNo.SetFocus
      
   End If
   
   lngLeft = grdDataList.Left + 20
   lngTop = grdDataList.Top + grdDataList.RowHeight(0) + 20
   For i = 0 To 4
      lngLeft = lngLeft + grdDataList.ColWidth(i)
   Next
   
   For i = grdDataList.TopRow To iNowRow - 1
      lngTop = lngTop + grdDataList.RowHeight(i)
   Next
   
   cboNo.Left = lngLeft: cboNo.Top = lngTop
   
End Sub

'檢查票據金額與案件規費
Public Function CheckData() As Boolean
   Dim i As Integer, j As Integer, k As Integer
   Dim bCancel As Boolean
   
   '檢查票據號碼
   For i = 1 To Me.grdDataList.Rows - 1
      If Me.grdDataList.TextMatrix(i, 5) = "" Then
         bCancel = True
         MsgBox "票據號碼不可空白", vbCritical, "注意"
         iNowRow = i
         SetBox
         Exit For
      End If
   Next i
   If bCancel = True Then
      Exit Function
   End If
   
   For j = LBound(sData, 1) To UBound(sData, 1)
      sData(j, 2) = "0"
      For i = 1 To Me.grdDataList.Rows - 1
         If Me.grdDataList.TextMatrix(i, 5) = cboNo.List(j - 1) Then
            sData(j, 2) = Format(Val(sData(j, 2)) + Val(Format(Me.grdDataList.TextMatrix(i, 4))))
            k = i
         End If
      Next i
      If Val(sData(j, 2)) > 0 And Val(sData(j, 1)) > Val(sData(j, 2)) Then
         bCancel = True
         MsgBox "票據金額檢查錯誤，票據號碼[" & sData(j, 0) & "]的金額不可大於設定案件的規費總額[$" & Format(sData(j, 2), DDollar) & "]！", vbCritical
         iNowRow = k
         SetBox
         Exit For
      End If
   Next j
   If bCancel = True Then
      Exit Function
   End If
   
   sNums = "" & sData(1, 0)
   sAmts = "" & sData(1, 1)
   sPays = "" & sData(1, 2)
   For j = 2 To UBound(sData, 1)
      sNums = sNums & "," & sData(j, 0)
      sAmts = sAmts & "," & sData(j, 1)
      sPays = sPays & "," & sData(j, 2)
   Next
   
   CheckData = True
End Function

Private Sub MatchData()
   Dim i As Integer, j As Integer
   
   For i = 1 To Me.grdDataList.Rows - 1
      For j = LBound(sData, 1) To UBound(sData, 1)
         If Format(sData(j, 1), DDollar) = Me.grdDataList.TextMatrix(i, 4) And sData(j, 3) = "" Then
            Me.grdDataList.TextMatrix(i, 6) = sData(j, 0)
            sData(j, 3) = "1"
            Exit For
         End If
      Next j
   Next i
   
End Sub

'讀取當天開給智慧局的支票資料
Private Function SetData() As Boolean
   
On Error GoTo ErrHnd
   
   cboNo.Clear
   Erase sData
   
   'Modify by Amy 2020/08/05 因a0e07改為 PKey,故需加a1p04
   strSql = "select a0e02,a0e11 from acc0e0 where a0e04='P' and a0e05='2' and a0e06='V0001' and a0e13=" & Text2 & _
      " and not exists(select * from acc1p0 where a1p09=a0e02 and a1p15=a0e06 and a1p04=a0e02||a0e01||a0e07||'2' )" & _
      " ORDER BY 1"
   
   CheckOC
   With adoRecordset
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount > 0 Then
         ReDim sData(1 To .RecordCount, 0 To 3)
         While Not .EOF
            cboNo.AddItem .Fields(0)
            sData(.AbsolutePosition, 0) = .Fields(0)
            sData(.AbsolutePosition, 1) = .Fields(1)
            .MoveNext
         Wend
         SetData = True
      End If
   End With
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   
End Function

Private Sub grdDataList_Scroll()
   grdDataList.TopRow = grdDataList.TopRow
   
   If grdDataList.TopRow > iNowRow Or grdDataList.TopRow < iNowRow - 9 Then
      cboNo.Visible = False
   Else
      SetBox
   End If
   
End Sub

Private Sub Text1_Change()
   tool3_enabled
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   If Text1 = "" Then
      MsgBox "發文日不可空白！"
      Cancel = True
   ElseIf ChkDate(Text1) = False Then
      Cancel = True
   End If
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
   If Text2 = "" Then
      MsgBox "開票日不可空白！"
      Cancel = True
   ElseIf ChkDate(Text2) = False Then
      Cancel = True
   End If
End Sub

'Amy 2014/11/05 由aacc_sav 搬回 'Add by Morgan 2008/9/23
Public Sub Frmacc31f0_Save()
   
   
   With Frmacc31f0
   
      Screen.MousePointer = vbHourglass
      
      Dim iRowNum As Integer, idx As Integer
      Dim Acc1p0(1 To 18) As String
      Dim iA1p03 As Integer '序號
      Dim idxNow As Integer '現在執行的支票號索引值
      Dim stCon As String 'SQL條件語法
      Dim sNum '支票號碼
      Dim sAmt '支票金額
      Dim sPay '支付規費
      
On Error GoTo Checking
   
      If .CheckData = True Then
         sNum = Split(.sNums, ",")
         sAmt = Split(.sAmts, ",")
         sPay = Split(.sPays, ",")
         
On Error GoTo Saving

         cnnConnection.BeginTrans
         
         '找出有選到的支票號
         For idxNow = LBound(sNum) To UBound(sNum)
            If Val(sPay(idxNow)) > 0 Then
               '借方
               iA1p03 = 0
               For iRowNum = 1 To .grdDataList.Rows - 1
                  If .grdDataList.TextMatrix(iRowNum, 5) = sNum(idxNow) Then
                     iA1p03 = iA1p03 + 1
                     Acc1p0(1) = "'1'"
                     Acc1p0(2) = "'L'"
                     Acc1p0(3) = "'" & Format(iA1p03, "000") & "'"
                     'Modify by Amy 2020/08/05 開票銀行改抓變數,並加帳號 原:011010075
                     Acc1p0(4) = "'" & .grdDataList.TextMatrix(iRowNum, 5) & 智慧局送件開票銀行 & 智慧局送件開票帳號 & "2'"
                     Acc1p0(5) = "'220102'"
                     Acc1p0(6) = "'TOT'"
                     Acc1p0(7) = Format(.grdDataList.TextMatrix(iRowNum, 4))
                     Acc1p0(8) = "0"
                     Acc1p0(9) = "'" & .grdDataList.TextMatrix(iRowNum, 5) & "'"
                     'Modify by Amy 2020/08/05 開票銀行改抓變數 原:011010075
                     Acc1p0(10) = "'" & 智慧局送件開票銀行 & "'"
                     '2010/6/22 MODIFY BY SONIA 用新的甲存帳號
                     'ACC1P0(11) = "'0149950'"
                     'Modify by Amy 2020/08/05 開票帳號改抓變數 原:1756650
                     Acc1p0(11) = "'" & 智慧局送件開票帳號 & "'"  'Modify by Amy 2020/07/24 改回1756650 'modify by sonia 2020/6/19 改帳號原為0149951(1756650)
                     Acc1p0(12) = .Text1
                     'Modify by Morgan 2010/1/22 摘要加本所案號
                     'ACC1P0(14) = "'" & .grdDataList.TextMatrix(iRowNum, 7) & "'"
                     Acc1p0(14) = "'" & .grdDataList.TextMatrix(iRowNum, 8) & "/" & .grdDataList.TextMatrix(iRowNum, 7) & "'"
                     Acc1p0(15) = "'V0001'"
                     Acc1p0(17) = "'" & .grdDataList.TextMatrix(iRowNum, 8) & "'"
                     Acc1p0(18) = .Text1
                     strSql = "insert into acc1p0(a1p01,a1p02,a1p03,a1p04,a1p05,a1p06,a1p07,a1p08,a1p09,a1p10,a1p11,a1p12,a1p14,a1p15,a1p17,a1p18 )" & _
                        "VALUES(" & Acc1p0(1) & "," & Acc1p0(2) & "," & Acc1p0(3) & "," & Acc1p0(4) & "," & Acc1p0(5) & "," & Acc1p0(6) & _
                        "," & Acc1p0(7) & "," & Acc1p0(8) & "," & Acc1p0(9) & "," & Acc1p0(10) & "," & Acc1p0(11) & "," & Acc1p0(12) & _
                        "," & Acc1p0(14) & "," & Acc1p0(15) & "," & Acc1p0(17) & "," & Acc1p0(18) & ")"
                     cnnConnection.Execute strSql, intI
                  End If
               Next
               
               '貸方
               iA1p03 = iA1p03 + 1
               Acc1p0(3) = "'" & Format(iA1p03, "000") & "'"
               Acc1p0(5) = "'2111'"
               Acc1p0(7) = "0"
               Acc1p0(8) = Val(sAmt(idxNow))
               Acc1p0(14) = "'0" & .Text1 & "/" & sNum(idxNow) & "/" & "經濟部智慧財產局'"
               Acc1p0(17) = "NULL"
               strSql = "insert into acc1p0(a1p01,a1p02,a1p03,a1p04,a1p05,a1p06,a1p07,a1p08,a1p09,a1p10,a1p11,a1p12,a1p14,a1p15,a1p17,a1p18 )" & _
                  "VALUES(" & Acc1p0(1) & "," & Acc1p0(2) & "," & Acc1p0(3) & "," & Acc1p0(4) & "," & Acc1p0(5) & "," & Acc1p0(6) & _
                  "," & Acc1p0(7) & "," & Acc1p0(8) & "," & Acc1p0(9) & "," & Acc1p0(10) & "," & Acc1p0(11) & "," & Acc1p0(12) & _
                  "," & Acc1p0(14) & "," & Acc1p0(15) & "," & Acc1p0(17) & "," & Acc1p0(18) & ")"
                  
               cnnConnection.Execute strSql, intI
               '票據金額不足部分用現金補
               If Val(sPay(idxNow)) > Val(sAmt(idxNow)) Then
                  iA1p03 = iA1p03 + 1
                  Acc1p0(3) = "'" & Format(iA1p03, "000") & "'"
                  Acc1p0(5) = "'1911'"
                  Acc1p0(7) = "0"
                  Acc1p0(8) = Val(sPay(idxNow)) - Val(sAmt(idxNow))
                  Acc1p0(14) = "'智慧局送件'"
                  Acc1p0(17) = "NULL"
                  strSql = "insert into acc1p0(a1p01,a1p02,a1p03,a1p04,a1p05,a1p06,a1p07,a1p08,a1p09,a1p10,a1p11,a1p12,a1p14,a1p15,a1p17,a1p18 )" & _
                     "VALUES(" & Acc1p0(1) & "," & Acc1p0(2) & "," & Acc1p0(3) & "," & Acc1p0(4) & "," & Acc1p0(5) & "," & Acc1p0(6) & _
                     "," & Acc1p0(7) & "," & Acc1p0(8) & "," & Acc1p0(9) & "," & Acc1p0(10) & "," & Acc1p0(11) & "," & Acc1p0(12) & _
                     "," & Acc1p0(14) & "," & Acc1p0(15) & "," & Acc1p0(17) & "," & Acc1p0(18) & ")"
                     
                  cnnConnection.Execute strSql, intI
               End If
            End If
         Next
         
         '更新APPLIST
         For iRowNum = 1 To .grdDataList.Rows - 1
            '時段
            stCon = ""
            Select Case .grdDataList.TextMatrix(iRowNum, 0)
               Case "上午": stCon = stCon & " and AL03='0'"
               Case "下午": stCon = stCon & " and AL03='1'"
            End Select
            
            '部門
            Select Case .grdDataList.TextMatrix(iRowNum, 1)
               Case "中所": stCon = stCon & " and AL02='2'"
               Case "南所": stCon = stCon & " and AL02='3'"
               Case "高所": stCon = stCon & " and AL02='4'"
            End Select
            
            stCon = stCon & " and AL04='1'"
            
            strSql = "UPDATE APPLIST SET AL06='" & .grdDataList.TextMatrix(iRowNum, 5) & "' WHERE AL01=" & DBDATE(.Text1) & " AND AL06 IS NULL" & stCon
            cnnConnection.Execute strSql, intI
         Next
               
         cnnConnection.CommitTrans
         MsgBox "轉傳票分錄資料產生完成！"
         KeyEnter vbKeyEscape
         
Saving:
      
         If Err.Number <> 0 Then
            cnnConnection.RollbackTrans
            MsgBox Err.Description, vbCritical
            Err.Clear
         End If
      End If
      
Checking:

      If Err.Number <> 0 Then
         MsgBox Err.Description, , MsgText(5)
         Err.Clear
      End If
      
      Screen.MousePointer = vbDefault
      
   End With
End Sub

