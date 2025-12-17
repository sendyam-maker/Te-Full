VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Frmacc31g0 
   AutoRedraw      =   -1  'True
   Caption         =   "智慧局電子送件付款作業"
   ClientHeight    =   4704
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4704
   ScaleWidth      =   8880
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3150
      Style           =   2  '單純下拉式
      TabIndex        =   1
      Top             =   143
      Width           =   1410
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5535
      MaxLength       =   7
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   150
      Width           =   1185
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1125
      MaxLength       =   7
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   150
      Width           =   1185
   End
   Begin VB.CommandButton Command1 
      Caption         =   "查詢"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6750
      TabIndex        =   3
      Top             =   143
      Width           =   675
   End
   Begin VB.ComboBox cboNo 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7020
      Style           =   2  '單純下拉式
      TabIndex        =   5
      Top             =   4230
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.CommandButton cmdDetail 
      Caption         =   "明細"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7650
      TabIndex        =   4
      Top             =   143
      Width           =   975
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   3615
      Left            =   225
      TabIndex        =   6
      Top             =   570
      Width           =   8400
      _ExtentX        =   14817
      _ExtentY        =   6371
      _Version        =   393216
      BackColor       =   -2147483624
      Cols            =   7
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   7
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "部門："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2475
      TabIndex        =   10
      Top             =   203
      Width           =   675
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "開票日："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4680
      TabIndex        =   9
      Top             =   203
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "請確認收到專業部清單紙本再按存檔!!!"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   225
      TabIndex        =   8
      Top             =   4230
      Width           =   5640
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "發文日："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   270
      TabIndex        =   7
      Top             =   203
      Width           =   900
   End
End
Attribute VB_Name = "Frmacc31g0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/10/20 Form2.0已修改 grdDataList
'Memo By Sonia 2012/12/4 智權人員欄已修改
'Modify by Morgan 2011/6/29 畫面加開票日
'Create by Morgan 2011/6/2 參考 frmacc31e0
Option Explicit

Public iNowRow As Integer '本次點選列數
Dim iLstRow As Integer '前次點選列數
Dim sData() As String '支票號碼,支票金額, 設定金額
Const iCol As Integer = 6 '金額行數
Public adoacc1p0 As New ADODB.Recordset '存檔用
Public adoquery As New ADODB.Recordset '查詢用
Public sNums As String '支票號碼字串,用逗點區隔
Public sAmts As String '支票金額字串,用逗點區隔

Private Sub cboNo_Click()
   grdDataList.TextMatrix(iNowRow, iCol) = cboNo.Text
End Sub

Private Sub cmdDetail_Click()
   Screen.MousePointer = vbHourglass
   If iNowRow > 0 Then
      Me.Enabled = False
      tool3_enabled
      Frmacc31g1.Show
   End If
   Screen.MousePointer = vbDefault
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

   PUB_InitForm Me, 9000, 5150 'Modify by Amy 2023/08/17 原:5000
   Text1.Text = strSrvDate(2)
   Text2.Text = Text1.Text
   Me.grdDataList.RowHeightMin = Me.cboNo.Height
   SetDataListWidth
   Me.grdDataList.Enabled = False
   SetGrid
   AddDeptCombo 'Add by Morgan 2011/8/8
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(133)
   
End Sub

Private Sub AddDeptCombo()
   Combo1.Clear
   Combo1.AddItem "ALL", 0
   Combo1.AddItem "P", 1
   Combo1.AddItem "FCP", 2
   Combo1.AddItem "T", 3
   Combo1.AddItem "FCT", 4
   Combo1.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   StatusClear
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc31g0 = Nothing
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
      .Rows = 2
      .row = 0:
      .col = 0: .ColWidth(.col) = 600: .Text = "時段"
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .col = 1: .ColWidth(.col) = 900: .Text = "部門別"
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .col = 2: .ColWidth(.col) = 600: .Text = "出名"
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .col = 3: .ColWidth(.col) = 1100: .Text = "清單別"
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .col = 4: .ColWidth(.col) = 1600: .Text = "金額"
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .col = 5: .ColWidth(.col) = 1200: .Text = "付款方式"
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .col = 6: .ColWidth(.col) = 1600: .Text = "票據號碼"
      cboNo.Width = .ColWidth(.col)
      .ColAlignment(.col) = flexAlignLeftCenter
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      
   End With
End Sub

Private Sub SetGrid()
   Dim stCon As String
   
   Select Case Combo1.ListIndex
      Case 1
         stCon = " and AL02='P1'"
      Case 2
         stCon = " and AL02='F2'"
      Case 3
         stCon = " and AL02='P2'"
      Case 4
         stCon = " and AL02='F1'"
   End Select
   
On Error GoTo ErrHnd
   'Modified by Morgan 2011/6/16 內商因作業方式不同改寫applistDETAILe
   'Modified by Morgan 2014/8/11 外商改同內商模式
   strSql = "SELECT AL02,AL03,AL04,ALD10,SUM(ALD09) AMT,3 Ord" & _
      " From appliste, applistDETAILe" & _
      " WHERE AL01=" & DBDATE(Text1) & " AND AL06 IS NULL AND AL04='8'" & stCon & _
      " AND ALD01(+)=AL01 AND ALD02(+)=AL02 AND ALD03(+)=AL03 and ald04(+)=al04 AND ALD10 IS NULL" & _
      " GROUP BY AL02,AL03,AL04,ALD10 HAVING SUM(ALD09)>0 "
      
   strSql = strSql & " union all SELECT AL02,AL03,AL04,ALD10,SUM(ALD09) AMT,DECODE(AL02,'P1',1,'F2',2,'P2',3,'F1',4) Ord" & _
      " From appliste, applisteDETAIL" & _
      " WHERE AL01=" & DBDATE(Text1) & " AND AL06 IS NULL AND AL04='8'" & stCon & _
      " AND ALD01(+)=AL01 AND ALD02(+)=AL02 AND ALD03(+)=AL03 AND ALD04(+)=AL04 AND ALD10 IS NULL" & stCon & _
      " GROUP BY AL02,AL03,AL04,ALD10 HAVING SUM(ALD09)>0 " & _
      " ORDER BY AL03,Ord,ALD10,AL04"
      
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
   
   With grdDataList
            
      '上午
      iRow = 0
      adoRecordset.MoveFirst
      While Not adoRecordset.EOF
         iRow = iRow + 1
         .Rows = iRow + 1
         '時段
         iCol = 0
         Select Case adoRecordset.Fields("AL03")
            Case "0": .TextMatrix(iRow, iCol) = "上午"
            Case "1": .TextMatrix(iRow, iCol) = "下午"
         End Select
         '部門
         iCol = 1
         Select Case adoRecordset.Fields("AL02")
            Case "P1": .TextMatrix(iRow, iCol) = "P"
            Case "F2": .TextMatrix(iRow, iCol) = "FCP"
            Case "P2": .TextMatrix(iRow, iCol) = "T"
            Case "F1": .TextMatrix(iRow, iCol) = "FCT"
         End Select
         
         '出名否
         iCol = 2
         Select Case "" & adoRecordset.Fields("ALD10")
            Case "": .TextMatrix(iRow, iCol) = "Y"
            Case "N": .TextMatrix(iRow, iCol) = "N"
         End Select
         
         '類別
         iCol = 3
         Select Case adoRecordset.Fields("AL04")
            Case "1": .TextMatrix(iRow, iCol) = "新案"
            Case "2": .TextMatrix(iRow, iCol) = "一般"
            Case "3": .TextMatrix(iRow, iCol) = "快速"
            Case "4": .TextMatrix(iRow, iCol) = "發明"
            Case "5": .TextMatrix(iRow, iCol) = "新型"
            Case "6": .TextMatrix(iRow, iCol) = "設計"
            Case "7": .TextMatrix(iRow, iCol) = "非智慧局"
            Case "8": .TextMatrix(iRow, iCol) = "電子送件"
         End Select
         
         '金額
         lngAmount = Val("" & adoRecordset.Fields("AMT"))
         iCol = 4: .TextMatrix(iRow, iCol) = Format(lngAmount, DDollar)
         '開票(固定)
         iCol = 5: .TextMatrix(iRow, iCol) = "開票"
      
         adoRecordset.MoveNext
      Wend
      
      If .TextMatrix(1, 0) <> "" Then
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
         For i = 0 To grdDataList.Cols - 1
            grdDataList.col = i
            grdDataList.CellForeColor = grdDataList.ForeColor
            grdDataList.CellBackColor = grdDataList.BackColor
         Next i
      End If
      
      '反白
      grdDataList.row = iNowRow
      For i = 0 To grdDataList.Cols - 1
         grdDataList.col = i
         grdDataList.CellForeColor = grdDataList.BackColor
         grdDataList.CellBackColor = grdDataList.BackColorSel
      Next i
      
      iLstRow = iNowRow
      
      If Me.grdDataList.TextMatrix(iNowRow, 5) = "開票" Then
         If grdDataList.TextMatrix(iNowRow, iCol) <> "" Then
            cboNo.Text = grdDataList.TextMatrix(iNowRow, iCol)
         Else
            cboNo.ListIndex = -1
         End If
         cboNo.Visible = True
         cboNo.SetFocus
      End If
      
   End If
   
   lngLeft = grdDataList.Left + 20
   lngTop = grdDataList.Top + grdDataList.RowHeight(0) + 20
   For i = 0 To 5
      lngLeft = lngLeft + grdDataList.ColWidth(i)
   Next
   
   For i = grdDataList.TopRow To iNowRow - 1
      lngTop = lngTop + grdDataList.RowHeight(i)
   Next
   
   cboNo.Left = lngLeft: cboNo.Top = lngTop
   
End Sub

Public Function CheckData() As Boolean

   Dim i As Integer, j As Integer, k As Integer
   
   '檢查票據號碼
   For i = 1 To Me.grdDataList.Rows - 1
      If Me.grdDataList.TextMatrix(i, 5) = "開票" Then
         If Me.grdDataList.TextMatrix(i, 6) = "" Then
            MsgBox "票據號碼不可空白", vbCritical, "注意"
            iNowRow = i
            SetBox
            Exit For
         End If
      End If
   Next i
   If i = Me.grdDataList.Rows Then
      For j = LBound(sData, 1) To UBound(sData, 1)
         sData(j, 2) = "0"
         For i = 1 To Me.grdDataList.Rows - 1
            If Me.grdDataList.TextMatrix(i, 6) = cboNo.List(j - 1) Then
               sData(j, 2) = Format(Val(sData(j, 2)) + Val(Replace(Me.grdDataList.TextMatrix(i, 4), ",", "")))
               k = i
            End If
         Next i
         If sData(j, 2) > 0 And sData(j, 1) <> sData(j, 2) Then
            MsgBox "票據金額檢查錯誤，票據號碼[" & sData(j, 0) & "]的金額應為[$" & Format(sData(j, 1), DDollar) & "]！", vbCritical
            iNowRow = k
            SetBox
            Exit For
         End If
      Next j
      If j = UBound(sData, 1) + 1 Then CheckData = True
   End If
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

   'Modify by Amy 2020/08/05 因a0e07改為 PKey,故需加a1p04
   strSql = "select a0e02,a0e11 from acc0e0 where a0e04='P' and a0e05='2' and a0e06='V0001' and a0e13=" & Text2 & _
      "  and not exists(select * from acc1p0 where a1p09=a0e02 and a1p15=a0e06 and a1p04=a0e02||a0e01||a0e07||'2' )" & _
      " ORDER BY 1"
   
   CheckOC
   With adoRecordset
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount > 0 Then
         cboNo.Clear
         Erase sData
         ReDim sData(1 To .RecordCount, 0 To 3)
         While Not .EOF
            cboNo.AddItem .Fields(0)
            sData(.AbsolutePosition, 0) = .Fields(0)
            sData(.AbsolutePosition, 1) = .Fields(1)
            sNums = sNums & IIf(sNums = "", "", ",") & .Fields(0)
            sAmts = sAmts & IIf(sAmts = "", "", ",") & .Fields(1)
            .MoveNext
         Wend
         SetData = True
      End If
   End With
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   
End Function
'取得明細語法
'若性質為延期303時:A類收文用CP43抓NP07；B類收文用CP43抓相關收文號的CP10
Public Function GetSql(Optional ByVal iRowNo As Integer = 0, Optional ByVal sType As String = "1") As String

   Dim strCon As String, stDept As String, strCont As String
   
   Dim strVTblX As String, strVTblY As String
   Dim strConX As String, strConY As String, stTemp As String

   If iRowNo = 0 Then iRowNo = iNowRow
   
   strCon = ""
   strCont = ""
   With grdDataList
   
      '上下午
      stTemp = .TextMatrix(iRowNo, 0)
      If stTemp = "上午" Then
         strCon = strCon & " AND ALD03='0' and CP82<AL05"
         strCont = strCont & " AND ALD03='0'"
      Else
         strCon = strCon & " AND ALD03='1' and CP82>=AL05"
         strCont = strCont & " AND ALD03='1'"
      End If
      
      '部門
      stDept = .TextMatrix(iRowNo, 1)
      Select Case stDept
         Case "P": strCon = strCon & " and AL02='P1'"
         Case "FCP": strCon = strCon & " and AL02='F2'"
         Case "T": strCon = strCon & " and AL02='P2'": strCont = strCont & " and AL02='P2'"
         Case "FCT": strCon = strCon & " and AL02='F1'": strCont = strCont & " and AL02='F1'"
      End Select
            
      '出名否
      stTemp = .TextMatrix(iRowNo, 2)
      If stTemp = "Y" Then
         strCon = strCon & " and ALD10 IS NULL"
         strCont = strCont & " and ALD10 IS NULL"
      Else
         strCon = strCon & " and ALD10='N'"
         strCont = strCont & " and ALD10='N'"
      End If
      
      '清單別
      stTemp = .TextMatrix(iRowNo, 3)
      Select Case stTemp
'         Case "新案": strCon = strCon & " and AL04='1'"
'         Case "一般": strCon = strCon & " and AL04='2'"
'         Case "快速": strCon = strCon & " and AL04='3'"
'         Case "發明": strCon = strCon & " and AL04='4'"
'         Case "新型": strCon = strCon & " and AL04='5'"
'         Case "設計": strCon = strCon & " and AL04='6'"
'         Case "非智慧局": strCon = strCon & " and AL04='7'"
         Case "電子送件": strCon = strCon & " and AL04='8'": strCont = strCont & " and AL04='8'"
      End Select

   End With
   
   If stDept = "P" Or stDept = "FCP" Then
         '虛擬表格語法
         strVTblX = " SELECT CP01,CP02,CP03,CP04,CP10,CP22,CP43,CP82,CP84,PA26,ALD09" & _
            " From appliste,applisteDETAIL,CASEPROGRESS, patent" & _
            " WHERE AL01=" & DBDATE(Text1) & strCon & _
            " AND ALD01(+)=AL01 AND ALD02(+)=AL02 AND ALD03(+)=AL03 AND ALD04(+)=AL04" & _
            " AND CP01(+)=ALD05 AND CP02(+)=ALD06 AND CP03(+)=ALD07 AND CP04(+)=ALD08" & _
            " AND CP27(+)=ALD01 AND CP84>0 AND CP118='Y'" & _
            " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 AND PA01 IS NOT NULL"
         If stDept = "FCP" Then
            strVTblX = strVTblX & " Union All" & _
               " SELECT CP01,CP02,CP03,CP04,CP10,CP22,CP43,CP82,CP84, SP08 PA26,ALD09" & _
               " From appliste,applisteDETAIL,CASEPROGRESS, servicepractice" & _
               " WHERE AL01=" & DBDATE(Text1) & strCon & _
               " AND ALD01(+)=AL01 AND ALD02(+)=AL02 AND ALD03(+)=AL03 AND ALD04(+)=AL04" & _
               " AND CP01(+)=ALD05 AND CP02(+)=ALD06 AND CP03(+)=ALD07 AND CP04(+)=ALD08" & _
               " AND CP27(+)=ALD01 AND CP84>0 AND CP118='Y'" & _
               " and sp01(+)=cp01 and sp02(+)=cp02 and sp03(+)=cp03 and sp04(+)=cp04 AND SP01 IS NOT NULL"
         End If
         
         'frmacc31g1 用
         If sType = "1" Then
            GetSql = "SELECT LPAD(X.CP01||'-'||X.CP02||'-'||X.CP03||'-'||X.CP04,15,' ') C01, X.CP84 C02, RPAD(CPM03,12,' ') C05, RPAD(NVL(cu04,' '),20,' ') C06" & _
               " FROM (" & strVTblX & ") X, customer" & _
               ",CASEPROGRESS B, NEXTPROGRESS,CASEPROPERTYMAP" & _
               " WHERE " & _
               " cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9,1)" & _
               " AND B.CP09(+)=DECODE(X.CP10,'404',X.CP43,NULL) AND NP01(+)=DECODE(X.CP10,'404',X.CP43,NULL)" & _
               " AND CPM01=X.CP01 AND CPM02=DECODE(X.CP10,'404',NVL(NP07,B.CP10),X.CP10)" & _
               " order by ALD09 DESC, 1, 2 DESC, 3"
               
         'acc_save 用
         Else
            'P,T的摘要加本所案號
            GetSql = "SELECT decode(X.cp01,'T','220101','P','220102','FCT','220103','220104') a1p05" & _
               ", X.cp84 a1p07,X.CP01||X.CP02||X.CP03||X.CP04||decode(instr('T,P',X.CP01),0,'','/'||substr(cu04,1,10))||'/'||CPM03 a1p14" & _
               ", X.CP01||X.CP02||X.CP03||X.CP04 a1p17" & _
               " FROM (" & strVTblX & ") X, customer" & _
               ",CASEPROGRESS B, NEXTPROGRESS,CASEPROPERTYMAP" & _
               " WHERE cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9,1)" & _
               " AND B.CP09(+)=DECODE(X.CP10,'404',X.CP43,NULL) AND NP01(+)=DECODE(X.CP10,'404',X.CP43,NULL)" & _
               " AND CPM01=X.CP01 AND CPM02=DECODE(X.CP10,'404',NVL(NP07,B.CP10),X.CP10)" & _
               " order by ALD09 DESC, a1p17, a1p07 DESC, a1p14"
         End If
   
'Modified by Morgan 2104/8/8
'     ElseIf stDept = "T" Then
     ElseIf stDept = "T" Or stDept = "FCT" Then
         '虛擬表格語法
         'Modified by Morgan 2025/3/25 +ALD11
         strVTblX = " SELECT CP01,CP02,CP03,CP04,CP10,CP22,CP43,CP82,CP84,TM23 PA26, ALD09,ALD11" & _
            " From  appliste,applistDETAILe,CASEPROGRESS, TRADEMARK" & _
            " WHERE AL01=" & DBDATE(Text1) & strCont & _
            " AND ALD01(+)=AL01 AND ALD02(+)=AL02 AND ALD03(+)=AL03 and ald04(+)=al04" & _
            " AND CP09(+)=ALD11 AND CP84>0 AND CP118='Y'" & _
            " and TM01(+)=cp01 and TM02(+)=cp02 and TM03(+)=cp03 and TM04(+)=cp04 AND TM01 IS NOT NULL"
               
         'frmacc31g1 用
         If sType = "1" Then
            GetSql = "SELECT LPAD(X.CP01||'-'||X.CP02||'-'||X.CP03||'-'||X.CP04,15,' ') C01, X.CP84 C02, RPAD(CPM03,12,' ') C05, RPAD(NVL(cu04,' '),20,' ') C06" & _
               " FROM (" & strVTblX & ") X, customer" & _
               ",CASEPROGRESS B, NEXTPROGRESS,CASEPROPERTYMAP" & _
               " WHERE cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9,1)" & _
               " AND B.CP09(+)=DECODE(X.CP10,'303',X.CP43,NULL) AND NP01(+)=DECODE(X.CP10,'303',X.CP43,NULL)" & _
               " AND CPM01=X.CP01 AND CPM02=DECODE(X.CP10,'303',NVL(NP07,B.CP10),X.CP10)" & _
               " order by ALD09 DESC, 1, 2 DESC, 3"
               
         'acc_save 用
         Else
            'P,T的摘要加本所案號
            'Modified by Morgan 2025/3/25 +ALD11
            GetSql = "SELECT decode(X.CP01,'T','220101','P','220102','FCT','220103','CFT','220105','220104') a1p05" & _
               ", X.cp84 a1p07,X.CP01||X.CP02||X.CP03||X.CP04||decode(instr('T,P',X.CP01),0,'','/'||substr(cu04,1,10))||'/'||CPM03 a1p14" & _
               ", X.CP01||X.CP02||X.CP03||X.CP04 a1p17,ALD11" & _
               " FROM (" & strVTblX & ") X, customer" & _
               ",CASEPROGRESS B, NEXTPROGRESS,CASEPROPERTYMAP" & _
               " WHERE cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9,1)" & _
               " AND B.CP09(+)=DECODE(X.CP10,'303',X.CP43,NULL) AND NP01(+)=DECODE(X.CP10,'303',X.CP43,NULL)" & _
               " AND CPM01=X.CP01 AND CPM02=DECODE(X.CP10,'303',NVL(NP07,B.CP10),X.CP10)" & _
               " order by ALD09 DESC, a1p17, a1p07 DESC, a1p14"
         End If
      
'Removed by Morgan 2014/8/8 與內商改共用
'      ElseIf stDept = "FCT" Then
'         '虛擬表格語法
'         strVTblX = " SELECT CP01,CP02,CP03,CP04,CP10,CP22,CP43,CP82,CP84,TM23 PA26, ALD09" & _
'            " From  appliste,applisteDETAIL,CASEPROGRESS, TRADEMARK" & _
'            " WHERE AL01=" & DBDATE(Text1) & strCon & _
'            " AND ALD01(+)=AL01 AND ALD02(+)=AL02 AND ALD03(+)=AL03 AND ALD04(+)=AL04" & _
'            " AND CP01(+)=ALD05 AND CP02(+)=ALD06 AND CP03(+)=ALD07 AND CP04(+)=ALD08" & _
'            " AND CP27(+)=ALD01 AND CP84>0 AND CP118='Y'" & _
'            " and TM01(+)=cp01 and TM02(+)=cp02 and TM03(+)=cp03 and TM04(+)=cp04 AND TM01 IS NOT NULL"
'         'frmacc31g1 用
'         If sType = "1" Then
'            GetSql = "SELECT LPAD(X.CP01||'-'||X.CP02||'-'||X.CP03||'-'||X.CP04,15,' ') C01, X.CP84 C02, RPAD(CPM03,12,' ') C05, RPAD(NVL(cu04,' '),20,' ') C06" & _
'               " FROM (" & strVTblX & ") X, customer" & _
'               ",CASEPROGRESS B, NEXTPROGRESS,CASEPROPERTYMAP" & _
'               " WHERE cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9,1)" & _
'               " AND B.CP09(+)=DECODE(X.CP10,'303',X.CP43,NULL) AND NP01(+)=DECODE(X.CP10,'303',X.CP43,NULL)" & _
'               " AND CPM01=X.CP01 AND CPM02=DECODE(X.CP10,'303',NVL(NP07,B.CP10),X.CP10)" & _
'               " order by ALD09 DESC, 1, 2 DESC, 3"
'
'         'acc_save 用
'         Else
'            'P,T的摘要加本所案號
'            GetSql = "SELECT decode(X.CP01,'T','220101','P','220102','FCT','220103','CFT','220105','220104') a1p05" & _
'               ", X.cp84 a1p07,X.CP01||X.CP02||X.CP03||X.CP04||decode(instr('T,P',X.CP01),0,'','/'||substr(cu04,1,10))||'/'||CPM03 a1p14" & _
'               ", X.CP01||X.CP02||X.CP03||X.CP04 a1p17" & _
'               " FROM (" & strVTblX & ") X, customer" & _
'               ",CASEPROGRESS B, NEXTPROGRESS,CASEPROPERTYMAP" & _
'               " WHERE cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9,1)" & _
'               " AND B.CP09(+)=DECODE(X.CP10,'303',X.CP43,NULL) AND NP01(+)=DECODE(X.CP10,'303',X.CP43,NULL)" & _
'               " AND CPM01=X.CP01 AND CPM02=DECODE(X.CP10,'303',NVL(NP07,B.CP10),X.CP10)" & _
'               " order by ALD09 DESC, a1p17, a1p07 DESC, a1p14"
'         End If
            
   End If
   
   GetSql = GetSql
      
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

'Amy 2014/11/05 由aacc_sav搬回 'Add by Morgan 2011/6/2 參考Frmacc31e0_Save
Public Sub Frmacc31g0_Save()
   
   With Frmacc31g0
   
      Screen.MousePointer = vbHourglass
      
      Dim iRowNum As Integer, idx As Integer
      Dim Acc1p0(1 To 18) As String
      Dim sNum '支票號碼
      Dim sAmt '支票金額
      Dim sData() As String '已用序號,累計金額
      Dim idxNow As Integer '現在執行的支票號索引值
      Dim stCon As String 'SQL條件語法
      Dim bolJAlert As Boolean, lngAmtJ As Long 'Added by Morgan 2025/3/25 對轉傳票金額
      
On Error GoTo Checking
   
   
      If .CheckData = True Then
      
         bolJAlert = False 'Added by Morgan 2025/3/25
   
         sNum = Split(.sNums, ",")
         sAmt = Split(.sAmts, ",")
      
         '序號,金額累計
         ReDim sData(UBound(sNum), 1 To 2)
      
On Error GoTo Saving

         cnnConnection.BeginTrans
         
         For iRowNum = 1 To .grdDataList.Rows - 1
            strSql = .GetSql(iRowNum, "2")
            CheckOC
            adoRecordset.CursorLocation = adUseClient
            adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset.RecordCount > 0 Then
               lngAmtJ = 0 'Added by Morgan 2025/3/25
               '找出該支票號的陣列位置
               For idx = 0 To UBound(sNum)
                  If sNum(idx) = .grdDataList.TextMatrix(iRowNum, 6) Then
                     idxNow = idx
                     Exit For
                  End If
               Next
               '因為序號及排序原因所以要跑回圈
               adoRecordset.MoveFirst
               Do While Not adoRecordset.EOF
                  sData(idxNow, 1) = Format(Val(sData(idxNow, 1)) + 1)
                  Acc1p0(1) = "'1'"
                  Acc1p0(2) = "'L'"
                  Acc1p0(3) = "'" & Format(sData(idxNow, 1), "000") & "'"
                  'Modify by Amy 2020/08/05 開票銀行改抓變數,並加帳號 原:011010075
                  Acc1p0(4) = "'" & .grdDataList.TextMatrix(iRowNum, 6) & 智慧局送件開票銀行 & 智慧局送件開票帳號 & "2'"
                  Acc1p0(5) = "'" & adoRecordset.Fields("a1p05") & "'"
                  Acc1p0(6) = "'TOT'"
                  Acc1p0(7) = adoRecordset.Fields("a1p07")
                  Acc1p0(8) = "0"
                  Acc1p0(9) = "'" & .grdDataList.TextMatrix(iRowNum, 6) & "'"
                  'Modify by Amy 2020/08/05 開票銀行改抓變數 原:011010075
                  Acc1p0(10) = "'" & 智慧局送件開票銀行 & "'"
                  'Modify by Amy 2020/08/05 開票帳號改抓變數 原:1756650
                  Acc1p0(11) = "'" & 智慧局送件開票帳號 & "'" 'Modify by Amy 2020/07/24 改回1756650 'modify by sonia 2020/6/19 改帳號原為0149951(1756650)
                  'Modified by Morgan 2020/11/17 應該用開票日--辜
                  'Acc1p0(12) = .Text1
                  Acc1p0(12) = .Text2
                  Acc1p0(14) = "'" & adoRecordset.Fields("a1p14") & "'"
                  Acc1p0(15) = "'V0001'"
                  Acc1p0(17) = "'" & adoRecordset.Fields("a1p17") & "'"
                  'Modified by Morgan 2020/11/17 應該用開票日--辜
                  'Acc1p0(18) = .Text1
                  Acc1p0(18) = .Text2
                  sData(idxNow, 2) = Format(Val(sData(idxNow, 2)) + Val(Acc1p0(7)))
                  strSql = "insert into acc1p0(a1p01,a1p02,a1p03,a1p04,a1p05,a1p06,a1p07,a1p08,a1p09,a1p10,a1p11,a1p12,a1p14,a1p15,a1p17,a1p18 )" & _
                     "VALUES(" & Acc1p0(1) & "," & Acc1p0(2) & "," & Acc1p0(3) & "," & Acc1p0(4) & "," & Acc1p0(5) & "," & Acc1p0(6) & _
                     "," & Acc1p0(7) & "," & Acc1p0(8) & "," & Acc1p0(9) & "," & Acc1p0(10) & "," & Acc1p0(11) & "," & Acc1p0(12) & _
                     "," & Acc1p0(14) & "," & Acc1p0(15) & "," & Acc1p0(17) & "," & Acc1p0(18) & ")"
                  cnnConnection.Execute strSql, intI
                  
                  'Added by Morgan 2025/3/25 CFT案的220105(應付規費－CFT)，若收據為智權公司時需產生對轉傳票--瑞婷
                  If Acc1p0(5) = "'220105'" Then
                     strExc(0) = "select a0k11 from acc0j0,acc0k0 where a0j01='" & adoRecordset.Fields("ALD11") & "' and a0k01(+)=a0j13 and a0k11='J'"
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                     If intI = 1 Then
                        bolJAlert = True
                        
                        lngAmtJ = lngAmtJ + Val(Acc1p0(7))
                        
                        strSql = "update acc1p0 set a1p07=" & lngAmtJ & " where a1p01=" & Acc1p0(1) & " and a1p02=" & Acc1p0(2) & " and a1p04=" & Acc1p0(4) & " and a1p05='1101' And A1P14 = " & Acc1p0(14) & " And A1P17 = " & Acc1p0(17)
                        cnnConnection.Execute strSql, intI
                        If intI = 0 Then
                           sData(idxNow, 1) = Format(Val(sData(idxNow, 1)) + 1)
                           Acc1p0(3) = "'" & Format(sData(idxNow, 1), "000") & "'"
                              
                           strSql = "insert into acc1p0(a1p01,a1p02,a1p03,a1p04,a1p05,a1p06,a1p07,a1p08,a1p09,a1p10,a1p11,a1p12,a1p14,a1p15,a1p17,a1p18 )" & _
                              "VALUES(" & Acc1p0(1) & "," & Acc1p0(2) & "," & Acc1p0(3) & "," & Acc1p0(4) & ",'1101'," & Acc1p0(6) & _
                              "," & lngAmtJ & ",0," & Acc1p0(9) & "," & Acc1p0(10) & "," & Acc1p0(11) & "," & Acc1p0(12) & _
                              "," & Acc1p0(14) & "," & Acc1p0(15) & "," & Acc1p0(17) & "," & Acc1p0(18) & ")"
                           cnnConnection.Execute strSql, intI
                        End If
                        
                        sData(idxNow, 1) = Format(Val(sData(idxNow, 1)) + 1)
                        Acc1p0(3) = "'" & Format(sData(idxNow, 1), "000") & "'"
                        
                        strSql = "insert into acc1p0(a1p01,a1p02,a1p03,a1p04,a1p05,a1p06,a1p07,a1p08,a1p09,a1p10,a1p11,a1p12,a1p14,a1p15,a1p17,a1p18 )" & _
                           "VALUES(" & Acc1p0(1) & "," & Acc1p0(2) & "," & Acc1p0(3) & "," & Acc1p0(4) & "," & Acc1p0(5) & "," & Acc1p0(6) & _
                           ",0," & Acc1p0(7) & "," & Acc1p0(9) & "," & Acc1p0(10) & "," & Acc1p0(11) & "," & Acc1p0(12) & _
                           "," & Acc1p0(14) & "," & Acc1p0(15) & "," & Acc1p0(17) & "," & Acc1p0(18) & ")"
                        cnnConnection.Execute strSql, intI
                        
                        sData(idxNow, 1) = Format(Val(sData(idxNow, 1)) + 1)
                        Acc1p0(3) = "'" & Format(sData(idxNow, 1), "000") & "'"
                        
                        strSql = "insert into acc1p0(a1p01,a1p02,a1p03,a1p04,a1p05,a1p06,a1p07,a1p08,a1p09,a1p10,a1p11,a1p12,a1p14,a1p15,a1p17,a1p18 )" & _
                           "VALUES('J'," & Acc1p0(2) & "," & Acc1p0(3) & "," & Acc1p0(4) & "," & Acc1p0(5) & "," & Acc1p0(6) & _
                           "," & Acc1p0(7) & ",0," & Acc1p0(9) & "," & Acc1p0(10) & "," & Acc1p0(11) & "," & Acc1p0(12) & _
                           "," & Acc1p0(14) & "," & Acc1p0(15) & "," & Acc1p0(17) & "," & Acc1p0(18) & ")"
                        cnnConnection.Execute strSql, intI
                        
                        
                        strSql = "update acc1p0 set a1p08=" & lngAmtJ & " where a1p01='J' and a1p02=" & Acc1p0(2) & " and a1p04=" & Acc1p0(4) & " and a1p05='1101' And A1P14 = " & Acc1p0(14) & " And A1P17 = " & Acc1p0(17)
                        cnnConnection.Execute strSql, intI
                        If intI = 0 Then
                           sData(idxNow, 1) = Format(Val(sData(idxNow, 1)) + 1)
                           Acc1p0(3) = "'" & Format(sData(idxNow, 1), "000") & "'"
                              
                           strSql = "insert into acc1p0(a1p01,a1p02,a1p03,a1p04,a1p05,a1p06,a1p07,a1p08,a1p09,a1p10,a1p11,a1p12,a1p14,a1p15,a1p17,a1p18 )" & _
                              "VALUES('J'," & Acc1p0(2) & "," & Acc1p0(3) & "," & Acc1p0(4) & ",'1101'," & Acc1p0(6) & _
                              ",0," & lngAmtJ & "," & Acc1p0(9) & "," & Acc1p0(10) & "," & Acc1p0(11) & "," & Acc1p0(12) & _
                              "," & Acc1p0(14) & "," & Acc1p0(15) & "," & Acc1p0(17) & "," & Acc1p0(18) & ")"
                           cnnConnection.Execute strSql, intI
                        End If
                        
                        
                     End If
                  End If
                  'end 2025/3/25
                  
                  adoRecordset.MoveNext
               Loop
              
               '時段
               stCon = ""
               Select Case .grdDataList.TextMatrix(iRowNum, 0)
                  Case "上午": stCon = stCon & " and AL03='0'"
                  Case "下午": stCon = stCon & " and AL03='1'"
               End Select
               '部門
               Select Case .grdDataList.TextMatrix(iRowNum, 1)
                  Case "P": stCon = stCon & " and AL02='P1'"
                  Case "T": stCon = stCon & " and AL02='P2'"
                  Case "FCT": stCon = stCon & " and AL02='F1'"
                  Case "FCP": stCon = stCon & " and AL02='F2'"
               End Select
               
               '類別
               Select Case .grdDataList.TextMatrix(iRowNum, 3)
                  Case "新案": stCon = stCon & " and AL04='1'"
                  Case "一般": stCon = stCon & " and AL04='2'"
                  Case "快速": stCon = stCon & " and AL04='3'"
                  Case "發明": stCon = stCon & " and AL04='4'"
                  Case "新型": stCon = stCon & " and AL04='5'"
                  Case "設計": stCon = stCon & " and AL04='6'"
                  Case "非智慧局": stCon = stCon & " and AL04='7'"
                  Case "電子送件": stCon = stCon & " and AL04='8'"
               End Select
               
               strSql = "UPDATE appliste SET AL06='" & sNum(idxNow) & "' WHERE AL01=" & DBDATE(.Text1) & stCon
               cnnConnection.Execute strSql
               
               '貸方
               If sData(idxNow, 2) = sAmt(idxNow) Then
                  sData(idxNow, 1) = Format(Val(sData(idxNow, 1)) + 1)
                  Acc1p0(3) = "'" & Format(sData(idxNow, 1), "000") & "'"
                  Acc1p0(5) = "'2111'"
                  Acc1p0(7) = "0"
                  Acc1p0(8) = sData(idxNow, 2)
                  'Modified by Morgan 2017/4/5 前面不必加 0 --瑞婷
                  'Acc1p0(14) = "'0" & .Text1 & "/" & sNum(idxNow) & "/" & "經濟部智慧財產局'"
                  Acc1p0(14) = "'" & .Text1 & "/" & sNum(idxNow) & "/" & "經濟部智慧財產局'"
                  Acc1p0(17) = "NULL"
                  strSql = "insert into acc1p0(a1p01,a1p02,a1p03,a1p04,a1p05,a1p06,a1p07,a1p08,a1p09,a1p10,a1p11,a1p12,a1p14,a1p15,a1p17,a1p18 )" & _
                     "VALUES(" & Acc1p0(1) & "," & Acc1p0(2) & "," & Acc1p0(3) & "," & Acc1p0(4) & "," & Acc1p0(5) & "," & Acc1p0(6) & _
                     "," & Acc1p0(7) & "," & Acc1p0(8) & "," & Acc1p0(9) & "," & Acc1p0(10) & "," & Acc1p0(11) & "," & Acc1p0(12) & _
                     "," & Acc1p0(14) & "," & Acc1p0(15) & "," & Acc1p0(17) & "," & Acc1p0(18) & ")"
                  
                  cnnConnection.Execute strSql
               '若進畫面後才改發文資料發現金額不符而無法產生貸方科目！
               Else
                  cnnConnection.RollbackTrans
                  MsgBox "資料已異動，票據號碼[" & sNum(idxNow) & "]金額[$" & Format(sAmt(idxNow), DDollar) & "]與實際發文金額[$" & Format(sData(idxNow, 2), DDollar) & "]不符，請確認後重新執行！", vbCritical
                  GoTo Checking
               End If
            End If
         Next
         
         cnnConnection.CommitTrans
         MsgBox "轉傳票分錄資料產生完成！", vbInformation
         
         'Added by Morgan 2025/3/25
         If bolJAlert Then
            MsgBox "本次有同時產生智權公司對轉傳票分錄！", vbExclamation
         End If
         'end 2025/3/25
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
