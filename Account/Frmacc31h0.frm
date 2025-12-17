VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Frmacc31h0 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "智慧局電子送件網路扣帳作業"
   ClientHeight    =   4596
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   4656
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4596
   ScaleWidth      =   4656
   Begin VB.TextBox Text1 
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
      Height          =   324
      Left            =   1125
      TabIndex        =   5
      Top             =   4200
      Width           =   1500
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
      Left            =   1125
      MaxLength       =   7
      TabIndex        =   0
      Text            =   "Text2"
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
      Left            =   2340
      TabIndex        =   1
      Top             =   143
      Width           =   675
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
      Left            =   3285
      TabIndex        =   2
      Top             =   143
      Width           =   975
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   3585
      Left            =   225
      TabIndex        =   3
      Top             =   570
      Width           =   4170
      _ExtentX        =   7345
      _ExtentY        =   6329
      _Version        =   393216
      BackColor       =   -2147483624
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label15 
      Alignment       =   2  '置中對齊
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
      Height          =   255
      Left            =   180
      TabIndex        =   6
      Top             =   4200
      Width           =   735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "扣帳日："
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
      TabIndex        =   4
      Top             =   210
      Width           =   900
   End
End
Attribute VB_Name = "Frmacc31h0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/10/20 Form2.0已修改 (無需修改)
'Memo By Sonia 2012/12/4 智權人員欄已修改
'Create by Morgan 2012/10/9 參考 frmacc31g0
Option Explicit

Public iNowRow As Integer '本次點選列數
Dim iLstRow As Integer '前次點選列數
Public adoacc1p0 As New ADODB.Recordset '存檔用
Public adoquery As New ADODB.Recordset '查詢用
Dim adoSum As New ADODB.Recordset, strSum As String, intI As Integer 'Add by Amy 2021/07/23 合計

Private Sub cmdDetail_Click()
   Screen.MousePointer = vbHourglass
   If iNowRow > 0 Then
      Me.Enabled = False
      tool3_enabled
      Frmacc31h1.Show
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Sub Command1_Click()
   Dim bCancel  As Boolean
   
   tool3_enabled
   'Removed by Morgan 2013/5/15 取消發文日
   'Text1_Validate bCancel
   'If bCancel Then Exit Sub
   'end 2013/5/15
   Text2_Validate bCancel
   If bCancel = False Then
      SetDataListWidth
      SetGrid
   End If
End Sub

Private Sub Form_Load()

   PUB_InitForm Me, Me.Width, Me.Height
   
   'Removed by Morgan 2013/5/15 取消發文日
   '發文日
   'Modified by Morgan 2012/12/28 改預設前一工作天
   'Text1.Text = TransDate(PUB_GetWorkDay1(strSrvDate(1) - 1, True), 1)
   'end 2013/5/15
   
   '扣帳日=發文日下一個工作天
   'Modified by Morgan 2013/5/15
   'Text2.Text = TransDate(PUB_GetWorkDay1(CompDate(2, 1, Text1.Text), False), 1)
   Text2.Text = strSrvDate(2)
   'end 2013/5/15
   
   'Modified by Morgan 2023/8/7
   'SetDataListWidth
   'SetGrid
   Command1.Value = True
   'end 2023/8/7
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(133)
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   StatusClear
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc31h0 = Nothing
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
      .Visible = False
      .Clear
      .Rows = 2
      .row = 0:
      .col = 0: .ColWidth(.col) = 900: .Text = "部門別"
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .col = 1: .ColWidth(.col) = 2600: .Text = "金額"
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .Visible = True
      .Enabled = False
      iNowRow = 0
   End With
End Sub

Private Sub SetGrid()
   Dim iRow As Integer
      
On Error GoTo ErrHnd
   'Modified by Morgan 2013/5/15
   'strExc(0) = "select * from acc1p0 where a1p04='電子送件網路扣帳" & Text1 & "'"
   strExc(0) = "select * from acc1p0 where a1p04='電子送件網路扣帳" & Text2 & "'"
   'end 2013/5/15
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      'Modified by Morgan 2013/5/15
      'MsgBox ChangeTStringToTDateString(Text1) & " 電子送件網路扣帳分錄資料已存在！"
      MsgBox ChangeTStringToTDateString(Text2) & " 電子送件網路扣帳分錄資料已存在！"
      'end 2013/5/15
      Exit Sub
   End If
   
   'Modified by Morgan 2013/5/15
   'strSql = "SELECT cp01,SUM(CP84) AMT" & _
      " From caseprogress" & _
      " WHERE cp27=" & DBDATE(Text1) & " AND cp118='A' and cp84>0 " & _
      " GROUP BY cp01"
   'Modify by Amy 2021/07/23 +strSum
   strSql = "SELECT cp01,SUM(CP84) AMT" & _
      " From caseprogress" & _
      " WHERE cp152=" & DBDATE(Text2) & " AND cp118='A' and cp84>0 "
   strSum = strSql
   strSql = strSql & " GROUP BY cp01"
   'end 2021/07/23
   'end 2013/5/13
   CheckOC
   With adoRecordset
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount > 0 Then
         iRow = 0
         Do While Not .EOF
            iRow = iRow + 1
            grdDataList.Rows = iRow + 1
            grdDataList.TextMatrix(iRow, 0) = "" & .Fields(0)
            grdDataList.TextMatrix(iRow, 1) = Format("" & .Fields(1), "#,###")
            .MoveNext
         Loop
         grdDataList.row = 1
         tool17_enabled
         grdDataList.Enabled = True
         ShowSum 'Add by Amy 2021/07/23
         Text2.Tag = Text2.Text 'Added by Morgan 2023/8/7
      Else
         MsgBox "無待作業資料！", vbInformation, "注意"
         
      End If
   End With
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical

End Sub

Private Sub GrdDataList_Click()
   If grdDataList.MouseRow > 0 Then
      iLstRow = iNowRow: iNowRow = grdDataList.MouseRow
      '若超出可選範圍則不作用
      If iNowRow = 0 Or iNowRow = grdDataList.Rows Then
         iNowRow = iLstRow
      Else
         'SelectRow
      End If
   End If
   
End Sub

Private Sub SelectRow()
   
   Dim i As Integer
   
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
End Sub

'取得明細語法
'若性質為延期303時:A類收文用CP43抓NP07；B類收文用CP43抓相關收文號的CP10
Public Function GetSql(Optional ByVal iRowNo As Integer = 0, Optional ByVal sType As String = "1") As String

   Dim strCon As String, stDept As String, strCont As String
   
   Dim strVTblX As String, strVTblY As String
   Dim strConX As String, strConY As String, stTemp As String

   If iRowNo = 0 Then iRowNo = iNowRow
   
   stDept = grdDataList.TextMatrix(iRowNo, 0)
   
   strCon = " and cp01='" & stDept & "'"
   
   If stDept = "P" Or stDept = "FCP" Then
      '虛擬表格語法
      'Modified by Morgan 2013/5/15
      'strVTblX = " SELECT CP01,CP02,CP03,CP04,CP10,CP22,CP43,CP82,CP84,PA26,CP14" & _
         " From CASEPROGRESS, patent" & _
         " WHERE cp27=" & DBDATE(Text1) & " AND cp118='A' and cp84>0" & strCon & _
         " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 AND PA01 IS NOT NULL"
      strVTblX = " SELECT CP01,CP02,CP03,CP04,CP10,CP22,CP43,CP82,CP84,PA26,CP14" & _
         " From CASEPROGRESS, patent" & _
         " WHERE cp152=" & DBDATE(Text2) & " AND cp118='A' and cp84>0" & strCon & _
         " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 AND PA01 IS NOT NULL"
      'end 2013/5/15
      
      'frmacc31h1 用
      If sType = "1" Then
         GetSql = "SELECT LPAD(X.CP01||'-'||X.CP02||'-'||X.CP03||'-'||X.CP04,15,' ') C01, X.CP84 C02, RPAD(CPM03,12,' ') C05, RPAD(NVL(cu04,' '),20,' ') C06" & _
            " FROM (" & strVTblX & ") X, customer" & _
            ",CASEPROGRESS B, NEXTPROGRESS,CASEPROPERTYMAP" & _
            " WHERE " & _
            " cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9,1)" & _
            " AND B.CP09(+)=DECODE(X.CP10,'404',X.CP43,NULL) AND NP01(+)=DECODE(X.CP10,'404',X.CP43,NULL)" & _
            " AND CPM01=X.CP01 AND CPM02=DECODE(X.CP10,'404',NVL(NP07,B.CP10),X.CP10)" & _
            " order by X.cp14,X.cp01,X.cp02"
            
      'acc_save 用
      Else
         GetSql = "SELECT decode(X.cp01,'T','220101','P','220102','FCT','220103','220104') a1p05" & _
            ", X.cp84 a1p07,X.CP01||X.CP02||X.CP03||X.CP04||decode(instr('T,P',X.CP01),0,'','/'||substr(cu04,1,10))||'/'||CPM03 a1p14" & _
            ", X.CP01||X.CP02||X.CP03||X.CP04 a1p17" & _
            " FROM (" & strVTblX & ") X, customer" & _
            ",CASEPROGRESS B, NEXTPROGRESS,CASEPROPERTYMAP" & _
            " WHERE cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9,1)" & _
            " AND B.CP09(+)=DECODE(X.CP10,'404',X.CP43,NULL) AND NP01(+)=DECODE(X.CP10,'404',X.CP43,NULL)" & _
            " AND CPM01=X.CP01 AND CPM02=DECODE(X.CP10,'404',NVL(NP07,B.CP10),X.CP10)" & _
            " order by X.cp14,X.cp01,X.cp02"
      End If
      
   'Added by Morgan 2020/1/14 +商標
   Else
      strVTblX = " SELECT CP01,CP02,CP03,CP04,CP10,CP22,CP43,CP82,CP84,tm23 PA26,CP14" & _
         " From CASEPROGRESS, trademark" & _
         " WHERE cp152=" & DBDATE(Text2) & " AND cp118='A' and cp84>0" & strCon & _
         " and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04 AND tm01 IS NOT NULL"
   
      'frmacc31h1 用
      If sType = "1" Then
         GetSql = "SELECT LPAD(X.CP01||'-'||X.CP02||'-'||X.CP03||'-'||X.CP04,15,' ') C01, X.CP84 C02, RPAD(CPM03,12,' ') C05, RPAD(NVL(cu04,' '),20,' ') C06" & _
            " FROM (" & strVTblX & ") X, customer" & _
            ",CASEPROGRESS B, NEXTPROGRESS,CASEPROPERTYMAP" & _
            " WHERE " & _
            " cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9,1)" & _
            " AND B.CP09(+)=DECODE(X.CP10,'303',X.CP43,NULL) AND NP01(+)=DECODE(X.CP10,'303',X.CP43,NULL)" & _
            " AND CPM01=X.CP01 AND CPM02=DECODE(X.CP10,'404',NVL(NP07,B.CP10),X.CP10)" & _
            " order by X.cp14,X.cp01,X.cp02"
            
      'acc_save 用
      Else
         GetSql = "SELECT decode(X.cp01,'T','220101','P','220102','FCT','220103','220104') a1p05" & _
            ", X.cp84 a1p07,X.CP01||X.CP02||X.CP03||X.CP04||decode(instr('T,P',X.CP01),0,'','/'||substr(cu04,1,10))||'/'||CPM03 a1p14" & _
            ", X.CP01||X.CP02||X.CP03||X.CP04 a1p17" & _
            " FROM (" & strVTblX & ") X, customer" & _
            ",CASEPROGRESS B, NEXTPROGRESS,CASEPROPERTYMAP" & _
            " WHERE cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9,1)" & _
            " AND B.CP09(+)=DECODE(X.CP10,'303',X.CP43,NULL) AND NP01(+)=DECODE(X.CP10,'303',X.CP43,NULL)" & _
            " AND CPM01=X.CP01 AND CPM02=DECODE(X.CP10,'404',NVL(NP07,B.CP10),X.CP10)" & _
            " order by X.cp14,X.cp01,X.cp02"
      End If
      
   End If
   GetSql = GetSql
      
End Function

'Removed by Morgan 2012/5/15
'Private Sub Text1_Change()
'   tool3_enabled
'   SetDataListWidth
'End Sub
'
'Private Sub Text1_GotFocus()
'   TextInverse Text1
'End Sub
'
'Private Sub Text1_Validate(Cancel As Boolean)
'   If Text1 = "" Then
'      MsgBox "發文日不可空白！"
'      Cancel = True
'   ElseIf ChkDate(Text1) = False Then
'      Cancel = True
'   Else
'      Text2.Text = TransDate(PUB_GetWorkDay1(CompDate(2, 1, Text1.Text), False), 1)
'   End If
'End Sub
'end 2013/5/15

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
   If Text2 = "" Then
      MsgBox "扣帳日不可空白！", vbExclamation
      Cancel = True
   ElseIf ChkDate(Text2) = False Then
      Cancel = True
   'Added by Morgan 2023/8/7
   '扣帳日當天北所放颱風假或為假日(5/1勞動節)，都不可以操作--秀玲
   ElseIf ChkWorkDay(DBDATE(Text2), , True, "1") = False Then
      MsgBox "扣帳日必須為工作日！", vbExclamation
      Cancel = True
   ElseIf DBDATE(Text2) > strSrvDate(1) Then
      MsgBox "扣帳日不可晚於系統日！", vbExclamation
      Cancel = True
   End If
   
End Sub


'Add by Morgan 2012/10/11 參考Frmacc31g0_Save
Public Sub Frmacc31h0_Save()
   
   With Frmacc31h0
   
      'Added by Morgan 2023/8/7
      If .Text2.Tag <> .Text2.Text Then
         MsgBox "扣帳日有變更，請重新查詢！", vbExclamation
         Exit Sub
      End If
      'end 2023/8/7
      
      Screen.MousePointer = vbHourglass
      
      Dim iRowNum As Integer, idx As Integer
      Dim Acc1p0(1 To 18) As String
      Dim lngAmt As Long, lngAmtTot As Long
      
On Error GoTo Saving

      cnnConnection.BeginTrans
      
      '應付
      Acc1p0(1) = "'1'"
      Acc1p0(2) = "'L'"
      'Modified by Morgan 2013/5/15
      'Acc1p0(4) = "'電子送件" & .Text1 & "'"
      Acc1p0(4) = "'電子送件" & .Text2 & "'"
      Acc1p0(6) = "'TOT'"
      Acc1p0(15) = "'V0001'"
      
      idx = 0
      lngAmtTot = 0
      '借方
      For iRowNum = 1 To .grdDataList.Rows - 1
         strSql = .GetSql(iRowNum, "2")
         CheckOC
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If adoRecordset.RecordCount > 0 Then
            adoRecordset.MoveFirst
            lngAmt = 0
            Do While Not adoRecordset.EOF
               idx = idx + 1
               Acc1p0(3) = "'" & Format(idx, "000") & "'"
               Acc1p0(5) = "'" & adoRecordset.Fields("a1p05") & "'"
               Acc1p0(7) = adoRecordset.Fields("a1p07")
               Acc1p0(8) = "0"
               Acc1p0(14) = "'" & adoRecordset.Fields("a1p14") & "'"
               Acc1p0(17) = "'" & adoRecordset.Fields("a1p17") & "'"
               'Modifed by Morgan 2013/5/15
               'Acc1p0(18) = .Text1
               Acc1p0(18) = .Text2
               
               strSql = "insert into acc1p0(a1p01,a1p02,a1p03,a1p04,a1p05,a1p06,a1p07,a1p08,a1p14,a1p15,a1p17,a1p18 )" & _
                  "VALUES(" & Acc1p0(1) & "," & Acc1p0(2) & "," & Acc1p0(3) & "," & Acc1p0(4) & "," & Acc1p0(5) & "," & Acc1p0(6) & _
                  "," & Acc1p0(7) & "," & Acc1p0(8) & _
                  "," & Acc1p0(14) & "," & Acc1p0(15) & "," & Acc1p0(17) & "," & Acc1p0(18) & ")"
               cnnConnection.Execute strSql, intI
               lngAmt = lngAmt + adoRecordset.Fields("a1p07")
               adoRecordset.MoveNext
            Loop
         End If
         If Format(.grdDataList.TextMatrix(iRowNum, 1)) <> lngAmt Then
            cnnConnection.RollbackTrans
            MsgBox "資料已異動，畫面金額[$" & .grdDataList.TextMatrix(iRowNum, 1) & "]與實際發文金額[$" & Format(lngAmt, DDollar) & "]不符，請確認後重新執行！", vbCritical
            GoTo Saving
         End If
         lngAmtTot = lngAmtTot + lngAmt
      Next
      
      '貸方(一天一次總扣款)
      idx = idx + 1
      Acc1p0(3) = "'" & Format(idx, "000") & "'"
      'modify by sonia 2013/12/6 科目由2115改2117,因外帳2115另有使用
      Acc1p0(5) = "'2117'"
      Acc1p0(7) = "0"
      Acc1p0(8) = lngAmtTot
      Acc1p0(14) = "'送件規費'"
      strSql = "insert into acc1p0(a1p01,a1p02,a1p03,a1p04,a1p05,a1p06,a1p07,a1p08,a1p14,a1p15,a1p18 )" & _
         "VALUES(" & Acc1p0(1) & "," & Acc1p0(2) & "," & Acc1p0(3) & "," & Acc1p0(4) & "," & Acc1p0(5) & "," & Acc1p0(6) & _
         "," & Acc1p0(7) & "," & Acc1p0(8) & _
         "," & Acc1p0(14) & "," & Acc1p0(15) & "," & Acc1p0(18) & ")"
      cnnConnection.Execute strSql
      
      '銀行出帳
      '借方
      'Modified by Morgan 2013/5/15
      'Acc1p0(4) = "'電子送件網路扣帳" & .Text1 & "'"
      Acc1p0(4) = "'電子送件網路扣帳" & .Text2 & "'"
      Acc1p0(3) = "'001'"
      'modify by sonia 2013/12/6 科目由2115改2117,因外帳2115另有使用
      Acc1p0(5) = "'2117'"
      Acc1p0(7) = lngAmtTot
      Acc1p0(8) = "0"
      Acc1p0(14) = "'電子送件網路扣帳'"
      Acc1p0(18) = .Text2
      strSql = "insert into acc1p0(a1p01,a1p02,a1p03,a1p04,a1p05,a1p06,a1p07,a1p08,a1p14,a1p15,a1p18 )" & _
         "VALUES(" & Acc1p0(1) & "," & Acc1p0(2) & "," & Acc1p0(3) & "," & Acc1p0(4) & "," & Acc1p0(5) & "," & Acc1p0(6) & _
         "," & Acc1p0(7) & "," & Acc1p0(8) & _
         "," & Acc1p0(14) & "," & Acc1p0(15) & "," & Acc1p0(18) & ")"
      cnnConnection.Execute strSql
      '手續費
      Acc1p0(3) = "'002'"
      Acc1p0(5) = "'611301'"
      Acc1p0(7) = "10"
      Acc1p0(8) = "0"
      Acc1p0(14) = "'華銀手續費'"
      Acc1p0(18) = .Text2
      strSql = "insert into acc1p0(a1p01,a1p02,a1p03,a1p04,a1p05,a1p06,a1p07,a1p08,a1p14,a1p15,a1p18 )" & _
         "VALUES(" & Acc1p0(1) & "," & Acc1p0(2) & "," & Acc1p0(3) & "," & Acc1p0(4) & "," & Acc1p0(5) & "," & Acc1p0(6) & _
         "," & Acc1p0(7) & "," & Acc1p0(8) & _
         "," & Acc1p0(14) & "," & Acc1p0(15) & "," & Acc1p0(18) & ")"
      cnnConnection.Execute strSql
      '貸方
      Acc1p0(3) = "'003'"
      Acc1p0(5) = "'110207'"
      Acc1p0(7) = "0"
      Acc1p0(8) = lngAmtTot + 10
      Acc1p0(14) = "'電子送件網路扣帳'"
      Acc1p0(18) = .Text2
      strSql = "insert into acc1p0(a1p01,a1p02,a1p03,a1p04,a1p05,a1p06,a1p07,a1p08,a1p14,a1p15,a1p18 )" & _
         "VALUES(" & Acc1p0(1) & "," & Acc1p0(2) & "," & Acc1p0(3) & "," & Acc1p0(4) & "," & Acc1p0(5) & "," & Acc1p0(6) & _
         "," & Acc1p0(7) & "," & Acc1p0(8) & _
         "," & Acc1p0(14) & "," & Acc1p0(15) & "," & Acc1p0(18) & ")"
      cnnConnection.Execute strSql
      
      cnnConnection.CommitTrans
      MsgBox "轉傳票分錄資料產生完成！"
      KeyEnter vbKeyEscape
         
Saving:
      If Err.Number <> 0 Then
         cnnConnection.RollbackTrans
         MsgBox Err.Description, vbCritical
         Err.Clear
      End If
      Screen.MousePointer = vbDefault
      
   End With
End Sub

'Add by Amy 2021/07/23 顯示合計
Private Sub ShowSum()
    strSum = Replace(strSum, "cp01,", "")
    intI = 1
    Set adoSum = ClsLawReadRstMsg(intI, strSum)
    If intI = 1 Then
        Text1 = Format(adoSum.Fields("Amt"), "#,###")
    End If
    If adoSum.State <> adStateClosed Then adoSum.Close
End Sub
