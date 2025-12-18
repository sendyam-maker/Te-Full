VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm210108 
   BorderStyle     =   1  '單線固定
   Caption         =   "業績達成月報表"
   ClientHeight    =   5520
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9432
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   9432
   Begin VB.CommandButton cmdPrint 
      Caption         =   "列印(&P)"
      Height          =   400
      Left            =   7650
      TabIndex        =   5
      Top             =   90
      Width           =   800
   End
   Begin VB.TextBox txtCloseDate 
      Height          =   285
      Left            =   1260
      MaxLength       =   5
      TabIndex        =   4
      Top             =   150
      Width           =   1140
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "查詢(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6750
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   90
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8520
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   90
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   4845
      Left            =   135
      TabIndex        =   2
      Top             =   540
      Width           =   9180
      _ExtentX        =   16193
      _ExtentY        =   8551
      _Version        =   393216
      BackColor       =   -2147483624
      Cols            =   3
      FixedCols       =   2
      WordWrap        =   -1  'True
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label1 
      Caption         =   "月報表之應收＝日報表目標之合計應收　月報表之達成＝日報表達成之本月實收"
      Height          =   360
      Left            =   3000
      TabIndex        =   6
      Top             =   105
      Width           =   3300
   End
   Begin VB.Label Label2 
      Caption         =   "統計年月"
      Height          =   180
      Left            =   225
      TabIndex        =   3
      Top             =   195
      Width           =   900
   End
End
Attribute VB_Name = "frm210108"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/07/06 更名為「業績達成月報表」 'Memo by Lydia 2021/08/27 上線
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo by Morgan2010/8/19 日期欄已修改
'Add by Morgan 2005/5/9
Option Explicit
Dim stWYM As String, stTYM As String
Dim iDays As Integer '月工作天數
'列印用變數
Dim iFixRowHeight As Integer, iFixColWidth As Integer '固定列高,固定行寬
Dim iRowHeight As Integer, iColWidth As Integer '一般列高,一般行寬
Dim iX0 As Integer, iY0 As Integer '表格起始位置
Dim iX1 As Integer, iY1 As Integer '表格終止位置
Dim iX As Integer, iY As Integer '現在列印位置
Dim iPageCols As Integer '單頁欄數
Dim strContent As String, iContent As Integer '列印內容,寬
Dim dblTemp As Double '暫存
Dim stVTable0  As String 'Add by Amy 2018/11/28

'設定工作天
Private Sub SetGrid1(ByRef p_Rst As ADODB.Recordset)
   Dim iDay As Integer
   Dim strNow As String  'Add by Amy2018/11/28
   
   'Add by Amy 2018/11/28
   strNow = Mid(Right("000000" & ServerTime, 6), 1, 4)
   With grdDataList
      .Visible = False
      .WordWrap = True
      .Clear
      iDays = p_Rst.RecordCount
      .Rows = 1 + 2 * iDays: .FixedRows = 1
      .Cols = 2: .FixedCols = 1
      .MergeCells = flexMergeRestrictColumns
      .MergeCol(0) = True
      .ColWidth(0) = 300
      .RowHeight(0) = 500
      .ColAlignmentFixed = flexAlignCenterCenter
      For iDay = 1 To iDays
         'Add by Amy 2018/11/28 S部門及國外每日業務點數記錄檔沒資料則新增一筆
         Call InsDailyFeat(p_Rst.Fields(0), strNow)
         'end 2018/11/28
         .TextMatrix(2 * iDay - 1, 0) = Val("" & p_Rst.Fields(0))
         .TextMatrix(2 * iDay, 0) = Val("" & p_Rst.Fields(0))
         p_Rst.MoveNext
      Next
      .Visible = True
   End With
End Sub
'設定目標點數
Private Sub SetGrid2(ByRef p_Rst As ADODB.Recordset)
   Dim stDTarget As String, stMTarget As String, iDay As Integer
   Dim stST15 As String, iPos As Integer
   Dim iRow As Integer, iCol As Integer
   With grdDataList
      .Visible = False
      .WordWrap = True
      .Cols = p_Rst.RecordCount + 1
      stST15 = ""
      iCol = 0
      For iPos = 1 To p_Rst.RecordCount
         iCol = iCol + 1
         .ColWidth(iCol) = 700
         .TextMatrix(0, iCol) = "" & p_Rst.Fields("C1")
         '加統計項目的已收欄位
         'Modify by Amy 2018/11/28 巨京不需應收欄位
         'Modify by Amy 2019/07/15 原:And p_Rst.Fields("C1") <> "巨京",巨京併入其他
         If "" & p_Rst.Fields("C5") = "X" Then
            .Cols = .Cols + 1
            iCol = iCol + 1
            .ColWidth(iCol - 1) = 780
            .ColWidth(iCol) = 780
            .TextMatrix(0, iCol) = "" & p_Rst.Fields("C1") & "應收"
            '應收點數
            stMTarget = "" & p_Rst.Fields("C2")
            stDTarget = Format(Val(stMTarget) / iDays, "#.00")
            For iDay = 1 To iDays - 1
               .TextMatrix(2 * iDay - 1, iCol) = stDTarget
               .TextMatrix(2 * iDay, iCol) = Format(iDay * Val(stDTarget), "#.00")
               .col = iCol
               .row = 2 * iDay - 1
               .CellBackColor = &H7FFFD4
               .row = 2 * iDay
               .CellBackColor = &H7FFFD4
            Next
            '最後一天
            .TextMatrix(2 * iDay - 1, iCol) = stDTarget
            .TextMatrix(2 * iDay, iCol) = Format(Val(stMTarget), "#.00")
            .col = iCol
            .row = 2 * iDay - 1
            .CellBackColor = &H7FFFD4
            .row = 2 * iDay
            .CellBackColor = &H7FFFD4
         End If
         p_Rst.MoveNext
      Next
      .Visible = True
   End With
End Sub
'設定收款點數
Private Sub SetGrid3(ByRef p_Rst As ADODB.Recordset)
   Dim iDay As Integer, dblAreaAmount As Double, bolNoData As Boolean
   Dim dblSum(1 To 4) '1北所2中所3國內4全所
   Dim iRow As Integer, iCol As Integer
   Dim ii As Integer
   
   With grdDataList
      .Visible = False
      .WordWrap = True
      iRow = 1
      Do While Not p_Rst.EOF
         '找日期對應的Row
         iDay = Val(Right("" & p_Rst.Fields(0), 2))
         Do While iRow < .Rows
            'Grid 不同日
            If Val(.TextMatrix(iRow, 0)) = iDay Then
               Exit Do
            End If
            iRow = iRow + 2
         Loop
         iCol = 1
         dblAreaAmount = 0
         bolNoData = True
         dblSum(4) = 0
         '找智權人員對應的位置
         Do While Not p_Rst.EOF
            If (Val(Right("" & p_Rst.Fields(0), 2)) <> iDay) Then
               Exit Do
            End If
            Do While iCol < .Cols
               '找到智權人員
               If "" & p_Rst.Fields(1) = .TextMatrix(0, iCol) Then
                  dblTemp = p_Rst.Fields(2)
                  .TextMatrix(iRow, iCol) = Format(dblTemp, "#.00")
                  If iRow = 1 Then
                     .TextMatrix(iRow + 1, iCol) = .TextMatrix(iRow, iCol)
                  Else
                     dblTemp = Val(.TextMatrix(iRow - 1, iCol)) + Val(.TextMatrix(iRow, iCol))
                     .TextMatrix(iRow + 1, iCol) = Format(dblTemp, "#.00")
                  End If
                  dblAreaAmount = dblAreaAmount + Val(p_Rst.Fields(2))
                  bolNoData = False
                  Exit Do
               ElseIf Right(.TextMatrix(0, iCol), 2) = "應收" Then
                  If bolNoData = False Then
                     dblTemp = dblAreaAmount
                     .TextMatrix(iRow, iCol - 1) = Format(dblTemp, "#.00")
                     If iRow = 1 Then
                        .TextMatrix(iRow + 1, iCol - 1) = .TextMatrix(iRow, iCol - 1)
                     Else
                        dblTemp = Val(.TextMatrix(iRow - 1, iCol - 1)) + Val(.TextMatrix(iRow, iCol - 1))
                        .TextMatrix(iRow + 1, iCol - 1) = Format(dblTemp, "#.00")
                     End If
                     'Add by Morgan 2010/3/15 沒資料時累計同前一日
                     For ii = iCol - 2 To 1 Step -1
                        If Right(.TextMatrix(0, ii), 2) = "應收" Then Exit For
                        If .TextMatrix(iRow + 1, ii) = "" Then
                           If iRow = 1 Then
                              .TextMatrix(iRow + 1, ii) = .TextMatrix(iRow, ii)
                           Else
                              .TextMatrix(iRow + 1, ii) = .TextMatrix(iRow - 1, ii)
                           End If
                        End If
                     Next
                        
                     dblAreaAmount = 0
                     bolNoData = True
                  End If
               'Added by Morgan 2018/5/15 沒資料時累計同前一日
               ElseIf .TextMatrix(iRow, iCol) = "" Then
                  dblTemp = 0
                  .TextMatrix(iRow, iCol) = Format(dblTemp, "#.00")
                  If iRow = 1 Then
                     .TextMatrix(iRow + 1, iCol) = .TextMatrix(iRow, iCol)
                  Else
                     dblTemp = Val(.TextMatrix(iRow - 1, iCol)) + Val(.TextMatrix(iRow, iCol))
                     .TextMatrix(iRow + 1, iCol) = Format(dblTemp, "#.00")
                  End If
               'end 2018/5/15
               End If
               iCol = iCol + 1
            Loop
            
            p_Rst.MoveNext
            If iCol = .Cols Then
               iCol = 1
            Else
               iCol = iCol + 1
            End If
         Loop
      Loop
      
      .Visible = True
   End With
End Sub
'設定加總欄位
Private Sub SetGrid4()
   Dim dblSum(1 To 7) As Double, dblTSum As Double
   Dim iRow As Integer, iCol As Integer
   
   With grdDataList
      .Visible = False
      .WordWrap = True
      For iRow = 1 To .Rows - 1
         Erase dblSum
         If .TextMatrix(iRow, 1) = "" Then Exit For
         '設定合計點數
         For iCol = 1 To .Cols - 1
            'Modify by Amy 2019/07/15 原: .TextMatrix(0, iCol) = "巨京" 因加客服組將巨京併入其他顯示
            If .TextMatrix(0, iCol) = "客服組" Then
                dblSum(7) = Val(.TextMatrix(iRow, iCol))
            'Modify by Amy 2018/11/28 +ＦＣＴ應收 判斷
            ElseIf Right(.TextMatrix(0, iCol), 2) = "應收" And .TextMatrix(0, iCol) <> "ＦＣＴ應收" Then
               Select Case .TextMatrix(0, iCol)
                  Case "北所應收"
                     dblSum(1) = dblTSum: dblTSum = 0
                     dblTemp = dblSum(1)
                     .TextMatrix(iRow, iCol - 1) = Format(dblTemp, "#.00")
                  Case "中所應收"
                     dblSum(2) = dblTSum: dblTSum = 0
                     dblTemp = dblSum(2)
                     .TextMatrix(iRow, iCol - 1) = Format(dblTemp, "#.00")
                  Case "台南所應收"
                     dblSum(3) = Val(.TextMatrix(iRow, iCol - 1))
                  Case "高雄所應收"
                     dblSum(4) = Val(.TextMatrix(iRow, iCol - 1))
                  Case "其他應收"
                     dblSum(5) = Val(.TextMatrix(iRow, iCol - 1))
                  Case "國內應收"
                     'Modify by Amy 2019/07/15 增加的客服組也算入國內應收
                     dblTemp = dblSum(1) + dblSum(2) + dblSum(3) + dblSum(4) + dblSum(5) + dblSum(7)
                     .TextMatrix(iRow, iCol - 1) = Format(dblTemp, "#.00")
                  'Add by Morgan 2007/10/18
                  'Mark by Amy 2018/11/28 巨京無目標,故無應收
'                  Case "巨京應收"
'                     dblSum(7) = Val(.TextMatrix(iRow, iCol - 1))
                  'Add by Amy 2019/07/15
                  Case "客服組應收"
                     dblSum(7) = Val(.TextMatrix(iRow, iCol - 1))
                  'Add by Amy 2018/11/28 10801開始國外部分為FCP及FCT拆出
                  Case "ＦＣＰ應收"
                    dblSum(6) = Val(.TextMatrix(iRow, iCol - 1)) + Val(.TextMatrix(iRow, iCol - 3))
                  Case "國外應收"
                     dblSum(6) = Val(.TextMatrix(iRow, iCol - 1))
                  Case "全所應收"
                     dblTemp = dblSum(1) + dblSum(2) + dblSum(3) + dblSum(4) + dblSum(5) + dblSum(6) + dblSum(7)
                     .TextMatrix(iRow, iCol - 1) = Format(dblTemp, "#.00")
                     
                  Case Else '各區
                     dblTSum = dblTSum + Val(.TextMatrix(iRow, iCol - 1))
               End Select
            End If
         Next
      Next
      .Visible = True
   End With
End Sub

Private Function doQuery() As Boolean
   Dim stVTable1 As String, stVTable2 As String, stVTable5 As String
   Dim stVTable3 As String, stVTable4 As String
   Dim stField As String, stQ As String 'Add by Amy 2018/11/28
   
   stTYM = Val(txtCloseDate)
   stWYM = Val(stTYM) + 191100
   
   'Add by Amy 2018/11/28 抓人員語法由下搬上來
   'Modify by Amy 2019/08/01 增加創新業務組用收入 420101 原:substr(ax205, 1, 2) = '41'
   stVTable0 = " SELECT DISTINCT ST02 C1,0 C2,1 C3,ST15 C4,AX209 C5 FROM ACC020,ACC021,STAFF WHERE A0205>" & stTYM & "00 AND A0205<" & stTYM & "99 AND A0201=AX201(+) AND A0202=AX202(+)" & _
                    "    AND (SUBSTR(AX205,1,1)='4' OR AX205='7191') AND AX209 IS NOT NULL AND NOT (AX205='4191' OR AX205='4192' OR AX205='4194' OR INSTR(AX213||' ','結餘')>0) AND AX209=ST01(+) AND SUBSTR(ST15,1,1)='S'"
   
   '抓工作天
   strSql = "SELECT substr(WD01,7,2) FROM WORKDAY WHERE WD01>" & stWYM & "00 AND WD01<" & stWYM & "99"
   
On Error GoTo ErrHnd

   CheckOC3
   With AdoRecordSet3
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount > 0 Then
         SetGrid1 AdoRecordSet3
      Else
         MsgBox "無法讀取工作天！"
         Exit Function
      End If
   End With
   
   'Modify by Morgan 2009/7/31 國外部判斷部門為 F 字頭( 原來抓 F41 ),每日業績點數一樣輸 F4100
   
   '抓智權人員名單+目標點數
   'Modify by Morgan 2007/4/2 不再控制員工是否在職否則過去資料將無法正確表示
   'Modify by Morgan 2007/10/18 加P29(巨京)
   '2009/12/9 MODIFY BY SONIA 國外部目標原只輸F4100,改為分F4101(FCL),F4102(FCP),F4103(FCT)
   'strSQL = " SELECT DISTINCT ST02 C1,0 C2,1 C3,ST15 C4,ST01 C5" & _
      " FROM DAILYFEAT, STAFF, ACC090 WHERE DF02>" & stTYM & "00 AND DF02 <" & stTYM & "99 AND DF01<>'F4100' AND ST01(+)=DF01" & _
      " AND A0901(+)=ST15" & _
      " Union" & _
      " SELECT MAX(DECODE(ST15,'F41','國外','M01','其他','P29','巨京',A0902)) C1,NVL(SUM(PE04),0) C2" & _
      ",DECODE( ST15,'F41',5,'M01',2,'P29',4,1) C3,ST15 C4,'X' C5" & _
      " From PERFORMANCE, STAFF, ACC090 WHERE PE02='TOT' AND PE03=" & stWYM & " AND ST01(+)=PE01 AND A0901(+)=ST15" & _
      " GROUP BY ST15"
   '2015/4/29 MODIFY BY SONIA 智權人員名單抓有輸入點數或有財務點數的,例D104012494李羽宸84045
   'strSql = " SELECT DISTINCT ST02 C1,0 C2,1 C3,ST15 C4,ST01 C5" & _
      " FROM DAILYFEAT, STAFF, ACC090 WHERE DF02>" & stTYM & "00 AND DF02 <" & stTYM & "99 AND DF01<>'F4100' AND ST01(+)=DF01" & _
      " AND A0901(+)=ST15" & _
      " Union" & _
      " SELECT MAX(DECODE(ST15,'F41','國外','F11','國外','F21','國外','M01','其他','P29','巨京',A0902)) C1,NVL(SUM(PE04),0) C2" & _
      ",DECODE( ST15,'F41',5,'F11',5,'F21',5,'M01',2,'P29',4,1) C3,DECODE(ST15,'F11','F41','F21','F41',ST15) C4,'X' C5" & _
      " From PERFORMANCE, STAFF, ACC090 WHERE PE02='TOT' AND PE03=" & stWYM & " AND ST01(+)=PE01 AND A0901(+)=ST15" & _
      " GROUP BY DECODE(ST15,'F11','F41','F21','F41',ST15),DECODE( ST15,'F41',5,'F11',5,'F21',5,'M01',2,'P29',4,1)"
   'Modify by Amy 2018/11/28 10801開始國外部分為FCP及FCT(原:DF01<>'F4100'),抓人員語法往上搬
   strSql = " SELECT DISTINCT ST02 C1,0 C2,1 C3,ST15 C4,ST01 C5" & _
      " FROM DAILYFEAT, STAFF, ACC090 WHERE DF02>" & stTYM & "00 AND DF02 <" & stTYM & "99 AND SubStr(DF01,1,3)<>'F41' AND ST01(+)=DF01" & _
      " AND A0901(+)=ST15"
  '抓人員語
   strSql = strSql & " Union" & stVTable0
   'end 201/11/28
    
   'Modify by Amy 2018/11/28 巨京沒目標另外拆出,10801開始國外部分為FCP及FCT拆出
'   strSql = strSql & " Union" & _
'      " SELECT MAX(DECODE(ST15,'F41','國外','F11','國外','F21','國外','M01','其他','P29','巨京',A0902)) C1,NVL(SUM(PE04),0) C2" & _
'      ",DECODE( ST15,'F41',5,'F11',5,'F21',5,'M01',2,'P29',4,1) C3,DECODE(ST15,'F11','F41','F21','F41',ST15) C4,'X' C5" & _
'      " From PERFORMANCE, STAFF, ACC090 WHERE PE02='TOT' AND PE03=" & stWYM & " AND ST01(+)=PE01 AND A0901(+)=ST15" & _
'      " GROUP BY DECODE(ST15,'F11','F41','F21','F41',ST15),DECODE( ST15,'F41',5,'F11',5,'F21',5,'M01',2,'P29',4,1)"
   'Modify by Amy 2019/07/15 因加客服組將巨京併入其他顯示,故拿掉P29
   strSql = strSql & " Union" & _
      " SELECT MAX(DECODE(ST15,'W10','客服組','M01','其他',A0902)) C1,NVL(SUM(PE04),0) C2" & _
      ",DECODE( ST15,'M01',2,1) C3,ST15 C4,'X' C5" & _
      " From PERFORMANCE, STAFF, ACC090 WHERE PE02='TOT' AND PE03=" & stWYM & " AND ST01(+)=PE01 AND A0901(+)=ST15" & _
      " And ST15 Not in ('F41','F11','F21')" & _
      " GROUP BY ST15,DECODE( ST15,'M01',2,1)"
  '國外部
  If Val(stTYM) >= Val(Left(每日業務點數FCPFCT啟用日, 5)) Then
    'modify by sonia 2021/1/14 +F4102加F4104及F4105,F4103加F4106及F4107
    strSql = strSql & " Union" & _
      " SELECT DECODE(PE01,'F4102','ＦＣＰ','F4104','ＦＣＰ','F4105','ＦＣＰ','ＦＣＴ') C1,NVL(SUM(PE04),0) C2" & _
      ",5 C3,ST15 C4,'X' C5" & _
      " From PERFORMANCE, STAFF, ACC090 WHERE PE02='TOT' AND PE03=" & stWYM & " AND ST01(+)=PE01 AND A0901(+)=ST15" & _
      " And PE01 in ('F4102','F4103','F4104','F4105','F4106','F4107')" & _
      " GROUP BY DECODE(PE01,'F4102','ＦＣＰ','F4104','ＦＣＰ','F4105','ＦＣＰ','ＦＣＴ'),ST15 "
  Else
    strSql = strSql & " Union" & _
      " SELECT '國外' C1,NVL(SUM(PE04),0) C2,5 C3,'F41' C4,'X' C5" & _
      " From PERFORMANCE, STAFF, ACC090 WHERE PE02='TOT' AND PE03=" & stWYM & " AND ST01(+)=PE01 AND A0901(+)=ST15" & _
      " And ST15 in ('F41','F11','F21')"
  End If
  'Mark by Amy 2019/0715 加客服組並將巨京併入其他顯示
'   strSql = strSql & " Union" & _
'      " SELECT '巨京' C1,0 C2,4 C3,'P29' C4,'X' C5 From Dual"
   'end 2018/11/28
   '2015/4/29 END
      
   
   strSql = strSql & " Union" & _
      " SELECT '北所' C1,NVL(SUM(PE04),0) C2,1 C3,'S1A' C4,'X' C5" & _
      " From PERFORMANCE, STAFF WHERE PE02='TOT' AND PE03=" & stWYM & " AND ST01(+)=PE01" & _
      " AND substr(ST15,1,2)='S1'"
      
   strSql = strSql & " Union" & _
      " SELECT '中所' C1,NVL(SUM(PE04),0) C2,1 C3,'S2A' C4,'X' C5" & _
      " From PERFORMANCE, STAFF WHERE PE02='TOT' AND PE03=" & stWYM & " AND ST01(+)=PE01" & _
      " AND substr(ST15,1,2)='S2'"
      
   'Modify by Morgan 2007/10/18 排除P29(巨京)
   strSql = strSql & " Union" & _
      " SELECT '國內' C1,NVL(SUM(PE04),0) C2,3 C3,'A' C4,'X' C5" & _
      " From PERFORMANCE, STAFF WHERE PE02='TOT' AND PE03=" & stWYM & " AND ST01(+)=PE01 AND SUBSTR(ST15,1,1)<>'F' AND ST15<>'P29'"
      
   strSql = strSql & " Union" & _
      " SELECT '全所' C1,NVL(SUM(PE04),0) C2,6 C3,'A' C4,'X' C5" & _
      " From PERFORMANCE WHERE PE02='TOT' AND PE03=" & stWYM
      
   strSql = strSql & " ORDER BY C3,C4,C5"
   
On Error GoTo ErrHnd

   CheckOC3
   With AdoRecordSet3
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount > 0 Then
         SetGrid2 AdoRecordSet3
      Else
         MsgBox "無法讀取智權人員名單！"
         Exit Function
      End If
   End With
   
   '其他收款：抓已入帳資料
   'Modify by Morgan 2007/10/18 排除P29(巨京)
   'Modified by Morgan 2011/12/29 改條件 SUBSTR(ST15,1,1)<>'S'->INSTR('S11,S13,S14,S15,S21,S22,S23,S29,S31,S41',ST15)=0 因 S10 也會有但算其他
   'modify by sonia 2014/1/21 取消a0201='1'條件
   'modify by sonia 2015/4/22 加不含'4194'科目,不含結餘傳票不要加科目限制
   'stVTable1 = "select a0205 V1C0,ROUND(sum(ax207)/1000,2) V1C1" & _
      " From acc020, acc021,STAFF" & _
      " Where A0205>" & stTYM & "00 AND A0205<=" & stTYM & "99" & _
      " and ax201(+) = a0201  and ax202(+) = a0202" & _
      " and ax209 Is Not Null and (substr(ax205, 1, 2) = '41' Or ax205 = '7121')" & _
      " and ax207>0 and not( ax205='4191' or ax205='4192'" & _
      " or ((ax205='410103' or ax205='411103' or ax205='4121' or ax205='4131') and instr(ax213||' ','結餘')>0))" & _
      " AND ST01=AX209 AND INSTR('S11,S13,S14,S15,S21,S22,S23,S29,S31,S41',ST15)=0 AND SUBSTR(ST15,1,1)<>'F' AND ST15<>'P29'" & _
      " group by a0205"
   'Modify by Amy 2018/11/28 原:INSTR('S11,S13,S14,S15,S21,S22,S23,S29,S31,S41',ST15)=0 ,後來加S24會被加進去
   'Modify by Amy 2018/07/15 原:AND ST15<>'P29' 因加客服組,將巨京併於其他顯示
   'Modify by Amy 2019/08/01 增加創新業務組用收入 420101 原:substr(ax205, 1, 2) = '41'
   stVTable1 = "select a0205 V1C0,ROUND(sum(ax207)/1000,2) V1C1" & _
      " From acc020, acc021,STAFF" & _
      " Where A0205>" & stTYM & "00 AND A0205<=" & stTYM & "99" & _
      " and ax201(+) = a0201  and ax202(+) = a0202" & _
      " and ax209 Is Not Null and (substr(ax205, 1, 1) = '4' Or ax205 = '7121')" & _
      " and ax207>0 and not( ax205='4191' or ax205='4192' or ax205='4194' or instr(ax213||' ','結餘')>0)" & _
      " AND ST01=AX209 AND SUBSTR(ST15,1,1)<>'S' AND SUBSTR(ST15,1,1)<>'F' AND ST15<>'W10'" & _
      " group by a0205"
   
   '其他扣點數資料 6/1以後才要
   'Modify by Morgan 2007/10/18 排除P29(巨京)
   'Modified by Morgan 2011/12/29 改條件 SUBSTR(ST15,1,1)<>'S'->INSTR('S11,S13,S14,S15,S21,S22,S23,S29,S31,S41',ST15)=0 因 S10 也會有但算其他
   'modify by sonia 2014/1/21 取消a0201='1'條件
   'modify by sonia 2015/4/22 加不含'4194'科目,不含結餘傳票不要加科目限制,71科目只抓7121故substr(ax205,1,2)='71'改為ax205 = '7121'
   'stVTable2 = "select a0205 V2C0,ROUND(sum(ax206)/1000,2) V2C1" & _
      " from acc020, acc021,STAFF" & _
      " where ax201(+) = a0201  and ax202(+) = a0202" & _
      " AND A0205>=940601 AND A0205>" & stTYM & "00 AND A0205<=" & stTYM & "99" & _
      " and ax209 is not null and ax206>0 and (substr(ax205, 1, 2) = '41' or substr(ax205,1,2)='71')" & _
      " and not (  (ax205='4191' or ax205='4192')" & _
      " or ( (ax205='410103' or ax205='411103' or ax205='4121' or ax205='4131') and instr(ax213||' ','結餘')>0 ) )" & _
      " AND ST01=AX209 AND INSTR('S11,S13,S14,S15,S21,S22,S23,S29,S31,S41',ST15)=0 AND SUBSTR(ST15,1,1)<>'F' AND ST15<>'P29'" & _
      " group by a0205"
   'Modify by Amy 2018/07/15 原:AND ST15<>'P29' 因加客服組,將巨京併於其他顯示
   'Modify by Amy 2019/08/01 增加創新業務組用收入 420101 原:substr(ax205, 1, 2) = '41'
   stVTable2 = "select a0205 V2C0,ROUND(sum(ax206)/1000,2) V2C1" & _
      " from acc020, acc021,STAFF" & _
      " where ax201(+) = a0201  and ax202(+) = a0202" & _
      " AND A0205>=940601 AND A0205>" & stTYM & "00 AND A0205<=" & stTYM & "99" & _
      " and ax209 is not null and ax206>0 and (substr(ax205, 1, 1) = '4' or ax205 = '7121')" & _
      " and not (ax205='4191' or ax205='4192' or ax205='4194' or instr(ax213||' ','結餘')>0)" & _
      " AND ST01=AX209 AND SUBSTR(ST15,1,1)<>'S' AND SUBSTR(ST15,1,1)<>'F' AND ST15<>'W10'" & _
      " group by a0205"
   'end 2018/11/28
   
   'Add by Morgan 2007/10/18
   '巨京收款：抓已入帳資料
   'modify by sonia 2014/1/21 取消a0201='1'條件
   'modify by sonia 2015/4/22 加不含'4194'科目,不含結餘傳票不要加科目限制
   'stVTable3 = "select a0205 V1C0,ROUND(sum(ax207)/1000,2) V1C1" & _
      " From acc020, acc021,STAFF" & _
      " Where A0205>" & stTYM & "00 AND A0205<=" & stTYM & "99" & _
      " and ax201(+) = a0201  and ax202(+) = a0202" & _
      " and ax209 Is Not Null and (substr(ax205, 1, 2) = '41' Or ax205 = '7121')" & _
      " and ax207>0 and not( ax205='4191' or ax205='4192'" & _
      " or ((ax205='410103' or ax205='411103' or ax205='4121' or ax205='4131') and instr(ax213||' ','結餘')>0))" & _
      " AND ST01=AX209 AND ST15='P29'" & _
      " group by a0205"
   'Modify by Amy 2019/07/15 原巨京改併入其他,此顯示「客服組」
   'Modify by Amy 2019/08/01 增加創新業務組用收入 420101 原:substr(ax205, 1, 2) = '41'
   stVTable3 = "select a0205 V1C0,ROUND(sum(ax207)/1000,2) V1C1" & _
      " From acc020, acc021,STAFF" & _
      " Where A0205>" & stTYM & "00 AND A0205<=" & stTYM & "99" & _
      " and ax201(+) = a0201  and ax202(+) = a0202" & _
      " and ax209 Is Not Null and (substr(ax205, 1, 1) = '4' Or ax205 = '7121')" & _
      " and ax207>0 and not( ax205='4191' or ax205='4192' or ax205='4194' or instr(ax213||' ','結餘')>0)" & _
      " AND ST01=AX209 AND ST15='W10'" & _
      " group by a0205"
   
   '巨京扣點數資料
   'modify by sonia 2014/1/21 取消a0201='1'條件
   'modify by sonia 2015/4/22 加不含'4194'科目,不含結餘傳票不要加科目限制,71科目只抓7121故substr(ax205,1,2)='71'改為ax205 = '7121'
   'stVTable4 = "select a0205 V2C0,ROUND(sum(ax206)/1000,2) V2C1" & _
      " from acc020, acc021,STAFF" & _
      " where ax201(+) = a0201  and ax202(+) = a0202" & _
      " AND A0205>=940601 AND A0205>" & stTYM & "00 AND A0205<=" & stTYM & "99" & _
      " and ax209 is not null and ax206>0 and (substr(ax205, 1, 2) = '41' or substr(ax205,1,2)='71')" & _
      " and not (  (ax205='4191' or ax205='4192')" & _
      " or ( (ax205='410103' or ax205='411103' or ax205='4121' or ax205='4131') and instr(ax213||' ','結餘')>0 ) )" & _
      " AND ST01=AX209 AND ST15='P29'" & _
      " group by a0205"
   'Modify by Amy 2019/07/15 原巨京改併入其他,此顯示「客服組」
   'Modify by Amy 2019/08/01 增加創新業務組用收入 420101 原:substr(ax205, 1, 2) = '41'
   stVTable4 = "select a0205 V2C0,ROUND(sum(ax206)/1000,2) V2C1" & _
      " from acc020, acc021,STAFF" & _
      " where ax201(+) = a0201  and ax202(+) = a0202" & _
      " AND A0205>=940601 AND A0205>" & stTYM & "00 AND A0205<=" & stTYM & "99" & _
      " and ax209 is not null and ax206>0 and (substr(ax205, 1, 1) = '4' or ax205 = '7121')" & _
      " and not (ax205='4191' or ax205='4192' or ax205='4194' or instr(ax213||' ','結餘')>0)" & _
      " AND ST01=AX209 AND ST15='W10'" & _
      " group by a0205"
      
   'end 2007/10/18
      
   '扣點數資料 6/1以後才要
   'modify by sonia 2014/1/21 取消a0201='1'條件
   'modify by sonia 2015/4/22 加不含'4194'科目,不含結餘傳票不要加科目限制,71科目只抓7121故substr(ax205,1,2)='71'改為ax205 = '7121'
   'stVTable5 = "select DECODE(SUBSTR(ax209,1,3),'F41','F4100',AX209) V5C0,a0205 V5C1,ROUND(sum(ax206)/1000,2) V5C2" & _
      " from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202" & _
      " AND A0205>=940601 AND A0205>" & stTYM & "00 AND A0205<=" & stTYM & "99" & _
      " and ax209 is not null and ax206>0 and (substr(ax205, 1, 2) = '41' or substr(ax205,1,2)='71')" & _
      " and not (  (ax205='4191' or ax205='4192')" & _
      " or ( (ax205='410103' or ax205='411103' or ax205='4121' or ax205='4131') and instr(ax213||' ','結餘')>0 ) )" & _
      " group by a0205,DECODE(SUBSTR(ax209,1,3),'F41','F4100',AX209)"
   'Modfiy by Amy 2018/11/28 10801開始國外部分為FCP及FCT拆出
   stField = "DECODE(SUBSTR(ax209,1,3),'F41','F4100',AX209)"
    If Val(stTYM) >= Val(Left(每日業務點數FCPFCT啟用日, 5)) Then
       stField = "ax209"
   End If
   'Modify by Amy 2019/08/01 增加創新業務組用收入 420101 原:substr(ax205, 1, 2) = '41'
   stVTable5 = "select " & stField & " V5C0,a0205 V5C1,ROUND(sum(ax206)/1000,2) V5C2" & _
      " from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202" & _
      " AND A0205>=940601 AND A0205>" & stTYM & "00 AND A0205<=" & stTYM & "99" & _
      " and ax209 is not null and ax206>0 and (substr(ax205, 1, 1) = '4' or ax205 = '7121')" & _
      " and not (ax205='4191' or ax205='4192' or ax205='4194' or instr(ax213||' ','結餘')>0) " & _
      " group by a0205," & stField
   
   '智權部+國外部
   '2015/4/29 MODIFY BY SONIA 智權人員名單抓有輸入點數或有財務點數的,例D104012494李羽宸84045
   'Modify by Amy 2018/11/28 10801開始國外部分為FCP及FCT拆出
   stField = ",decode(ST15,'F41','國外','F11','國外','F21','國外',ST02) 智權人員" & _
                   ",NVL(DF03,0)-NVL(V5C2,0) 點數, DECODE( ST15,'F41',5,'F11',5,'F21',5,1) 排序1"
   If Val(stTYM) >= Val(Left(每日業務點數FCPFCT啟用日, 5)) Then
       stField = ",ST02 智權人員" & _
                   ",NVL(DF03,0)-NVL(V5C2,0) 點數, DECODE(SubStr(ST01,1,3),'F41',5,1) 排序1"
   End If
   'modify by sonia 2021/1/19 110/1起F4102改用F4104及F4105,F4103改用F4106及F4107
   'strSql = " SELECT DF02 日期" & stField & ",ST15,ST01" & _
      " FROM DAILYFEAT, STAFF, ACC090,(" & stVTable5 & ") VT5" & _
      " WHERE DF02>" & stTYM & "00 AND DF02 <" & stTYM & "99" & _
      " AND ST01(+)=DF01 AND A0901(+)=ST15 AND V5C1(+)=DF02 AND V5C0(+)=DF01"
   strSql = " SELECT DF02 日期" & stField & ",ST15,ST01" & _
      " FROM STAFF, ACC090, " & _
      "(SELECT DF02,decode(DF01,'F4104','F4102','F4105','F4102','F4106','F4103','F4107','F4103',DF01) DF01,SUM(DF03) DF03 FROM DAILYFEAT " & _
      " GROUP BY DF02,decode(DF01,'F4104','F4102','F4105','F4102','F4106','F4103','F4107','F4103',DF01)),(" & stVTable5 & ") VT5" & _
      " WHERE DF02>" & stTYM & "00 AND DF02 <" & stTYM & "99" & _
      " AND ST01(+)=DF01 AND A0901(+)=ST15 AND V5C1(+)=DF02 AND V5C0(+)=DF01"
   'end 2021/1/19
   
   'strSql = " SELECT WD01-19110000 日期,decode(ST15,'F41','國外','F11','國外','F21','國外',ST02) 智權人員" & _
      ",NVL(DF03,0)-NVL(V5C2,0) 點數, DECODE( ST15,'F41',5,'F11',5,'F21',5,1) 排序1,ST15,ST01" & _
      " FROM (SELECT WD01,ST01,ST02,ST15,NVL(DF03,0) DF03 FROM WORKDAY,STAFF WHERE WD01 >" & stWYM & "00 AND WD01 <" & stWYM & "99 and WD01 <=" & strSrvDate(1) & "), " & _
      " (SELECT DF02,DF01,DF03 FROM DAILYFEAT WHERE DF02>" & stTYM & "00 AND DF02<" & stTYM & "99), " & _
      " (" & stVTable5 & ") VT5" & _
      " WHERE ST01=V5C0(+) AND WD01-19110000=V5C1(+) "
   
   '其他
   'Modfiy by Amy 2019/07/15 原:2 排序1, SetGrid3 才會抓到正確人員
   strSql = strSql & " UNION SELECT WD01-19110000 日期,'其他' 智權人員,NVL(V1C1,0)-NVL(V2C1,0) 點數, 3 排序1,NULL,NULL" & _
      " FROM WORKDAY,(" & stVTable1 & ") VT1,(" & stVTable2 & ") VT2" & _
      " WHERE WD01 >" & stWYM & "00 AND WD01 <" & stWYM & "99 and WD01<=" & strSrvDate(1) & _
      " AND V1C0(+)=WD01-19110000 AND V2C0(+)=WD01-19110000"
   
   '巨京->客服組
   'Modify by Amy 2019/07/15 增加客服組,巨京併入其他顯示,排序1設為2
   strSql = strSql & " UNION SELECT WD01-19110000 日期,'客服組' 智權人員,NVL(V1C1,0)-NVL(V2C1,0) 點數, 2 排序1,NULL,NULL" & _
      " FROM WORKDAY,(" & stVTable3 & ") VT1,(" & stVTable4 & ") VT2" & _
      " WHERE WD01 >" & stWYM & "00 AND WD01 <" & stWYM & "99 and WD01<=" & strSrvDate(1) & _
      " AND V1C0(+)=WD01-19110000 AND V2C0(+)=WD01-19110000"
      
   strSql = strSql & " ORDER BY 1,4,5,6"
      
On Error GoTo ErrHnd

   CheckOC3
   With AdoRecordSet3
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount > 0 Then
         SetGrid3 AdoRecordSet3
         txtCloseDate.Tag = txtCloseDate
      Else
         MsgBox "無法讀取收款資料！"
         Exit Function
      End If
   End With
   
   SetGrid4
   
   doQuery = True
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Function

Private Sub cmdExit_Click()
 Unload Me
End Sub

Private Sub cmdPrint_Click()
If txtCloseDate.Tag = "" Then
      MsgBox "請先查詢後再按列印！"
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   If DoPrint = True Then
      MsgBox "列印完成", vbInformation
   End If
   Screen.MousePointer = vbDefault
End Sub

'畫格線
'Modify by Amy 2018/11/28 +bolTitleOnly
Private Sub printTable(p_iPage As Integer, Optional ByVal bolTitleOnly As Boolean = False)
   Dim iXp As Integer, iYp As Integer '前次列印起算位置
   Dim iRow As Integer, iCol As Integer
         
   'Modfiy by Amy 2018/11/28
   If bolTitleOnly = False Then
        '第1,2條橫線
        iX = iX0: iY = iY0
        Printer.Line (iX, iY)-(iX1, iY)
        iY = iY + iFixRowHeight
        Printer.Line (iX, iY)-(iX1, iY)
        '第2以後橫線
        For iRow = 1 To 46
           iY = iY + iRowHeight
           If iRow Mod 2 = 0 Then
              iX = iX0
           Else
              iX = iX0 + iFixColWidth
           End If
           Printer.Line (iX, iY)-(iX1, iY)
        Next
        
        '第1,2條豎線
        iX = iX0: iY = iY0
        Printer.Line (iX, iY)-(iX, iY1)
        iX = iX + iFixColWidth
        Printer.Line (iX, iY)-(iX, iY1)
        '第2以後豎線
        For iCol = 1 To iPageCols
           iX = iX + iColWidth
           Printer.Line (iX, iY)-(iX, iY1)
        Next
        '填日期
        With grdDataList
           .Visible = False
           iXp = iX0: iYp = iY0 + iFixRowHeight
           Printer.FontSize = 8 'Modify by Amy 2018/12/24 原:12 1080100後國外部分兩欄會多一頁,不要多一頁-秀玲
           For iRow = 1 To .Rows - 1 Step 2
              strContent = .TextMatrix(iRow, 0)
              '置中
              iX = iXp + (iFixColWidth / 2 - Printer.TextWidth(strContent) / 2)
              iY = iYp + (iRowHeight - Printer.TextHeight(strContent) / 2)
              Printer.CurrentX = iX: Printer.CurrentY = iY
              Printer.Print strContent
              iYp = iYp + 2 * iRowHeight
           Next
           .Visible = True
        End With
   End If
   'end 2018/11/28
   
   '頁首
   strContent = Format(Left(stWYM, 4) - 1911) & "年" & Val(Mid(stWYM, 5)) & "月業績達成月報表"
   Printer.FontSize = 16
   iX = iX0 + iX1 / 2 - Printer.TextWidth(strContent) / 2
   iY = iY0 - Printer.TextHeight(strContent) - 50
   Printer.CurrentX = iX: Printer.CurrentY = iY
   Printer.Print strContent
   
   '頁尾
   strContent = "第 " & p_iPage & " 頁"
   Printer.FontSize = 9
   iX = iX0 + iX1 / 2 - Printer.TextWidth(strContent) / 2
   iY = iY1 + 50
   'Add by Amy 2018/11/28 只印頁首頁尾
   If bolTitleOnly = True Then
        iX = 7831
        iY = 10760
   End If
   Printer.CurrentX = iX: Printer.CurrentY = iY
   Printer.Print strContent
   
End Sub

Private Function DoPrint() As Boolean
   Dim iRow As Integer, iCol As Integer, iPage As Integer
   Dim iXp As Integer, iYp As Integer '前次列印起算位置
   Dim iPageCol As Integer '相對欄位

On Error GoTo ErrHnd
   
   Printer.Orientation = 2
   Printer.FontName = "標楷體"
   Printer.DrawStyle = vbSolid
   Printer.DrawWidth = 2
   iPageCols = 19 'Modify by Amy 2018/12/24 原:18 1080100後國外部分兩欄會多一頁,不要多一頁-秀玲
   iFixRowHeight = 550: iRowHeight = 210
   iFixColWidth = 200 'Modify by Amy 2018/12/24 原:400 1080100後國外部分兩欄會多一頁,不要多一頁-秀玲
   iColWidth = 850
   iX0 = 200: iY0 = 500
   iX1 = iX0 + iFixColWidth + iPageCols * iColWidth
   iY1 = iY0 + iFixRowHeight + 2 * 23 * iRowHeight
   
   With grdDataList
      .Visible = False
      iPage = 1
      printTable iPage
      iPageCol = 1
      iXp = iX0 + iFixColWidth '起始X點
      For iCol = 1 To .Cols - 1
         If iPageCol > iPageCols Then
            Printer.NewPage
            iPage = iPage + 1
            printTable iPage
            iPageCol = 1
            iXp = iX0 + iFixColWidth
         End If
         iRow = 0 '固定列
         Printer.FontSize = 12
         iYp = iY0
         '印兩列
         'Modify by Amy 2018/12/24 中區其他 字會壓線
         If Right(.TextMatrix(iRow, iCol), 2) = "應收" Or .TextMatrix(iRow, iCol) = "中區其他" Then
            If .TextMatrix(iRow, iCol) = "中區其他" Then
                strContent = "中區"
            Else
                iContent = Len(.TextMatrix(iRow, iCol))
                strContent = Left(.TextMatrix(iRow, iCol), iContent - 2)
            End If
            iX = iXp + (iColWidth / 2 - Printer.TextWidth(strContent) / 2)
            iY = iYp + (iFixRowHeight / 4 - Printer.TextHeight(strContent) / 2)
            Printer.CurrentX = iX: Printer.CurrentY = iY
            Printer.Print strContent
            If .TextMatrix(iRow, iCol) = "中區其他" Then
                strContent = "其他"
            Else
                strContent = "應收"
            End If
            iX = iXp + (iColWidth / 2 - Printer.TextWidth(strContent) / 2)
            iY = iYp + (iFixRowHeight * 3 / 4 - Printer.TextHeight(strContent) / 2)
            Printer.CurrentX = iX: Printer.CurrentY = iY
            Printer.Print strContent
            '反白
         'end 2018/12/24
         
         '印一列
         Else
            strContent = .TextMatrix(iRow, iCol)
            iX = iXp + (iColWidth / 2 - Printer.TextWidth(strContent) / 2)
            iY = iYp + (iFixRowHeight / 2 - Printer.TextHeight(strContent) / 2)
            Printer.CurrentX = iX: Printer.CurrentY = iY
            Printer.Print strContent
         End If
         
         Printer.FontSize = 9
         iYp = iY0 + iFixRowHeight '起始Y點
         For iRow = 1 To .Rows - 1
            strContent = .TextMatrix(iRow, iCol)
            '數字靠右
            iX = iXp + (iColWidth - Printer.TextWidth(strContent)) - 20
            iY = iYp + (iRowHeight / 2 - Printer.TextHeight(strContent) / 2)
            Printer.CurrentX = iX: Printer.CurrentY = iY
            Printer.Print strContent
            iYp = iYp + iRowHeight '下移一列
         Next
         iPageCol = iPageCol + 1
         iXp = iXp + iColWidth '右移一行
      Next
      .Visible = True
   End With
   Call PrintMemo(iPage) 'Add by Amy 2018/11/28
   Printer.EndDoc
   DoPrint = True
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   
End Function

'Add by Amy 2018/11/28
Private Sub PrintMemo(ByVal iPage As Integer)
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String, strField As String, strOldD As String, strData As String
    Dim iRow As Integer, intQ As Integer
 
    strQ = "Select Distinct df02,a0902,a0901 From DailyFeat,Staff,acc090 " & _
              "Where df02>" & stTYM & "00 And df02<=" & stTYM & "99 And df01=st01(+) And st15=a0901(+)" & _
              "And df06='QPGMR' And df09 is null And Substr(df01,1,3)<>'F41' " & _
              "Order by df02,a0901"
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        strContent = ""
        Printer.FontSize = 12
        Printer.NewPage
        iPage = iPage + 1
        printTable iPage, True
        
        iX = 200  '起始X點
        iY = 500
        Printer.CurrentX = iX: Printer.CurrentY = iY
        Printer.Print "每日未輸入業務點數如下："
        Do While Not RsQ.EOF
            If iRow > 23 Then
                Printer.NewPage
                iPage = iPage + 1
                printTable iPage
                iRow = 1
                iX = 200: iY = 500
            ElseIf strOldD <> RsQ.Fields("df02") And strOldD <> MsgText(601) Then
                iY = iY + 300 '下移一列
                Printer.CurrentX = iX: Printer.CurrentY = iY
                Printer.Print strOldD & "：" & Mid(strData, 2)
                strData = ""
            End If
            strData = strData & "," & RsQ.Fields("a0902")
            
            strOldD = "" & RsQ.Fields("df02")
            RsQ.MoveNext
        Loop
    End If
    
End Sub

Private Sub cmdSearch_Click()
   Screen.MousePointer = vbHourglass
   If ConstrainCheck = True Then
      Call doQuery
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   txtCloseDate = TransDate(CompWorkDay(2, strSrvDate(1), 1), 1) \ 100
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm210108 = Nothing
End Sub

Private Function ConstrainCheck() As Boolean
   ConstrainCheck = False
   '因為會寫一整個月資料,故控制不可查當月
   If Val(txtCloseDate) >= Val(Left(strSrvDate(1), 6)) - 191100 Then
      MsgBox "目前尚未開放可查" & Val(Left(strSrvDate(1), 6)) - 191100 & "月以後資料！"
      Exit Function
   End If
   '日期格式
   If ChkDate(txtCloseDate & "01") = False Then
      txtCloseDate.SetFocus
      txtCloseDate_GotFocus
      Exit Function
   End If
   ConstrainCheck = True
End Function

Private Sub txtCloseDate_Change()
   txtCloseDate.Tag = ""
End Sub

Private Sub txtCloseDate_GotFocus()
   TextInverse txtCloseDate
   'edit by nickc 2007/07/11 切換輸入法改用API
   'txtCloseDate.IMEMode = 2
   CloseIme
End Sub

 'Add by Amy 2018/11/28 新增當月未輸之每日業務點數 S部門及國外
Private Sub InsDailyFeat(ByVal stDate As String, ByVal stNow As String)
    Dim strIns As String, intI As Integer
    Dim strQ As String, strQ2 As String, strAll As String, strField As String
    
    'Memo 此有修改frm210107 也要修改
    'S部門(未輸入值新增0)
    strIns = "Insert Into DailyFeat (DF01,DF02,DF03,DF04,DF06,DF07,DF08) " & _
                     "Select c5," & txtCloseDate & stDate & ",0,0,'QPGMR'," & strSrvDate(1) & "," & stNow & " " & _
                     "From (" & stVTable0 & "),DailyFeat " & _
                     "Where C5=df01(+)  And c5 is not null And df01 is null And df02(+)=" & txtCloseDate & stDate
    cnnConnection.Execute strIns, intI
    
    '國外部(1080101後國外部改區分FCP及FCT)
    strQ = "SubStr(st01,1,3)||'00'"
    strQ2 = "And SubStr(ax209,1,4)='F410'"
    If Val(txtCloseDate) >= Val(Left(每日業務點數FCPFCT啟用日, 5)) Then
       strQ = "st01"
       strQ2 = "And ax209 in('F4102','F4103')"
       'add by sonia 2021/1/19 110/1起改用F4104~F4107
       If Val(txtCloseDate) >= 11001 Then
         strQ2 = "And ax209 in('F4104','F4105','F4106','F4107')"
       End If
       'end 2021/1/19
    End If
    
    '每日V1C1-扣點數V5C2(940601以後再扣)
    'Modify by Amy 2019/08/01 增加創新業務組用收入 420101 原:substr(ax205, 1, 2) = '41'
    strAll = "Where a0205= " & txtCloseDate & stDate & " And ax201(+) = a0201  And ax202(+) = a0202 " & _
              "And ax209 Is Not Null And (substr(ax205, 1, 1) = '4' Or ax205 = '7121') " & _
              "And not( ax205='4191' or ax205='4192' or ax205='4194' or instr(ax213||' ','結餘')>0) "
    
    'strQ = "Select " & strQ & " as st01,TO_CHAR(ROUND((sum(nvl(V1C1,0))-sum(nvl(V5C2,0)))/1000,2),'9999999.00') V1 From Staff," & _
              "(Select ax209,sum(nvl(ax207,0)) V1C1 From Acc020, Acc021 " & strAll & strQ2 & " Group by ax209) VT1, " & _
              "(Select ax209,sum(nvl(ax206,0)) V5C2 From Acc020, Acc021 " & strAll & strQ2 & " And a0205>=940601 Group by ax209) VT5 " & _
              "Where st01=VT1.ax209(+) And st01=VT5.ax209(+) " & Replace(strQ2, "ax209", "st01") & " Group by " & strQ
              
  strQ = "Select " & strQ & " as st01,TO_CHAR(ROUND((sum(nvl(V1C1,0)))/1000,2),'9999999.00') V1 From Staff," & _
              "(Select ax209,sum(nvl(ax207,0)) V1C1 From Acc020, Acc021 " & strAll & strQ2 & " Group by ax209) VT1 " & _
              "Where st01=VT1.ax209(+) " & Replace(strQ2, "ax209", "st01") & " Group by " & strQ

   strIns = "Insert Into DailyFeat (DF01,DF02,DF03,DF04,DF06,DF07,DF08) " & _
                "Select st01," & txtCloseDate & stDate & ",V1,V1,'QPGMR'," & strSrvDate(1) & "," & stNow & " " & _
                "From (" & strQ & "),DailyFeat " & _
                "Where st01=df01(+) And df01 is null And df02(+)=" & txtCloseDate & stDate
    cnnConnection.Execute strIns, intI
End Sub
