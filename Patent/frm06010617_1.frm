VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm06010617_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "信件沖銷記錄統計"
   ClientHeight    =   5730
   ClientLeft      =   5450
   ClientTop       =   3400
   ClientWidth     =   8950
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   8950
   Begin VB.CommandButton cmdExcel 
      Caption         =   "Excel(&E)"
      Height          =   375
      Left            =   7320
      TabIndex        =   10
      Top             =   4560
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   5490
      Style           =   2  '單純下拉式
      TabIndex        =   5
      Top             =   5430
      Width           =   3450
   End
   Begin VB.CommandButton cmdPrinter 
      Caption         =   "列印(&P)"
      Height          =   375
      Left            =   7320
      TabIndex        =   6
      Top             =   5010
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   6510
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   10
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   7740
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   10
      Width           =   756
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList1 
      Height          =   4725
      Left            =   30
      TabIndex        =   0
      Top             =   660
      Width           =   6900
      _ExtentX        =   12171
      _ExtentY        =   8343
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '透明
      Caption         =   "印表機："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4590
      TabIndex        =   9
      Top             =   5460
      Width           =   855
   End
   Begin VB.Label lbl1 
      Caption         =   " "
      Height          =   180
      Index           =   4
      Left            =   7005
      TabIndex        =   8
      Top             =   3225
      Width           =   1695
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "總計："
      Height          =   180
      Left            =   7005
      TabIndex        =   7
      Top             =   3015
      Width           =   540
   End
   Begin VB.Label lbl1 
      Caption         =   " "
      Height          =   180
      Index           =   0
      Left            =   990
      TabIndex        =   2
      Top             =   435
      Width           =   7920
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "查詢期間："
      Height          =   180
      Left            =   60
      TabIndex        =   1
      Top             =   435
      Width           =   900
   End
End
Attribute VB_Name = "frm06010617_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create By Sindy 2024/4/25
Option Explicit

Dim strSql As String
Dim i As Integer, j As Long
Dim PLeft(0 To 8) As Integer, iPrint As Integer, Page As Integer, strTemp3(0 To 10) As String
Public cmdState As Integer '紀錄作用按鍵
Dim strPrinter As String
Dim xlsAnnuity As New Excel.Application
Dim wksAnnuity As New Worksheet
Dim intCounter As Integer


Private Sub SetDataListWidth()
   If frm06010617.txt1(6) = "1" Then '1.承辦同仁
      grdDataList1.Cols = 6
      grdDataList1.row = 0
      grdDataList1.col = 0: grdDataList1.Text = "部門"
      grdDataList1.ColWidth(0) = 1200
      grdDataList1.CellAlignment = flexAlignLeftCenter
      grdDataList1.col = 1: grdDataList1.Text = "統計條件"
      grdDataList1.ColWidth(1) = 1000
      grdDataList1.CellAlignment = flexAlignLeftCenter
      grdDataList1.col = 2: grdDataList1.Text = "沖銷類別"
      grdDataList1.ColWidth(2) = 2500
      grdDataList1.CellAlignment = flexAlignLeftCenter
      grdDataList1.col = 3: grdDataList1.Text = "數量"
      grdDataList1.ColWidth(3) = 800
      grdDataList1.CellAlignment = flexAlignRightCenter
      grdDataList1.col = 4: grdDataList1.Text = "排序"
      grdDataList1.ColWidth(4) = 0
      grdDataList1.CellAlignment = flexAlignRightCenter
      grdDataList1.col = 5: grdDataList1.Text = "ST03"
      grdDataList1.ColWidth(5) = 0
   Else
      grdDataList1.Cols = 3
      grdDataList1.row = 0
      grdDataList1.col = 0: grdDataList1.Text = "部門"
      grdDataList1.ColWidth(0) = 1000
      grdDataList1.Text = "代碼"
      grdDataList1.CellAlignment = flexAlignLeftCenter
      grdDataList1.col = 1: grdDataList1.Text = "統計條件"
      grdDataList1.ColWidth(1) = 3000
      grdDataList1.CellAlignment = flexAlignLeftCenter 'flexAlignCenterCenter
      grdDataList1.col = 2: grdDataList1.Text = "數量"
      grdDataList1.ColWidth(2) = 800
      grdDataList1.CellAlignment = flexAlignRightCenter
   End If
End Sub

Public Sub PubShowNextData()
   Select Case cmdState
      Case 0
         tmpBol = fnCancelNowFormAndShowParentForm(Me)
      Case 1
         Unload frm06010617
         Unload Me
      Case Else
   End Select
End Sub

Private Sub SetExcelWorksheets()
   xlsAnnuity.Visible = True
   xlsAnnuity.SheetsInNewWorkbook = 1 '預設工作表數量
   xlsAnnuity.Workbooks.add
   Set wksAnnuity = xlsAnnuity.Worksheets(1)
   xlsAnnuity.ActiveWindow.Zoom = 75 '畫面比例100%太大了,調整為75%
   'wksAnnuity.PageSetup.Orientation = xlLandscape '橫印
   wksAnnuity.PageSetup.Orientation = wdOrientLandscape '直印
   wksAnnuity.PageSetup.LeftMargin = 28.34 'Application.InchesToPoints(0.393700787401575)
   wksAnnuity.PageSetup.RightMargin = 28.34 'Application.InchesToPoints(0.393700787401575)
   wksAnnuity.PageSetup.TopMargin = 42.51 'Application.InchesToPoints(0.590551181102362)
   wksAnnuity.PageSetup.BottomMargin = 42.51 'Application.InchesToPoints(0.590551181102362)
   wksAnnuity.PageSetup.HeaderMargin = 28.34 'Application.InchesToPoints(0.393700787401575)
   wksAnnuity.PageSetup.FooterMargin = 28.34 'Application.InchesToPoints(0.393700787401575)
   '設定各欄位長度
   wksAnnuity.Columns("A:A").ColumnWidth = 10
   wksAnnuity.Columns("B:B").ColumnWidth = 10
   wksAnnuity.Columns("C:C").ColumnWidth = 10
   wksAnnuity.Columns("D:D").ColumnWidth = 11
   'intCounter = 1 'intCounter + 1
End Sub

Private Sub cmdExcel_Click()
Dim intItem As Integer
   
   Set xlsAnnuity = New Excel.Application
   Call SetExcelWorksheets
   intCounter = 0
   With grdDataList1
      For intItem = 0 To grdDataList1.Rows - 1
         intCounter = intCounter + 1
         For i = 0 To grdDataList1.Cols - IIf(frm06010617.txt1(6) = "1", 2, 1)
            xlsAnnuity.Range(Chr(65 + i) & intCounter).Value = grdDataList1.TextMatrix(intItem, i)
         Next i
      Next intItem
      intCounter = intCounter + 1
      xlsAnnuity.Range(Chr(65 + (grdDataList1.Cols - IIf(frm06010617.txt1(6) = "1", 3, 2))) & intCounter).Value = "總計"
      xlsAnnuity.Range(Chr(65 + (grdDataList1.Cols - IIf(frm06010617.txt1(6) = "1", 2, 1))) & intCounter).Value = lbl1(4).Caption
   End With
   
   Set xlsAnnuity = Nothing
   Set wksAnnuity = Nothing
End Sub

Private Sub cmdOK_Click(Index As Integer)
   cmdState = Index
   PubShowNextData
End Sub

'列印    第二階段
Private Sub cmdPrinter_Click()
Dim intItem As Integer
   
   Screen.MousePointer = vbHourglass
   If grdDataList1.Rows <> 1 Then
      PUB_RestorePrinter Combo1
      Page = 1
      PrintTitle
      With grdDataList1
         For intItem = 1 To IIf(.Rows - 1 < 12, 12, .Rows - 1)   '右邊合計欄之行數要全印
            If iPrint >= 16000 Then
               Printer.NewPage
               Page = Page + 1
               PrintTitle
            End If
            If intItem > .Rows - 1 Then
               Erase strTemp3
            Else
               Erase strTemp3
               For i = 0 To .Cols - 1
                  strTemp3(i) = Me.grdDataList1.TextMatrix(intItem, i)
               Next i
            End If
            
            If Page = 1 And intItem = 1 Then
               Printer.CurrentX = 9000
               Printer.CurrentY = iPrint
               Printer.Print "總　計：" & lbl1(4).Caption
            End If
            PrintDatil
         Next intItem
      End With
      
      Printer.EndDoc
      PUB_RestorePrinter strPrinter
      ShowPrintOk
   Else
      MsgBox "沒有資料可以列印 !", vbCritical
   End If
   Screen.MousePointer = vbDefault
End Sub

Sub PrintTitle()
   GetPleft
   iPrint = 500
   Printer.Orientation = 1
   Printer.Font.Name = "細明體"
   Printer.Font.Size = 22
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.CurrentX = 4500
   Printer.CurrentY = iPrint
   Printer.Print "信件沖銷記錄統計"
   iPrint = iPrint + 500
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   
   Printer.CurrentX = 4300
   Printer.CurrentY = iPrint
   Printer.Print "轉入日期：" & Format(ChangeTStringToTDateString(frm06010617.txt1(1)) & " ", "@@@@@@@@@") & "－" & ChangeTStringToTDateString(frm06010617.txt1(2))
   
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print "列印人：" & strUserName
   Printer.CurrentX = 8500
   Printer.CurrentY = iPrint
   Printer.Print "列印日期：" & Format(strSrvDate(2), "##/##/##")
   iPrint = iPrint + 300
   Printer.CurrentX = 8500
   Printer.CurrentY = iPrint
   Printer.Print "頁　　次：" & str(Page)
   iPrint = iPrint + 300
   
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print String(90, "-")
   iPrint = iPrint + 300
   
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = iPrint
   Printer.Print Me.grdDataList1.TextMatrix(0, 0)
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iPrint
   Printer.Print Me.grdDataList1.TextMatrix(0, 1)
   If frm06010617.txt1(6) = "1" Then
      Printer.CurrentX = PLeft(2)
      Printer.CurrentY = iPrint
      Printer.Print Me.grdDataList1.TextMatrix(0, 2)
      Printer.CurrentX = PLeft(3)
      Printer.CurrentY = iPrint
      Printer.Print "數量"
   Else
      Printer.CurrentX = PLeft(2)
      Printer.CurrentY = iPrint
      Printer.Print "數量"
   End If
   
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print String(90, "-")
   iPrint = iPrint + 300
End Sub

Sub GetPleft()
   Erase PLeft
   If frm06010617.txt1(6) = "2" Or frm06010617.txt1(6) = "3" Then
      PLeft(0) = 500
      PLeft(1) = 2000
      PLeft(2) = 6500
   Else
      PLeft(0) = 500
      PLeft(1) = 2500
      PLeft(2) = 4000
      PLeft(3) = 7500
   End If
End Sub

Sub PrintDatil()
   For i = 0 To 3
      Printer.CurrentX = PLeft(i)
      Printer.CurrentY = iPrint
      If (frm06010617.txt1(6) = "2" Or frm06010617.txt1(6) = "3") And i = 1 Then
         Printer.Print Mid(strTemp3(i), 1, 40)
      Else
         Printer.Print strTemp3(i)
      End If
   Next i
   iPrint = iPrint + 300
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   SetDataListWidth
   cmdState = -1
   PUB_SetPrinter Me.Name, Combo1, strPrinter
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm06010617_1 = Nothing
End Sub

Sub StrMenu()
Dim jj As Integer
Dim m_Condition As String
Dim strText As String, dblCnt As Double
Dim intMaxRow As Integer, intSortRow As String

   Me.Enabled = False
   '讀出資料
   If DoTemp = False Then
      tmpBol = fnCancelNowFormAndShowParentForm(Me)
      Screen.MousePointer = vbDefault
      Exit Sub
   End If
   
   '顯示表單資料
   lbl1(0).Caption = frm06010617.txt1(1) + "－" + frm06010617.txt1(2)
   
   '清除畫面右邊統計數
   lbl1(4).Caption = "" '總計
   
   m_Condition = ""
   '+補 Order by 語法(O12不會自動以 group by 欄位排序)
   Select Case frm06010617.txt1(6)
      Case "1"
           m_Condition = "承辦同仁"
           strSql = "select nvl(a0922,nvl(a0902,st03)) as 部門 ,st02 as 承辦同仁" & _
                    ",decode(R02013," & 外專信件處理結果 & ") as 沖銷類別,count(*) as 數量, 0 as 排序,st03,st93" & _
                    " from r100105,staff,acc090,acc090new" & _
                    " where id='" & strUserNum & "' and substr(R02003,1,5)=st01(+) and st03=a0901(+) and st93=a0921(+)" & _
                    " group by nvl(a0922,nvl(a0902,st03)),st02,st03,st93,R02003,R02013" & _
                    " order by st03 desc,st93 asc,substr(R02003,1,5),to_number(R02013)"
      Case "2"
           m_Condition = "沖銷類別"
           strSql = "SELECT R02013 AS 代碼,decode(R02013," & 外專信件處理結果 & ") AS 沖銷類別,COUNT(*) AS 數量" & _
                    " FROM R100105 WHERE ID='" & strUserNum & "'" & _
                    " GROUP BY R02013 ORDER BY to_number(R02013)"
      Case "3"
           m_Condition = "來函對象"
           strSql = "SELECT R02011 AS 來函對象,decode(R02013," & 外專信件處理結果 & ") AS 沖銷類別,COUNT(*) AS 數量" & _
                    " FROM R100105 WHERE ID='" & strUserNum & "'" & _
                    " GROUP BY R02011,R02013" & _
                    " ORDER BY R02011,to_number(R02013)"
      Case Else
   End Select
   
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   Set grdDataList1.Recordset = adoRecordset
   SetDataListWidth '表頭置中
   If frm06010617.txt1(6) = "3" Then '更換標題
      Me.grdDataList1.TextMatrix(0, 0) = m_Condition
      Me.grdDataList1.ColWidth(0) = 3500
      Me.grdDataList1.TextMatrix(0, 1) = "沖銷類別"
      Me.grdDataList1.ColWidth(1) = 1500
   End If
   
   j = 0
   For i = 1 To grdDataList1.Rows - 1
      '統計條件統一放 col=1
'      grdDataList1.col = 1
      grdDataList1.row = i
      
      '總計
      If frm06010617.txt1(6) = "1" Then
         grdDataList1.col = 3
      Else
         grdDataList1.col = 2
      End If
      j = j + Val(grdDataList1.Text)
   Next i
   lbl1(4).Caption = str(j) '總計
   
   '加小計
   If frm06010617.txt1(6) = "1" Then '承辦同仁
      intMaxRow = grdDataList1.Rows - 1: strText = ""
      dblCnt = 0: jj = intMaxRow
      intSortRow = 0
      For i = 1 To intMaxRow
         If strText <> grdDataList1.TextMatrix(i, 1) And strText <> "" Then
            grdDataList1.AddItem ""
            intSortRow = intSortRow + 1
            jj = jj + 1
            grdDataList1.TextMatrix(jj, 2) = "小計"
            grdDataList1.TextMatrix(jj, 3) = dblCnt
            grdDataList1.TextMatrix(jj, 4) = intSortRow '排序
            dblCnt = 0
         End If
         intSortRow = intSortRow + 1
         dblCnt = dblCnt + Val(grdDataList1.TextMatrix(i, 3))
         strText = grdDataList1.TextMatrix(i, 1)
         grdDataList1.TextMatrix(i, 4) = intSortRow '排序
      Next i
      If dblCnt > 0 Then
         grdDataList1.AddItem ""
         intSortRow = intSortRow + 1
         jj = jj + 1
         grdDataList1.TextMatrix(jj, 2) = "小計"
         grdDataList1.TextMatrix(jj, 3) = dblCnt
         grdDataList1.TextMatrix(jj, 4) = intSortRow '排序
         dblCnt = 0
      End If
      grdDataList1.col = 4
      grdDataList1.Sort = 3 '數值昇冪
   End If
   
   Me.Enabled = True
End Sub

Function DoTemp() As Boolean
Dim strSQL1 As String
   
   frm06010617.Hide
   
   j = 0
   cnnConnection.Execute "DELETE FROM R100105 where id='" & strUserNum & "' "
   
   '組合條件:
   '轉入日期
   If Len(Trim(frm06010617.txt1(1))) <> 0 Then
      strSQL1 = strSQL1 + " AND ii01>=" & Val(ChangeTStringToWString(frm06010617.txt1(1))) & " "
   End If
   If Len(Trim(frm06010617.txt1(2))) <> 0 Then
      strSQL1 = strSQL1 & " AND ii01<=" & Val(ChangeTStringToWString(frm06010617.txt1(2))) & " "
   End If
   If Len(Trim(frm06010617.txt1(1))) <> 0 Or Len(Trim(frm06010617.txt1(2))) <> 0 Then
      pub_QL05 = pub_QL05 & ";轉入日期：" & frm06010617.txt1(1) & "-" & frm06010617.txt1(2)
   End If
   '來函對象
   If Len(Trim(frm06010617.txt1(3))) <> 0 Then
      strSQL1 = strSQL1 + " AND instr(ii11,'" & frm06010617.txt1(3) & "')>0 "
   End If
'   If Len(Trim(frm06010617.txt1(4))) <> 0 Then
'      strSQL1 = strSQL1 + " AND ii11<='" & frm06010617.txt1(4) & "' "
'   End If
'   If Len(Trim(frm06010617.txt1(3))) <> 0 Or Len(Trim(frm06010617.txt1(4))) <> 0 Then
   If Len(Trim(frm06010617.txt1(3))) <> 0 Then
      'pub_QL05 = pub_QL05 & ";" & frm06010617.Label7 & frm06010617.txt1(3) & "-" & frm06010617.txt1(4)
      pub_QL05 = pub_QL05 & ";" & frm06010617.Label7 & frm06010617.txt1(3)
   End If
   '沖銷類別
   If Len(Trim(frm06010617.cboType(0))) <> 0 Then
      strSQL1 = strSQL1 + " AND ii27>=" & Trim(Left(frm06010617.cboType(0), 2))
   End If
   If Len(Trim(frm06010617.cboType(1))) <> 0 Then
      strSQL1 = strSQL1 + " AND ii27<=" & Trim(Left(frm06010617.cboType(1), 2))
   End If
   If Len(Trim(frm06010617.cboType(0))) <> 0 Or Len(Trim(frm06010617.cboType(1))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & frm06010617.Label6 & Trim(Left(frm06010617.cboType(0), 2)) & "-" & Trim(Left(frm06010617.cboType(1), 2))
   End If
   '承辦同仁
   If Len(Trim(frm06010617.txt1(5))) <> 0 Then
      strSQL1 = strSQL1 + " AND ir19='" & frm06010617.txt1(5) & "' "
   End If
   If Len(Trim(frm06010617.txt1(5))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & frm06010617.Label17 & frm06010617.txt1(5)
   End If
   pub_QL05 = pub_QL05 & ";統計條件：" & frm06010617.txt1(6) & IIf(frm06010617.txt1(6) = "1", ".承辦同仁", IIf(frm06010617.txt1(6) = "2", ".沖銷類別", IIf(frm06010617.txt1(6) = "3", ".來函對象", "")))
   
   '部門別
   If Len(Trim(frm06010617.txt1(7))) <> 0 Then
      strSQL1 = strSQL1 + " AND st93>='" & Left(frm06010617.txt1(7), 3) & "' "
   End If
   If Len(Trim(frm06010617.txt1(8))) <> 0 Then
      strSQL1 = strSQL1 + " AND st93<='" & Left(frm06010617.txt1(8), 3) & "' "
   End If
   If Len(Trim(frm06010617.txt1(7))) <> 0 Or Len(Trim(frm06010617.txt1(8))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & frm06010617.Label8 & frm06010617.txt1(8) & "-" & frm06010617.txt1(7)
   End If
   
   strSql = "select '" & strUserNum & "',ii01||'-'||ii03||'-'||ir04,SqlDateT(ii01),decode(st03,'F23',decode(ii27,'12','8',ii27),decode(ii27,'7','12',ii27)),ir19,ii11" & _
            " From ipdeptinput, inputrecord, staff" & _
            " where ii27 is not null and ii27<>'11'" & _
            " and ir19 is not null and ir08 is not null and ir22 is null and ir16 not in('3','4','8')" & _
            " and ii01=ir01 and ii03=ir03 and ir19=st01(+) and st03 in('F22','F23')" & strSQL1 & _
            " Union " & _
            "select '" & strUserNum & "',ii01||'-'||ii03||'-'||ir04,SqlDateT(ii01),decode(st03,'F23',decode(ii27,'12','8',ii27),decode(ii27,'7','12',ii27)),ir19,ii11" & _
            " From ipdeptinput, inputrecord, staff" & _
            " where ii27 is not null and ii27<>'11'" & _
            " and ir19 is not null and ir08 is not null and ir22 is not null and ir16 in('7','9')" & _
            " and ii01=ir01 and ii03=ir03 and ir19=st01(+) and st03 in('F22','F23')" & strSQL1
   cnnConnection.Execute "insert into r100105(ID,R02008,R02002,R02013,R02003,R02011) " & strSql
   
   strSql = "select * from r100105 where id ='" & strUserNum & "' And RowNum <= 1 "
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
      InsertQueryLog (adoRecordset.RecordCount)
   Else
      Call InsertQueryLog(0)
      ShowNoData
      Screen.MousePointer = vbDefault
      DoTemp = False
      Exit Function
   End If
   CheckOC
   DoTemp = True
End Function
