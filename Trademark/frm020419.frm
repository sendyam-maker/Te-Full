VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm020419 
   BorderStyle     =   1  '單線固定
   Caption         =   "台灣商標爭議案件補充資料次數明細及統計"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6705
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   6705
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      Height          =   2715
      Left            =   630
      TabIndex        =   14
      Top             =   150
      Width           =   5985
      Begin VB.CommandButton cmdSearch 
         Caption         =   "確定(&O)"
         Default         =   -1  'True
         Height          =   375
         Left            =   4320
         Style           =   1  '圖片外觀
         TabIndex        =   8
         Top             =   30
         Width           =   800
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "結束(&X)"
         CausesValidation=   0   'False
         Height          =   375
         Left            =   5160
         Style           =   1  '圖片外觀
         TabIndex        =   9
         Top             =   30
         Width           =   800
      End
      Begin VB.TextBox txtSalesArea1 
         Height          =   285
         Left            =   2505
         MaxLength       =   3
         TabIndex        =   3
         Top             =   630
         Width           =   435
      End
      Begin VB.TextBox txtSalesArea 
         Height          =   285
         Left            =   1920
         MaxLength       =   3
         TabIndex        =   2
         Top             =   630
         Width           =   435
      End
      Begin VB.TextBox txtSales 
         Height          =   285
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   4
         Top             =   960
         Width           =   915
      End
      Begin VB.TextBox txtDate 
         Height          =   285
         Index           =   1
         Left            =   2985
         MaxLength       =   7
         TabIndex        =   1
         Top             =   300
         Width           =   915
      End
      Begin VB.TextBox txtDate 
         Height          =   285
         Index           =   0
         Left            =   1920
         MaxLength       =   7
         TabIndex        =   0
         Top             =   300
         Width           =   915
      End
      Begin VB.TextBox txtQueryType 
         Height          =   264
         Left            =   1920
         MaxLength       =   1
         TabIndex        =   5
         Text            =   "1"
         Top             =   1290
         Width           =   240
      End
      Begin VB.TextBox txtSortType 
         Height          =   264
         Left            =   1920
         MaxLength       =   1
         TabIndex        =   6
         Text            =   "1"
         Top             =   1590
         Width           =   240
      End
      Begin VB.TextBox txtPrintType 
         Height          =   264
         Left            =   1920
         MaxLength       =   1
         TabIndex        =   7
         Text            =   "1"
         Top             =   1890
         Width           =   240
      End
      Begin VB.Line Line2 
         X1              =   2325
         X2              =   2595
         Y1              =   750
         Y2              =   750
      End
      Begin VB.Label lblSalesName 
         Height          =   180
         Left            =   2880
         TabIndex        =   21
         Top             =   1020
         Width           =   1440
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "智權人員："
         Height          =   180
         Left            =   990
         TabIndex        =   20
         Top             =   1020
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "業務區："
         Height          =   180
         Index           =   0
         Left            =   1185
         TabIndex        =   19
         Top             =   675
         Width           =   720
      End
      Begin VB.Line Line5 
         X1              =   2835
         X2              =   3105
         Y1              =   450
         Y2              =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "補充資料日期："
         Height          =   180
         Left            =   645
         TabIndex        =   18
         Top             =   360
         Width           =   1260
      End
      Begin VB.Label Label1 
         Caption         =   "查詢結果： 　   ( 1.明細 2.統計 )"
         Height          =   180
         Index           =   122
         Left            =   1005
         TabIndex        =   17
         Top             =   1350
         Width           =   3090
      End
      Begin VB.Label Label1 
         Caption         =   "統計資料排序順序： 　   ( 1.次數 2.業務區 )"
         Height          =   180
         Index           =   1
         Left            =   270
         TabIndex        =   16
         Top             =   1650
         Width           =   3510
      End
      Begin VB.Label Label1 
         Caption         =   "列印別： 　   ( 1.查詢 2.報表 )"
         Height          =   180
         Index           =   2
         Left            =   1170
         TabIndex        =   15
         Top             =   1950
         Width           =   3090
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "回前畫面(&U)"
      Height          =   375
      Left            =   7650
      Style           =   1  '圖片外觀
      TabIndex        =   10
      Top             =   60
      Width           =   1185
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   5205
      Left            =   30
      TabIndex        =   11
      Top             =   495
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   9181
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
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
   Begin VB.Label Label1 
      Caption         =   "共　0　件"
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   3
      Left            =   60
      TabIndex        =   13
      Top             =   195
      Width           =   1125
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   105
      TabIndex        =   12
      Top             =   5430
      Width           =   45
   End
End
Attribute VB_Name = "frm020419"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Create By Sindy 2012/5/16
Option Explicit

Dim i As Integer, j As Integer
Dim PLeft(1 To 12) As Integer
Dim strTemp(1 To 12) As String
Dim iPgae As Integer, iLine As Integer


Private Sub SetDataListWidth()
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer

If txtQueryType = "1" Then '明細
   arrGridHeadText = Array("業務區", "智權人員", "本所案號", "案件性質", "承辦人", "本所期限", _
                           "異動項目", "異動日", "齊備日", "補充資料日", "承辦期限", "指定會稿日")
   arrGridHeadWidth = Array(850, 750, 1000, 1000, 750, 850 _
                            , 1000, 850, 1000, 1000, 850, 900)
Else '統計
   'Modify By Sindy 2012/7/16 增加異動項目
   arrGridHeadText = Array("業務區", "智權人員", "異動項目", "次數")
   arrGridHeadWidth = Array(1200, 900, 1700, 900)
End If
grdDataList.MergeCells = flexMergeRestrictColumns
grdDataList.Cols = UBound(arrGridHeadText) + 1
For iRow = 0 To grdDataList.Cols - 1
   grdDataList.row = 0
   grdDataList.col = iRow
   grdDataList.Text = arrGridHeadText(iRow)
   grdDataList.ColWidth(iRow) = arrGridHeadWidth(iRow)
   grdDataList.CellAlignment = flexAlignLeftCenter
Next
End Sub

Private Function ConstrainCheck() As Boolean
   Dim bolCancel As Boolean
   ConstrainCheck = True
   
'   If txtDate(0) = "" And txtDate(1) = "" And _
'      txtSalesArea = "" And txtSalesArea1 = "" And _
'      txtSales = "" Then
'      MsgBox "至少輸入一項查詢條件！", vbExclamation
'      txtDate(0).SetFocus
'      txtDate_GotFocus (0)
'      ConstrainCheck = False
'      Exit Function
'   End If
   
   If txtDate(0) = "" Then
      MsgBox "請輸入補充資料日期(起)！", vbExclamation
      txtDate(0).SetFocus
      txtDate_GotFocus (0)
      ConstrainCheck = False
      Exit Function
   End If
   If txtDate(1) = "" Then
      MsgBox "請輸入補充資料日期(迄)！", vbExclamation
      txtDate(1).SetFocus
      txtDate_GotFocus (1)
      ConstrainCheck = False
      Exit Function
   End If
   If txtDate(0) <> "" And txtDate(1) <> "" Then
      bolCancel = False
      Call txtDate_Validate(0, bolCancel)
      If bolCancel = True Then
         ConstrainCheck = False
         Exit Function
      End If
      bolCancel = False
      Call txtDate_Validate(1, bolCancel)
      If bolCancel = True Then
         ConstrainCheck = False
         Exit Function
      End If
   End If
   
   If txtSalesArea = "" And txtSalesArea1 <> "" Then
      MsgBox "請輸入業務區(起)！", vbExclamation
      txtSalesArea.SetFocus
      txtSalesArea_GotFocus
      ConstrainCheck = False
      Exit Function
   End If
   If txtSalesArea <> "" And txtSalesArea1 = "" Then
      MsgBox "請輸入業務區(迄)！", vbExclamation
      txtSalesArea1.SetFocus
      txtSalesArea1_GotFocus
      ConstrainCheck = False
      Exit Function
   End If
   If txtSalesArea1 <> "" Then
      bolCancel = False
      Call txtSalesArea1_Validate(bolCancel)
      If bolCancel = True Then
         ConstrainCheck = False
         Exit Function
      End If
   End If
   
   If txtQueryType = "" Then
      MsgBox "查詢結果不可空白！", vbExclamation
      txtQueryType.SetFocus
      txtQueryType_GotFocus
      ConstrainCheck = False
      Exit Function
   Else
      If txtQueryType = "2" And txtSortType = "" Then
         MsgBox "統計資料排序順序不可空白！", vbExclamation
         txtSortType.SetFocus
         txtSortType_GotFocus
         ConstrainCheck = False
         Exit Function
      End If
   End If
   
   If txtPrintType = "" Then
      MsgBox "列印別不可空白！", vbExclamation
      txtPrintType.SetFocus
      txtPrintType_GotFocus
      ConstrainCheck = False
      Exit Function
   End If
End Function

Public Function doQuery() As Boolean
Dim stCon As String

On Error GoTo ErrHnd
   
   doQuery = False
   stCon = ""
   
   If txtDate(0) <> "" And txtDate(1) <> "" Then
      'Modify By Sindy 2012/7/16 tcd02異動類別原只有2.智權人員補充資料日,增加4.電腦中心取消齊備
      'stCon = " and tcd06>=" & DBDATE(txtDate(0)) & " and tcd06<=" & DBDATE(txtDate(1))
      'Modify By Sindy 2012/10/22 增加5.通知補充資料
      'stCon = " and ((tcd02='2' and tcd06>=" & DBDATE(txtDate(0)) & " and tcd06<=" & DBDATE(txtDate(1)) & ") or (tcd02='4' and TCD04>=" & DBDATE(txtDate(0)) & " and TCD04<=" & DBDATE(txtDate(1)) & "))"
      stCon = stCon & " and ((tcd02='2' and tcd06>=" & DBDATE(txtDate(0)) & " and tcd06<=" & DBDATE(txtDate(1)) & ") or ((tcd02='4' or tcd02='5') and TCD04>=" & DBDATE(txtDate(0)) & " and TCD04<=" & DBDATE(txtDate(1)) & "))"
   End If
   If txtSalesArea <> "" And txtSalesArea1 <> "" Then
      stCon = stCon & " and cp12>='" & txtSalesArea & "' and cp12<='" & txtSalesArea1 & "'"
   End If
   If txtSales <> "" Then
      stCon = stCon & " and cp13='" & txtSales & "'"
   End If
   
   '查詢SQL
   If txtQueryType = "1" Then '明細
      'Modify By Sindy 2012/7/16 sqldatet(ep06) as 齊備日==>decode(tcd02,'4','電腦中心取消齊備',sqldatet(ep06)) as 齊備日
      'Modify By Sindy 2012/10/22 decode(tcd02,'4','電腦中心取消齊備',sqldatet(ep06)) as 齊備日==>decode(tcd02,'4','電腦中心取消齊備','5','通知補充資料',sqldatet(ep06)) as 齊備日
      strSql = "select A0902 as 業務區,s1.st02 as 智權人員,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as 本所案號,cpm03 as 案件性質,s2.st02 as 承辦人,sqldatet(cp06) as 本所期限,decode(tcd02,'2','業務補充','4','電腦中心取消齊備','5','通知補充資料') as 異動項目,sqldatet(tcd04) as 異動日,sqldatet(ep06) as 齊備日,sqldatet(tcd06) as 補充資料日,sqldatet(cp48) as 承辦期限,sqldatet(ep28) as 指定會稿日" & _
               " from tmctldate,caseprogress,engineerprogress,staff s1,staff s2,casepropertymap,acc090" & _
               " where tcd01=cp09(+)" & _
               " and tcd01=ep02(+)" & stCon & _
               " and cp13=s1.st01(+)" & _
               " and cp14=s2.st01(+)" & _
               " and cp01=cpm01(+) and cp10=cpm02(+)" & _
               " and cp12=a0901(+)" & _
               " order by cp12,cp13,cp01,cp02,cp03,cp04,cp09,tcd06"
   Else '統計
      'Modify By Sindy 2012/7/16 增加異動項目:電腦中心取消齊備
      'Modify By Sindy 2012/10/22 增加異動項目:通知補充資料
      strSql = "select A0902 as 業務區,s1.st02 as 智權人員,decode(tcd02,'2','業務補充','4','電腦中心取消齊備','5','通知補充資料') as 作業狀況,count(*) as 次數" & _
               " from tmctldate,caseprogress,engineerprogress,staff s1,staff s2,casepropertymap,acc090" & _
               " where tcd01=cp09(+)" & _
               " and tcd01=ep02(+)" & stCon & _
               " and cp13=s1.st01(+)" & _
               " and cp14=s2.st01(+)" & _
               " and cp01=cpm01(+) and cp10=cpm02(+)" & _
               " and cp12=a0901(+)" & _
               " group by cp12,cp13,A0902,s1.st02,tcd02 "
      If txtSortType = "1" Then '次數
         strSql = strSql & " order by 次數 desc,cp12 asc,cp13 asc"
      Else
         strSql = strSql & " order by cp12 asc,cp13 asc"
      End If
   End If
   CheckOC3
   grdDataList.Rows = 2
   grdDataList.Clear
   SetDataListWidth
   With AdoRecordSet3
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount > 0 Then
         Set grdDataList.Recordset = AdoRecordSet3.Clone
         SetDataListWidth
         Label1(3).Caption = "共　" & .RecordCount & "　件"
         If txtPrintType = "1" Then
            Call getFormType(1) '查詢
         Else
            Call PrintData '列印
         End If
      Else
         Label1(3).Caption = "共　0　件"
         MsgBox "無符合資料！", vbInformation
      End If
   End With
   
   doQuery = True
   Exit Function
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Function

Private Sub cmdExit_Click()
   Unload Me
End Sub

Public Sub cmdSearch_Click()
   Screen.MousePointer = vbHourglass
   grdDataList.MousePointer = flexHourglass
   If ConstrainCheck = True Then
      SetDataListWidth
      Call doQuery
   End If
   grdDataList.MousePointer = flexDefault
   Screen.MousePointer = vbDefault
End Sub

'回前畫面
Private Sub Command1_Click()
   SetDataListWidth
   Call getFormType(0)
End Sub

Private Sub getFormType(Optional strType As Integer = 0)
   If strType = 0 Then '查詢條件畫面
      Me.Height = 3400
      Me.Width = 6800
      Frame1.Visible = True
      Command1.Value = False
      grdDataList.Visible = False
      Label1(3).Visible = False
   Else '明細畫面
      Me.Height = 6120
      Me.Width = 9045
      Frame1.Visible = False
      Command1.Visible = True
      grdDataList.Visible = True
      Label1(3).Visible = True
   End If
   MoveFormToCenter Me
End Sub

Private Sub Form_Load()
   Call getFormType(0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm020419 = Nothing
End Sub

Private Sub txtDate_GotFocus(Index As Integer)
   TextInverse txtDate(Index)
   CloseIme
End Sub

Private Sub txtDate_Validate(Index As Integer, Cancel As Boolean)
   If txtDate(Index) <> "" Then
      If ChkDate(txtDate(Index)) = False Then
         Cancel = True
         txtDate(Index).SetFocus
         txtDate_GotFocus Index
         Exit Sub
      End If
      If Index = 1 Then
         If RunNick2(txtDate(0), txtDate(1)) = True Then
            txtDate(Index).SetFocus
            txtDate_GotFocus Index
            Cancel = True
            Exit Sub
         End If
      End If
   End If
End Sub

Private Sub txtPrintType_GotFocus()
   TextInverse txtPrintType
   CloseIme
End Sub

Private Sub txtPrintType_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtQueryType_GotFocus()
   TextInverse txtQueryType
   CloseIme
End Sub

Private Sub txtQueryType_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtSales_Change()
   If Len(txtSales) > 4 Then
      lblSalesName = GetStaffName(txtSales, True)
   Else
      lblSalesName = ""
   End If
End Sub

Private Sub txtSales_GotFocus()
   TextInverse txtSales
   CloseIme
End Sub

Private Sub txtSales_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtSalesArea_GotFocus()
   TextInverse txtSalesArea
   CloseIme
End Sub

Private Sub txtSalesArea_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtSalesArea1_GotFocus()
   TextInverse txtSalesArea1
   CloseIme
End Sub

Private Sub txtSalesArea1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtSalesArea1_Validate(Cancel As Boolean)
If Trim(txtSalesArea1) <> "" Then
   If RunNick(txtSalesArea, txtSalesArea1) = True Then
      Cancel = True
      Exit Sub
   End If
End If
End Sub

Private Sub txtSortType_GotFocus()
   TextInverse txtSortType
   CloseIme
End Sub

Private Sub txtSortType_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub PrintData()
If txtQueryType = "1" Then '明細
   Printer.Orientation = 2 '2.橫印
   iLine = 1: iPgae = 0
   For i = 1 To grdDataList.Rows - 1
      For j = 1 To 12
          strTemp(j) = ""
      Next j
      strTemp(1) = StrToStr(grdDataList.TextMatrix(i, 0), 5)
      strTemp(2) = grdDataList.TextMatrix(i, 1)
      strTemp(3) = grdDataList.TextMatrix(i, 2)
      strTemp(4) = StrToStr(grdDataList.TextMatrix(i, 3), 5)
      strTemp(5) = grdDataList.TextMatrix(i, 4)
      strTemp(6) = grdDataList.TextMatrix(i, 5)
      strTemp(7) = StrToStr(grdDataList.TextMatrix(i, 6), 8)
      strTemp(8) = grdDataList.TextMatrix(i, 7)
      strTemp(9) = grdDataList.TextMatrix(i, 8)
      strTemp(10) = grdDataList.TextMatrix(i, 9)
      strTemp(11) = grdDataList.TextMatrix(i, 10)
      strTemp(12) = grdDataList.TextMatrix(i, 11)
      If iLine > 36 Or iLine = 1 Then
         If iPgae <> 0 Then Printer.NewPage
         iLine = 1
         PrintTitle '列印表頭
      End If
      PrintDetail
   Next i
Else
   Printer.Orientation = 1 '1.直印
   iLine = 1: iPgae = 0
   For i = 1 To grdDataList.Rows - 1
      For j = 1 To 4 '3
          strTemp(j) = ""
      Next j
      strTemp(1) = grdDataList.TextMatrix(i, 0)
      strTemp(2) = grdDataList.TextMatrix(i, 1)
      strTemp(3) = grdDataList.TextMatrix(i, 2)
      strTemp(4) = grdDataList.TextMatrix(i, 3) 'Add By Sindy 2012/7/16
      If iLine > 51 Or iLine = 1 Then
         If iPgae <> 0 Then Printer.NewPage
         iLine = 1
         PrintTitle '列印表頭
      End If
      PrintDetail
   Next i
End If
Printer.EndDoc
ShowPrintOk
End Sub

Sub GetPleft()
If txtQueryType = "1" Then '明細
   PLeft(1) = 500
   PLeft(2) = 1800
   PLeft(3) = 2700
   PLeft(4) = 4300
   PLeft(5) = 5600
   PLeft(6) = 6500
   PLeft(7) = 7500
   PLeft(8) = 9300
   PLeft(9) = 10400
   PLeft(10) = 11400
   PLeft(11) = 12700
   PLeft(12) = 13900
Else
   PLeft(1) = 500
   PLeft(2) = 2500
   PLeft(3) = 4000
   PLeft(4) = 7000 'Add By Sindy 2012/7/16
End If
End Sub

Sub PrintTitle()
GetPleft
iPgae = iPgae + 1

Printer.Font.Size = 18
Printer.Font.Underline = False
Printer.FontBold = False

If txtQueryType = "1" Then '明細
   Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("台灣商標爭議案件補充資料明細") / 2)
   Printer.CurrentY = iLine * 300
   Printer.Print "台灣商標爭議案件補充資料明細"
Else
   Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("台灣商標爭議案件補充資料次數統計") / 2)
   Printer.CurrentY = iLine * 300
   Printer.Print "台灣商標爭議案件補充資料次數統計"
End If

Printer.Font.Size = 12
Printer.CurrentX = PLeft(1)
Printer.CurrentY = 900
Printer.Print "列印人員：" & strUserName
Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("補充資料日期：" & ChangeTStringToTDateString(txtDate(0)) & " - " & ChangeTStringToTDateString(txtDate(1))) / 2)
Printer.CurrentY = 900
Printer.Print "補充資料日期：" & ChangeTStringToTDateString(txtDate(0)) & " - " & ChangeTStringToTDateString(txtDate(1))
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
Printer.CurrentY = 600
Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
Printer.CurrentY = 900
Printer.Print "頁　　次：" & Printer.Page

iLine = 5
If txtQueryType = "1" Then '明細
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iLine * 300
   Printer.Print "業務區"
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iLine * 300
   Printer.Print "智權人員"
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = iLine * 300
   Printer.Print "本所案號"
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = iLine * 300
   Printer.Print "案件性質"
   Printer.CurrentX = PLeft(5)
   Printer.CurrentY = iLine * 300
   Printer.Print "承辦人"
   Printer.CurrentX = PLeft(6)
   Printer.CurrentY = iLine * 300
   Printer.Print "本所期限"
   'Add By Sindy 2012/10/25
   Printer.CurrentX = PLeft(7)
   Printer.CurrentY = iLine * 300
   Printer.Print "異動項目"
   Printer.CurrentX = PLeft(8)
   Printer.CurrentY = iLine * 300
   Printer.Print "異動日"
   '2012/10/25 End
   Printer.CurrentX = PLeft(9)
   Printer.CurrentY = iLine * 300
   Printer.Print "齊備日"
   Printer.CurrentX = PLeft(10)
   Printer.CurrentY = iLine * 300
   Printer.Print "補充資料日"
   Printer.CurrentX = PLeft(11)
   Printer.CurrentY = iLine * 300
   Printer.Print "承辦期限"
   Printer.CurrentX = PLeft(12)
   Printer.CurrentY = iLine * 300
   Printer.Print "指定會稿日"
   iLine = iLine + 1
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iLine * 300
   Printer.Print String(215, "-")
   iLine = iLine + 1
Else
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iLine * 300
   Printer.Print "業務區"
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iLine * 300
   Printer.Print "智權人員"
   'Add By Sindy 2012/7/16
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = iLine * 300
   Printer.Print "異動項目"
   '2012/7/16 End
   Printer.CurrentX = PLeft(4) - Printer.TextWidth("次數")
   Printer.CurrentY = iLine * 300
   Printer.Print "次數"
   iLine = iLine + 1
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iLine * 300
   Printer.Print String(140, "-")
   iLine = iLine + 1
End If
End Sub

Sub PrintDetail()
Dim m_j As Integer

If txtQueryType = "1" Then '明細
   For m_j = 1 To 12
      Printer.CurrentX = PLeft(m_j)
      Printer.CurrentY = iLine * 300
      Printer.Print strTemp(m_j)
   Next m_j
   iLine = iLine + 1
Else
   For m_j = 1 To 4 '3
      If m_j = 4 Then 'Modify By Sindy 2012/7/16
         Printer.CurrentX = PLeft(m_j) - Printer.TextWidth(strTemp(m_j))
      Else
         Printer.CurrentX = PLeft(m_j)
      End If
      Printer.CurrentY = iLine * 300
      Printer.Print strTemp(m_j)
   Next m_j
   iLine = iLine + 1
End If
End Sub
