VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm050207_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "員工查詢印表記錄資料查詢"
   ClientHeight    =   5750
   ClientLeft      =   170
   ClientTop       =   960
   ClientWidth     =   8950
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5750
   ScaleWidth      =   8950
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   3
      Left            =   8040
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   60
      Width           =   756
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "列印(P)"
      Default         =   -1  'True
      Height          =   400
      Index           =   1
      Left            =   6030
      TabIndex        =   2
      Top             =   60
      Width           =   756
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   6810
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   60
      Width           =   1200
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   4005
      Left            =   60
      TabIndex        =   1
      Top             =   1680
      Width           =   8850
      _ExtentX        =   15610
      _ExtentY        =   7056
      _Version        =   393216
      BackColor       =   -2147483624
      Cols            =   6
      FixedCols       =   0
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      HighLight       =   0
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9
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
   Begin MSForms.TextBox Text1 
      Height          =   915
      Left            =   1110
      TabIndex        =   13
      Top             =   690
      Width           =   7755
      VariousPropertyBits=   -1466941413
      BackColor       =   -2147483637
      ScrollBars      =   2
      Size            =   "13679;1614"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Left            =   1110
      TabIndex        =   12
      Top             =   150
      Width           =   1005
      BackColor       =   -2147483637
      VariousPropertyBits=   27
      Size            =   "1773;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label9 
      Caption         =   "條件："
      Height          =   225
      Left            =   150
      TabIndex        =   11
      Top             =   750
      Width           =   915
   End
   Begin VB.Label Label8 
      Height          =   225
      Left            =   5100
      TabIndex        =   10
      Top             =   450
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "結果筆數："
      Height          =   225
      Left            =   4110
      TabIndex        =   9
      Top             =   450
      Width           =   945
   End
   Begin VB.Label Label6 
      Height          =   225
      Left            =   1110
      TabIndex        =   8
      Top             =   450
      Width           =   2925
   End
   Begin VB.Label Label5 
      Caption         =   "程式名稱："
      Height          =   225
      Left            =   150
      TabIndex        =   7
      Top             =   450
      Width           =   915
   End
   Begin VB.Label Label4 
      Height          =   225
      Left            =   3330
      TabIndex        =   6
      Top             =   150
      Width           =   2355
   End
   Begin VB.Label Label3 
      Caption         =   "日期／時間："
      Height          =   225
      Left            =   2190
      TabIndex        =   5
      Top             =   150
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "操作人員："
      Height          =   225
      Left            =   150
      TabIndex        =   4
      Top             =   150
      Width           =   915
   End
End
Attribute VB_Name = "frm050207_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/22 改成Form2.0 ; grdDataList改字型=新細明體-ExtB、Label2、Text1
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/2 日期欄已修改
'2010/01/06 CREATE BY Sindy
Option Explicit

Dim strSQL1 As String, strSQL2 As String, StrSQL3 As String, StrSQL4 As String, strSQL5 As String
Dim strSql As String, i As Integer, j As Integer, s As Integer
Dim StrTag As String, intK As Integer
'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer
Dim m_i As Integer
Dim PLeft(1 To 6) As Integer
Dim strTemp(1 To 6) As String
Dim iPgae As Integer, iLine As Integer
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序


Private Sub SetDataListWidth()
grdDataList.row = 0
grdDataList.col = 0: grdDataList.Text = "操作人員"
grdDataList.ColWidth(0) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 1: grdDataList.Text = "日　期 "
grdDataList.ColWidth(1) = 750
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 2: grdDataList.Text = "時　間"
grdDataList.ColWidth(2) = 750
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 3: grdDataList.Text = "程式名稱"
grdDataList.ColWidth(3) = 1400
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 4: grdDataList.Text = "條　件"
grdDataList.ColWidth(4) = 4500
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 5: grdDataList.Text = "結果筆數"
grdDataList.ColWidth(5) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
End Sub

Public Sub PubShowNextData()
Select Case cmdState
Case 1 '列印
      Call StrPrintMenu
Case 2
      frm050207.Show
      Unload Me
Case 3
      Unload frm050207
      Unload Me
Case Else
End Select
End Sub

'報表
Sub StrPrintMenu()
Dim i As Long
Printer.Orientation = 2
'Printer.FontName = "標楷體"
If grdDataList.Rows - 1 > 0 Then
   PrintTitle
   For i = 1 To grdDataList.Rows - 1
      If grdDataList.RowHeight(i) > 0 Then 'Add By Sindy 2012/7/31
         For m_i = 1 To 6
             strTemp(m_i) = ""
         Next m_i
         
         strTemp(1) = Trim(grdDataList.TextMatrix(i, 0))
         strTemp(2) = Trim(grdDataList.TextMatrix(i, 1))
         strTemp(3) = Trim(grdDataList.TextMatrix(i, 2))
         strTemp(4) = Left(Trim(grdDataList.TextMatrix(i, 3)), 9)
         strTemp(5) = Left(Trim(grdDataList.TextMatrix(i, 4)), 64)
         strTemp(6) = Trim(grdDataList.TextMatrix(i, 5))
         
         If iLine > 36 Then
            Printer.NewPage
            iLine = 1
            PrintTitle '列印表頭
         End If
         PrintDetail
      End If
   Next i
Else
   MsgBox "無符合列印的資料!!!", vbExclamation + vbOKOnly
   Exit Sub
End If
Printer.EndDoc
ShowPrintOk
End Sub

Sub GetPleft()
PLeft(1) = 500
PLeft(2) = 1500
PLeft(3) = 2250
PLeft(4) = 3000
PLeft(5) = 5000
PLeft(6) = 15500
End Sub

Sub PrintTitle()
GetPleft

Printer.Font.Size = 12
Printer.Font.Underline = False
Printer.FontBold = False
Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("員工查詢印表記錄資料明細表") / 2)
Printer.CurrentY = 300
Printer.Print "員工查詢印表記錄資料明細表"
Printer.Font.Size = 10
Printer.Font.Underline = False
Printer.FontBold = False
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
Printer.CurrentY = 600
Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
Printer.CurrentY = 900
Printer.Print "頁　　次：" & Printer.Page
iLine = 4
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print "操作人員"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iLine * 300
Printer.Print "日期"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iLine * 300
Printer.Print "時間"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iLine * 300
Printer.Print "程式名稱"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iLine * 300
Printer.Print "條件"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iLine * 300
Printer.Print "結果筆數"
iLine = iLine + 1
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print String(255, "-")
iLine = iLine + 1
End Sub

Sub PrintDetail()
Dim m_j As Integer
For m_j = 1 To 6
   Printer.CurrentX = PLeft(m_j)
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(m_j)
Next m_j
iLine = iLine + 1
End Sub

Private Sub cmdok_Click(Index As Integer)
'92.04.16 nick 紀錄作用按鍵
cmdState = Index
PubShowNextData
Exit Sub
End Sub

Private Sub Form_Load()
   bolToEndByNick = False
   MoveFormToCenter Me
   SetDataListWidth
   '92.04.16 nick
   cmdState = -1
End Sub

Sub StrMenu()
Dim m_i As Integer
Dim rsTmp As New ADODB.Recordset
Dim strSR04 As String
Dim arrTemp As Variant
Dim arrTemp2 As Variant
Dim intRow As Integer, bolCheck As Boolean

Me.Enabled = False

strSQL1 = ""
'部門
If Len(frm050207.Combo1(0).Text) > 0 Then
   strSQL1 = strSQL1 + " and st03='" & Trim(Left(frm050207.Combo1(0).Text, 5)) & "' "
End If
'員工編號
If Len(frm050207.Combo1(1).Text) > 0 Then
   strSQL1 = strSQL1 + " and ql01='" & Trim(Left(frm050207.Combo1(1).Text, 7)) & "' "
End If
'不含離職人員
If frm050207.Check1.Value = 0 Then
   strSQL1 = strSQL1 + " and st04='1' "
End If
'操作日期
If Len(frm050207.txt1(0).Text) > 0 Then
   strSQL1 = strSQL1 + " and ql02 >= " & Val(Trim(frm050207.txt1(0).Text)) + 19110000
End If
If Len(frm050207.txt1(1).Text) > 0 Then
   strSQL1 = strSQL1 + " and ql02 <= " & Val(Trim(frm050207.txt1(1).Text)) + 19110000
End If
'操作時間
If Len(frm050207.txt1(2).Text) > 0 Then
   strSQL1 = strSQL1 + " and ql03 >= " & Trim(frm050207.txt1(2).Text)
End If
If Len(frm050207.txt1(3).Text) > 0 Then
   strSQL1 = strSQL1 + " and ql03 <= " & Trim(frm050207.txt1(3).Text)
End If
'程式名稱
If frm050207.Combo1(2).Text <> "" Then
   strSQL1 = strSQL1 + " and fo02 = '" & frm050207.Combo1(2).Text & "' "
End If
'系統類別
If Len(frm050207.txt1(4).Text) > 0 And UCase(frm050207.txt1(4).Text) <> "ALL" Then
   'strSQL1 = strSQL1 + " and INSTR(ql05||ql07,'系統類別：" & Trim(frm050207.txt1(4).Text) & "')>0"
   'Modify By Sindy 2012/7/31
   strSQL1 = strSQL1 + " and INSTR(ql05||ql07,'系統類別')>0"
   '2012/7/31 End
End If
'本所案號
If Len(frm050207.txtSystem.Text) > 0 And Len(frm050207.txtCode(0).Text) > 0 Then
   strSQL1 = strSQL1 + " and INSTR(ql05||ql07,'本所案號：" & Trim(frm050207.txtSystem.Text) & "-" & Trim(frm050207.txtCode(0).Text) & "-" & Trim(frm050207.txtCode(1).Text) & "-" & Trim(frm050207.txtCode(2).Text) & "')>0"
End If
'strSql = "select st02 AS 操作人員," & SQLDate("ql02", True) & " AS 日期,SqlTime(ql03) AS 時間,substr(fo03,1,4)||fo02 AS 程式名稱,ql05||ql07 AS 條件,ql06 AS 結果筆數 "
strSql = "select st02 AS 操作人員," & SQLDate("ql02", True) & " AS 日期,SqlTime(ql03) AS 時間,nvl(fo02,ql04) AS 程式名稱,ql05||ql07 AS 條件,ql06 AS 結果筆數 " & _
               " From querylog, staff, Form " & _
               " where ql01=st01(+) and ql04=fo01(+) " & strSQL1 & _
               " ORDER BY 1,2,3"
CheckOC
Dim StrTest1 As String, StrTest2 As String
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
   Set grdDataList.Recordset = adoRecordset
   'Modify By Sindy 2012/7/31
   If Len(frm050207.txt1(4).Text) > 0 And UCase(frm050207.txt1(4).Text) <> "ALL" Then
      For intRow = 1 To grdDataList.Rows - 1
         If InStr(UCase(grdDataList.TextMatrix(intRow, 4)), "系統類別：ALL") = 0 And _
            InStr(UCase(grdDataList.TextMatrix(intRow, 4)), "系統類別:ALL") = 0 Then
            '過濾系統類別
            intI = InStr(grdDataList.TextMatrix(intRow, 4), "系統類別")
            strSQL1 = Mid(Trim(grdDataList.TextMatrix(intRow, 4)), intI + 5, Len(Trim(grdDataList.TextMatrix(intRow, 4))))
            If InStr(strSQL1, ";") > 0 Then
               strSQL1 = Mid(strSQL1, 1, InStr(strSQL1, ";") - 1)
            End If
            arrTemp = Split(Trim(frm050207.txt1(4).Text), ",")
            arrTemp2 = Split(Trim(strSQL1), ",")
            bolCheck = False
            For i = 0 To UBound(arrTemp2)
               For j = 0 To UBound(arrTemp)
                  If arrTemp2(i) = arrTemp(j) Then
                     bolCheck = True
                     Exit For
                  End If
               Next j
               If bolCheck = True Then Exit For
            Next i
            If bolCheck = False Then
               grdDataList.RowHeight(intRow) = 0
            End If
         End If
      Next intRow
   End If
   '2012/7/31 End
Else
   ShowNoData
   Me.Enabled = True
   Screen.MousePointer = vbDefault
   tmpBol = fnCancelNowFormAndShowParentForm(Me)
   Exit Sub
End If
CheckOC
Me.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm050207_1 = Nothing
End Sub

Private Sub grdDataList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
getGrdColRow grdDataList, x, y, nCol, nRow
If nRow < 0 Then Exit Sub
grdDataList.col = nCol
grdDataList.row = nRow
If grdDataList.row = 0 And grdDataList.Rows > 1 Then
   Select Case Me.grdDataList.col
      Case 5
          If m_blnColOrderAsc = True Then
              Me.grdDataList.Sort = 3 '數字昇冪
              m_blnColOrderAsc = False
          Else
              Me.grdDataList.Sort = 4 '數字降冪
              m_blnColOrderAsc = True
          End If
      Case Else
          If m_blnColOrderAsc = True Then
              Me.grdDataList.Sort = 5 '字串昇冪
              m_blnColOrderAsc = False
          Else
              Me.grdDataList.Sort = 6 '字串降冪
              m_blnColOrderAsc = True
          End If
   End Select
End If
End Sub

Private Sub grdDataList_SelChange()
Dim TmpRow As Long
   TmpRow = grdDataList.MouseRow
   grdDataList.col = 0
   If TmpRow <> 0 Then
      Label2 = grdDataList.TextMatrix(TmpRow, 0)
      Label4 = grdDataList.TextMatrix(TmpRow, 1) & " " & grdDataList.TextMatrix(TmpRow, 2)
      Label6 = grdDataList.TextMatrix(TmpRow, 3)
      If grdDataList.TextMatrix(TmpRow, 5) = "" Then
         Label8 = "無法統計"
      Else
         Label8 = grdDataList.TextMatrix(TmpRow, 5)
      End If
      Text1 = grdDataList.TextMatrix(TmpRow, 4)
   Else
      Label2 = ""
      Label4 = ""
      Label6 = ""
      Label8 = ""
      Text1 = ""
   End If
   grdDataList.Visible = True
End Sub
