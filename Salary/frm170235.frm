VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm170235 
   BorderStyle     =   1  '單線固定
   Caption         =   "其他所得人補充保費查詢及列印"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9450
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   9450
   Begin VB.TextBox txtPayDate 
      Height          =   285
      Index           =   1
      Left            =   2520
      MaxLength       =   7
      TabIndex        =   3
      Top             =   450
      Width           =   915
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   2
      Left            =   1260
      MaxLength       =   1
      TabIndex        =   5
      Text            =   "1"
      Top             =   1110
      Width           =   375
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "查詢(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   7590
      Style           =   1  '圖片外觀
      TabIndex        =   6
      Top             =   120
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8490
      Style           =   1  '圖片外觀
      TabIndex        =   9
      Top             =   120
      Width           =   800
   End
   Begin VB.TextBox txtStaffNo 
      Height          =   285
      Left            =   1260
      MaxLength       =   6
      TabIndex        =   4
      Top             =   780
      Width           =   1365
   End
   Begin VB.TextBox txtPayDate 
      Height          =   285
      Index           =   0
      Left            =   1260
      MaxLength       =   7
      TabIndex        =   2
      Top             =   450
      Width           =   915
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   0
      Left            =   1305
      MaxLength       =   1
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   1
      Left            =   1980
      MaxLength       =   1
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "列印(&P)"
      Enabled         =   0   'False
      Height          =   400
      Left            =   6120
      Style           =   1  '圖片外觀
      TabIndex        =   8
      Top             =   120
      Width           =   800
   End
   Begin VB.Frame Frame1 
      Caption         =   "印表機"
      Height          =   570
      Left            =   4185
      TabIndex        =   7
      Top             =   4950
      Width           =   5070
      Begin VB.ComboBox cmbPrinter 
         Height          =   300
         Left            =   135
         Style           =   2  '單純下拉式
         TabIndex        =   10
         Top             =   210
         Width           =   4815
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   3525
      Left            =   90
      TabIndex        =   11
      Top             =   1380
      Width           =   9180
      _ExtentX        =   16193
      _ExtentY        =   6218
      _Version        =   393216
      BackColor       =   -2147483624
      Cols            =   7
      ScrollTrack     =   -1  'True
      SelectionMode   =   1
      AllowUserResizing=   1
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
   Begin VB.Line Line1 
      Index           =   0
      X1              =   2205
      X2              =   2520
      Y1              =   570
      Y2              =   570
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "查  詢  別：           ( 1.統計 2.明細 )"
      Height          =   180
      Left            =   315
      TabIndex        =   16
      Top             =   1170
      Width           =   2640
   End
   Begin VB.Label lblName 
      Height          =   180
      Left            =   2745
      TabIndex        =   15
      Top             =   825
      Width           =   2250
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "所得人編號："
      Height          =   180
      Left            =   135
      TabIndex        =   14
      Top             =   825
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "日　　期："
      Height          =   180
      Index           =   0
      Left            =   315
      TabIndex        =   13
      Top             =   480
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "公  司  別："
      Height          =   180
      Index           =   1
      Left            =   315
      TabIndex        =   12
      Top             =   180
      Width           =   900
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   1710
      X2              =   2025
      Y1              =   240
      Y2              =   240
   End
End
Attribute VB_Name = "frm170235"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Created by Morgan 2013/3/1
Option Explicit
Dim PLeft() As Integer, PColName() As String
Dim iPrint As Integer, iPage As Integer
Dim m_iStartX As Integer, m_iStartY As Integer, m_iColGap As Integer
Dim m_iPageHeight As Long, m_iLineHeight As Long, m_iMargin As Long
Dim stLstComp As String, stLstCompName As String

Dim m_DefaultPrinter As String

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdSearch_Click()
   Screen.MousePointer = vbHourglass
   If TxtValidate = True Then
      SetDataListWidth
      doQuery
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmdPrint_Click()
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   If cmbPrinter <> Printer.DeviceName Then
      PUB_RestorePrinter cmbPrinter
   End If
   PrintSheet
   '若印表機變動, 則更新列印設定
   If cmbPrinter.Tag <> cmbPrinter Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.cmbPrinter.Name, 0, 0, Me.cmbPrinter.Text
   End If
   If Printer.DeviceName <> m_DefaultPrinter Then
      PUB_RestorePrinter m_DefaultPrinter
   End If
   Me.Enabled = True
   Screen.MousePointer = vbDefault

End Sub

Private Sub PrintSheet()
   
   Dim ii As Integer, iCompCol As Integer, stLstGroup As String, iLstPage As Integer
   
   GetPleft
   
   With grdDataList
   stLstComp = .TextMatrix(1, 0)
   stLstCompName = CompNameQuery(stLstComp)
   
   iPage = 0
   PrintPageHeader
   PrintPageHeader1
   
   If txt1(2).Tag = "1" Then
      iCompCol = 4
   Else
      iCompCol = 0
   End If
   
   stLstGroup = .TextMatrix(1, 0)
   For ii = 1 To .Rows - 1
      If txt1(2).Tag = "2" Then
         If stLstGroup <> .TextMatrix(ii, 0) Then
            If stLstComp <> .TextMatrix(ii, 0) Then
               stLstComp = .TextMatrix(ii, 0)
               stLstCompName = CompNameQuery(stLstComp)
            End If
            
            Printer.NewPage
            PrintPageHeader
            PrintPageHeader1
            
            stLstGroup = .TextMatrix(ii, 0)
         End If
      End If
      
      iLstPage = iPage
      PrintNewLine
      If (.TextMatrix(ii, 1) = "" Or .TextMatrix(ii, 1) = "合計:") Then
         If iLstPage = iPage Then
            DrawLine
            PrintNewLine
         End If
         Printer.FontBold = True
      Else
         Printer.FontBold = False
      End If
      PrintDetail ii
      If (.TextMatrix(ii, 1) = "" Or .TextMatrix(ii, 1) = "合計:") Then
         PrintNewLine
         If txt1(2).Tag = "1" And ii <> .Rows - 1 Then
            PrintPageHeader1
         End If
      End If
   Next
   Printer.EndDoc
   End With
   MsgBox "列印完畢！"
End Sub

Private Sub PrintDetail(iRow As Integer)
    Dim iCol As Integer
    Printer.Font.Size = 11
    
    With grdDataList
    If txt1(2).Tag = "1" Then
      For iCol = 1 To UBound(PColName)
        Select Case iCol
        Case 3 To 4
           If .TextMatrix(iRow, iCol - 1) <> "" Then
              strExc(0) = Format(.TextMatrix(iRow, iCol - 1), DDollar2)
              Printer.CurrentX = PLeft(iCol + 1, 1) - Printer.TextWidth(strExc(0)) - m_iColGap
              Printer.CurrentY = iPrint
              Printer.Print strExc(0)
           End If
        Case Else
           Printer.CurrentX = PLeft(iCol, 1)
           Printer.CurrentY = iPrint
           Printer.Print .TextMatrix(iRow, iCol - 1)
        End Select
      Next
   Else
      For iCol = 1 To UBound(PColName)
        Select Case iCol
        Case 5 To 7
           If .TextMatrix(iRow, iCol) <> "" Then
              strExc(0) = Format(.TextMatrix(iRow, iCol), DDollar2)
              Printer.CurrentX = PLeft(iCol + 1, 1) - Printer.TextWidth(strExc(0)) - m_iColGap
              Printer.CurrentY = iPrint
              Printer.Print strExc(0)
           End If
        Case Else
           Printer.CurrentX = PLeft(iCol, 1)
           Printer.CurrentY = iPrint
           Printer.Print .TextMatrix(iRow, iCol)
        End Select
      Next
   End If
   End With
End Sub

   
Private Sub doQuery()
   Dim stVTB As String
   Dim stConNhi As String
   
   '公司
   If txt1(0) <> "" Then
      stConNhi = stConNhi & " and nhi11>='" & txt1(0) & "'"
   End If
   If txt1(0) <> "" Then
      stConNhi = stConNhi & " and nhi11<='" & txt1(1) & "'"
   End If
   
   '日期
   If txtPayDate(0) <> "" Then
      stConNhi = stConNhi & " and nhi02>=" & (txtPayDate(0) + 19110000)
   End If
   
   If txtPayDate(1) <> "" Then
      stConNhi = stConNhi & " and nhi02<=" & (txtPayDate(1) + 19110000)
   End If
   
   '所得人
   If txtStaffNo <> "" Then
      stConNhi = stConNhi & " and nhi01='" & txtStaffNo & "'"
   End If
    
   
   '紀錄最後條件
   txt1(0).Tag = txt1(0)
   txt1(1).Tag = txt1(1)
   txtPayDate(0).Tag = txtPayDate(0)
   txtPayDate(1).Tag = txtPayDate(1)
   txtStaffNo.Tag = txtStaffNo
   lblName.Tag = lblName
   txt1(2).Tag = txt1(2)
   
   '只要抓有補充保費的
   If txt1(2) = "1" Then
      strExc(0) = "select nhi11 公司,nhi03 格式,sum(nhi07) 所得總額,sum(nhi06) 補充保費,nhi11" & _
         " From nhi2nd" & _
         " Where nhi05 = 0" & stConNhi & _
         " group by nhi11,nhi03"
   Else
      strExc(0) = "select nhi11 公司,nhi01 所得人編號,nvl(st02,oi04) 名稱,sqldatet(nhi02) 日期,nhi03 格式,nhi07 所得總額,nhi08 費基,nhi06 補充保費" & _
         " From nhi2nd, staff, otherincomer" & _
         " Where nhi05 = 0" & stConNhi & _
         " and st01(+)=nhi01 and oi01(+)=nhi01" & _
         " order by 1,2,3,4"
   End If
   
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   Set grdDataList.Recordset = RsTemp.Clone
   SetDataListWidth True
   If intI = 1 Then
      AddSubTotal
      cmdPrint.Enabled = True
   Else
      cmdPrint.Enabled = False
   End If
   
End Sub

Private Sub AddSubTotal()
   Dim ii As Integer, jj As Integer, stComp As String, stGroup As String, stGroupName As String, strAddItem As String, lngColor As Long, lngColor1 As Long
   Dim dblSub1(9) As Double, dblTot(9) As Double
   
   lngColor = &H90EE90
   lngColor1 = &H90EE40
   
   
   With grdDataList
      .Visible = False
      '統計
      If txt1(2) = "1" Then
         stGroup = .TextMatrix(1, 0)
         ii = 1
         Do While ii < .Rows
            If stGroup <> .TextMatrix(ii, 0) Then
               strAddItem = stGroup & vbTab & "合計:"
               For jj = 2 To 3
                  strAddItem = strAddItem & vbTab & dblSub1(jj)
                  dblSub1(jj) = 0
               Next
               .AddItem strAddItem, ii
               
               .row = ii
               For jj = 0 To .Cols - 1
                  .col = jj: .CellBackColor = lngColor
               Next
      
               ii = ii + 1
               stGroup = .TextMatrix(ii, 0)
            End If
            For jj = 2 To 3
               dblSub1(jj) = dblSub1(jj) + Val(.TextMatrix(ii, jj))
            Next
            ii = ii + 1
         Loop
         strAddItem = stGroup & vbTab & "合計:"
         For jj = 2 To 3
            strAddItem = strAddItem & vbTab & dblSub1(jj)
            dblSub1(jj) = 0
         Next
         .AddItem strAddItem, ii
         
         .row = ii
         For jj = 0 To .Cols - 1
            .col = jj: .CellBackColor = lngColor
         Next
      '明細
      Else
         stComp = .TextMatrix(1, 0)
         stGroup = .TextMatrix(1, 1)
         stGroupName = .TextMatrix(1, 2)
         ii = 1
         Do While ii < .Rows
            If stGroup <> .TextMatrix(ii, 1) Then
               strAddItem = stComp & vbTab & vbTab & stGroupName & vbTab & "小計:" & vbTab
               For jj = 5 To 7
                  strAddItem = strAddItem & vbTab & dblSub1(jj)
                  dblTot(jj) = dblTot(jj) + dblSub1(jj)
                  dblSub1(jj) = 0
               Next
               .AddItem strAddItem, ii
               .row = ii
               For jj = 0 To .Cols - 1
                  .col = jj: .CellBackColor = lngColor
               Next
               ii = ii + 1
               
               'Added by Morgan 2013/3/14 +合計
               If stComp <> .TextMatrix(ii, 0) Then
                  strAddItem = stComp & vbTab & vbTab & vbTab & "合計:" & vbTab
                  For jj = 5 To 7
                     strAddItem = strAddItem & vbTab & dblTot(jj)
                     dblTot(jj) = 0
                  Next
                  .AddItem strAddItem, ii
                  .row = ii
                  For jj = 0 To .Cols - 1
                     .col = jj: .CellBackColor = lngColor1
                  Next
                  ii = ii + 1
               End If
               'end 2013/3/14
            
               stComp = .TextMatrix(ii, 0)
               stGroup = .TextMatrix(ii, 1)
               stGroupName = .TextMatrix(ii, 2)
            End If
               
            dblSub1(5) = dblSub1(5) + Val(.TextMatrix(ii, 5))
            dblSub1(6) = dblSub1(6) + Val(.TextMatrix(ii, 6))
            dblSub1(7) = dblSub1(7) + Val(.TextMatrix(ii, 7))
            
            ii = ii + 1
         Loop
         
         strAddItem = stComp & vbTab & vbTab & stGroupName & vbTab & "小計:" & vbTab
         For jj = 5 To 7
            strAddItem = strAddItem & vbTab & dblSub1(jj)
            dblTot(jj) = dblTot(jj) + dblSub1(jj)
            dblSub1(jj) = 0
         Next
         .AddItem strAddItem, ii
         
         .row = ii
         For jj = 0 To .Cols - 1
            .col = jj: .CellBackColor = lngColor
         Next
         
         ii = ii + 1
         strAddItem = stComp & vbTab & vbTab & vbTab & "合計:" & vbTab
         For jj = 5 To 7
            strAddItem = strAddItem & vbTab & dblTot(jj)
            dblTot(jj) = 0
         Next
         .AddItem strAddItem, ii
         .row = ii
         For jj = 0 To .Cols - 1
            .col = jj: .CellBackColor = lngColor1
         Next
         
      
      
      End If
      .Visible = True
   End With
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   PUB_SetPrinter Me.Name, cmbPrinter, m_DefaultPrinter
   
   If strUserNum = "86021" Then
      txt1(0) = "A": txt1(0).Enabled = False
      txt1(1) = "A": txt1(1).Enabled = False
   End If
   
   '預設最小未繳費月份
   'Modified by Morgan 2013/5/2 改15號以前預設上上月,以後預設上月
   'strExc(0) = "select min(nhi02) from nhi2nd where nhi12 is null"
   'intI = 1
   'Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   'If intI = 1 Then
   '   strExc(1) = TransDate(RsTemp(0), 1) \ 100
   '   txtPayDate(0) = strExc(1) & "01"
   '   txtPayDate(1) = TransDate(GetLastDay(txtPayDate(0)), 1)
   'End If
   If Val(Right(strSrvDate(1), 2)) > 15 Then
      strExc(1) = CompDate(1, -1, strSrvDate(1))
   Else
      strExc(1) = CompDate(1, -2, strSrvDate(1))
   End If
   strExc(2) = (strExc(1) \ 100) - 191100
   txtPayDate(0) = strExc(2) & "01"
   txtPayDate(1) = TransDate(GetLastDay(txtPayDate(0)), 1)
   'end 2013/5/2
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm170235 = Nothing
End Sub

Private Sub SetDataListWidth(Optional ByVal p_bolHeaderOnly As Boolean = False)
   Dim ii As Integer
   With grdDataList
      .Visible = False
      
      '統計
      If txt1(2) = "1" Then
         If p_bolHeaderOnly = False Then
            .Clear
            .Rows = 2: .Cols = 4: .FixedRows = 1: .FixedCols = 0
         End If
         
         .row = 0
         .col = 0: .ColWidth(.col) = 500: .Text = "公司"
         .ColAlignment(.col) = flexAlignCenterCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 1: .ColWidth(.col) = 500: .Text = "格式"
         .ColAlignment(.col) = flexAlignLeftCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 2: .ColWidth(.col) = 1200: .Text = "所得總額"
         .ColAlignment(.col) = flexAlignRightCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 3: .ColWidth(.col) = 1200: .Text = "補充保費"
         .ColAlignment(.col) = flexAlignRightCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .ColWidth(4) = 0
      '明細
      Else
         If p_bolHeaderOnly = False Then
            .Clear
            .Rows = 2: .Cols = 8: .FixedRows = 1: .FixedCols = 0
         End If
         
         .row = 0
         .col = 0: .ColWidth(.col) = 500: .Text = "公司"
         .ColAlignment(.col) = flexAlignCenterCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 1: .ColWidth(.col) = 1200: .Text = "所得人編號"
         .ColAlignment(.col) = flexAlignLeftCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 2: .ColWidth(.col) = 1500: .Text = "名稱"
         .ColAlignment(.col) = flexAlignLeftCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 3: .ColWidth(.col) = 900: .Text = "日期"
         .ColAlignment(.col) = flexAlignCenterCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 4: .ColWidth(.col) = 500: .Text = "格式"
         .ColAlignment(.col) = flexAlignCenterCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 5: .ColWidth(.col) = 1000: .Text = "所得總額"
         .ColAlignment(.col) = flexAlignRightCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 6: .ColWidth(.col) = 900: .Text = "費基"
         .ColAlignment(.col) = flexAlignRightCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 7: .ColWidth(.col) = 900: .Text = "補充保費"
         .ColAlignment(.col) = flexAlignRightCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      End If
      .Refresh
      .Visible = True
   End With
End Sub

Private Function TxtValidate() As Boolean
   Dim bCancel As Boolean
   
   If txt1(2) = "1" Then
      If txtPayDate(0) = "" Then
         MsgBox "請輸入年月!"
         txtPayDate(0).SetFocus
         Exit Function
      End If
   Else
      If txtStaffNo = "" And txtPayDate(0) = "" Then
         MsgBox "請輸入年月或所得人編號!"
         txtPayDate(0).SetFocus
         Exit Function
      End If
   End If
   
   If txt1(2) = "" Then
      MsgBox "請輸入統計別!"
      txt1(2).SetFocus
      Exit Function
   End If
   TxtValidate = True
End Function

Private Sub txt1_GotFocus(Index As Integer)
   TextInverse txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   Select Case Index
   Case 2
      If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") Then
         KeyAscii = 0
      End If
   End Select
End Sub

Private Sub txtPayDate_GotFocus(Index As Integer)
   TextInverse txtPayDate(Index)
   CloseIme
End Sub

Private Sub txtPayDate_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
   End If
End Sub

Private Sub txtStaffNo_Change()
   lblName = ""
End Sub

Private Sub txtStaffNo_GotFocus()
   TextInverse txtStaffNo
   CloseIme
End Sub

Private Sub txtStaffNo_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtStaffNo_Validate(Cancel As Boolean)
   If txtStaffNo <> "" Then
      lblName = StaffQuery(txtStaffNo)
   End If
End Sub

Private Sub GetPleft()
   Dim ii As Integer
   
   Printer.PaperSize = 9
   Printer.Orientation = 1
   Printer.FontSize = 12
   m_iStartX = 300
   m_iStartY = 500
   m_iPageHeight = Printer.ScaleHeight
   m_iLineHeight = 300
   m_iMargin = (Printer.Height - Printer.ScaleHeight) / 2
   m_iColGap = 150 '欄位間隔
   
   With grdDataList
   '統計
   If txt1(2) = "1" Then
      ReDim PLeft(5, 2)
      ReDim PColName(4)
      ii = 1
      PColName(ii) = .TextMatrix(0, ii - 1) '"公司"
      PLeft(ii, 1) = m_iStartX
      PLeft(ii, 2) = 2
      
      ii = ii + 1
      PColName(ii) = .TextMatrix(0, ii - 1) '"格式"
      PLeft(ii, 1) = PLeft(ii - 1, 1) + Printer.TextWidth(String(PLeft(ii - 1, 2), "　")) + m_iColGap
      PLeft(ii, 2) = 2
   
      ii = ii + 1
      PColName(ii) = .TextMatrix(0, ii - 1) '"所得總額"
      PLeft(ii, 1) = PLeft(ii - 1, 1) + Printer.TextWidth(String(PLeft(ii - 1, 2), "　")) + m_iColGap
      PLeft(ii, 2) = 5
      
      ii = ii + 1
      PColName(ii) = .TextMatrix(0, ii - 1) '"補充保費"
      PLeft(ii, 1) = PLeft(ii - 1, 1) + Printer.TextWidth(String(PLeft(ii - 1, 2), "　")) + m_iColGap
      PLeft(ii, 2) = 5
      
   Else
      ReDim PLeft(8, 2)
      ReDim PColName(7)
      ii = 1
      PColName(ii) = .TextMatrix(0, ii) '"所得人編號"
      PLeft(ii, 1) = m_iStartX
      PLeft(ii, 2) = 5
      
      ii = ii + 1
      PColName(ii) = .TextMatrix(0, ii) '"名稱"
      PLeft(ii, 1) = PLeft(ii - 1, 1) + Printer.TextWidth(String(PLeft(ii - 1, 2), "　")) + m_iColGap
      PLeft(ii, 2) = 4
      
      ii = ii + 1
      PColName(ii) = .TextMatrix(0, ii) '"日期"
      PLeft(ii, 1) = PLeft(ii - 1, 1) + Printer.TextWidth(String(PLeft(ii - 1, 2), "　")) + m_iColGap
      PLeft(ii, 2) = 5
      
      ii = ii + 1
      PColName(ii) = .TextMatrix(0, ii) '"格式"
      PLeft(ii, 1) = PLeft(ii - 1, 1) + Printer.TextWidth(String(PLeft(ii - 1, 2), "　")) + m_iColGap
      PLeft(ii, 2) = 2
   
      ii = ii + 1
      PColName(ii) = .TextMatrix(0, ii) '"所得總額"
      PLeft(ii, 1) = PLeft(ii - 1, 1) + Printer.TextWidth(String(PLeft(ii - 1, 2), "　")) + m_iColGap
      PLeft(ii, 2) = 5
      
      ii = ii + 1
      PColName(ii) = .TextMatrix(0, ii) '"費基"
      PLeft(ii, 1) = PLeft(ii - 1, 1) + Printer.TextWidth(String(PLeft(ii - 1, 2), "　")) + m_iColGap
      PLeft(ii, 2) = 5
      
      ii = ii + 1
      PColName(ii) = .TextMatrix(0, ii) '"補充保費"
      PLeft(ii, 1) = PLeft(ii - 1, 1) + Printer.TextWidth(String(PLeft(ii - 1, 2), "　")) + m_iColGap
      PLeft(ii, 2) = 5
   End If
   ii = ii + 1
   PLeft(ii, 1) = PLeft(ii - 1, 1) + Printer.TextWidth(String(PLeft(ii - 1, 2), "　")) + m_iColGap
   End With
End Sub


Private Sub PrintPageHeader()
   Dim strTmp As String
   
   If txt1(2).Tag = "1" Then
      strTmp = "台一關係企業 補充保費統計表(其他所得)"
   Else
      strTmp = "台一關係企業 補充保費明細表(其他所得)"
   End If
   
   iPrint = m_iStartY
   Printer.FontName = "細明體"
   Printer.Font.Size = 20
   Printer.Font.Bold = True
   Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(strTmp)) / 2
   Printer.CurrentY = iPrint
   Printer.Print strTmp
   iPrint = iPrint + 600
   
   Printer.Font.Size = 11
   Printer.Font.Bold = False
   
   strTmp = "公  司  別：" & txt1(0).Tag & " - " & txt1(1).Tag
   Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(String(10, "　"))) / 2
   Printer.CurrentY = iPrint
   Printer.Print strTmp
   
   iPrint = iPrint + m_iLineHeight
   strTmp = "日　　期：" & txtPayDate(0).Tag & " - " & txtPayDate(1).Tag
   Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(String(10, "　"))) / 2
   Printer.CurrentY = iPrint
   Printer.Print strTmp
   
   
   iPrint = iPrint + m_iLineHeight
   Printer.CurrentX = m_iStartX
   Printer.CurrentY = iPrint
   Printer.Print "列印人：" & strUserName
   
   
   strTmp = "員工編號：" & txtStaffNo.Tag & " " & lblName.Tag
   Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(String(10, "　"))) / 2
   Printer.CurrentY = iPrint
   Printer.Print strTmp
   
   Printer.CurrentX = Printer.ScaleWidth - m_iMargin - 2500
   Printer.CurrentY = iPrint
   Printer.Print "列印日期：" & Format(strSrvDate(2), "##/##/##")
   
   iPrint = iPrint + m_iLineHeight
   
   iPage = iPage + 1
   Printer.CurrentX = Printer.ScaleWidth - m_iMargin - 2500
   Printer.CurrentY = iPrint
   Printer.Print "頁    次：" & str(iPage)
   
   If txt1(2) = "2" Then
      iPrint = iPrint + m_iLineHeight
      Printer.CurrentX = m_iStartX
      Printer.CurrentY = iPrint
      Printer.Print "公司別：" & stLstComp & "　" & stLstCompName
   End If
   
   PrintNewLine
   
End Sub


Private Sub PrintPageHeader1()
   
   iPrint = iPrint + m_iLineHeight
   Printer.Font.Size = 11
   Printer.FontBold = True
   
   If txt1(2).Tag = "1" Then
      For intI = 1 To UBound(PColName)
         Select Case intI
         Case 3 To 4
            Printer.CurrentX = PLeft(intI + 1, 1) - Printer.TextWidth(PColName(intI)) - m_iColGap
            Printer.CurrentY = iPrint
            Printer.Print PColName(intI)
         Case Else
            Printer.CurrentX = PLeft(intI, 1)
            Printer.CurrentY = iPrint
            Printer.Print PColName(intI)
         End Select
      Next
   Else
      For intI = 1 To UBound(PColName)
         Select Case intI
         Case 5, 6, 7
            Printer.CurrentX = PLeft(intI + 1, 1) - Printer.TextWidth(PColName(intI)) - m_iColGap
            Printer.CurrentY = iPrint
            Printer.Print PColName(intI)
         Case Else
            Printer.CurrentX = PLeft(intI, 1)
            Printer.CurrentY = iPrint
            Printer.Print PColName(intI)
         End Select
      Next
   End If
   iPrint = iPrint + m_iLineHeight
   DrawLine
End Sub

Private Sub DrawLine()
   Printer.DrawWidth = 5
   Printer.Line (PLeft(1, 1), iPrint)-(PLeft(UBound(PLeft, 1), 1), iPrint)
   iPrint = iPrint - m_iLineHeight / 2
End Sub

Private Sub PrintNewLine(Optional ByVal p_iExtraLines As Integer = 2)

   iPrint = iPrint + m_iLineHeight
   If iPrint >= (m_iPageHeight - m_iMargin - p_iExtraLines * m_iLineHeight) Then
      DrawLine
      Printer.NewPage
      PrintPageHeader
      PrintPageHeader1
      iPrint = iPrint + m_iLineHeight
   End If
   
End Sub




