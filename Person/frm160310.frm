VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm160310 
   BorderStyle     =   1  '單線固定
   Caption         =   "尾牙抽獎名條"
   ClientHeight    =   5796
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5040
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "新細明體-ExtB"
      Size            =   9
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5796
   ScaleWidth      =   5040
   Begin VB.TextBox txtStaffNo 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1536
      MaxLength       =   6
      TabIndex        =   3
      Top             =   1200
      Width           =   705
   End
   Begin VB.TextBox txtInput 
      Appearance      =   0  '平面
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1992
      TabIndex        =   11
      Top             =   3384
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.OptionButton Option1 
      Caption         =   "抽獎名條"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   360
      TabIndex        =   4
      Top             =   1656
      Width           =   1200
   End
   Begin VB.OptionButton Option1 
      Caption         =   "同仁名條"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Value           =   -1  'True
      Width           =   1200
   End
   Begin VB.TextBox txtZone 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1272
      MaxLength       =   1
      TabIndex        =   2
      Top             =   864
      Width           =   324
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "列印(&P)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   3000
      TabIndex        =   6
      Top             =   60
      Width           =   915
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "離開(&X)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   1
      Left            =   3945
      TabIndex        =   7
      Top             =   60
      Width           =   915
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   60
      TabIndex        =   8
      Top             =   5136
      Width           =   4875
      Begin VB.ComboBox cmbPrinter 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   276
         Left            =   705
         Style           =   2  '單純下拉式
         TabIndex        =   0
         Top             =   180
         Width           =   3870
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   75
         TabIndex        =   9
         Top             =   240
         Width           =   765
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3072
      Left            =   624
      TabIndex        =   5
      Top             =   1944
      Width           =   3792
      _ExtentX        =   6689
      _ExtentY        =   5419
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      AllowBigSelection=   0   'False
      FormatString    =   "獎項名稱|金額|張數"
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
   End
   Begin MSForms.Label lblName 
      Height          =   180
      Left            =   2328
      TabIndex        =   13
      Top             =   1248
      Width           =   1452
      Caption         =   "XXX"
      Size            =   "2561;317"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "員工編號："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   0
      Left            =   600
      TabIndex        =   12
      Top             =   1236
      Width           =   900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "所別：           (0:全所 1:北 2:中 3:南 4:高)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   624
      TabIndex        =   10
      Top             =   912
      Width           =   3120
   End
End
Attribute VB_Name = "frm160310"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Created by Morgan 2023/11/22
Option Explicit

Dim m_DefaultPrinter As String
Dim iRow As Integer, iCol As Integer '本次點選列數,行數
Dim iLstRow1 As Integer '前次點選列數1
Const clColorSel As Long = &HFFC0C0
Dim lX As Long, lY As Long, loX As Long, loY As Long, ldX As Long, ldY As Long

Private Sub cmdok_Click(Index As Integer)
   Select Case Index
      Case 0
         If Option1(0).Value = True Then
            If txtZone & txtStaffNo = "" Then
               MsgBox "請輸入列印條件!!", vbExclamation
               txtZone.SetFocus
               Exit Sub
            End If
            
         ElseIf Option1(1).Value = True Then
            With MSHFlexGrid1
            For intI = 1 To .Rows - 1
               If Val(.TextMatrix(intI, 2)) > 0 Then
                  Exit For
               End If
            Next
            If intI = .Rows Then
               MsgBox "請輸入列印張數!!", vbExclamation
               Exit Sub
            End If
            End With
         End If
         
         Screen.MousePointer = vbHourglass
         StrMenu
         PUB_RestorePrinter m_DefaultPrinter
         '若印表機變動, 則更新列印設定
         If cmbPrinter.Tag <> cmbPrinter Then
             PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.cmbPrinter.Name, 0, 0, Me.cmbPrinter.Text
         End If
         Screen.MousePointer = vbDefault
         
      Case 1
         Unload Me
   End Select
End Sub

'明細表
Sub StrMenu()
   Dim rsQuery As ADODB.Recordset
   Dim stCon As String
   'Modified by Morgan 2024/2/6 +排除第4碼是9的及抓薪資檔有基本薪資的(Ex:98099)
   stCon = " and st01>'6' and st01<'F' and substr(st03,1,1)<>'R' and substr(st01,4,1)<>'9'"
   If Option1(0).Value = True Then
      If txtZone <> "0" And txtZone <> "" Then stCon = stCon & " and st06='" & txtZone & "'"
      If txtStaffNo <> "" Then stCon = stCon & " and st01='" & txtStaffNo & "'"
      strExc(0) = "select st01,st02 from staff where st04='1'" & stCon & _
         " and exists(select * from salarydata where sd01=st01 and sd20>0)" & _
         " order by st93,st01"
      Set rsQuery = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 0 Then
         ShowPrintOk
      ElseIf intI = 1 Then
         If doPrint(rsQuery) = True Then
             ShowPrintOk
         End If
      End If
   Else
      doPrint2
   End If
   Set rsQuery = Nothing
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   PUB_SetPrinter Me.Name, cmbPrinter, m_DefaultPrinter
   LoadGrid1
   lblName = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm160310 = Nothing
End Sub

Private Sub MSHFlexGrid1_Click()
   With MSHFlexGrid1
   .row = .MouseRow
   .col = .MouseCol
   End With
   GridClick
End Sub

Private Sub txtStaffNo_Change()
   lblName = ""
   If Len(txtStaffNo) = 5 Then
      If ChkStaffID(txtStaffNo) = False Then
         If ClsPDGetStaffN(txtStaffNo, strExc(1)) = True Then
            lblName = strExc(1)
         End If
      End If
   End If
End Sub

Private Sub txtStaffNo_GotFocus()
   TextInverse txtStaffNo
End Sub

Private Sub txtStaffNo_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtInput_KeyDown(KeyCode As Integer, Shift As Integer)
   '向下
   With MSHFlexGrid1
   If KeyCode = 40 Then
      If .row < .Rows - 1 Then
         .TextMatrix(.row, .col) = Format(txtInput.Text, "#,##0")
         .row = .row + 1
         GridClick
      End If
   '向上
   ElseIf KeyCode = 38 Then
      If .row > 1 Then
         .TextMatrix(.row, .col) = Format(txtInput.Text, "#,##0")
         .row = .row - 1
         GridClick
      End If
   End If
   End With
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
   
   If Not ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 8 Or KeyAscii = vbKeyReturn Or KeyAscii = vbKeyEscape) Then
      KeyAscii = 0
      Beep
   Else
   
      If KeyAscii = vbKeyReturn Then
         With MSHFlexGrid1
         If .col = 2 Then
            .TextMatrix(iRow, iCol) = Format(txtInput.Text, "#,##0")
         Else
            .TextMatrix(iRow, iCol) = txtInput.Text
         End If
         If iCol > 0 And iCol < 2 Then
            .col = iCol + 1
            SetBox MSHFlexGrid1, txtInput
         ElseIf iCol = 2 Then
            If .row < .Rows - 1 Then
               .row = .row + 1
            Else
               .row = 1
            End If
            GridClick
         Else
            txtInput.Visible = False
         End If
         End With
      ElseIf KeyAscii = vbKeyEscape Then
         txtInput = txtInput.Tag
         TextInverse txtInput
      End If
      
   End If
End Sub

Private Sub GridClick()
   If Option1(1).Value = True Then
      With MSHFlexGrid1
      If .col = 2 Then
         SetBox MSHFlexGrid1, txtInput
      End If
   
      If .TextMatrix(.row, 0) <> "" Then
         SetGridColor MSHFlexGrid1, iLstRow1
         iLstRow1 = .row
      End If
      End With
   End If
End Sub

Private Sub txtInput_Validate(Cancel As Boolean)
   
   With MSHFlexGrid1
   If .col = 2 Then
      .TextMatrix(iRow, iCol) = Format(txtInput.Text, "#,##0")
   Else
      .TextMatrix(iRow, iCol) = txtInput.Text
   End If
   End With
   
   'Modified by Morgan 2023/11/22 改設定不可見位置,否則會自我觸發導致堆疊空間不足錯誤28
   'txtInput.Visible = False
   txtInput.Top = -1000
   'end 2023/11/22
End Sub

Private Sub txtUserNo_Change()

End Sub

Private Sub txtZone_GotFocus()
   TextInverse txtZone
End Sub

Private Sub txtZone_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("4")) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub SetGridHead1(Optional bolReset As Boolean = False)
   With MSHFlexGrid1
   .Visible = False
   If bolReset Then
      .Clear
      .Rows = 2
      .Cols = 3
      .col = 2
   End If
   .TextMatrix(0, 0) = "獎項名稱"
   .ColWidth(0) = 1500
   .ColAlignmentFixed(0) = flexAlignCenterCenter
   .ColAlignment(0) = flexAlignLeftCenter
   .TextMatrix(0, 1) = "金額"
   .ColWidth(1) = 1000
   .ColAlignmentFixed(1) = flexAlignCenterCenter
   .ColAlignment(1) = flexAlignRightCenter
   .TextMatrix(0, 2) = "張數"
   .ColWidth(2) = 1000
   .ColAlignmentFixed(2) = flexAlignCenterCenter
   .ColAlignment(2) = flexAlignRightCenter
   If .Rows = 1 Then .col = 2
   .Visible = True
   End With
End Sub

Private Sub LoadGrid1()
   Dim stSQL As String, intQ As Integer
   Dim rsQuery As ADODB.Recordset
   
   SetGridHead1 True
   stSQL = "select ac03,trim(to_char(ac10,'999,990')) ac10,0 Amt from allcode where ac01='14' order by ac02"
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      Set MSHFlexGrid1.Recordset = rsQuery.Clone
      SetGridHead1
   ElseIf intQ = 0 Then
      MsgBox "沒有獎項資料！", vbExclamation
   End If
   
End Sub

Private Sub SetBox(pGrid As MSHFlexGrid, pText As TextBox, Optional pValue As String = "")
   Dim ii As Integer
   Dim lngLeft As Long, lngTop As Long
   
   With pGrid
      If .row > 0 And .col > 0 Then
         pText.FontName = .CellFontName
         pText.FontSize = .CellFontSize
         pText.Alignment = .CellAlignment \ 5
         If pValue <> "" Then
            pText.Text = pValue
         Else
            pText.Text = Format(.TextMatrix(.row, .col))
         End If
         pText.Tag = pText.Text
         pText.Width = .ColWidth(.col)
         pText.Height = .RowHeight(.row)
         
         If .CellAlignment < 3 Then
            pText.Alignment = 0
         ElseIf .CellAlignment < 6 Then
            pText.Alignment = 2
         Else
            pText.Alignment = 1
         End If
         lngLeft = .Left + 25
         lngTop = .Top + .RowHeight(0) + 25
         For ii = 0 To .col - 1
            lngLeft = lngLeft + .ColWidth(ii)
         Next
         For ii = .TopRow To .row - 1
            lngTop = lngTop + .RowHeight(ii)
         Next
         pText.Left = lngLeft: pText.Top = lngTop
         pText.Visible = True
         If pText.Locked = False Then
            pText.SetFocus
            TextInverse pText
         End If
         iRow = .row: iCol = .col
      End If
   End With
End Sub

Private Sub SetGridColor(pGrid As MSHFlexGrid, pLstRow As Integer)
   Dim ii As Integer
   Dim lColor As Long
   Dim iRow As Integer
   
   With pGrid
   If pLstRow <> .row Then
      iRow = .row
      For ii = 0 To .Cols - 1
         .col = ii
         .CellBackColor = clColorSel
      Next
      If pLstRow > 0 Then
         .row = pLstRow
         For ii = 0 To .Cols - 1
            .col = ii
            .CellBackColor = .BackColor
         Next
      End If
      .row = iRow
   End If
   .Refresh
   End With
End Sub

Private Function doPrint(pRst As ADODB.Recordset) As Boolean
   
   Dim iPages As Integer
   Dim strText As String
   Dim iRows As Integer, iCols As Integer
   
On Error GoTo ErrHnd
   
   PUB_RestorePrinter cmbPrinter
   Printer.PaperSize = 9 'A4
   Printer.Orientation = 1 '直印
   Printer.FontName = "標楷體"
   
   loX = (Printer.Width - Printer.ScaleWidth) / 2 '上邊界(不可列印區)
   loY = (Printer.Height - Printer.ScaleHeight) / 2 '左邊界(不可列印區)
   
   iRows = 4
   iCols = 2
   PrintDotLine iRows, iCols '畫裁切線
   
   iPages = 0
   With pRst
   .MoveFirst
   Do While Not .EOF
      
      iPages = iPages + 1
      If iPages > iRows * iCols Then
         Printer.NewPage
         PrintDotLine iRows, iCols
         iPages = 1
      End If
      
      Printer.Font.Size = 72
      strText = .Fields("st02")
      'Modified by Morgan 2024/2/6
      'lX = (ldX - Printer.TextWidth(strText)) / 2 + ldX * ((iPages - 1) Mod iCols) - loX
      lX = (ldX - PUB_GetPrintWidth(strText)) / 2 + ldX * ((iPages - 1) Mod iCols) - loX
      'end 2024/2/6
      lY = ldY / 2 - Printer.TextHeight(strText) / 2 + ldY * ((iPages - 1) \ iCols) - loY
      PUB_PrintUnicodeText strText, lX, lY, 0

      Printer.Font.Size = 16
      strText = "員工編號：" & .Fields("st01") & " 金額：　　"
      lX = (ldX - Printer.TextWidth(strText)) / 2 + ldX * ((iPages - 1) Mod iCols) - loX
      Printer.CurrentX = lX
      Printer.CurrentY = Printer.CurrentY + 100
      Printer.Print strText
      .MoveNext
   Loop
   End With
   Printer.EndDoc
   doPrint = True
   Exit Function
   
ErrHnd:
   MsgBox Err.Description, vbCritical
   
End Function

Private Function doPrint2() As Boolean
   Dim ii As Integer, jj As Integer, kk As Integer
   Dim iPages As Integer
   Dim strText As String
   Dim iRows As Integer, iCols As Integer
   
   PUB_RestorePrinter cmbPrinter
   Printer.PaperSize = 9 'A4
   Printer.Orientation = 1 '直印
   Printer.FontName = "標楷體"
   
   loX = (Printer.Width - Printer.ScaleWidth) / 2 '上邊界(不可列印區)
   loY = (Printer.Height - Printer.ScaleHeight) / 2 '左邊界(不可列印區)
   
   iRows = 4
   iCols = 1
   PrintDotLine iRows, iCols '畫裁切線
   
   iPages = 0
   With MSHFlexGrid1
   For ii = 1 To .Rows - 1
      kk = Val(.TextMatrix(ii, 2))
      If kk > 0 Then
         For jj = 1 To kk
            iPages = iPages + 1
            If iPages > iRows * iCols Then
               Printer.NewPage
               PrintDotLine iRows, iCols
               iPages = 1
            End If
            
            Printer.Font.Size = 60
            strText = "恭喜您抽中" & .TextMatrix(ii, 0)
            lX = (ldX - Printer.TextWidth(strText)) / 2 + ldX * ((iPages - 1) Mod iCols) - loX
            lY = 650 + ldY * ((iPages - 1) \ iCols) - loY
            Printer.CurrentX = lX
            Printer.CurrentY = lY
            Printer.Print strText
            
            strText = "紅包" & .TextMatrix(ii, 1) & "元"
            lX = (ldX - Printer.TextWidth(strText)) / 2 + ldX * ((iPages - 1) Mod iCols) - loX
            Printer.CurrentX = lX
            Printer.Print strText
            
            Printer.Font.Size = 18
            strText = "員工編號：　　　姓名：　　"
            lX = (ldX - Printer.TextWidth(strText)) / 2 + ldX * ((iPages - 1) Mod iCols) - loX
            Printer.CurrentX = lX
            Printer.CurrentY = Printer.CurrentY + 200
            Printer.Print strText
         Next
      End If
   Next
   End With
   Printer.EndDoc
   doPrint2 = True
   Exit Function
   
ErrHnd:
   MsgBox Err.Description, vbCritical
End Function

Private Sub PrintDotLine(Optional pRows As Integer = 4, Optional pCols As Integer = 2)
   Dim iDrawStyle As Integer, iL As Integer
      
   iDrawStyle = Printer.DrawStyle
   
   ldX = Printer.Width / pCols
   ldY = Printer.Height / pRows
   
   Printer.DrawStyle = vbDot
   
   If pRows > 1 Then
      For iL = 1 To pRows - 1
         lY = iL * ldY - loY
         Printer.Line (0, lY)-(Printer.ScaleWidth, lY)
      Next
   End If
   
   If pCols > 1 Then
      For iL = 1 To pCols - 1
         lX = iL * ldX - loX
         Printer.Line (lX, 0)-(lX, Printer.ScaleHeight)
      Next
   End If
   
   Printer.DrawStyle = iDrawStyle
End Sub
