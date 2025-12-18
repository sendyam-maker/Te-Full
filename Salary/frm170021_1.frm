VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm170021_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "健保費明細"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7545
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   7545
   StartUpPosition =   1  '所屬視窗中央
   Begin VB.CommandButton cmdExit 
      Caption         =   "確定(&O)"
      Height          =   405
      Index           =   2
      Left            =   3870
      TabIndex        =   5
      Top             =   60
      Width           =   1125
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "取消(&C)"
      Height          =   405
      Index           =   1
      Left            =   5040
      TabIndex        =   4
      Top             =   60
      Width           =   1125
   End
   Begin VB.ComboBox cboHL05 
      Appearance      =   0  '平面
      Height          =   300
      Left            =   1035
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   90
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.TextBox txtInput 
      Appearance      =   0  '平面
      Height          =   375
      Left            =   180
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   60
      Width           =   825
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "回前畫面(&U)"
      Height          =   405
      Index           =   0
      Left            =   6210
      TabIndex        =   1
      Top             =   60
      Width           =   1125
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   2535
      Left            =   180
      TabIndex        =   0
      Top             =   570
      Width           =   7150
      _ExtentX        =   12621
      _ExtentY        =   4471
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   5
      FixedCols       =   0
      ForeColorSel    =   16777215
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      FormatString    =   "姓名|稱謂|健保費|超額眷口|補助說明"
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
      _Band(0).Cols   =   5
   End
End
Attribute VB_Name = "frm170021_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/20 Form2.0已修改
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Memo by Morgan 2010/12/2 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
'Create by Morgan 2009/6/23
Option Explicit

Public strSM01 As String
Public strSM02 As String
Public m_EditMode As Integer '1:新增 2:修改 3:刪除 4:查詢

Dim iRow As Integer
Dim iCol As Integer


Private Sub cboHL05_Click()
   GRD1.TextMatrix(iRow, iCol) = cboHL05.Text
   GRD1.Recordset.Move iRow - 1, 1
   GRD1.Recordset.Fields(iCol) = GRD1.TextMatrix(iRow, iCol)
   txtInput.Visible = False
End Sub

Private Sub cmdExit_Click(Index As Integer)
   If Index = 2 Then
      If txtInput.Visible = True Then
         txtInput_KeyPress vbKeyReturn
      End If
      GRD1.Recordset.UpdateBatch
   End If
   Unload Me
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm170021_1 = Nothing
End Sub

Public Sub SetGrid(p_AdoRst As ADODB.Recordset)
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iCol As Integer
   
   Set GRD1.Recordset = p_AdoRst
      
   arrGridHeadText = Array("姓名", "稱謂", "健保費", "超額眷口", "補助說明", "序")
   arrGridHeadWidth = Array(850, 600, 820, 850, 3650, 0)
   GRD1.Visible = False
   GRD1.Cols = UBound(arrGridHeadText) + 1
   For iCol = 0 To GRD1.Cols - 1
      GRD1.row = 0
      GRD1.col = iCol
      GRD1.Text = arrGridHeadText(iCol)
      GRD1.ColWidth(iCol) = arrGridHeadWidth(iCol)
      GRD1.CellAlignment = flexAlignCenterCenter
   Next
   GRD1.ColAlignment(1) = flexAlignCenterCenter
   GRD1.ColAlignment(2) = flexAlignRightCenter
   GRD1.ColAlignment(3) = flexAlignCenterCenter
   GRD1.Visible = True
   
   txtInput.Visible = False
   SetcboHL05
   If m_EditMode <> 2 Then
      cmdExit(1).Visible = False
      cmdExit(2).Visible = False
   Else
      cmdExit(0).Visible = False
   End If
End Sub

Private Sub GRD1_Click()
   If m_EditMode <> 2 Then Exit Sub
   With GRD1
      .row = .MouseRow
      .col = .MouseCol
      SetBox
   End With
End Sub

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim nCol As Long, nRow As Long
   getGrdColRow GRD1, x, y, nCol, nRow
   GRD1.col = nCol
   GRD1.row = nRow
End Sub

Private Sub grd1_SelChange()
   Dim tmpMouseRow
   Dim i, j
   
   tmpMouseRow = GRD1.row
   GRD1.Visible = True
   If tmpMouseRow <> 0 Then
       GRD1.row = tmpMouseRow
       GRD1.col = 0
       If GRD1.CellBackColor = QBColor(15) Then
             GRD1.Visible = False
             For j = 1 To GRD1.Rows - 1
                 GRD1.row = j
                 For i = 0 To GRD1.Cols - 1
                      GRD1.col = i
                      GRD1.CellBackColor = QBColor(15)
                 Next i
            Next j
            GRD1.row = tmpMouseRow
            For i = 0 To GRD1.Cols - 1
                GRD1.col = i
                GRD1.CellBackColor = &HFFC0C0
            Next i
            GRD1.Visible = True
       End If
   End If
End Sub

Private Sub SetBox()
   
   Dim lngLeft As Long, lngTop As Long
   Dim ii As Integer
   
   With GRD1
      If .row > 0 Then
         Select Case .col
         Case 2
            cboHL05.Visible = False
            txtInput.FontName = .CellFontName
            txtInput.FontSize = .CellFontSize
            txtInput.Alignment = .CellAlignment \ 5
            txtInput.Text = .TextMatrix(.row, .col)
            txtInput.Tag = txtInput.Text
            txtInput.Width = .ColWidth(.col)
            txtInput.Height = .RowHeight(.row)
            iRow = .row: iCol = .col
            txtInput.Visible = True
            txtInput.SetFocus
            TextInverse txtInput
            lngLeft = .Left + 25
            lngTop = .Top + .RowHeight(0) + 25
            For ii = 0 To .col - 1
               lngLeft = lngLeft + .ColWidth(ii)
            Next
            For ii = .TopRow To .row - 1
               lngTop = lngTop + .RowHeight(ii)
            Next
            txtInput.Left = lngLeft: txtInput.Top = lngTop
         Case 4
            txtInput.Visible = False
            cboHL05.FontName = .CellFontName
            cboHL05.FontSize = .CellFontSize
            cboHL05.Text = .TextMatrix(.row, .col)
            cboHL05.Tag = cboHL05.ListIndex
            cboHL05.Width = .ColWidth(.col)
            'cboHL05.Height = .RowHeight(.row)
            iRow = .row: iCol = .col
            cboHL05.Visible = True
            cboHL05.SetFocus
            
            lngLeft = .Left + 25
            lngTop = .Top + .RowHeight(0) + 25
            For ii = 0 To .col - 1
               lngLeft = lngLeft + .ColWidth(ii)
            Next
            For ii = .TopRow To .row - 1
               lngTop = lngTop + .RowHeight(ii)
            Next
            cboHL05.Left = lngLeft: cboHL05.Top = lngTop
         Case Else
         
         End Select
      End If
   End With
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
   If Not ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = Asc(".") Or KeyAscii = 8 Or KeyAscii = vbKeyReturn Or KeyAscii = vbKeyEscape) Then
      KeyAscii = 0
      Beep
   Else
      If KeyAscii = vbKeyReturn Then
         UpdateData
         txtInput.Visible = False
      ElseIf KeyAscii = vbKeyEscape Then
         txtInput = txtInput.Tag
         TextInverse txtInput
      End If
   End If
End Sub

Private Sub SetcboHL05()
   cboHL05.Clear
   cboHL05.AddItem "無"
   strSql = "select HR01||' '||HR04 from HiReduce order by 1"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      With RsTemp
      Do While Not .EOF
         cboHL05.AddItem "" & .Fields(0).Value
         .MoveNext
      Loop
      End With
   End If
End Sub

Private Sub txtInput_Validate(Cancel As Boolean)
   If txtInput.Text <> txtInput.Tag Then
      UpdateData
   End If
End Sub

Private Sub UpdateData()
   GRD1.TextMatrix(iRow, iCol) = Format(txtInput.Text, "#")
   GRD1.Recordset.Move iRow - 1, 1
   GRD1.Recordset.Fields(iCol) = GRD1.TextMatrix(iRow, iCol)
End Sub
