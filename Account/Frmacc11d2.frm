VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Frmacc11d2 
   AutoRedraw      =   -1  'True
   Caption         =   "拆收據金額分配"
   ClientHeight    =   3510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3510
   ScaleWidth      =   5175
   StartUpPosition =   3  '系統預設值
   Begin VB.TextBox Text5 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "新細明體-ExtB"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1125
      Locked          =   -1  'True
      MaxLength       =   14
      TabIndex        =   13
      Top             =   510
      Width           =   1572
   End
   Begin VB.TextBox Text7 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "新細明體-ExtB"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1125
      Locked          =   -1  'True
      MaxLength       =   14
      TabIndex        =   12
      Top             =   870
      Width           =   1572
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "新細明體-ExtB"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1125
      Locked          =   -1  'True
      MaxLength       =   9
      TabIndex        =   9
      Top             =   150
      Width           =   1572
   End
   Begin VB.CommandButton Command1 
      Caption         =   "確定"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   0
      Left            =   3195
      TabIndex        =   3
      Top             =   150
      Width           =   840
   End
   Begin VB.CommandButton Command1 
      Caption         =   "取消"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   1
      Left            =   4095
      TabIndex        =   2
      Top             =   150
      Width           =   840
   End
   Begin VB.TextBox txtInput 
      Appearance      =   0  '平面
      Height          =   375
      Left            =   2655
      TabIndex        =   1
      Text            =   "Text3"
      Top             =   2310
      Visible         =   0   'False
      Width           =   1635
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
      Height          =   1875
      Left            =   225
      TabIndex        =   0
      Top             =   1500
      Width           =   4740
      _ExtentX        =   8361
      _ExtentY        =   3307
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      HighLight       =   0
      AllowUserResizing=   1
      FormatString    =   "收據編號|服務費|規費|收據抬頭"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "服務費"
      BeginProperty Font 
         Name            =   "新細明體-ExtB"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   270
      TabIndex        =   11
      Top             =   510
      Width           =   810
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "規費"
      BeginProperty Font 
         Name            =   "新細明體-ExtB"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   270
      TabIndex        =   10
      Top             =   870
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "收文號"
      BeginProperty Font 
         Name            =   "新細明體-ExtB"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   270
      TabIndex        =   8
      Top             =   150
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "未分配服務費"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   270
      TabIndex        =   7
      Top             =   1230
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "未分配規費"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   2835
      TabIndex        =   6
      Top             =   1230
      Width           =   1050
   End
   Begin VB.Label lblService 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "999,999"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1665
      TabIndex        =   5
      Top             =   1230
      Width           =   690
   End
   Begin VB.Label lblFee 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "999,999"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4050
      TabIndex        =   4
      Top             =   1230
      Width           =   690
   End
End
Attribute VB_Name = "Frmacc11d2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/14 Form2.0已修改
'Memo By Sonia 2012/12/4 智權人員欄已修改
'Created by Morgan 2011/9/22
Option Explicit

Dim iRow As Integer, iCol As Integer


Private Sub Command1_Click(Index As Integer)
   If Index = 0 Then
      If TxtValidate = True Then
         Frmacc11d0.Tag = "Y"
      Else
         Exit Sub
      End If
   End If
   Unload Me
End Sub

Private Function TxtValidate() As Boolean
   Dim ii As Integer
   If Val(lblService) <> 0 Then
      MsgBox "[" & Label1(4) & "] 必須為 0！"
      Exit Function
   End If
   If Val(lblFee) <> 0 Then
      MsgBox "[" & Label1(5) & "] 必須為 0！"
      Exit Function
   End If
   
   With MSHFlexGrid2
   For ii = 1 To .Rows - 1
      '服務費規費都是0時要檢查該收據是否還有其他收文號
      If Val(.TextMatrix(ii, 1)) = 0 And Val(.TextMatrix(ii, 2)) = 0 Then
         strExc(0) = "select * from acc0j0 where a0j13='" & .TextMatrix(ii, 0) & "' and  a0j01<>'" & Text1 & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 0 Then
            MsgBox "收據 [" & .TextMatrix(ii, 0) & "] 只含本收文號，服務費規費不可同時為 0 !!"
            Exit Function
         End If
      End If
   Next
   End With
   TxtValidate = True
End Function

Private Sub Form_Activate()
   GridHead2
   SetRestValue
   txtInput.Visible = False
   MsgBox "若 [智權人員] 或 [是否合併] 欄位有修改將會一併更新相關收據編號!!", vbExclamation, "更新提醒"
End Sub

Private Sub Form_Load()
   PUB_InitForm Me, Me.Width, Me.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set Frmacc11d2 = Nothing
End Sub

Private Sub MSHFlexGrid2_Click()
Dim iCurCol As Integer, iCurRow As Integer
      
   With MSHFlexGrid2
   If .MouseRow > 0 And .MouseRow < .Rows And .MouseCol < .Cols Then
      iCurRow = .MouseRow
      iCurCol = .MouseCol
      .Visible = False
      If .col > 0 And .col < 3 Then SetBox
      .Visible = True
   End If
   End With
End Sub

Private Sub GridHead2()
   Dim ii As Integer
   With MSHFlexGrid2
      .Visible = False
      .FormatString = .FormatString
      .row = 0
      .col = 0
      .CellAlignment = flexAlignCenterCenter
      .ColWidth(0) = 1000
      .ColAlignment(0) = flexAlignLeftCenter
      .col = 1
      .CellAlignment = flexAlignCenterCenter
      .ColWidth(1) = 800
      .ColAlignment(1) = flexAlignRightCenter
      .col = 2
      .CellAlignment = flexAlignCenterCenter
      .ColWidth(2) = 800
      .ColAlignment(2) = flexAlignRightCenter
      .col = 3
      .CellAlignment = flexAlignCenterCenter
      .ColWidth(3) = 1850
      .ColAlignment(3) = flexAlignLeftCenter
      .Visible = True
   End With
End Sub

Private Sub SetBox(Optional pbolSetValue As Boolean = True)
   
   Dim lngLeft As Long, lngTop As Long
   Dim ii As Integer
   
   With MSHFlexGrid2
   txtInput.FontName = .CellFontName
   txtInput.FontSize = .CellFontSize
   If .CellAlignment < 3 Then
      txtInput.Alignment = 0 '靠左
   ElseIf .CellAlignment < 6 Then
      txtInput.Alignment = 2 '置中
   ElseIf .CellAlignment < 9 Then
      txtInput.Alignment = 1 '靠右
   Else
      txtInput.Alignment = 0 '靠左
   End If
   If pbolSetValue = True Then
      txtInput.Text = .TextMatrix(.row, .col)
   End If
   txtInput.Tag = txtInput.Text
   txtInput.Width = .ColWidth(.col) + 10
   txtInput.Height = .RowHeight(.row) - 5
   iRow = .row: iCol = .col
   lngLeft = .Left + 20
   lngTop = .Top + .RowHeight(0) + 20
   For ii = .LeftCol To .col - 1
      lngLeft = lngLeft + .ColWidth(ii)
   Next
   For ii = .TopRow To .row - 1
      lngTop = lngTop + .RowHeight(ii)
   Next
   txtInput.Left = lngLeft: txtInput.Top = lngTop
   If txtInput.Left + txtInput.Width < .Left + .Width Then
      txtInput.Visible = True
      txtInput.SetFocus
      TextInverse txtInput
   Else
      txtInput.Visible = False
   End If
   End With
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
   If KeyAscii = 8 Then Exit Sub
   
   If KeyAscii = vbKeyReturn Then
      If txtInput <> txtInput.Tag Then
         If iCol = 1 Or iCol = 2 Then
            MSHFlexGrid2.TextMatrix(iRow, iCol) = Val(txtInput.Text)
         Else
            MSHFlexGrid2.TextMatrix(iRow, iCol) = txtInput.Text
         End If
         MSHFlexGrid2.Recordset.Move iRow - 1, adBookmarkFirst
         MSHFlexGrid2.Recordset(iCol).Value = MSHFlexGrid2.TextMatrix(iRow, iCol)
         MSHFlexGrid2.Recordset.UpdateBatch
         SetRestValue
      End If
      GoNext
   ElseIf KeyAscii = vbKeyEscape Then
      txtInput = txtInput.Tag
      TextInverse txtInput
   '服務費規費欄位
   ElseIf iCol = 1 Or iCol = 2 Then
      If Not ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9"))) Then
         KeyAscii = 0
         Beep
         Exit Sub
      End If
   End If
End Sub

Private Sub GoNext()
   With MSHFlexGrid2
   If .col < 2 Then
      .col = .col + 1
   Else
      .col = 1
      If .row < .Rows - 1 Then
         .row = .row + 1
      Else
         .row = 1
      End If
   End If
   SetBox
   End With
End Sub

Private Sub SetRestValue()
   Dim lngService As Long, lngFee As Long
   
   lngService = Val(Text5)
   lngFee = Val(Text7)
   
   With MSHFlexGrid2
   For intI = 1 To .Rows - 1
      lngService = lngService - Val(.TextMatrix(intI, 1))
      lngFee = lngFee - Val(.TextMatrix(intI, 2))
   Next
   End With
   lblService = Format(lngService, "#,##0")
   lblFee = Format(lngFee, "#,##0")
End Sub

Private Sub txtInput_LostFocus()
   txtInput.Visible = False
   If txtInput <> txtInput.Tag Then
      If iCol = 1 Or iCol = 2 Then
         MSHFlexGrid2.TextMatrix(iRow, iCol) = Val(txtInput.Text)
      Else
         MSHFlexGrid2.TextMatrix(iRow, iCol) = txtInput.Text
      End If
      MSHFlexGrid2.Recordset.Move iRow - 1, adBookmarkFirst
      MSHFlexGrid2.Recordset(iCol).Value = MSHFlexGrid2.TextMatrix(iRow, iCol)
      MSHFlexGrid2.Recordset.UpdateBatch
      SetRestValue
   End If
End Sub
