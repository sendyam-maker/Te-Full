VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm010011 
   BorderStyle     =   1  '單線固定
   Caption         =   "主管機關來函查詢"
   ClientHeight    =   5715
   ClientLeft      =   210
   ClientTop       =   735
   ClientWidth     =   9315
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   9315
   Begin VB.CommandButton cmdOK 
      Caption         =   "查詢收文資料(&F)"
      CausesValidation=   0   'False
      Enabled         =   0   'False
      Height          =   400
      Index           =   0
      Left            =   5600
      TabIndex        =   0
      Top             =   70
      Width           =   1500
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   400
      Index           =   1
      Left            =   7124
      TabIndex        =   1
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   8352
      TabIndex        =   2
      Top             =   70
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   5028
      Left            =   120
      TabIndex        =   3
      Top             =   564
      Width           =   9072
      _ExtentX        =   16007
      _ExtentY        =   8864
      _Version        =   393216
      FixedCols       =   0
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      ScrollBars      =   2
      AllowUserResizing=   1
      RowSizingMode   =   1
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
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "合計共筆"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   780
   End
End
Attribute VB_Name = "frm010011"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/16 Form2.0已修改 grdDataList
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Modify By Sindy 2010/7/22 日期欄已修改
Option Explicit

'intLastRow上一次反白的Row
'blnOKtoShow決定是否要反白
Dim intLastRow As Integer, blnOKtoShow As Boolean
'intChoose被選擇之收文號總數
Dim intChoose As Integer
'intLeaveKind離開時，是0:結束1:回上一畫面
Dim intLeaveKind As Integer

Public Sub ClearOneRow()
Dim i As Integer

   intChoose = intChoose - 1
   If intChoose = 0 Then cmdOK(0).Enabled = False
   For i = 1 To grdDataList.Rows - 1
          If grdDataList.TextMatrix(i, 0) <> "" Then
             grdDataList.TextMatrix(i, 0) = ""
             Exit For
          End If
   Next
End Sub

Private Sub cmdOK_Click(Index As Integer)
Dim i As Integer

   If Index = 0 Then
      frm010002.bolIsQuery = True
      frm010002.Caption = Me.Caption
      Me.Hide
   Else
     If Index = 2 Then
        intLeaveKind = 0
     Else
        intLeaveKind = 1
     End If
     Unload Me
     intChoose = 0
   End If
End Sub

Private Sub Form_Activate()
   If grdDataList.Rows = 1 Then
      ShowMsg MsgText(1030)
      intLeaveKind = 1
      Unload Me
   End If
End Sub

Private Sub Form_Load()
Dim strSql As String, i As Integer, strKind1 As String, strKind2 As String, varSaveCursor

   MoveFormToCenter Me
   varSaveCursor = Screen.MousePointer
   Screen.MousePointer = vbHourglass
   For i = 0 To 4
          If frm010008.OptChoose(i) Then
             Exit For
          End If
   Next
   If i = 4 Then
      If frm010008.txtSystem = 馬德里案 Then
         strKind1 = frm010008.txtSystem + frm010008.txtTFCode(0) + IIf(frm010008.txtTFCode(1) = "", "0", frm010008.txtTFCode(1)) + IIf(frm010008.txtTFCode(2) = "", "0", frm010008.txtTFCode(2)) + IIf(frm010008.txtTFCode(3) = "", "00", frm010008.txtTFCode(3))
         strKind2 = frm010008.txtSystem + frm010008.txtTFCode(4) + IIf(frm010008.txtTFCode(5) = "", "0", frm010008.txtTFCode(5)) + IIf(frm010008.txtTFCode(6) = "", "0", frm010008.txtTFCode(6)) + IIf(frm010008.txtTFCode(7) = "", "00", frm010008.txtTFCode(7))
         
      Else
          strKind1 = frm010008.txtSystem + frm010008.txtCode(0) + IIf(frm010008.txtCode(1) = "", "0", frm010008.txtCode(1)) + IIf(frm010008.txtCode(2) = "", "00", frm010008.txtCode(2))
          strKind2 = frm010008.txtSystem + frm010008.txtCode(3) + IIf(frm010008.txtCode(4) = "", "0", frm010008.txtCode(4)) + IIf(frm010008.txtCode(5) = "", "00", frm010008.txtCode(5))
      End If
   Else
      strKind1 = frm010008.txtKeyIn(i * 2)
      strKind2 = frm010008.txtKeyIn(i * 2 + 1)
   End If
   'edit by nickc 2007/02/06 不用 dll 了
   'Set grdDataList.Recordset = obj001.ReadOrgRst(i, strKind1, strKind2)
   Set grdDataList.Recordset = Cls001ReadOrgRst(i, strKind1, strKind2, frm010008.Check1.Value)
   grdDataList.Refresh
   SetDataListWidth
   intLastRow = 0
   Me.Show
   If grdDataList.Rows > 1 Then
      ShowBar grdDataList, intLastRow, 11
   End If
   intChoose = 0
   Screen.MousePointer = varSaveCursor
End Sub

Private Sub SetDataListWidth()
Dim varGridWidth() As Variant

   varGridWidth = Array(300, 1000, 900, 400, 650, 0, 200, 250, 800, 800, 1200, 2300)
   SetGridDataListWidth grdDataList, varGridWidth()
   SetDataListVision grdDataList, True, True
   blnOKtoShow = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If intLeaveKind = 1 Then
      frm010008.Show
   Else
      Unload frm010008
   End If
   intLeaveKind = 0
   'Add By Cheng 2002/07/18
   Set frm010011 = Nothing
End Sub

Private Sub grdDataList_KeyPress(KeyAscii As Integer)
   If KeyAscii = 32 Then GrdDataList_Click
End Sub

Private Sub grdDataList_RowColChange()
   If intLastRow <> grdDataList.row Then
      If blnOKtoShow Then
         blnOKtoShow = False
         ShowBar grdDataList, intLastRow, 11
         blnOKtoShow = True
      End If
   End If
End Sub

Private Sub grdDataList_GotFocus()
   GridGotFocus grdDataList
End Sub

Private Sub grdDataList_LostFocus()
   GridLostFocus grdDataList
End Sub

Private Sub GrdDataList_Click()
   If grdDataList.TextMatrix(grdDataList.row, 0) = "ˇ" Then
      grdDataList.TextMatrix(grdDataList.row, 0) = ""
      intChoose = intChoose - 1
   Else
      grdDataList.TextMatrix(grdDataList.row, 0) = "ˇ"
      intChoose = intChoose + 1
   End If
   If intChoose = 0 Then cmdOK(0).Enabled = False Else cmdOK(0).Enabled = True
End Sub
