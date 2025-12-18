VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm100137 
   BorderStyle     =   1  '單線固定
   Caption         =   "查詢特殊置換字對照表"
   ClientHeight    =   5640
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4836
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   4836
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      Height          =   372
      Index           =   1
      Left            =   3696
      TabIndex        =   2
      Top             =   120
      Width           =   900
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "畫面更新(&E)"
      Height          =   372
      Index           =   0
      Left            =   2496
      TabIndex        =   1
      Top             =   120
      Width           =   1164
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MGrid1 
      Height          =   4860
      Left            =   168
      TabIndex        =   0
      Top             =   576
      Width           =   4500
      _ExtentX        =   7938
      _ExtentY        =   8573
      _Version        =   393216
      FixedCols       =   0
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   10.2
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frm100137"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create by Amy 2023/08/17
Option Explicit

Public UpForm As Form '前畫面
Dim rsMain As New ADODB.Recordset, strQ As String, intM As Integer
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序

Private Sub cmdOK_Click(Index As Integer)
   Select Case Index
      Case 0 '畫面更新
         Call doQuery
      Case 1 '結束
         tmpBol = fnCancelNowFormAndShowParentForm(Me)
   End Select
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   Call doQuery
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set rsMain = Nothing
   Set frm100137 = Nothing
End Sub

Private Sub doQuery()

   Call SetGrd(True) '清空
   
   strQ = "Select SW01,SW02,SW03||'.'||Decode(Sw03,'1','符號','2','文字',SW03) As MClass,SW03 " & _
               "From SpecWord Order by SW03 Desc,SW01 "
   intM = 1
   Set rsMain = ClsLawReadRstMsg(intM, strQ)
   If intM = 1 Then
      Set MGrid1.Recordset = rsMain
      Call SetGrd
   Else
      MsgBox "查無資料!!"
   End If
   
End Sub

Private Sub SetGrd(Optional ByVal pReset As Boolean = False)
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer

   arrGridHeadText = Array("置換字", "統一字", "類別", "SW03")
   arrGridHeadWidth = Array(1000, 1000, 1200, 0)
   
   MGrid1.Visible = False
   MGrid1.Cols = UBound(arrGridHeadText) + 1
   If pReset = True Then
      MGrid1.Clear
      MGrid1.Rows = 2
   End If
       
   For iRow = 0 To MGrid1.Cols - 1
      MGrid1.row = 0
      MGrid1.col = iRow
      MGrid1.Text = arrGridHeadText(iRow)
      MGrid1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      MGrid1.CellAlignment = flexAlignCenterCenter
   Next

   MGrid1.Visible = True
   
End Sub

Private Sub MGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim nCol As Long, nRow As Long
   
   getGrdColRow MGrid1, x, y, nCol, nRow
   If nCol < 0 Or nRow < 0 Then Exit Sub
   
   MGrid1.col = nCol
   MGrid1.row = nRow
   If Me.MGrid1.row < 1 Then
         If m_blnColOrderAsc = True Then
            Me.MGrid1.Sort = 5 '字串昇冪
            m_blnColOrderAsc = False
         Else
            Me.MGrid1.Sort = 6 '字串降冪
            m_blnColOrderAsc = True
         End If
   End If
End Sub
