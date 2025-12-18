VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm140415_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "各項指示分類查詢"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8190
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   8190
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MGrid1 
      Height          =   3465
      Left            =   120
      TabIndex        =   14
      Top             =   2010
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   6112
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      AllowUserResizing=   3
      FormatString    =   "V|分類代號|說　　　　　明|使用部門|基本檔設定"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "結束(&X)"
      Height          =   360
      Left            =   7050
      TabIndex        =   2
      Top             =   150
      Width           =   915
   End
   Begin VB.CommandButton CmdSearch 
      Caption         =   "查詢(&Q)"
      Default         =   -1  'True
      Height          =   360
      Left            =   6120
      TabIndex        =   1
      Top             =   150
      Width           =   915
   End
   Begin VB.TextBox txtFind 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   4
      Left            =   1200
      MaxLength       =   1
      TabIndex        =   6
      Top             =   1560
      Width           =   585
   End
   Begin VB.TextBox txtFind 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   3
      Left            =   1200
      TabIndex        =   5
      Top             =   1155
      Width           =   6015
   End
   Begin VB.TextBox txtFind 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   2
      Left            =   2070
      MaxLength       =   2
      TabIndex        =   4
      Top             =   750
      Width           =   585
   End
   Begin VB.TextBox txtFind 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   1200
      MaxLength       =   2
      TabIndex        =   3
      Top             =   750
      Width           =   585
   End
   Begin VB.TextBox txtFind 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   1200
      MaxLength       =   1
      TabIndex        =   0
      Top             =   165
      Width           =   585
   End
   Begin VB.Label Label1 
      Caption         =   "模糊比對"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   5
      Left            =   7290
      TabIndex        =   13
      Top             =   1208
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(P專利、T商標)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   1860
      TabIndex        =   12
      Top             =   1620
      Width           =   1305
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   1860
      X2              =   2010
      Y1              =   930
      Y2              =   930
   End
   Begin VB.Label lblTitle 
      Caption         =   "(A通用、C通訊、D個案或特殊指示、E業拓、F財務、L法務、P專利、T商標、U承辦人)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   1860
      TabIndex        =   11
      Top             =   150
      Width           =   4020
   End
   Begin VB.Label Label1 
      Caption         =   "使用部門："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   3
      Left            =   120
      TabIndex        =   10
      Top             =   1613
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "說　　明："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   2
      Left            =   120
      TabIndex        =   9
      Top             =   1208
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "分類代號："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   803
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "分類部門："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   218
      Width           =   1035
   End
End
Attribute VB_Name = "frm140415_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2021/11/23 Form2.0已檢查 (無需修改的物件)
'Create by Lydia 2020/08/11 各項指示分類查詢
Option Explicit
Dim strKind(1 To 3) As String '分類第1碼: 改用Table控制
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSearch_Click()
Dim rsQuery As New ADODB.Recordset
Dim intR As Integer, strCon As String, stSQL As String

    If txtFind(0) <> "" Then strCon = strCon & " and IT01=" & CNULL(txtFind(0))
    
    If txtFind(1) <> "" And txtFind(2) <> "" Then
       If txtFind(1) > txtFind(2) Then
           MsgBox "分類代號起值不可大於迄值！", vbCritical, "檢核資料"
           txtFind(1).SetFocus
           txtFind_GotFocus (1)
           Exit Sub
       End If
    End If
    If txtFind(1) <> "" Then strCon = strCon & " and IT02>=" & CNULL(txtFind(1))
    If txtFind(2) <> "" Then strCon = strCon & " and IT02<=" & CNULL(txtFind(2))
    
    If txtFind(3) <> "" Then
       If txtFind(3) = "%" Then  'ex.折扣率%
           strCon = strCon & " and REPLACE(IT03,'%','％') LIKE '%％'"
       Else
           strCon = strCon & " and IT03 LIKE '%" & UCase(Trim(txtFind(3))) & "%'"
       End If
    End If
    '使用部門
    If txtFind(4) <> "" Then
        strCon = strCon & " and (IT10 IS NULL OR IT10=" & CNULL(txtFind(4)) & ") "
    End If

    stSQL = "SELECT IT01||DECODE(IT01," & GetAddStr(strKind(3)) & ",IT01) TYPE," & _
                     "IT01,IT02,IT03,IT10||DECODE(IT10,'P','專利','T','商標','') AS IT10,DECODE(IT11,NULL,NULL,'Y') IT11T " & _
                     "From INSTTYPE " & IIf(strCon <> "", "WHERE" & Mid(strCon, 5), "")
    stSQL = stSQL & " ORDER BY IT01,IT02 "
    
    Call SetGrd(True) '清空
    
    intR = 0
    Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
    If intR = 1 Then
       Set MGrid1.Recordset = rsQuery
       Call SetGrd
    End If
    
    Set rsQuery = Nothing
    Exit Sub
    
ErrorHand2:
   If Err.Number > 0 Then
      MsgBox Err.Description
      Exit Sub
   End If
         
End Sub

Private Sub SetGrd(Optional ByVal pReset As Boolean = False)
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer

    arrGridHeadText = Array("分類部門", "IT01", "分類代號", "說　　明", "使用部門", "基本檔")
    arrGridHeadWidth = Array(1000, 0, 1000, 3200, 1000, 800)
   
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

    For iRow = 1 To MGrid1.Rows - 1
         MGrid1.row = iRow
         MGrid1.col = 5 '基本檔設定：置中
         MGrid1.CellAlignment = flexAlignCenterCenter
    Next iRow
    MGrid1.Visible = True
   
End Sub

Private Sub Form_Load()

   MoveFormToCenter Me

   strKind(1) = PUB_GetInType("1")
   strKind(2) = PUB_GetInType("2")
   strKind(3) = PUB_GetInType("3")
   lblTitle.Caption = "(" & strKind(2) & ")"
   
   Call cmdSearch_Click  '預設全部資料
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm140415_1 = Nothing
End Sub

Private Sub txtFind_GotFocus(Index As Integer)
   TextInverse txtFind(Index)
End Sub

Private Sub txtFind_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
        Case 0, 4
              KeyAscii = UpperCase(KeyAscii)
   End Select
End Sub

Private Sub txtFind_Validate(Index As Integer, Cancel As Boolean)
Dim iLen As Integer

   Cancel = False
   If txtFind(Index) = "" Then Exit Sub
   
   Select Case Index
        Case 0 '分類部門
           If InStr(strKind(1), txtFind(Index)) = 0 Or Trim(txtFind(Index)) = "、" Then
               MsgBox "請輸入" & strKind(1) & " !", vbCritical, "輸入錯誤"
               GoTo JumpSet
           End If
        Case 1, 2 '分類代號(起迄)
           If Val(txtFind(Index)) < 0 Or Val(txtFind(Index)) > 99 Then
               MsgBox "請輸入00~99!", vbCritical, "輸入錯誤"
               GoTo JumpSet
           End If
        Case 3 '說明
            txtFind(Index).Text = PUB_StringFilter(txtFind(Index).Text)  'Added by Lydia 2020/05/14 清除字串中的enter
        Case 4 '使用部門IT10
            If txtFind(Index) <> "P" And txtFind(Index) <> "T" Then
                MsgBox "請輸入P、T !", vbCritical, "輸入錯誤"
                GoTo JumpSet
            End If
   End Select
   
   If Not CheckLengthIsOK(txtFind(Index), iLen) Then
      Cancel = True
   End If
   
   Exit Sub

JumpSet:
   txtFind(Index).SetFocus
   txtFind_GotFocus (Index)
   Cancel = True
End Sub

Private Sub MGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
   getGrdColRow MGrid1, x, y, nCol, nRow
   If nCol < 0 Or nRow < 0 Then Exit Sub
   MGrid1.col = nCol
   MGrid1.row = nRow
   If Me.MGrid1.row < 1 And Me.MGrid1.Text <> "V" Then
      If InStr("分類代號", Me.MGrid1.Text) > 0 Then
         If m_blnColOrderAsc = True Then
            Me.MGrid1.Sort = 3  '數值昇冪
            m_blnColOrderAsc = False
         Else
            Me.MGrid1.Sort = 4 '數值降冪
            m_blnColOrderAsc = True
         End If
      Else
         If m_blnColOrderAsc = True Then
            Me.MGrid1.Sort = 5 '字串昇冪
            m_blnColOrderAsc = False
         Else
            Me.MGrid1.Sort = 6 '字串降冪
            m_blnColOrderAsc = True
         End If
      End If
   End If
End Sub
