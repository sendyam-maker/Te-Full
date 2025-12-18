VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm04060304_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "國內公開後實審輸入"
   ClientHeight    =   5460
   ClientLeft      =   -465
   ClientTop       =   930
   ClientWidth     =   9330
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   9330
   Begin VB.CommandButton buttonClear 
      Caption         =   "清除查詢結果(&C)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   4920
      TabIndex        =   7
      Top             =   600
      Width           =   1500
   End
   Begin VB.CommandButton buttonQuery 
      Caption         =   "查詢(&F)"
      Height          =   400
      Left            =   7560
      TabIndex        =   4
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton buttonSearch 
      Caption         =   "同卷期多筆查詢(&S)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   3012
      TabIndex        =   6
      Top             =   600
      Width           =   1860
   End
   Begin VB.CommandButton buttonAdd 
      Caption         =   "新增(&A)"
      Default         =   -1  'True
      Height          =   400
      Left            =   5076
      TabIndex        =   1
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton buttonMod 
      Caption         =   "修改(&M)"
      Height          =   400
      Left            =   5904
      TabIndex        =   2
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton buttonDel 
      Caption         =   "刪除(&D)"
      Height          =   400
      Left            =   6732
      TabIndex        =   3
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton buttonExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8388
      TabIndex        =   5
      Top             =   70
      Width           =   800
   End
   Begin VB.TextBox textQuery 
      Height          =   264
      Left            =   1320
      MaxLength       =   11
      TabIndex        =   0
      Top             =   630
      Width           =   1452
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
      Height          =   4356
      Left            =   96
      TabIndex        =   9
      Top             =   1008
      Width           =   9108
      _ExtentX        =   16060
      _ExtentY        =   7673
      _Version        =   393216
      Cols            =   13
      FixedCols       =   0
      BackColorBkg    =   16772048
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      MergeCells      =   1
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
      _Band(0).Cols   =   13
      _Band(0).GridLinesBand=   2
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "公開編號 :"
      Height          =   180
      Left            =   120
      TabIndex        =   8
      Top             =   648
      Width           =   816
   End
End
Attribute VB_Name = "frm04060304_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/3 改成Form2.0 (grdList)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/12/28 申請案號欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
'Add by Morgan 2005/9/20
Option Explicit
Dim m_CurrSel As Integer
Dim m_TPG13 As String
Dim m_stMsg As String
Dim m_stTit As String
Dim m_Resp As VbMsgBoxResult

Private Sub Form_Load()
   MoveFormToCenter Me
   buttonClear_Click
End Sub

Private Sub InitialGrdList()
   FixGrid grdList
   grdList.row = 0
   grdList.col = 0
   grdList.Text = ""
   grdList.ColWidth(0) = 300
   grdList.col = 1
   grdList.Text = "公開編號"
   grdList.ColWidth(1) = 1200
   grdList.ColAlignment(1) = flexAlignLeftCenter
   grdList.col = 2
   grdList.Text = "申請案號"
   grdList.ColWidth(2) = 1000
   grdList.ColAlignment(2) = flexAlignCenterCenter
   grdList.col = 3
   grdList.Text = "實審申請日"
   grdList.ColWidth(3) = 1000
   grdList.ColAlignment(3) = flexAlignLeftCenter
   grdList.col = 4
   grdList.Text = "是否本人申請"
   grdList.ColWidth(4) = 1200
   grdList.ColAlignment(4) = flexAlignLeftCenter
   grdList.col = 5
   grdList.Text = "實審公開日"
   grdList.ColWidth(5) = 1000
   grdList.ColAlignment(5) = flexAlignCenterCenter
   grdList.col = 6
   grdList.Text = "實審公開卷期"
   grdList.ColWidth(6) = 1200
   grdList.ColAlignment(6) = flexAlignCenterCenter
   grdList.col = 7
   grdList.Text = "代理人"
   grdList.ColWidth(7) = 1000
   grdList.ColAlignment(7) = flexAlignLeftCenter
   grdList.col = 8
   grdList.Text = "本所案號"
   grdList.ColWidth(8) = 1200
   grdList.ColAlignment(8) = flexAlignLeftCenter
   grdList.col = 9
   grdList.Text = "事務所名稱"
   grdList.ColWidth(9) = 1200
   grdList.ColAlignment(9) = flexAlignLeftCenter
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm04060304_1 = Nothing
End Sub

Private Sub textQuery_GotFocus()
    TextInverse Me.textQuery
End Sub

Private Sub textQuery_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 檢查此筆資料是否存在
Private Function IsDataExist(ByVal p_TPG02 As String, Optional ByRef p_TPG13 As String) As Boolean
   
On Error GoTo ErrHnd

   p_TPG13 = ""
   CheckOC3
   With AdoRecordSet3
      strSql = "SELECT TPG13 FROM TPGAZETTE WHERE TPG02 = '" & p_TPG02 & "' AND TPG09='N'"
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenDynamic
      If .RecordCount > 0 Then
         IsDataExist = True
         If IsNull(.Fields("TPG13")) = False Then
            p_TPG13 = .Fields("TPG13")
         End If
      End If
   End With
   
ErrHnd:
   If Err.NUMBER <> 0 Then MsgBox Err.Description, vbCritical

End Function

' 按下新增按紐
Private Sub buttonAdd_Click()
   
   Dim bCancel As Boolean
   
   If Trim(textQuery) = "" Then
      m_stTit = "檢查公開編號"
      m_stMsg = "公開編號不可空白"
      m_Resp = MsgBox(m_stMsg, vbOKOnly, m_stTit)
      textQuery.SetFocus
      GoTo EXITSUB
   End If
    
   If IsDataExist(textQuery, m_TPG13) = False Then
      m_stTit = "新增資料"
      m_stMsg = "該公開編號未輸入公開無提實審資料！"
      m_Resp = MsgBox(m_stMsg, vbOKOnly + vbExclamation, m_stTit)
      
   ElseIf m_TPG13 <> "" Then
      m_stTit = "新增資料"
      m_stMsg = "該公開編號公開後實審資料已存在, 是否要修改公開編號"
      m_Resp = MsgBox(m_stMsg, vbYesNo + vbQuestion + vbDefaultButton2, m_stTit)
      '若取消作業
      If m_Resp = vbNo Then
         textQuery.SetFocus
         textQuery_GotFocus
         GoTo EXITSUB
      '若繼續作業
      Else
         '強制按下修改按鈕
         frm04060304_2.m_Force = True
         buttonMod_Click
      End If
      
   Else
      frm04060304_2.Show
      frm04060304_2.m_EditMode = strAdd
      frm04060304_2.m_Multi = False
      frm04060304_2.txtTPG02 = textQuery
      frm04060304_2.UpdateData
      Me.Hide
   End If
   
EXITSUB:

End Sub

' 按下修改按鈕
Public Sub buttonMod_Click()
   m_stTit = "修改資料"
   If IsDataExist(textQuery, m_TPG13) = True Then
      If m_TPG13 <> "" Then
         frm04060304_2.m_EditMode = strEdit
         If grdList.Rows > 3 Then
            frm04060304_2.m_Multi = True
         End If
         frm04060304_2.Show
         frm04060304_2.txtTPG02 = textQuery
         frm04060304_2.UpdateData
         Me.Hide
      Else
         m_stMsg = "該公開編號未輸入公開後實審資料！"
         m_Resp = MsgBox(m_stMsg, vbOKOnly + vbExclamation, m_stTit)
         textQuery.SetFocus
         textQuery_GotFocus
      End If
   Else
      m_stMsg = "該公開編號未輸入公開無提實審資料！"
      m_Resp = MsgBox(m_stMsg, vbOKOnly + vbExclamation, m_stTit)
      textQuery.SetFocus
      textQuery_GotFocus
   End If
End Sub

' 按下離開按紐
Private Sub buttonExit_Click()
   Unload frm04060304_2
   Unload Me
End Sub

' 按下多筆查詢按紐
Public Sub buttonSearch_Click()
   If Trim(textQuery) = "" Then
      m_stTit = "檢核資料"
      m_stMsg = "請輸入公開編號"
      m_Resp = MsgBox(m_stMsg, vbOKOnly, m_stTit)
   Else
      ExecuteQuery
   End If
End Sub

Private Sub ExecuteQuery()

   Screen.MousePointer = vbHourglass

On Error GoTo ErrHnd

   Dim i As Integer
   
   strSql = "SELECT '',T1.TPG02,T1.TPG01," & SQLDate("T1.TPG13") & ",T1.TPG14" & _
      "," & SQLDate("T1.TPG10") & ",T1.TPG11||'-'||T1.TPG12,TA03" & _
      "," & ChgPatent("", 1) & ",T1.TPG08 " & _
      "FROM TPGAZETTE T1, PATENT, TAGENT " & _
      "WHERE (T1.TPG11,T1.TPG12) in (SELECT T2.TPG11,T2.TPG12 FROM TPGAZETTE T2 WHERE T2.TPG02 = '" & textQuery & "' AND T2.TPG13>0 )" & _
      " AND T1.TPG13>0 AND PA11(+)=T1.TPG01 AND PA09(+)='000'" & _
      " AND TA02(+)=T1.TPG07 AND TA01(+)='P' ORDER BY TPG02 ASC"

   CheckOC3
   With AdoRecordSet3
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenForwardOnly, adLockReadOnly
      Set grdList.Recordset = .Clone
      InitialGrdList
      If .RecordCount > 0 Then
         For i = 0 To grdList.Rows - 1
            If InStr(grdList.TextMatrix(i, 1), textQuery) > 0 Then
               grdList.TopRow = i
               grdList.row = i
               m_CurrSel = 0
               grdList_ShowSelection
               Exit For
            End If
         Next
      Else
         m_stTit = "多筆查詢"
         m_stMsg = "無符合資料"
         m_Resp = MsgBox(m_stMsg, vbOKOnly, m_stTit)
         textQuery.SetFocus
         textQuery_GotFocus
      End If
   End With
      
ErrHnd:
   If Err.NUMBER <> 0 Then MsgBox Err.Description, vbCritical

   Screen.MousePointer = vbDefault
End Sub

Private Sub buttonClear_Click()
   m_CurrSel = 0
   InitGrid 10, grdList
   InitialGrdList
End Sub

Public Sub buttonQuery_Click()

   m_stTit = "查詢資料"
   If IsDataExist(textQuery, m_TPG13) = True Then
      If m_TPG13 <> "" Then
         frm04060304_2.m_EditMode = strFind
         If grdList.Rows > 3 Then
            frm04060304_2.m_Multi = True
         End If
         frm04060304_2.Show
         frm04060304_2.txtTPG02 = textQuery
         frm04060304_2.UpdateData
         Me.Hide
      Else
         m_stMsg = "該申請號未輸入公開後實審資料！"
         m_Resp = MsgBox(m_stMsg, vbOKOnly + vbExclamation, m_stTit)
         textQuery.SetFocus
         textQuery_GotFocus
      End If
   Else
      m_stMsg = "該申請號未輸入公開無提實審資料！"
      m_Resp = MsgBox(m_stMsg, vbOKOnly + vbExclamation, m_stTit)
      textQuery.SetFocus
      textQuery_GotFocus
   End If
   
End Sub

Public Sub grdList_SelChange()
   If grdList.row > 0 Then
      textQuery.Text = grdList.TextMatrix(grdList.row, 1)
   End If
   grdList_ShowSelection
End Sub

' 將GridList所選取的列反白, 並將未選取的列設成一般顏色
Private Sub grdList_ShowSelection()
   Dim nCurrSel As Integer
   Dim nCol As Integer
   
   nCurrSel = grdList.row
   
   ' 與前一選擇的列位置相同則不處理
   If m_CurrSel = grdList.row Then
      GoTo EXITSUB
   End If
   
   ' 將原先選取的列回復到正常的顏色
   If m_CurrSel > 0 And m_CurrSel < grdList.Rows Then
      grdList.row = m_CurrSel
      grdList.col = 1
      If grdList.CellBackColor <> &H80000005 Then
         For nCol = 1 To grdList.Cols - 1
            grdList.col = nCol
            If grdList.CellBackColor <> &H80000005 Then: grdList.CellBackColor = &H80000005
            If grdList.CellForeColor <> &H80000008 Then: grdList.CellForeColor = &H80000008
         Next nCol
      End If
      grdList.col = 0
   End If
   ' 設定成所選取的列
   m_CurrSel = nCurrSel
   ' 將所選取的列反白
   If m_CurrSel > 0 And m_CurrSel < grdList.Rows Then
      grdList.row = m_CurrSel
      grdList.col = 1
      For nCol = 1 To grdList.Cols - 1
         grdList.col = nCol
         grdList.CellBackColor = &H8000000D
         grdList.CellForeColor = &H80000005
      Next nCol
      grdList.col = 0
   End If
EXITSUB:
End Sub

' 按下刪除按紐
Public Sub buttonDel_Click()

   m_stTit = "刪除資料"
   ' 檢查該筆資料是否存在
   If IsDataExist(textQuery, m_TPG13) = True Then
      If m_TPG13 <> "" Then
         frm04060304_2.m_EditMode = strDel
         frm04060304_2.m_Multi = False
         frm04060304_2.Show
         frm04060304_2.txtTPG02 = textQuery
         frm04060304_2.UpdateData
         Me.Hide
      Else
         m_stMsg = "該申請號未輸入公開後實審資料！"
         m_Resp = MsgBox(m_stMsg, vbOKOnly + vbExclamation, m_stTit)
         textQuery.SetFocus
         textQuery_GotFocus
      End If
   Else
      m_stMsg = "該申請號未輸入公開資料！"
      m_Resp = MsgBox(m_stMsg, vbOKOnly + vbExclamation, m_stTit)
      textQuery.SetFocus
      textQuery_GotFocus
   End If
End Sub

Public Sub SetInputTPG02(Optional ByVal p_bEmpty As Boolean = True)
   textQuery.SetFocus
   If p_bEmpty Then
      textQuery = Empty
   Else
      textQuery_GotFocus
   End If
End Sub
