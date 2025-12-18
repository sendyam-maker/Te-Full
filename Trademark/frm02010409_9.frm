VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm02010409_9 
   BorderStyle     =   1  '單線固定
   Caption         =   "服務業務結果輸入(正片號碼)"
   ClientHeight    =   5664
   ClientLeft      =   276
   ClientTop       =   996
   ClientWidth     =   9132
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5664
   ScaleWidth      =   9132
   Begin VB.TextBox textSPKey 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   660
      Width           =   2292
   End
   Begin VB.TextBox textSP22 
      Height          =   264
      Left            =   1200
      MaxLength       =   13
      TabIndex        =   0
      Top             =   960
      Width           =   2292
   End
   Begin VB.TextBox textSP23 
      Height          =   264
      Left            =   1200
      MaxLength       =   1
      TabIndex        =   1
      Top             =   1260
      Width           =   612
   End
   Begin VB.CommandButton cmdMod 
      Caption         =   "修改(&M)"
      Height          =   400
      Left            =   6144
      TabIndex        =   2
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6972
      TabIndex        =   3
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   7800
      TabIndex        =   4
      Top             =   70
      Width           =   1200
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
      Height          =   4032
      Left            =   96
      TabIndex        =   10
      Top             =   1560
      Width           =   8952
      _ExtentX        =   15790
      _ExtentY        =   7112
      _Version        =   393216
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
      Caption         =   "本所案號 :"
      Height          =   252
      Left            =   120
      TabIndex        =   8
      Top             =   660
      Width           =   1092
   End
   Begin VB.Label Label31 
      Caption         =   "正片號碼 :"
      Height          =   252
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   1092
   End
   Begin VB.Label Label32 
      Caption         =   "是否合格 :"
      Height          =   252
      Left            =   120
      TabIndex        =   6
      Top             =   1260
      Width           =   1092
   End
   Begin VB.Label Label33 
      Caption         =   "( Y:合格 N:不合格 )"
      Height          =   252
      Left            =   1920
      TabIndex        =   5
      Top             =   1260
      Width           =   1572
   End
End
Attribute VB_Name = "frm02010409_9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/10/18 Form2.0已修改 ; grdList從MSFlexGrid改為MSHFlexGrid
'Memo by Amy 2021/12/29 Form2.0已修改 grdList
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/5 日期欄已修改
Option Explicit

' 本所案號
Dim m_SP01 As String
Dim m_SP02 As String
Dim m_SP03 As String
Dim m_SP04 As String


' 宣告
Private Type SITEM
   Name As String
   Value As String
   '911016 nick 新增
   strCP09 As String
End Type
Dim m_SItemList() As SITEM
Dim m_SItemCount As Integer

Private Sub Form_Load()
   MoveFormToCenter Me
   textSPKey.BackColor = &H8000000F
End Sub

' 清除串列
Private Sub ClearSItem()
   If m_SItemCount > 0 Then
      Erase m_SItemList
   End If
   m_SItemCount = 0
End Sub

' 新增一個項目
'911017 nick
'Private Sub AddSItem(ByVal strName As String, ByVal strItem As String)
Private Sub AddSItem(ByVal strName As String, ByVal strItem As String, ByVal strCP09 As String)
   Dim nPos As Integer
   Dim bFind As Boolean
   bFind = False
   For nPos = 0 To m_SItemCount - 1
      If m_SItemList(nPos).Name = strName Then
         m_SItemList(nPos).Value = strItem
         '911017 nick
         m_SItemList(nPos).strCP09 = strCP09
         bFind = True
         Exit For
      End If
   Next nPos
   
   If bFind = False Then
      ReDim Preserve m_SItemList(m_SItemCount + 1)
      m_SItemList(m_SItemCount).Name = strName
      m_SItemList(m_SItemCount).Value = strItem
      '911017 nick
      m_SItemList(m_SItemCount).strCP09 = strCP09
      m_SItemCount = m_SItemCount + 1
   End If
End Sub
' 設定一個項目
Private Sub SetSItem(ByVal strName As String, ByVal strItem As String)
   Dim nPos As Integer
   Dim bFind As Boolean
   For nPos = 0 To m_SItemCount - 1
      If m_SItemList(nPos).Name = strName Then
         m_SItemList(nPos).Value = strItem
         Exit For
      End If
   Next nPos
End Sub
' 刪除一個項目
Private Sub DeleteSItem(ByVal strName As String, ByVal strItem As String)
   Dim nPos As Integer
   Dim bFind As Boolean
   For nPos = 0 To m_SItemCount - 1
      If m_SItemList(nPos).Name = strName Then
         m_SItemList(nPos).Name = Empty
         m_SItemList(nPos).Value = Empty
         '911017 nick
         m_SItemList(nPos).strCP09 = Empty
         Exit For
      End If
   Next nPos
End Sub
' 檢查此項目是否存在
Private Function IsExistSItem(ByVal strName As String) As Boolean
   Dim nPos As Integer
   IsExistSItem = False
   For nPos = 0 To m_SItemCount - 1
      If m_SItemList(nPos).Name = strName Then
         IsExistSItem = True
         Exit For
      End If
   Next nPos
End Function

Public Sub SetData(ByVal nType As Integer, ByVal strData As String, Optional ByVal bClear As Boolean = False)
   ' 清除搜尋的Key
   If bClear = True Then
      m_SP01 = Empty
      m_SP02 = Empty
      m_SP03 = Empty
      m_SP04 = Empty
   End If
   
   Select Case nType
      ' 本所案號 欄位1
      Case 0: m_SP01 = strData
      ' 本所案號 欄位2
      Case 1: m_SP02 = strData
      ' 本所案號 欄位3
      Case 2: m_SP03 = strData
      ' 本所案號 欄位4
      Case 3: m_SP04 = strData
   End Select
End Sub

' 初始化 GridList
Private Sub InitialGrdList()
   grdList.Clear
   grdList.Rows = 1
   '911016 nick
   'grdList.Cols = 3
   grdList.Cols = 4
   grdList.ColWidth(0) = 300
   grdList.row = 0
   grdList.col = 1
   grdList.Text = "正片號碼"
   grdList.ColWidth(1) = 2000
   grdList.col = 2
   grdList.Text = "是否合格"
   grdList.ColWidth(2) = 1000
   '911016 nick 新增
   grdList.col = 3
   grdList.Text = "收文號"
   grdList.ColWidth(3) = 1000
   
End Sub

Private Sub RefreshData()
   Dim nPos As Integer
   
   InitialGrdList
   For nPos = 0 To m_SItemCount - 1
      If m_SItemList(nPos).Name <> Empty Then
         grdList.Rows = grdList.Rows + 1
         grdList.row = grdList.Rows - 1
         grdList.col = 1
         grdList.Text = m_SItemList(nPos).Name
         grdList.col = 2
         grdList.Text = m_SItemList(nPos).Value
         '911016 nick 新增
         grdList.col = 3
         grdList.Text = m_SItemList(nPos).strCP09
      End If
   Next nPos
   
   'Added by Lydia 2023/10/18
   If grdList.Rows >= 2 Then
      grdList.FixedRows = 1
   End If
   'end 2023/10/18
   
End Sub

Public Sub QueryData()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim strSP22 As String
   Dim strSP23 As String
   Dim nCount As Integer
   Dim nIndex As Integer
   Dim strTemp As String
   Dim strName As String
   Dim strValue As String
   
   ' 顯示本所案號
   textSPKey = m_SP01 & m_SP02 & m_SP03 & m_SP04
   
   strSP22 = Empty
   strSP23 = Empty
   
   '911017 nick 新增     是否合格 預設值代'Y'
   Dim nick911017rs As New ADODB.Recordset
   Dim nickstrsql As String
   nickstrsql = "select bc02,nvl(decode(length(bc03),0,'Y',bc03),'Y'),bc01 from barcode,caseprogress where cp01='" & m_SP01 & "' " & _
            " and cp02='" & m_SP02 & "' and cp03='" & m_SP03 & "' " & _
            " and cp04='" & m_SP04 & "' and cp09=bc01 order by bc01,bc02,bc03 "
   Set nick911017rs = New ADODB.Recordset
   nick911017rs.CursorLocation = adUseClient
   nick911017rs.Open nickstrsql, cnnConnection, adOpenStatic, adLockReadOnly
   If nick911017rs.RecordCount <> 0 Then
       nick911017rs.MoveFirst
       Do While nick911017rs.EOF = False
            AddSItem CheckStr(nick911017rs.Fields(0).Value), CheckStr(nick911017rs.Fields(1).Value), CheckStr(nick911017rs.Fields(2).Value)
            nick911017rs.MoveNext
       Loop
   End If
   
   RefreshData
   
   Set rsTmp = Nothing
End Sub
   
Private Sub OnSaveData()
   Dim nIndex As Integer
   Dim nCount As Integer
End Sub

Private Sub cmdCancel_Click()
   frm02010409_5.Show
   Unload Me
End Sub

Private Sub cmdMod_Click()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   If CheckDataValid = False Then
      GoTo EXITSUB
   End If
   
   If IsEmptyText(textSP22) = False Then
      If IsExistSItem(Trim(textSP22)) = False Then
         strTit = "資料檢核"
         strMsg = "該正片號碼不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      Else
         SetSItem Trim(textSP22), Trim(textSP23)
         RefreshData
      End If
   Else
      strTit = "資料檢核"
      strMsg = "請輸入正片號碼"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   End If
EXITSUB:
End Sub

Private Sub cmdOK_Click()
   Dim strSP22 As String
   Dim strSP23 As String
   Dim bFirst As Boolean
   Dim nPos As Integer
   Dim strSql As String
   
   strSP22 = Empty
   strSP23 = Empty
   bFirst = True

'911017 nick 先刪除
Dim nickstrsql As String
Dim nickIndex As Integer
Dim strBC01 As String
Dim strBC02 As String
Dim strBC03 As String

'Add By Cheng 2002/11/07
On Error GoTo ErrorHandler

cnnConnection.BeginTrans

nickstrsql = "delete barcode where bc01 in (select cp09 from caseprogress where cp01='" & m_SP01 & "' and cp02='" & m_SP02 & "' and cp03='" & IIf(Len(m_SP03) = 0, "0", m_SP03) & "' and cp04='" & IIf(Len(m_SP04) = 0, "00", m_SP04) & "') "
cnnConnection.Execute nickstrsql
'911017 nick 再新增
With grdList
    For nickIndex = 1 To .Rows - 1
        .row = nickIndex
        .col = 1
        strBC02 = .Text
        .col = 2
        strBC03 = .Text
        .col = 3
        strBC01 = .Text
        nickstrsql = "insert into barcode (bc01,bc02,bc03) values ('" & strBC01 & "','" & strBC02 & "','" & strBC03 & "') "
        cnnConnection.Execute nickstrsql
    Next nickIndex
End With

'Add By Cheng 2002/11/07
cnnConnection.CommitTrans
   
' 回到前一個畫面
frm02010409_5.Show
Unload Me

'Add By Cheng 2002/11/07
Exit Sub
ErrorHandler:
    cnnConnection.RollbackTrans
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/18
   Set frm02010409_9 = Nothing
End Sub

Private Sub grdList_SelChange()
   If grdList.row > 0 Then
      grdList.col = 1
      textSP22 = grdList.Text
      grdList.col = 2
      textSP23 = grdList.Text
   End If
   grdList_ShowSelection
End Sub

' 將GridList所選取的列反白, 並將未選取的列設成一般顏色
Private Sub grdList_ShowSelection()
   Dim nRow As Integer
   Dim nCol As Integer
   Dim nCurrSel As Integer
   nCurrSel = grdList.row
   For nRow = 1 To grdList.Rows - 1
      grdList.row = nRow
      If nRow = nCurrSel Then
         For nCol = 1 To grdList.Cols - 1
            grdList.col = nCol
            grdList.CellBackColor = &H8000000D
            grdList.CellForeColor = &H80000005
         Next nCol
      Else
         grdList.col = 1
         If grdList.CellBackColor <> &H80000005 Then
            For nCol = 1 To grdList.Cols - 1
               grdList.col = nCol
               If grdList.CellBackColor <> &H80000005 Then: grdList.CellBackColor = &H80000005
               If grdList.CellForeColor <> &H80000008 Then: grdList.CellForeColor = &H80000008
            Next nCol
         End If
      End If
   Next nRow
   grdList.row = nCurrSel
   grdList.col = 0
End Sub

Private Sub textSP23_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textSP23_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textSP23) = False Then
      Select Case textSP23
         Case "Y", "N":
         Case Else
            Cancel = True
            strTit = "資料檢核"
            strMsg = "是否合格欄位只可輸入Y或N"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textSP23_GotFocus
      End Select
   End If
End Sub

Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   CheckDataValid = False
   If IsEmptyText(textSP22) = True Then
      strTit = "資料檢核"
      strMsg = "請輸入正片號碼"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   If IsEmptyText(textSP23) = True Then
      strTit = "資料檢核"
      strMsg = "請輸入是否合格"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textSP22_GotFocus()
   InverseTextBox textSP22
End Sub

Private Sub textSP23_GotFocus()
   InverseTextBox textSP23
End Sub

