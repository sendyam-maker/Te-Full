VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmCCCCode 
   Caption         =   "CCC Code"
   ClientHeight    =   5172
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5172
   ScaleWidth      =   6720
   Begin VB.TextBox textSPKey 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Width           =   2292
   End
   Begin VB.TextBox textSP24 
      Height          =   264
      Left            =   1200
      MaxLength       =   15
      TabIndex        =   0
      Top             =   480
      Width           =   2292
   End
   Begin VB.TextBox textSP25 
      Height          =   264
      Left            =   1200
      MaxLength       =   1
      TabIndex        =   1
      Top             =   840
      Width           =   612
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "新增"
      Height          =   492
      Left            =   120
      TabIndex        =   2
      Top             =   4560
      Width           =   1212
   End
   Begin VB.CommandButton cmdMod 
      Caption         =   "修改"
      Height          =   492
      Left            =   1440
      TabIndex        =   3
      Top             =   4560
      Width           =   1212
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "刪除"
      Height          =   492
      Left            =   2760
      TabIndex        =   4
      Top             =   4560
      Width           =   1212
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   492
      Left            =   4080
      TabIndex        =   5
      Top             =   4560
      Width           =   1212
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Default         =   -1  'True
      Height          =   492
      Left            =   5400
      TabIndex        =   6
      Top             =   4560
      Width           =   1212
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
      Height          =   3252
      Left            =   72
      TabIndex        =   12
      Top             =   1176
      Width           =   6492
      _ExtentX        =   11451
      _ExtentY        =   5736
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
      TabIndex        =   11
      Top             =   120
      Width           =   1092
   End
   Begin VB.Label Label31 
      Caption         =   "CCC Code :"
      Height          =   252
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   972
   End
   Begin VB.Label Label32 
      Caption         =   "是否授權 :"
      Height          =   252
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   1092
   End
   Begin VB.Label Label33 
      Caption         =   "( Y:授權 N:不授權 )"
      Height          =   252
      Left            =   1920
      TabIndex        =   8
      Top             =   840
      Width           =   1572
   End
End
Attribute VB_Name = "frmCCCCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/10/20 Form2.0已修改 ; grdList從MSFlexGrid改為MSHFlexGrid
'Memo by Sonia 2022/2/24 改成Form2.0 不用改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/12 日期欄已修改
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 本程式可以設定CCC Code
' 當要使用本程式時, 請呼叫 SetData 來設定本所案號, CCC Code等參數
' 當要取回資料時, 請呼叫 GetData 來取得使用者按下的是OK還是Cancel
' 再使用GetData來取回 CCC Code
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

' 本所案號
Dim m_SP01 As String
Dim m_SP02 As String
Dim m_SP03 As String
Dim m_SP04 As String
' CCC Code
Dim m_SP24 As String
Dim m_SP25 As String
' 該程式主控或父程式主控 0:表該程式主控 1:表父程式主控
Dim m_Control As Integer
' 使用者選擇的是OK還是Cancel
Dim m_OKCancel As String

' 宣告
Private Type SITEM
   Name As String
   Value As String
End Type
Dim m_SItemList() As SITEM
Dim m_SItemCount As Integer

Private Sub Form_Load()
   MoveFormToCenter Me
   textSPKey.BackColor = &H8000000F
   m_OKCancel = "0"
End Sub

' 清除串列
Private Sub ClearSItem()
   If m_SItemCount > 0 Then
      Erase m_SItemList
   End If
   m_SItemCount = 0
End Sub

' 新增一個項目
Private Sub AddSItem(ByVal strName As String, ByVal strItem As String)
   Dim nPos As Integer
   Dim bFind As Boolean
   bFind = False
   For nPos = 0 To m_SItemCount - 1
      If m_SItemList(nPos).Name = strName Then
         m_SItemList(nPos).Value = strItem
         bFind = True
         Exit For
      End If
   Next nPos
   
   If bFind = False Then
      ReDim Preserve m_SItemList(m_SItemCount + 1)
      m_SItemList(m_SItemCount).Name = strName
      m_SItemList(m_SItemCount).Value = strItem
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

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 取得使用者修改後的資料
' Input : nType = 0 表取得使用者按下的是OK還是Cancel
'                 "1"表OK, "0"表Cancel
'                 1 表取得欄位SP24 (CCC Code)
'                 2 表取得欄位SP25 (是否授權)
Public Function GetData(ByVal nType As Integer) As String
   GetData = Empty
   Select Case nType
      Case 0: GetData = m_OKCancel
      Case 1: GetData = m_SP24
      Case 2: GetData = m_SP25
   End Select
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 設定資料
' Input : nType = 0 表設定的資料是本所案號的第一個欄位
'                 1 表設定的資料是本所案號的第二個欄位
'                 2 表設定的資料是本所案號的第三個欄位
'                 3 表設定的資料是本所案號的第四個欄位
'                 4 表設定的資料是 CCC Code
'                 5 表設定的資料是 是否授權
'         strData = 資料內容
'         bClear = True 表清除所有資料
'                  False 表不清除所有資料
' 說明 : 設定預設值時, 第一次設定的欄位需清除所有資料
'        程式預設由本程式自動讀取服務業務基本檔的CCC Code, 且自動更新
'        但若有設定CCC Code時, 程式會變成不讀基本檔, 由父程式自行控制寫入的動作
Public Sub SetData(ByVal nType As Integer, ByVal strData As String, Optional ByVal bClear As Boolean = False)
   ' 清除搜尋的Key
   If bClear = True Then
      m_SP01 = Empty
      m_SP02 = Empty
      m_SP03 = Empty
      m_SP04 = Empty
      m_SP24 = Empty
      m_SP25 = Empty
      m_Control = 0
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
      ' CCC Code
      Case 4:
         m_SP24 = strData
'         m_Control = 1
      Case 5: m_SP25 = strData
   End Select
End Sub

' 初始化 GridList
Private Sub InitialGrdList()
   grdList.Clear
   grdList.Rows = 1
   grdList.Cols = 3
   grdList.ColWidth(0) = 300
   grdList.row = 0
   grdList.col = 1
   grdList.Text = "CCC Code"
   grdList.ColWidth(1) = 2000
   grdList.col = 2
   grdList.Text = "是否授權"
   grdList.ColWidth(2) = 1000
End Sub

Private Sub RefreshData()
   Dim nPos As Integer
   
   InitialGrdList
   For nPos = 0 To m_SItemCount - 1
      If m_SItemList(nPos).Name <> Empty Then
         grdList.Rows = grdList.Rows + 1
         grdList.row = grdList.Rows - 1
         grdList.TextMatrix(grdList.row, 1) = m_SItemList(nPos).Name
         grdList.TextMatrix(grdList.row, 2) = m_SItemList(nPos).Value
      End If
   Next nPos
   
   'Added by Lydia 2023/10/20
   If grdList.Rows >= 2 Then
      grdList.FixedRows = 1
   End If
   'end 2023/10/20
End Sub

Public Sub QueryData()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim nCount As Integer
   Dim nIndex As Integer
   Dim strTemp As String
   Dim strName As String
   Dim strValue As String
   
   ' 顯示本所案號
   textSPKey = m_SP01 & m_SP02 & m_SP03 & m_SP04
   
   m_SP24 = Empty
   m_SP25 = Empty
   ClearSItem        '2008/11/10 ADD BY SONIA TM-000047改基本檔後未清
   
   If m_Control = 0 Then
      strSql = "SELECT * FROM ServicePractice " & _
            "WHERE SP01 = '" & m_SP01 & "' AND " & _
                  "SP02 = '" & m_SP02 & "' AND " & _
                  "SP03 = '" & m_SP03 & "' AND " & _
                  "SP04 = '" & m_SP04 & "'"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         If IsNull(rsTmp.Fields("SP24")) = False Then
            m_SP24 = rsTmp.Fields("SP24")
         End If
         If IsNull(rsTmp.Fields("SP25")) = False Then
            m_SP25 = rsTmp.Fields("SP25")
         End If
      End If
      rsTmp.Close
   End If
         
   nCount = GetSubStringCount(m_SP24)
   For nIndex = 1 To nCount
      strTemp = GetSubString(m_SP24, nIndex)
      If IsEmptyText(strTemp) = False Then
         strName = strTemp
         ' 91.10.15 modify by louis
         'If Len(m_SP25) >= nIndex Then
         '   strValue = Mid(m_SP25, nIndex, 1)
         'Else
         '   strValue = "N"
         'End If
         strValue = GetSubString(m_SP25, nIndex)
         If IsEmptyText(strValue) Then strValue = "N"
         AddSItem strName, strValue
      End If
   Next nIndex
   
   RefreshData
   
   Set rsTmp = Nothing
End Sub

' 儲存資料 若讀寫資料庫的控制權在本程式則將會儲存到資料庫中, 否則只會更新m_SP24及m_SP25
'edit b nick 2004/11/03
'Private sub OnSaveData()
Public Function OnSaveData() As Boolean
OnSaveData = True
On Error GoTo ErrorHandler

   Dim nIndex As Integer
   Dim nCount As Integer
   Dim strSP24 As String
   Dim strSP25 As String
   strSP24 = Empty
   strSP25 = Empty
   For nIndex = 0 To m_SItemCount - 1
      If IsEmptyText(m_SItemList(nIndex).Name) = False Then
         If IsEmptyText(strSP24) = False Then: strSP24 = strSP24 & ","
         strSP24 = strSP24 & m_SItemList(nIndex).Name
      End If
   Next nIndex
   For nIndex = 0 To m_SItemCount - 1
      If IsEmptyText(m_SItemList(nIndex).Name) = False Then
         ' 91.10.15 modify by louis
         If IsEmptyText(strSP25) = False Then: strSP25 = strSP25 & ","
         strSP25 = strSP25 & m_SItemList(nIndex).Value
      End If
   Next nIndex
   m_SP24 = strSP24
   m_SP25 = strSP25
   
   If m_Control = 0 Then
      strSql = "Update ServicePractice SET SP24 = '" & m_SP24 & "', " & _
                                          "SP25 = '" & m_SP25 & "' " & _
               "WHERE SP01 = '" & m_SP01 & "' AND " & _
                     "SP02 = '" & m_SP02 & "' AND " & _
                     "SP03 = '" & m_SP03 & "' AND " & _
                     "SP04 = '" & m_SP04 & "'"
      cnnConnection.Execute strSql
   End If
Exit Function
ErrorHandler:
    OnSaveData = False
End Function

Private Sub cmdAdd_Click()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   If CheckDataValid = False Then
      GoTo EXITSUB
   End If
   
   If IsEmptyText(textSP24) = False Then
      If IsExistSItem(Trim(textSP24)) = False Then
         AddSItem Trim(textSP24), Trim(textSP25)
         RefreshData
      Else
         strTit = "資料檢核"
         strMsg = "該CCC Code已存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      End If
   Else
      strTit = "資料檢核"
      strMsg = "請輸入CCC Code"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   End If
   
EXITSUB:
End Sub

Private Sub cmdCancel_Click()
   Me.Hide
End Sub

Private Sub cmdDel_Click()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   If CheckDataValid = False Then
      GoTo EXITSUB
   End If
   
   If IsEmptyText(textSP24) = False Then
      If IsExistSItem(Trim(textSP24)) = False Then
         strTit = "資料檢核"
         strMsg = "該CCC Code不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      Else
         DeleteSItem Trim(textSP24), Trim(textSP25)
         RefreshData
      End If
   Else
      strTit = "資料檢核"
      strMsg = "請輸入CCC Code"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   End If
EXITSUB:
End Sub

Private Sub cmdMod_Click()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   If CheckDataValid = False Then
      GoTo EXITSUB
   End If
   
   If IsEmptyText(textSP24) = False Then
      If IsExistSItem(Trim(textSP24)) = False Then
         strTit = "資料檢核"
         strMsg = "該CCC Code不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      Else
         SetSItem Trim(textSP24), Trim(textSP25)
         RefreshData
      End If
   Else
      strTit = "資料檢核"
      strMsg = "請輸入CCC Code"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   End If
EXITSUB:
End Sub

Private Sub cmdOK_Click()
   m_OKCancel = "1"
   'edit by nick 2004/11/3
   'OnSaveData
   If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
   Me.Hide
End Sub

Private Sub grdList_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeySpace Then
      If grdList.row > 0 Then
         If (grdList.TextMatrix(grdList.row, 2)) = "Y" Then
            grdList.TextMatrix(grdList.row, 2) = "N"
         Else
            grdList.TextMatrix(grdList.row, 2) = "Y"
         End If
         SetSItem grdList.TextMatrix(grdList.row, 1), grdList.TextMatrix(grdList.row, 2)
         grdList_SelChange
      End If
   End If
End Sub

Private Sub grdList_SelChange()
   If grdList.row > 0 Then
      textSP24 = grdList.TextMatrix(grdList.row, 1)
      textSP25 = grdList.TextMatrix(grdList.row, 2)
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

Private Sub textSP25_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textSP25_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textSP25) = False Then
      Select Case textSP25
         Case "Y", "N":
         Case Else
            Cancel = True
            strTit = "資料檢核"
            strMsg = "是否授權欄位只可輸入Y或N"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      End Select
   End If
End Sub

Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   CheckDataValid = False
   If IsEmptyText(textSP24) = True Then
      strTit = "資料檢核"
      strMsg = "請輸入CCC Code"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   If IsEmptyText(textSP25) = True Then
      strTit = "資料檢核"
      strMsg = "請輸入是否授權"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   'edit by nickc 2006/04/26
   'If grdList.Rows > 15 Then
   If grdList.Rows > 50 Then
      strTit = "資料檢核"
      'edit by nickc 2006/04/26
      'strMsg = "資料內容不可超過15筆"
      strMsg = "資料內容不可超過50筆"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textSP24_GotFocus()
   InverseTextBox textSP24
End Sub

Private Sub textSP25_GotFocus()
   InverseTextBox textSP25
End Sub


