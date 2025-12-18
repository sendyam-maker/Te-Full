VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm04060201_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "大陸專利公報資料維護"
   ClientHeight    =   5745
   ClientLeft      =   5040
   ClientTop       =   2010
   ClientWidth     =   9345
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   9345
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   660
      Left            =   168
      TabIndex        =   10
      Top             =   72
      Width           =   3732
      Begin VB.ComboBox cmbPrinter 
         Height          =   276
         Left            =   765
         Style           =   2  '單純下拉式
         TabIndex        =   11
         Top             =   240
         Width           =   2880
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "印表機"
         Height          =   180
         Index           =   3
         Left            =   120
         TabIndex        =   12
         Top             =   264
         Width           =   540
      End
   End
   Begin VB.CommandButton buttonClear 
      Caption         =   "清除查詢結果(&C)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   4248
      TabIndex        =   7
      Top             =   816
      Width           =   1500
   End
   Begin VB.TextBox textQuery 
      Height          =   264
      Left            =   1380
      MaxLength       =   14
      TabIndex        =   0
      Top             =   840
      Width           =   1452
   End
   Begin VB.CommandButton buttonExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8412
      TabIndex        =   6
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton buttonDel 
      Caption         =   "刪除(&D)"
      Height          =   400
      Left            =   6756
      TabIndex        =   4
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton buttonMod 
      Caption         =   "修改(&M)"
      Height          =   400
      Left            =   5928
      TabIndex        =   3
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton buttonAdd 
      Caption         =   "新增(&A)"
      Default         =   -1  'True
      Height          =   400
      Left            =   5100
      TabIndex        =   2
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton buttonSearch 
      Caption         =   "多筆查詢(&S)"
      Height          =   350
      Left            =   3024
      TabIndex        =   1
      Top             =   816
      Width           =   1200
   End
   Begin VB.CommandButton buttonQuery 
      Caption         =   "查詢(&F)"
      Height          =   400
      Left            =   7584
      TabIndex        =   5
      Top             =   60
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
      Height          =   4356
      Left            =   168
      TabIndex        =   9
      Top             =   1248
      Width           =   8964
      _ExtentX        =   15822
      _ExtentY        =   7673
      _Version        =   393216
      Cols            =   12
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
      _Band(0).Cols   =   12
      _Band(0).GridLinesBand=   2
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label1 
      Caption         =   "申請案號 :"
      Height          =   252
      Left            =   180
      TabIndex        =   8
      Top             =   840
      Width           =   1092
   End
End
Attribute VB_Name = "frm04060201_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/24 改成Form2.0 (grdList)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/12/28 申請案號欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit
' 變數宣告區
Dim m_CurrSel As Integer

'910709 Sieg 413
Dim SeekPrint As Integer, SeekPrintL As Integer
Dim Prn As Printer
Dim m_CPB01List() As String
Dim m_CPB01ListCount As Integer
'Add by Morgan 2004/3/26
'判斷是否為PCT
Dim m_PA46 As String
Private Sub buttonClear_Click()
   InitialGrdList
End Sub

Private Sub Form_Load()
'910709 Sieg 413
Dim i As Integer, j As Integer
   
    MoveFormToCenter Me
    InitGrid 8, grdList
    InitialGrdList
    
    strExc(0) = Printer.DeviceName
    SeekPrintL = Printer.Orientation
    j = 0
    For i = 0 To Printers.Count - 1
        Set Printer = Printers(i)
        If Printer.DeviceName = strExc(0) Then
            SeekPrint = i
        Else
            cmbPrinter.AddItem Printer.DeviceName, j
            j = j + 1
        End If
        'Add By Cheng 2002/11/06
        If Printer.DeviceName = strExc(0) Then
            SeekPrint = i
        End If
    Next i
    'Add By Cheng 2003/02/14
    Set Printer = Printers(SeekPrint)
    If cmbPrinter.ListCount > 0 Then cmbPrinter.ListIndex = 0
    ClearCPB01List
End Sub

Private Sub ExecuteQuery()
 Dim i As Integer
   ' 查詢
   Screen.MousePointer = vbHourglass
   strExc(0) = "SELECT '',C1.CPB01,C1.CPB02," & SQLDate("C1.CPB03") & "," & _
      "C1.CPB04||'-'||C1.CPB05,FNM02," & ChgPatent("", 1) & "," & _
      "C1.CPB08 AS CPB08 FROM CPBULLETIN C1,CPBulletin C2,CAgent,PATENT WHERE " & _
      "C2.CPB01 = '" & textQuery & "' AND C1.CPB03=C2.CPB03 AND " & _
      "C1.CPB06=FNM01(+) AND C1.CPB01=PA11(+) AND '020'=PA09(+) ORDER BY CPB02 DESC"
   intI = 0
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   Screen.MousePointer = vbDefault
   
   ' 檢查是否有資料傳回來
   Set grdList.Recordset = RsTemp
   
   InitialGrdList
   
   If intI = 1 Then
      For i = 0 To grdList.Rows - 1
         If InStr(grdList.TextMatrix(i, 1), textQuery) > 0 Then
            grdList.TopRow = i
            grdList.row = i
            m_CurrSel = 0
            grdList_ShowSelection
            Exit For
         End If
      Next
   End If
End Sub
' 檢查此筆資料是否存在
Public Function IsDataExist(ByVal strKey As String) As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   IsDataExist = False
   strSql = "SELECT * FROM CPBulletin WHERE CPB01 = '" & strKey & "'"
   ' 查詢
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenDynamic
   
   ' 檢查是否有資料傳回來
   If rsTmp.RecordCount <= 0 Then
      IsDataExist = False
   Else
      IsDataExist = True
   End If
   
   rsTmp.Close
   Set rsTmp = Nothing
End Function
' 按下多筆查詢按紐
Private Sub buttonQuery_Click()
   Dim strMsg As String
   Dim strTit As String
   Dim nResponse
   If IsDataExist(textQuery) = True Then
      frm04060201_2.SetMode (2)
      frm04060201_2.SetData (textQuery)
      frm04060201_2.Show
      frm04060201_2.UpdateData
      Me.Hide
   Else
      strTit = "查詢資料"
      strMsg = "資料庫無此筆資料"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   End If
End Sub
' 按下新增按紐
Private Sub buttonAdd_Click()
   Dim strMsg As String
   Dim strTit As String
   Dim nResponse
   'Add by Morgan 2004/3/26
   m_PA46 = ""
   
   If IsDataExist(textQuery) = True Then
      strMsg = "申請案號已存在, 請輸入其它的申請案號"
      strTit = "新增資料"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
    'Add By Cheng 2003/03/11
    '判斷是否為本所案號
    If IsOurCase(Me.textQuery.Text) = False Then
        strMsg = "非本所案件不可新增"
        strTit = "非本所案件"
        nResponse = MsgBox(strMsg, vbOKOnly, strTit)
        'Add By Cheng 2003/03/14
        Me.textQuery.Text = ""
        Me.textQuery.SetFocus
        GoTo EXITSUB
    End If
    '93.1.14 modify by sonia
    'Modify by Morgan 2004/3/26
    '大陸PCT案不檢查申請案號
   'If ChkAppNo(textQuery.Text, Val(Mid(textQuery.Text, 3, 1)), 1) = False Then
   '大陸2003以後申請案號為14碼,故專利種類需抓第五碼
   If m_PA46 <> "Y" And ChkAppNo(textQuery.Text, Val(Mid(textQuery.Text, IIf(Len(textQuery.Text) = 14, 5, 3), 1)), 1) = False Then
   
   'If IsValidData(textQuery.Text) = False Then
      strMsg = "請輸入正確的申請案號"
      strTit = "申請案號"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
    
   frm04060201_2.SetMode (0)
   frm04060201_2.SetData (textQuery)
   frm04060201_2.Show
   frm04060201_2.UpdateData
   'Add By Cheng 2003/02/14
   frm04060201_2.text02.Text = "CN"
   Me.Hide
EXITSUB:
End Sub
' 按下變更按紐
Private Sub buttonMod_Click()
   Dim strMsg As String
   Dim strTit As String
   Dim nResponse
   Dim strKey As String
   If IsDataExist(textQuery) = True Then
      frm04060201_2.SetMode (1)
      frm04060201_2.SetData (textQuery)
      frm04060201_2.Show
      frm04060201_2.UpdateData
      Me.Hide
   Else
      strTit = "修改資料"
      strMsg = "資料庫無此筆資料"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   End If
End Sub
' 搜尋資料
Private Sub buttonSearch_Click()
   ExecuteQuery
End Sub
' 按下離開按紐
Private Sub buttonExit_Click()
   Unload Me
End Sub
' 按下刪除按紐
Private Sub buttonDel_Click()
   Dim strSql As String
   Dim strMsg As String
   Dim strTit As String
   
   If IsDataExist(textQuery) = True Then
      frm04060201_2.SetMode (3)
      frm04060201_2.SetData (textQuery)
      frm04060201_2.Show
      frm04060201_2.UpdateData
      Me.Hide
   Else
      strTit = "刪除資料"
      strMsg = "資料庫中無此筆資料"
      MsgBox strMsg, vbOKOnly, strTit
   End If
End Sub

Private Sub grdList_SelChange()
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

Public Sub AddCPB01(strCPB01 As String)
   If IsEmptyText(strCPB01) = False Then
      ReDim Preserve m_CPB01List(m_CPB01ListCount + 1)
      m_CPB01List(m_CPB01ListCount) = strCPB01
      m_CPB01ListCount = m_CPB01ListCount + 1
   End If
End Sub

' 清除申請人代碼暫存區
Private Sub ClearCPB01List()
   If m_CPB01ListCount > 0 Then
      Erase m_CPB01List
   End If
   m_CPB01ListCount = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim nPageNo As Integer
Dim strCust As String
Dim nPos As Integer

    Set Printer = Printers(SeekPrint)
    Printer.Orientation = SeekPrintL
    Unload frm04060201_2
    Set frm04060201_1 = Nothing
End Sub

Public Function IsValidData(ByVal strData As String)
   Dim nLength As Integer
   Dim nCount As Integer
   Dim nAmount As Integer
   Dim cCheck As String
   Dim nRest As Integer
   
   nLength = Len(strData)
   IsValidData = True
   
   If nLength <> 10 Then
      IsValidData = False
      GoTo EXITSUB
   End If
   
   If IsNumeric(Mid(strData, 1, 8)) = False Then
      IsValidData = False
      GoTo EXITSUB
   End If
   
   ' 90.07.06 modify by louis (不檢查需小於85)
   'If Val(Mid(strData, 1, 2)) < 85 Then
   '   IsValidData = False
   '   GoTo ExitSub
   'End If
   
   If Mid(strData, 9, 1) <> "." Then
      IsValidData = False
      GoTo EXITSUB
   End If
   
   If Val(Mid(strData, 3, 1)) > 3 Then
      IsValidData = False
      GoTo EXITSUB
   End If

   nAmount = 0
   For nCount = 1 To 8
      nAmount = nAmount + Val(Mid(strData, nCount, 1)) * (nCount + 1)
   Next nCount
   nRest = nAmount Mod 11
   If nRest = 10 Then
      cCheck = "X"
   Else
      cCheck = CStr(nRest)
   End If
   If cCheck <> Mid(strData, 10, 1) Then
      IsValidData = False
      GoTo EXITSUB
   End If
EXITSUB:
End Function

Private Sub InitialGrdList()
   FixGrid grdList
   grdList.row = 0
   grdList.col = 0
   grdList.Text = ""
   grdList.ColWidth(0) = 300
   grdList.col = 1
   grdList.Text = "申請案號"
   grdList.ColWidth(1) = 1000
   grdList.ColAlignment(1) = flexAlignCenterCenter
   grdList.col = 2
   grdList.Text = "公告號"
   grdList.ColWidth(2) = 1200
   grdList.ColAlignment(2) = flexAlignCenterCenter
   grdList.col = 3
   grdList.Text = "公告日"
   grdList.ColWidth(3) = 800
   grdList.ColAlignment(3) = flexAlignCenterCenter
   grdList.col = 4
   grdList.Text = "卷號"
   grdList.ColWidth(4) = 700
   grdList.ColAlignment(4) = flexAlignCenterCenter
   grdList.col = 5
   grdList.Text = "代理事務所"
   grdList.ColWidth(5) = 1000
   grdList.ColAlignment(5) = flexAlignLeftCenter
   grdList.col = 6
   grdList.Text = "本所案號"
   grdList.ColWidth(6) = 1200
   grdList.ColAlignment(6) = flexAlignLeftCenter
   grdList.col = 7
   grdList.Text = "申請人"
   grdList.ColWidth(7) = 2500
   grdList.ColAlignment(7) = flexAlignLeftCenter
End Sub

Private Function GetCPBNumber(ByVal strKey As String) As String
   Dim rsTmp As ADODB.Recordset
   Dim strSql As String
   
   GetCPBNumber = Empty
   Set rsTmp = New ADODB.Recordset
   strSql = "SELECT * FROM Patent " & _
            "WHERE PA11 = '" & strKey & "' AND " & _
            "PA09 = '020'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenDynamic
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      GetCPBNumber = rsTmp.Fields("PA01") & "-" & rsTmp.Fields("PA02") & "-" & rsTmp.Fields("PA03") & "-" & rsTmp.Fields("PA04")
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

Private Function GetCPBAgentCompany(ByVal strKey As String) As String
   Dim rsTmp As ADODB.Recordset
   Dim strSql As String
   
   GetCPBAgentCompany = Empty
   Set rsTmp = New ADODB.Recordset
   strSql = "SELECT * FROM CAgent WHERE FNM01 = '" & strKey & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenDynamic
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      GetCPBAgentCompany = rsTmp.Fields("FNM02")
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

Private Sub textQuery_GotFocus()
   InverseAll textQuery
End Sub

' 轉換成大寫
Private Sub textQuery_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 將所有的文字反白
Private Sub InverseAll(ByRef tb As TextBox)
   tb.SelStart = 0
   tb.SelLength = Len(tb.Text)
End Sub

Public Sub SetInputCPB01()
   textQuery = Empty
   textQuery.SetFocus
End Sub

' 更新列表中的資料
Public Sub UpdateRecord(ByVal strKey As String)
   Dim nIndex As Integer
   Dim Str1 As String
   Dim Str2 As String
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   For nIndex = 0 To grdList.Rows - 1
      If grdList.TextMatrix(nIndex, 1) = strKey Then
         ' 組成SQL語法
         strSql = "SELECT * FROM CPBulletin WHERE CPB01 = '" & strKey & "'"
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            ' 若在資料庫中找到該筆資料則更新此筆資料的內容
            rsTmp.MoveFirst
            If IsNull(rsTmp.Fields("CPB01")) = False Then
               grdList.TextMatrix(nIndex, 1) = rsTmp.Fields("CPB01")
            End If
            If IsNull(rsTmp.Fields("CPB02")) = False Then
               grdList.TextMatrix(nIndex, 2) = rsTmp.Fields("CPB02")
            End If
            If IsNull(rsTmp.Fields("CPB03")) = False Then
               grdList.TextMatrix(nIndex, 3) = ChangeWStringToTString(rsTmp.Fields("CPB03"))
            End If
            Str1 = Empty
            Str2 = Empty
            If IsNull(rsTmp.Fields("CPB04")) = False Then
               Str1 = rsTmp.Fields("CPB04")
            End If
            If IsNull(rsTmp.Fields("CPB05")) = False Then
               Str2 = rsTmp.Fields("CPB05")
            End If
            If Str1 = Empty Then: Str1 = "  "
            If Str2 = Empty Then: Str2 = "  "
            grdList.TextMatrix(nIndex, 4) = Str1 & " - " & Str2
            If IsNull(rsTmp.Fields("CPB06")) = False Then
               grdList.TextMatrix(nIndex, 5) = GetCPBAgentCompany(rsTmp.Fields("CPB06"))
            End If
            ' 本所案號
            If IsNull(rsTmp.Fields("CPB01")) = False Then
               grdList.TextMatrix(nIndex, 5) = GetCPBNumber(rsTmp.Fields("CPB01"))
            End If
            If IsNull(rsTmp.Fields("CPB08")) = False Then
               grdList.TextMatrix(nIndex, 5) = rsTmp.Fields("CPB08")
            End If
         Else
            ' 資料庫中無該筆資料表示此筆已被刪除
            grdList.RemoveItem (nIndex)
         End If
         Exit For
      End If
   Next nIndex
End Sub

'Add By Cheng 2003/03/11
'判斷是否為本所案件
Private Function IsOurCase(StrPA11 As String) As Boolean
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

IsOurCase = False
StrSQLa = "Select * From Patent Where PA09='020' And PA11='" & StrPA11 & "' "
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
   'Add by Morgan 2004/3/26
   m_PA46 = "" & rsA("PA46")
    
    IsOurCase = True
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
End Function
