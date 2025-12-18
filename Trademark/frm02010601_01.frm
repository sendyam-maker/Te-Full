VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm02010601_01 
   BorderStyle     =   1  '單線固定
   Caption         =   "代理人已收達/已提申"
   ClientHeight    =   5760
   ClientLeft      =   72
   ClientTop       =   1056
   ClientWidth     =   9336
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   9336
   Begin VB.CommandButton cmdQuery 
      Caption         =   "尋找(&F)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6180
      TabIndex        =   8
      Top             =   60
      Width           =   972
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   400
      Left            =   7200
      TabIndex        =   5
      Top             =   60
      Width           =   972
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8220
      TabIndex        =   6
      Top             =   60
      Width           =   972
   End
   Begin VB.TextBox textTM01 
      Height          =   264
      Left            =   1320
      MaxLength       =   3
      TabIndex        =   0
      Top             =   540
      Width           =   732
   End
   Begin VB.TextBox textTM03 
      Height          =   264
      Left            =   3120
      MaxLength       =   1
      TabIndex        =   3
      Top             =   540
      Width           =   372
   End
   Begin VB.TextBox textTM04 
      Height          =   264
      Left            =   3480
      MaxLength       =   2
      TabIndex        =   4
      Top             =   540
      Width           =   732
   End
   Begin VB.TextBox textTM02_2 
      Height          =   264
      Left            =   2760
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   540
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.TextBox textTM02 
      Height          =   264
      Left            =   2040
      MaxLength       =   6
      TabIndex        =   1
      Top             =   540
      Width           =   1092
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
      Height          =   4752
      Left            =   96
      TabIndex        =   9
      Top             =   888
      Width           =   9132
      _ExtentX        =   16108
      _ExtentY        =   8382
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
      TabIndex        =   7
      Top             =   540
      Width           =   1092
   End
End
Attribute VB_Name = "frm02010601_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/10/16 Form2.0已修改 ; grdList從MSFlexGrid改為MSHFlexGrid
'Memo By Lydia 2021/08/17 Form2.0已檢查 (無需修改的物件)
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/9 日期欄已修改
Option Explicit

Public m_TM01 As String
Public m_TM02 As String
Public m_TM03 As String
Public m_TM04 As String
Dim m_CurrSel As Integer
' 0: 表內商, 1: 表外商
'Dim m_SysKind As Integer
'Add By Sindy 2019/5/22
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Public m_RDate As String
Dim m_Done As Boolean
Dim m_PrevForm As Form
'2019/5/22 END


'Add By Sindy 2019/5/22
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

'Added by Sindy 2019/5/22
Private Sub Form_Activate()
   If m_strIR01 <> "" And m_Done = False Then
      textTM01.Text = m_TM01
      If m_TM01 = "TF" Then
         textTM02.Text = Left(m_TM02, 5)
         textTM02_2.Text = Mid(m_TM02, 6)
      Else
         textTM02.Text = m_TM02
      End If
      textTM03.Text = m_TM03
      textTM04.Text = m_TM04
      cmdQuery.Value = True
      m_Done = True
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
End Sub
'2019/5/22 END

' 設定其為外商還是內商的系統
' Input : nSys 系統類別
'         0 ==> 內商
'         1 ==> 外商
'Public Sub SetSystem(ByVal nSys As Integer)
'   If nSys = 1 Then
'      m_SysKind = 1
'   Else
'      m_SysKind = 0
'   End If
'End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdOK_Click()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   'Add By Sindy 2019/5/22
   If m_strIR01 <> "" Then
      If m_TM01 & m_TM02 & m_TM03 & m_TM04 <> textTM01 & IIf(textTM01 = "TF", textTM02 & textTM02_2, textTM02) & textTM03 & textTM04 Then
         MsgBox "信件輸入必須與信件本所案號(" & m_TM01 & m_TM02 & m_TM03 & m_TM04 & ")一致！"
         Exit Sub
      End If
   End If
   '2019/5/22 END
   
   If IsEmptyText(m_TM01) = True Or IsEmptyText(m_TM02) = True Or IsEmptyText(m_TM03) = True Or IsEmptyText(m_TM04) = True Then
      strTit = "資料檢核"
      strMsg = "請先選取一筆資料"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   Else
      DisplayNextForm
   End If
End Sub

Private Sub cmdQuery_Click()
   If CheckDataValid() = True Then
      QueryData
   End If
End Sub
   
Public Sub QueryData()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   'Add By Sindy 2019/5/22
   If m_strIR01 <> "" Then
      If m_TM01 & m_TM02 & m_TM03 & m_TM04 <> textTM01 & IIf(textTM01 = "TF", textTM02 & textTM02_2, textTM02) & textTM03 & textTM04 Then
         MsgBox "信件輸入必須與信件本所案號(" & m_TM01 & m_TM02 & m_TM03 & m_TM04 & ")一致！"
         Exit Sub
      End If
   End If
   '2019/5/22 END
   
   InitialGrdList
   
   m_TM01 = Trim(textTM01)
   m_TM02 = Trim(textTM02)
   If m_TM01 = "TF" Then: m_TM02 = m_TM02 & Trim(textTM02_2)
   m_TM03 = Trim(textTM03)
   If IsEmptyText(m_TM03) = True Then: m_TM03 = "0"
   m_TM04 = Trim(textTM04)
   If IsEmptyText(m_TM04) = True Then: m_TM04 = "00"
   
   Select Case m_TM01
      Case "T", "TF", "CFT", "FCT":
         strSql = "SELECT * FROM TradeMark " & _
                  "WHERE TM01 = '" & m_TM01 & "' AND " & _
                        "TM02 = '" & m_TM02 & "' AND " & _
                        "TM03 = '" & m_TM03 & "' AND " & _
                        "TM04 = '" & m_TM04 & "' "
      Case Else:
         strSql = "SELECT * FROM ServicePractice " & _
                  "WHERE SP01 = '" & m_TM01 & "' AND " & _
                        "SP02 = '" & m_TM02 & "' AND " & _
                        "SP03 = '" & m_TM03 & "' AND " & _
                        "SP04 = '" & m_TM04 & "' "
   End Select
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      ' 檢查申請國家必須不為台灣
      Select Case m_TM01
         Case "T", "TF", "CFT", "FCT":
            If IsNull(rsTmp.Fields("TM10")) = False Then
               If rsTmp.Fields("TM10") < "010" Then
                  strTit = "資料檢核"
                  strMsg = "該案件的申請國家為台灣, 不可執行此項功能"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  rsTmp.Close
                  GoTo EXITSUB
               End If
            End If
         Case Else:
            If IsNull(rsTmp.Fields("SP09")) = False Then
               If rsTmp.Fields("SP09") < "010" Then
                  strTit = "資料檢核"
                  strMsg = "該案件的申請國家為台灣, 不可執行此項功能"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  rsTmp.Close
                  GoTo EXITSUB
               End If
            End If
      End Select
      ListData rsTmp
      grdList.FixedRows = 1 'Added by Lydia 2023/10/16
   Else
      strTit = "資料檢核"
      strMsg = "無符合條件的資料"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      rsTmp.Close
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   grdList_SetSelection 1
   
   ' 90.06.19
   ' 若只有一筆時直接進入下一個畫面
   If grdList.Rows = 2 Then
      DisplayNextForm
   End If
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

Private Sub ListData(ByRef rsTmp As ADODB.Recordset)
   Dim nIndex As Integer
   Dim strTemp As String
   Dim strNation As String

   rsTmp.MoveFirst
   Select Case Trim(textTM01)
      Case "T", "TF", "CFT", "FCT":
         Do While rsTmp.EOF = False
            grdList.Rows = grdList.Rows + 1
            nIndex = grdList.Rows - 1
            ' 本所案號
            grdList.TextMatrix(nIndex, 1) = rsTmp.Fields("TM01") & "-" & rsTmp.Fields("TM02") & "-" & rsTmp.Fields("TM03") & "-" & rsTmp.Fields("TM04")
            ' 案件名稱
            strTemp = Empty
            If IsNull(rsTmp.Fields("TM05")) = False Then: strTemp = rsTmp.Fields("TM05")
            If IsEmptyText(strTemp) = True And IsNull(rsTmp.Fields("TM06")) = False Then: strTemp = rsTmp.Fields("TM06")
            If IsEmptyText(strTemp) = True And IsNull(rsTmp.Fields("TM07")) = False Then: strTemp = rsTmp.Fields("TM07")
            grdList.TextMatrix(nIndex, 2) = strTemp
            ' 申請國家
            If IsNull(rsTmp.Fields("TM10")) = False Then
               strNation = rsTmp.Fields("TM10")
               grdList.TextMatrix(nIndex, 6) = GetNationName(rsTmp.Fields("TM10"), 0)
            End If
            ' 商標種類
            If IsNull(rsTmp.Fields("TM08")) = False Then
               If strNation < "010" Then
                  grdList.TextMatrix(nIndex, 3) = GetTradeMarkName(rsTmp.Fields("TM08"), 0)
               Else
                  grdList.TextMatrix(nIndex, 3) = GetTradeMarkName(rsTmp.Fields("TM08"), 1)
               End If
            End If
            ' 商品類別
            If IsNull(rsTmp.Fields("TM09")) = False Then
               grdList.TextMatrix(nIndex, 4) = rsTmp.Fields("TM09")
            End If
            ' 申請人
            If IsNull(rsTmp.Fields("TM23")) = False Then
               grdList.TextMatrix(nIndex, 5) = GetCustomerName(rsTmp.Fields("TM23"), 0)
            End If
            ' 本所案號
            If IsNull(rsTmp.Fields("TM01")) = False Then
               grdList.TextMatrix(nIndex, 7) = rsTmp.Fields("TM01")
            End If
            If IsNull(rsTmp.Fields("TM02")) = False Then
               grdList.TextMatrix(nIndex, 8) = rsTmp.Fields("TM02")
            End If
            If IsNull(rsTmp.Fields("TM03")) = False Then
               grdList.TextMatrix(nIndex, 9) = rsTmp.Fields("TM03")
            End If
            If IsNull(rsTmp.Fields("TM04")) = False Then
               grdList.TextMatrix(nIndex, 10) = rsTmp.Fields("TM04")
            End If
            rsTmp.MoveNext
         Loop
      Case Else:
         Do While rsTmp.EOF = False
            grdList.Rows = grdList.Rows + 1
            nIndex = grdList.Rows - 1
            ' 本所案號
            grdList.TextMatrix(nIndex, 1) = rsTmp.Fields("SP01") & "-" & rsTmp.Fields("SP02") & "-" & rsTmp.Fields("SP03") & "-" & rsTmp.Fields("SP04")
            ' 案件名稱
            strTemp = Empty
            If IsNull(rsTmp.Fields("SP05")) = False Then: strTemp = rsTmp.Fields("SP05")
            If IsEmptyText(strTemp) = True And IsNull(rsTmp.Fields("SP06")) = False Then: strTemp = rsTmp.Fields("SP06")
            If IsEmptyText(strTemp) = True And IsNull(rsTmp.Fields("SP07")) = False Then: strTemp = rsTmp.Fields("SP07")
            grdList.TextMatrix(nIndex, 2) = strTemp
            ' 申請國家
            If IsNull(rsTmp.Fields("SP09")) = False Then
               strNation = rsTmp.Fields("SP09")
               grdList.TextMatrix(nIndex, 6) = GetNationName(rsTmp.Fields("SP09"), 0)
            End If
            ' 申請人
            If IsNull(rsTmp.Fields("SP08")) = False Then
               grdList.TextMatrix(nIndex, 5) = GetCustomerName(rsTmp.Fields("SP08"), 0)
            End If
            ' 本所案號
            If IsNull(rsTmp.Fields("SP01")) = False Then
               grdList.TextMatrix(nIndex, 7) = rsTmp.Fields("SP01")
            End If
            If IsNull(rsTmp.Fields("SP02")) = False Then
               grdList.TextMatrix(nIndex, 8) = rsTmp.Fields("SP02")
            End If
            If IsNull(rsTmp.Fields("SP03")) = False Then
               grdList.TextMatrix(nIndex, 9) = rsTmp.Fields("SP03")
            End If
            If IsNull(rsTmp.Fields("SP04")) = False Then
               grdList.TextMatrix(nIndex, 10) = rsTmp.Fields("SP04")
            End If
            rsTmp.MoveNext
         Loop
   End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Sindy 2019/5/22
   If Not m_PrevForm Is Nothing Then
      Set m_PrevForm = Nothing
   End If
   'Add By Cheng 2002/07/18
   Set frm02010601_01 = Nothing
End Sub

Private Sub textTM01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 本所案號中的系統類別
Private Sub textTM01_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textTM01) = False Then
      If IsUserHasRightOfSystem(strUserNum, textTM01) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "您沒有權限使用該系統類別的案件"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM01_GotFocus
         GoTo EXITSUB
      End If
      
      Select Case textTM01
         Case "TF":
            textTM02_2.Visible = True
            textTM02_2.Locked = False
            textTM02_2.TabStop = True
            textTM02.MaxLength = 5
         Case Else
            textTM02_2.Visible = False
            textTM02_2.Locked = True
            textTM02_2.TabStop = False
            textTM02.MaxLength = 6
      End Select
   Else
      textTM02_2.Visible = False
      textTM02_2.Locked = True
      textTM02_2.TabStop = False
      textTM02.MaxLength = 6
   End If
   
EXITSUB:
End Sub

' 設定Grid List的一列為選取的狀態
Private Sub grdList_SetSelection(ByVal nSel As Integer)
   If grdList.Rows >= 2 Then
      grdList.row = nSel
      grdList_SelChange
      grdList_ShowSelection
   End If
End Sub

' 初始化 GridList
Private Sub InitialGrdList()
   grdList.Clear
   grdList.Rows = 1
   grdList.Cols = 11
   grdList.ColWidth(0) = 300
   grdList.row = 0
   grdList.col = 1
   grdList.Text = "本所案號"
   grdList.ColWidth(1) = 1200
   grdList.col = 2
   grdList.Text = "商標名稱"
   grdList.ColWidth(2) = 1000
   grdList.col = 3
   grdList.Text = "商標種類"
   grdList.ColWidth(3) = 1200
   grdList.col = 4
   grdList.Text = "商品類別"
   grdList.ColWidth(4) = 1000
   grdList.col = 5
   grdList.Text = "申請人"
   grdList.ColWidth(5) = 1200
   grdList.col = 6
   grdList.Text = "申請國家"
   grdList.ColWidth(6) = 1200
   grdList.col = 7
   grdList.Text = "本所案號第一欄"
   grdList.ColWidth(7) = 0
   grdList.col = 8
   grdList.Text = "本所案號第二欄"
   grdList.ColWidth(8) = 0
   grdList.col = 9
   grdList.Text = "本所案號第三欄"
   grdList.ColWidth(9) = 0
   grdList.col = 10
   grdList.Text = "本所案號第四欄"
   grdList.ColWidth(10) = 0
End Sub

Private Sub grdList_SelChange()
   If grdList.row > 0 Then
      m_TM01 = grdList.TextMatrix(grdList.row, 7)
      m_TM02 = grdList.TextMatrix(grdList.row, 8)
      m_TM03 = grdList.TextMatrix(grdList.row, 9)
      m_TM04 = grdList.TextMatrix(grdList.row, 10)
   End If
   grdList_ShowSelection
   cmdOK.SetFocus
End Sub

' 將GridList所選取的列反白, 並將未選取的列設成一般顏色
Private Sub grdList_ShowSelection()
   Dim nCurrSel As Integer
   Dim nCol As Integer
   
   nCurrSel = grdList.row
   
   ' 與前一選擇的列位置相同則不處理
   If m_CurrSel = grdList.row Then
      Dim nOldCol As Integer
      nOldCol = grdList.col
      grdList.col = 1
      If grdList.CellBackColor <> &H8000000D Then
         For nCol = 1 To grdList.Cols - 1
            grdList.col = nCol
            If grdList.CellBackColor <> &H8000000D Then grdList.CellBackColor = &H8000000D
            If grdList.CellForeColor <> &H80000005 Then grdList.CellForeColor = &H80000005
         Next nCol
      End If
      grdList.col = nOldCol
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

Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse

   CheckDataValid = False

   If IsEmptyText(textTM01) = True Or IsEmptyText(textTM02) = True Then
      strTit = "資料檢核"
      strMsg = "本所案號輸入不完整"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If

   'Select Case textTM01
   '   Case "TF":
   '      If IsEmptyText(textTM01) = True Or IsEmptyText(textTM02) = True Then
   '         strTit = "資料檢核"
   '         strMsg = "本所案號輸入不完整"
   '         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   '         GoTo ExitSub
   '      End If
   '   Case Else:
   '      If IsEmptyText(textTM01) = True Or IsEmptyText(textTM02) = True Or IsEmptyText(textTM03) = True Or IsEmptyText(textTM04) = True Then
   '         strTit = "資料檢核"
   '         strMsg = "本所案號輸入不完整"
   '         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   '         GoTo ExitSub
   '      End If
   'End Select
   CheckDataValid = True
EXITSUB:
End Function

Private Sub DisplayNextForm()
   frm02010601_02.SetData 0, m_TM01, True
   frm02010601_02.SetData 1, m_TM02, False
   frm02010601_02.SetData 2, m_TM03, False
   frm02010601_02.SetData 3, m_TM04, False
   'Add By Sindy 2019/5/22
   If Not m_PrevForm Is Nothing Then
      Call frm02010601_02.SetParent(m_PrevForm)
   End If
   frm02010601_02.m_strIR01 = m_strIR01
   frm02010601_02.m_strIR02 = m_strIR02
   frm02010601_02.m_strIR03 = m_strIR03
   frm02010601_02.m_strIR04 = m_strIR04
   '2019/5/22 END
   frm02010601_02.Show
   frm02010601_02.QueryData
   cmdQuery.SetFocus
   Me.Hide
End Sub

Private Sub textTM01_GotFocus()
   CloseIme
   InverseTextBox textTM01
End Sub

Private Sub textTM02_GotFocus()
   InverseTextBox textTM02
End Sub

Private Sub textTM02_2_GotFocus()
   InverseTextBox textTM02_2
End Sub

Private Sub textTM03_GotFocus()
   InverseTextBox textTM03
End Sub

Private Sub textTM04_GotFocus()
   InverseTextBox textTM04
End Sub



