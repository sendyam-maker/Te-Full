VERSION 5.00
Begin VB.Form frm090131 
   BorderStyle     =   1  '單線固定
   Caption         =   "圖形路徑輸入"
   ClientHeight    =   4584
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8328
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4984.384
   ScaleMode       =   0  '使用者自訂
   ScaleWidth      =   8328
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&E)"
      Default         =   -1  'True
      Height          =   303
      Left            =   7300
      TabIndex        =   15
      Top             =   96
      Width           =   900
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame2"
      Height          =   492
      Left            =   7608
      TabIndex        =   11
      Top             =   816
      Visible         =   0   'False
      Width           =   4188
      Begin VB.CommandButton cmdAdd 
         Caption         =   "<- 新增"
         Height          =   285
         Index           =   0
         Left            =   0
         TabIndex        =   13
         Top             =   24
         Width           =   735
      End
      Begin VB.TextBox txtData 
         Height          =   324
         Index           =   0
         Left            =   840
         MaxLength       =   7
         TabIndex        =   12
         Text            =   "01-B-00"
         Top             =   0
         Width           =   780
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "女人"
         Height          =   228
         Left            =   1680
         TabIndex        =   14
         Top             =   48
         Width           =   2436
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame1"
      Height          =   2851
      Left            =   24
      TabIndex        =   2
      Top             =   1704
      Width           =   8280
      Begin VB.CommandButton cmdRemove 
         Caption         =   "移除"
         Height          =   330
         Left            =   7368
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   0
         Width           =   780
      End
      Begin VB.ListBox List1 
         Height          =   2388
         Index           =   1
         ItemData        =   "frm090131.frx":0000
         Left            =   48
         List            =   "frm090131.frx":0037
         TabIndex        =   6
         Top             =   420
         Width           =   1932
      End
      Begin VB.ListBox List1 
         Height          =   2388
         Index           =   2
         ItemData        =   "frm090131.frx":0132
         Left            =   2412
         List            =   "frm090131.frx":0134
         TabIndex        =   5
         Top             =   420
         Width           =   2268
      End
      Begin VB.ListBox List1 
         Height          =   2388
         Index           =   3
         ItemData        =   "frm090131.frx":0136
         Left            =   5112
         List            =   "frm090131.frx":0138
         TabIndex        =   4
         Top             =   420
         Width           =   3036
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "新增"
         Height          =   330
         Index           =   1
         Left            =   6528
         TabIndex        =   3
         Top             =   0
         Width           =   780
      End
      Begin VB.Label Label3 
         Caption         =   "大分類："
         Height          =   228
         Index           =   0
         Left            =   72
         TabIndex        =   10
         Top             =   180
         Width           =   756
      End
      Begin VB.Label Label3 
         Caption         =   "中分類："
         Height          =   228
         Index           =   1
         Left            =   2472
         TabIndex        =   9
         Top             =   180
         Width           =   756
      End
      Begin VB.Label Label3 
         Caption         =   "小分類："
         Height          =   228
         Index           =   2
         Left            =   5136
         TabIndex        =   8
         Top             =   180
         Width           =   756
      End
   End
   Begin VB.ListBox List1 
      Height          =   1488
      Index           =   0
      ItemData        =   "frm090131.frx":013A
      Left            =   1200
      List            =   "frm090131.frx":0141
      TabIndex        =   1
      Top             =   96
      Width           =   4500
   End
   Begin VB.Label lblCnt 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   228
      Left            =   672
      TabIndex        =   17
      Top             =   456
      Width           =   324
   End
   Begin VB.Label Label4 
      Caption         =   "數量："
      Height          =   228
      Left            =   48
      TabIndex        =   16
      Top             =   456
      Width           =   600
   End
   Begin VB.Label Label1 
      Caption         =   "已輸入路徑："
      Height          =   228
      Left            =   48
      TabIndex        =   0
      Top             =   144
      Width           =   1140
   End
End
Attribute VB_Name = "frm090131"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create by Lydia 2024/06/21 無需改成Form2.0 ;
Option Explicit
Private Declare Function SendMessageByNum Lib "user32" _
  Alias "SendMessageA" (ByVal hWnd As Long, ByVal _
  wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Const LB_SETHORIZONTALEXTENT = &H194

Dim m_Role As String
Dim m_InputList As String
Dim m_PrevForm As Form
Dim m_MaxLen As Integer '路徑欄位最大數量
Dim intQ As Integer, strQuery As String, rsQuery As New ADODB.Recordset
Dim m_KeyNo(0 To 4) As String  '0=查名單號；1~3 選取的大,中,小分類代號, 4-小分類名稱
Private Const cntLen As Integer = 7 '完整的圖形路徑長度

Public Sub SetParent(ByVal pRole As String, ByVal pForm As Form, ByVal pKeyNo As String, pInputVal As String)
  
   Set m_PrevForm = pForm
   m_Role = pRole 'Q-查詢,M-維護
   m_KeyNo(0) = pKeyNo    '查名單號
   m_InputList = pInputVal
   
   If TypeName(m_PrevForm) = "frm090128" Then
      m_MaxLen = 5
   Else
      m_MaxLen = 6
   End If
      
End Sub

Private Sub cmdAdd_Click(Index As Integer)
Dim tmpBol As Boolean, tmpKey As String, tmplist As String
Dim intN As Integer

   If Index = 0 Then
      '(保留)：只提供點選清單輸入
      Call Txtdata_Validate(0, tmpBol)
      If Trim(Label2.Caption) = "" Then
         MsgBox "請先輸入正確的圖形路徑！", vbCritical, "圖形路徑輸入檢查"
         txtData(0).SetFocus
         Txtdata_GotFocus 0
         Exit Sub
      Else
         tmpKey = txtData(0) & " " & Trim(Label2.Caption)
      End If
   Else
      If Trim(m_KeyNo(1)) = "" Or Trim(m_KeyNo(2)) = "" Or Trim(m_KeyNo(3)) = "" Then
         MsgBox "請先選取正確的圖形路徑！", vbCritical, "圖形路徑輸入檢查"
         Exit Sub
      Else
         tmpKey = m_KeyNo(1) & "-" & m_KeyNo(2) & "-" & m_KeyNo(3) & " " & Trim(m_KeyNo(4))
      End If
   End If
   If tmpKey = "" Then Exit Sub
   
   tmplist = ""
   intN = 0
   If List1(0).ListCount > 0 Then
      For intI = 0 To List1(0).ListCount - 1
         If Trim(Left(List1(0).List(intI), cntLen)) = Trim(Left(tmpKey, cntLen)) Then
            MsgBox "已存在相同的圖形路徑：" & List1(0).List(intI), vbCritical, "圖形路徑輸入檢查"
            Exit Sub
         End If
         tmplist = tmplist & "," & Trim(Left(List1(0).List(intI), cntLen))
         intN = intN + 1
      Next intI
   End If
   tmplist = tmplist & "," & Trim(Left(tmpKey, cntLen))
   tmplist = Mid(tmplist, 2)
   intN = intN + 1
   If intN > m_MaxLen Then
      MsgBox "圖形路徑不可超過" & m_MaxLen & "組！", vbCritical, "圖形路徑輸入檢查"
      Exit Sub
   End If
   
   List1(0).AddItem tmpKey
   If Index = 0 Then
      txtData(0) = ""
      Label2.Caption = ""
   End If
   SetListScroll List1(0)
   m_InputList = tmplist
   lblCnt = List1(0).ListCount
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdRemove_Click()
   m_InputList = PUB_RemoveListBox2(List1(0), m_InputList)
   lblCnt = List1(0).ListCount
End Sub

Private Sub Form_Load()
Dim tmplist As Variant
Dim oObj As Object

   Frame1.BackColor = &H8000000F
   Frame2.BackColor = &H8000000F
   
   For Each oObj In List1
      oObj.Clear
   Next
   For Each oObj In txtData
      oObj.Text = ""
   Next
   m_KeyNo(1) = ""
   m_KeyNo(2) = ""
   m_KeyNo(3) = ""
   
   Label2.Caption = ""
   If m_Role = "Q" Then
      Me.Width = 6800
      Me.Height = 2100
      cmdExit.Left = 5760
      Frame1.Visible = False
      'Frame2.Visible = False
   Else
      Me.Width = 8400
      Me.Height = 5000
      cmdExit.Left = 7300
      Frame1.Visible = True
      'Frame2.Visible = True
      Call SetListR1(List1(1))
      Call SetListR2(List1(2))
      Call SetListR3(List1(3))
   End If
   MoveFormToCenter Me
   
   '帶入已輸入的路徑
   List1(0).Clear
   lblCnt = "0"
   If m_InputList <> "" Then
      tmplist = Split(m_InputList, ",")
      For intI = 0 To UBound(tmplist)
         If Trim(tmplist(intI)) <> "" Then
            If Pub_ChkTMR3isExist(Trim(tmplist(intI)), False, strExc(1), strExc(2)) Then
               List1(0).AddItem tmplist(intI) & " " & strExc(2)
            Else
               List1(0).AddItem tmplist(intI) & " (不存在的路徑)"
            End If
         End If
      Next intI
      lblCnt = UBound(tmplist) + 1
   End If
   List1(0).Tag = m_InputList
End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   strQuery = ""
   If List1(0).ListCount > 0 Then
      For intI = 0 To List1(0).ListCount - 1
         strQuery = strQuery & "," & Trim(Left(List1(0).List(intI), cntLen))
      Next intI
      strQuery = Mid(strQuery, 2)
   End If
   
   If TypeName(m_PrevForm) <> "Nothing" Then
      If m_Role = "M" Then
         m_PrevForm.SetData strQuery
      End If
      m_PrevForm.Enabled = True
   End If

   Set rsQuery = Nothing
   Set frm090131 = Nothing
End Sub

'*****圖形路徑：大分類****
Private Sub SetListR1(ByRef pLBox As ListBox)
   
   pLBox.Clear
   strQuery = "select * from TMQAppR1 order by 1"
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, strQuery)
   If intQ = 1 Then
      rsQuery.MoveFirst
      Do While Not rsQuery.EOF
         pLBox.AddItem rsQuery.Fields("TMR101") & " " & rsQuery.Fields("TMR102")
         rsQuery.MoveNext
      Loop
      SetListScroll pLBox
   End If
End Sub

'*****圖形路徑：中分類****
Private Sub SetListR2(ByRef pLBox As ListBox, Optional ByVal pKey01 As String)
   
   pLBox.Clear
   If Trim(pKey01) = "" Then Exit Sub
   
   strQuery = "select * from TMQAppR2 where TMR201='" & pKey01 & "' order by 1,2 "
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, strQuery)
   If intQ = 1 Then
      rsQuery.MoveFirst
      Do While Not rsQuery.EOF
         pLBox.AddItem rsQuery.Fields("TMR202") & " " & rsQuery.Fields("TMR203")
         rsQuery.MoveNext
      Loop
      SetListScroll pLBox
   End If
End Sub

'*****圖形路徑：小分類****
Private Sub SetListR3(ByRef pLBox As ListBox, Optional ByVal pKey01 As String, Optional ByVal pKey02 As String)
   
   pLBox.Clear
   If Trim(pKey01) = "" Or Trim(pKey02) = "" Then Exit Sub
   
   strQuery = "select * from TMQAppR3 where TMR301='" & pKey01 & "' and TMR302='" & pKey02 & "' order by 1,2,3 "
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, strQuery)
   If intQ = 1 Then
      rsQuery.MoveFirst
      Do While Not rsQuery.EOF
         pLBox.AddItem rsQuery.Fields("TMR303") & " " & rsQuery.Fields("TMR304")
         rsQuery.MoveNext
      Loop
      SetListScroll pLBox
   End If
End Sub

Private Sub SetListScroll(oList As ListBox)
   Dim ii As Integer
   Dim lWnow As Long, lWmax As Long
   
   lWmax = 0
   For ii = 0 To oList.ListCount - 1
      lWnow = TextWidth(oList.List(ii) & " ")
      If lWnow > lWmax Then
         lWmax = lWnow
      End If
   Next
  
   If ScaleMode = vbTwips Then lWmax = lWmax / Screen.TwipsPerPixelX  ' if twips change to pixels
   SendMessageByNum oList.hWnd, LB_SETHORIZONTALEXTENT, lWmax, 0
End Sub

Private Sub List1_Click(Index As Integer)
Dim strTmpA As String, strTmpB As String
   
   strTmpA = GetListVal(List1(Index), Index, strTmpB)
   If Index = 0 Then
   Else
      Select Case Index
         Case 1  '大分類->中分類
            If m_KeyNo(Index) <> strTmpA And strTmpA <> "" Then
               Call SetListR2(List1(2), strTmpA)
               m_KeyNo(2) = ""
               List1(3).Clear
               m_KeyNo(3) = ""
            End If
         Case 2  '中分類->小分類
            If m_KeyNo(Index) <> strTmpA And strTmpA <> "" Then
               Call SetListR3(List1(3), m_KeyNo(1), strTmpA)
               m_KeyNo(3) = ""
            End If
         Case 3 '小分類
            m_KeyNo(4) = strTmpB
      End Select
      m_KeyNo(Index) = strTmpA
   End If
   
End Sub

Private Sub Txtdata_GotFocus(Index As Integer)
   TextInverse txtData(Index)
End Sub

Private Sub txtdata_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Txtdata_Validate(Index As Integer, Cancel As Boolean)
Dim strTmp1 As String, strTmp2 As String

   Select Case Index
      Case 0
         strTmp1 = ""
         If Trim(txtData(Index)) <> "" Then
            strTmp2 = ""
            If Pub_ChkTMR3isExist(Trim(txtData(Index)), True, strTmp1, strTmp2) = True Then
               txtData(Index) = strTmp1
            End If
         End If
         Label2.Caption = strTmp2
   End Select
End Sub

Private Function GetListVal(ByRef pLBox As ListBox, ByVal pInx As Integer, Optional ByRef pCName As String) As String
Dim intP As Integer
   
   GetListVal = ""
   pCName = ""
   If pLBox.ListCount > 0 Then
      For intP = 0 To pLBox.ListCount - 1
         If pLBox.Selected(intP) = True Then
            If pInx = 0 Then '完整的圖形路徑
               GetListVal = Trim(Left(pLBox.List(intP), cntLen))
               pCName = Trim(Mid(pLBox.List(intP), cntLen + 1))
            Else
               GetListVal = Trim(Left(pLBox.List(intP), 2))
               pCName = Trim(Mid(pLBox.List(intP), 3))
            End If
            Exit For
         End If
      Next intP
   End If
   
End Function
