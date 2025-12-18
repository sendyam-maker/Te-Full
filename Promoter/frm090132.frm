VERSION 5.00
Begin VB.Form frm090132 
   BorderStyle     =   1  '單線固定
   Caption         =   "3519組群輸入"
   ClientHeight    =   4584
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7332
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4584
   ScaleWidth      =   7332
   Begin VB.ListBox List1 
      Height          =   1848
      Index           =   0
      ItemData        =   "frm090132.frx":0000
      Left            =   1176
      List            =   "frm090132.frx":0002
      TabIndex        =   10
      Top             =   48
      Width           =   4932
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame1"
      Height          =   2600
      Left            =   24
      TabIndex        =   5
      Top             =   1944
      Width           =   7236
      Begin VB.CommandButton cmdAdd 
         Caption         =   "新增"
         Height          =   330
         Index           =   1
         Left            =   2784
         TabIndex        =   8
         Top             =   0
         Width           =   780
      End
      Begin VB.ListBox List1 
         Height          =   2208
         Index           =   1
         ItemData        =   "frm090132.frx":0004
         Left            =   1152
         List            =   "frm090132.frx":0006
         TabIndex        =   7
         Top             =   360
         Width           =   6000
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "移除"
         Height          =   330
         Left            =   3648
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   0
         Width           =   780
      End
      Begin VB.Label Label3 
         Caption         =   "3519組群："
         Height          =   228
         Index           =   0
         Left            =   48
         TabIndex        =   9
         Top             =   648
         Width           =   876
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame2"
      Height          =   492
      Left            =   4680
      TabIndex        =   1
      Top             =   1152
      Visible         =   0   'False
      Width           =   4188
      Begin VB.TextBox txtData 
         Height          =   324
         Index           =   0
         Left            =   840
         MaxLength       =   6
         TabIndex        =   3
         Top             =   0
         Width           =   780
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "<- 新增"
         Height          =   285
         Index           =   0
         Left            =   0
         TabIndex        =   2
         Top             =   24
         Width           =   735
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Height          =   228
         Left            =   1680
         TabIndex        =   4
         Top             =   48
         Width           =   2484
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&E)"
      Default         =   -1  'True
      Height          =   303
      Left            =   6300
      TabIndex        =   0
      Top             =   100
      Width           =   900
   End
   Begin VB.Label lblCnt2 
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
      Left            =   936
      TabIndex        =   15
      Top             =   696
      Visible         =   0   'False
      Width           =   228
   End
   Begin VB.Label Label5 
      Caption         =   "3519組群："
      Height          =   228
      Left            =   48
      TabIndex        =   14
      Top             =   696
      Visible         =   0   'False
      Width           =   912
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
      Left            =   840
      TabIndex        =   13
      Top             =   384
      Width           =   300
   End
   Begin VB.Label Label4 
      Caption         =   "總數量："
      Height          =   228
      Left            =   48
      TabIndex        =   12
      Top             =   384
      Width           =   768
   End
   Begin VB.Label Label1 
      Caption         =   "已輸入組群："
      Height          =   228
      Left            =   48
      TabIndex        =   11
      Top             =   96
      Width           =   1140
   End
End
Attribute VB_Name = "frm090132"
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
Dim m_MaxLen As Integer '組群欄位最大長度
Dim intQ As Integer, strQuery As String, rsQuery As New ADODB.Recordset
Dim m_Status As String '檢索方式：W-文字,P-圖形/文字及圖形
Private Const cntLen3519 As Integer = 6

Public Sub SetParent(ByVal pRole As String, ByVal pForm As Form, pInputVal As String, ByVal pStatus As String)
  
   Set m_PrevForm = pForm
   m_Role = pRole 'Q-查詢,M-維護
   m_InputList = pInputVal
   m_Status = pStatus
   
   If TypeName(m_PrevForm) = "frm090126" Then
      m_MaxLen = 80
   Else
      m_MaxLen = 500
   End If
   
End Sub

Private Sub cmdAdd_Click(Index As Integer)
Dim tmpBol As Boolean, tmpKey As String, tmplist As String
Dim intN As Integer, strTmp1 As String

   If Index = 0 Then
      '(保留)：只提供點選清單輸入
      If Pub_ChkTMQCisExist(m_PrevForm.Name, Trim(txtData(0)), "2", m_Status, strTmp1, m_InputList) = True Then
         tmpKey = PUB_StrToStr(txtData(0), cntLen3519, True) & " " & strTmp1
      Else
         Label2.Caption = ""
         txtData(0).SetFocus
         Txtdata_GotFocus 0
         Exit Sub
      End If
   Else
      tmpKey = GetListVal(List1(Index), Index, strTmp1)
      tmpKey = PUB_StrToStr(tmpKey, cntLen3519, True) & " " & strTmp1
   End If
   If tmpKey = "" Then Exit Sub
   
   tmplist = ""
   intN = 0
   If List1(0).ListCount + 1 > m_MaxLen \ 7 Then
      MsgBox "組群數量不可超過" & m_MaxLen \ 7 & "個！", vbCritical, "組群輸入檢查"
      Exit Sub
   End If
   If List1(0).ListCount > 0 Then
      For intI = 0 To List1(0).ListCount - 1
         If Trim(Left(List1(0).List(intI), cntLen3519)) <> Trim(Left(tmpKey, cntLen3519)) Then
            If Trim(Left(List1(0).List(intI), 4)) = "3519" Then
               intN = intN + 1
            End If
         Else
            MsgBox "已存在相同的組群：" & List1(0).List(intI), vbCritical, "組群輸入檢查"
            Exit Sub
         End If
         tmplist = tmplist & "," & Trim(Left(List1(0).List(intI), cntLen3519))
      Next intI
   End If

   If Left(tmpKey, 4) = "3519" Then
      intN = intN + 1
      '(保留)
      'If TypeName(m_PrevForm) = "frm090126" And intN = 6 Then
      '   MsgBox "3519組群不可超過" & intN - 1 & "個！", vbCritical, "組群輸入檢查"
      '   Exit Sub
      'End If
   End If
   
   tmplist = tmplist & "," & Trim(Left(tmpKey, cntLen3519))
   tmplist = Mid(tmplist, 2)

   List1(0).AddItem tmpKey
   If Index = 0 Then
      txtData(0) = ""
      Label2.Caption = ""
   End If
   SetListScroll List1(0)
   m_InputList = tmplist
   lblCnt = List1(0).ListCount
   lblCnt2 = intN
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdRemove_Click()
Dim strTmp1 As String
   
   strTmp1 = GetListVal(List1(0), 0, strExc(1))
   m_InputList = PUB_RemoveListBox2(List1(0), m_InputList)
   lblCnt = List1(0).ListCount
   If Trim(Left(strTmp1, 4)) = "3519" Then
      lblCnt2 = Val(lblCnt2) - 1
   End If
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

   Label2.Caption = ""
   If m_Role = "Q" Then
     ' Me.Width = 7400
      Me.Height = 2400
      'cmdExit.Left = 4600
      Frame1.Visible = False
      'Frame2.Visible = False
   Else
      'Me.Width = 7400
      Me.Height = 5100
      'cmdExit.Left = 6300
      Frame1.Visible = True
      'Frame2.Visible = True
      Call SetList3519(List1(1))
   End If
   MoveFormToCenter Me
   
   '帶入已輸入的路徑
   List1(0).Clear
   lblCnt = "0"
   intQ = 0
   If m_InputList <> "" Then
      tmplist = Split(m_InputList, ",")
      For intI = 0 To UBound(tmplist)
         If Trim(tmplist(intI)) <> "" Then
            If Pub_ChkTMQCisExist(m_PrevForm.Name, Trim(tmplist(intI)), "2", m_Status, strExc(1), , False) = True Then
               List1(0).AddItem PUB_StrToStr(tmplist(intI) & " ", cntLen3519, True) & " " & strExc(1)
            Else
               List1(0).AddItem PUB_StrToStr(tmplist(intI) & " ", cntLen3519, True) & " (不存在的組群)"
            End If
            If Trim(Left(tmplist(intI), 4)) = "3519" Then
               intQ = intQ + 1
            End If
         End If
      Next intI
      lblCnt = UBound(tmplist) + 1
      lblCnt2 = intQ
   End If
   List1(0).Tag = m_InputList
End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   strQuery = ""
   If List1(0).ListCount > 0 Then
      For intI = 0 To List1(0).ListCount - 1
         strQuery = strQuery & "," & Trim(Left(List1(0).List(intI), cntLen3519))
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
   Set frm090132 = Nothing
End Sub

'*****3519組群清單****
Private Sub SetList3519(ByRef pLBox As ListBox)
   
   pLBox.Clear
   strQuery = "select tmqc01,tmqc06 from TMQclass where substr(tmqc01,1,4)='3519' and length(tmqc01)=6 order by 1"
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, strQuery)
   If intQ = 1 Then
      rsQuery.MoveFirst
      Do While Not rsQuery.EOF
         pLBox.AddItem rsQuery.Fields("TMQC01") & " " & rsQuery.Fields("TMQC06")
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
         If Trim(txtData(Index)) <> "" Then
            Label2.Caption = ""
            If Left(txtData(0), 4) <> "3519" And Len(txtData(0)) <> 6 Then
               MsgBox "請輸入3519的6碼組群代號!!", vbCritical, "輸入檢查"
               txtData(0).SetFocus
               Txtdata_GotFocus 0
               Cancel = True
               Exit Sub
            Else
               If Pub_ChkTMQCisExist(m_PrevForm.Name, Trim(txtData(Index)), "2", m_Status, strTmp1) = True Then
                  Label2.Caption = strTmp1
               End If
            End If
         End If
   End Select
End Sub

Private Function GetListVal(ByRef pLBox As ListBox, ByVal pInx As Integer, Optional ByRef pCName As String) As String
Dim intP As Integer
   
   GetListVal = ""
   pCName = ""
   If pLBox.ListCount > 0 Then
      For intP = 0 To pLBox.ListCount - 1
         If pLBox.Selected(intP) = True Then
            If pInx = 0 Then
               GetListVal = Trim(Left(pLBox.List(intP), cntLen3519))
               pCName = Trim(Mid(pLBox.List(intP), cntLen3519 + 1))
            Else
               GetListVal = Trim(Left(pLBox.List(intP), cntLen3519))
               pCName = Trim(Mid(pLBox.List(intP), cntLen3519 + 1))
            End If
            Exit For
         End If
      Next intP
   End If
   
End Function

