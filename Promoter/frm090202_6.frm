VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090202_6 
   BorderStyle     =   1  '單線固定
   Caption         =   "英文核稿人欄修改權限設定"
   ClientHeight    =   3480
   ClientLeft      =   6090
   ClientTop       =   1545
   ClientWidth     =   5610
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   5610
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   1
      Left            =   4590
      TabIndex        =   4
      Top             =   90
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "存檔(&S)"
      Height          =   345
      Index           =   0
      Left            =   3750
      TabIndex        =   0
      Top             =   90
      Width           =   800
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  '沒有框線
      Height          =   2865
      Left            =   1620
      TabIndex        =   6
      Top             =   510
      Width           =   2805
      Begin VB.CommandButton cmdRemove 
         Caption         =   "移除 ->"
         Height          =   285
         Left            =   1815
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "<- 新增"
         Height          =   285
         Left            =   1815
         TabIndex        =   2
         Top             =   420
         Width           =   735
      End
      Begin MSForms.ListBox lstUsers 
         Height          =   2400
         Left            =   30
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   390
         Width           =   1725
         VariousPropertyBits=   746586139
         ScrollBars      =   2
         DisplayStyle    =   2
         Size            =   "3043;4233"
         MatchEntry      =   0
         MultiSelect     =   1
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox cboUser 
         Height          =   300
         Left            =   30
         TabIndex        =   1
         Top             =   90
         Width           =   2535
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "4471;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
   End
   Begin VB.TextBox TextEPA01 
      Height          =   285
      Left            =   360
      TabIndex        =   5
      Text            =   "TextEPA01"
      Top             =   2190
      Visible         =   0   'False
      Width           =   3435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "有權限的人員："
      Height          =   180
      Index           =   3
      Left            =   330
      TabIndex        =   7
      Top             =   600
      Width           =   1410
   End
End
Attribute VB_Name = "frm090202_6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/17 改成Form2.0 (cboUser,lstUsers)
'Create By Sindy 2015/4/10
Option Explicit


Private Sub cmdOK_Click(Index As Integer)
   Select Case Index
      Case 0 '存檔
         Call SaveData
      Case 1 '結束
         Unload Me
   End Select
End Sub

' 存檔
Private Function SaveData() As Boolean
Dim strSql As String
Dim varTemp As Variant
Dim ii As Integer
        
On Error GoTo ErrHand
   
   SaveData = False
   cnnConnection.BeginTrans
   
   '先檢查有沒有要刪除的
   strSql = "select EPA01 from EP14Authority order by EPA01 asc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      RsTemp.MoveFirst
      Do While Not RsTemp.EOF
         If InStr(TextEPA01, Trim(RsTemp.Fields(0))) = 0 Then
            strSql = "delete from EP14Authority where EPA01='" & Trim(RsTemp.Fields(0)) & "'"
            cnnConnection.Execute strSql
         End If
         RsTemp.MoveNext
      Loop
   End If
   '再檢查有沒有要新增的
   varTemp = Split(TextEPA01, ",")
   For ii = 0 To UBound(varTemp)
      strSql = "select EPA01 from EP14Authority where EPA01='" & varTemp(ii) & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 0 Then
         strSql = "insert into EP14Authority(EPA01,EPA02) values('" & Trim(varTemp(ii)) & "','" & strSrvDate(1) & "')"
         cnnConnection.Execute strSql
      End If
   Next ii
   
   cnnConnection.CommitTrans
   SaveData = True
   MsgBox "存檔完成！"
   Exit Function
   
ErrHand:
   cnnConnection.RollbackTrans
   MsgBox " 存檔失敗！" & vbCrLf & Err.Description
End Function

Private Sub Form_Load()
   MoveFormToCenter Me
   
   Call SetComboData '組下拉式選單
   QueryData
End Sub

Private Sub SetComboData()
Dim rs As New ADODB.Recordset
   
   '有承辦人工作進度資料查詢作業權限的人員
   Me.cboUser.Clear
   If rs.State <> adStateClosed Then rs.Close
   rs.CursorLocation = adUseClient
   rs.Open "select st01,st02,st03,st05 from staff_right,staff where sr01=st05(+)" & _
           "and st04='1' and sr02='frm090614' and st03>='P10' and st03<='P19' order by st01 asc", _
            cnnConnection, adOpenStatic, adLockReadOnly
   Me.cboUser.AddItem ""
   While Not rs.EOF
      Me.cboUser.AddItem Trim(rs.Fields("st01").Value) & " " & rs.Fields("st02").Value
      rs.MoveNext
   Wend
   Set rs = Nothing
   cboUser.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm090202_6 = Nothing
End Sub

'查詢資料
Private Function QueryData() As Boolean
Dim rsTmp As New ADODB.Recordset
   
   strSql = "SELECT EPA01||' '||ST02 FROM EP14Authority,staff WHERE EPA01=ST01(+) order by EPA01 desc"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   lstUsers.Clear
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      Do While Not rsTmp.EOF
         lstUsers.AddItem rsTmp.Fields(0), 0
         rsTmp.MoveNext
      Loop
   End If
   rsTmp.Close
   TextEPA01 = ComposeList(lstUsers)
   
EXITSUB:
   Set rsTmp = Nothing
End Function

'新增人員
Private Sub cmdAdd_Click()
   AddLstFrmCbo cboUser, lstUsers
   TextEPA01 = ComposeList(lstUsers)
End Sub

'移除人員
Private Sub cmdRemove_Click()
   RemoveList lstUsers
   TextEPA01 = ComposeList(lstUsers)
End Sub

Private Function ComposeList(oList As Object) As String
Dim varTemp As Variant
   
   strExc(1) = ""
   If oList.ListCount > 0 Then
      varTemp = Split(oList.List(0), " ")
      strExc(1) = varTemp(0)
      For intI = 1 To oList.ListCount - 1
         varTemp = Split(oList.List(intI), " ")
         strExc(1) = strExc(1) & "," & varTemp(0)
      Next
   End If
   ComposeList = strExc(1)
End Function

Private Sub AddLstFrmCbo(oCombo As Object, oList As Object)
   Dim idx As Integer, bFound As Boolean
   
   If oCombo <> "" And oCombo.ListIndex >= 0 Then
      For idx = 0 To oList.ListCount - 1
         If oList.List(idx) = oCombo Then
            MsgBox "資料已存在！"
            oCombo.SetFocus
            bFound = True
            Exit For
         End If
      Next
      If bFound = False Then
         oList.AddItem oCombo, 0
         oCombo = ""
      End If
   End If
End Sub

Private Sub RemoveList(oList As Object)
   Dim idx As Integer, ii As Integer
   
   If oList.ListCount > 0 Then
      ii = 0
      For idx = 0 To oList.ListCount - 1
         If oList.Selected(ii) = True Then
            oList.RemoveItem ii
            ii = ii - 1
         End If
         ii = ii + 1
      Next
   End If
End Sub

Private Sub cbouser_GotFocus()
   If cboUser.Locked = False Then
      CloseIme
      'Modified by Morgan 2022/1/17
      'SendMessage cboUser.hWnd, CB_SHOWDROPDOWN, 1, 0
      cboUser.DropDown
      'end 2022/1/17
   End If
End Sub

Private Sub cboUser_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub
