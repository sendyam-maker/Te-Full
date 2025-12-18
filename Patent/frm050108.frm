VERSION 5.00
Begin VB.Form frm050108 
   BorderStyle     =   1  '單線固定
   Caption         =   "國內外案件刪除作業"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4605
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   4605
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   3645
      TabIndex        =   3
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2820
      TabIndex        =   2
      Top             =   60
      Width           =   800
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   1
      Left            =   3180
      MaxLength       =   7
      TabIndex        =   1
      Top             =   1020
      Width           =   1215
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   0
      Left            =   1620
      MaxLength       =   7
      TabIndex        =   0
      Top             =   1020
      Width           =   1215
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "－"
      Height          =   180
      Index           =   1
      Left            =   2910
      TabIndex        =   5
      Top             =   1080
      Width           =   180
   End
   Begin VB.Label lbl 
      Caption         =   "國外案號發文日："
      Height          =   255
      Index           =   0
      Left            =   90
      TabIndex        =   4
      Top             =   1050
      Width           =   1575
   End
End
Attribute VB_Name = "frm050108"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/3 改成Form2.0 (無)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit

Private Sub cmdOK_Click(Index As Integer)
Dim ii As Integer
Select Case Index
Case 0 '確定
   For ii = 0 To 1
      If CheckKeyIn(ii) = False Then
         Me.txt(ii).SetFocus
         txt_GotFocus ii
         Exit Sub
      End If
   Next ii
   If Len(Me.txt(0).Text) <= 0 Then
      MsgBox "請輸入國外案號發文起日!!!", vbExclamation + vbOKOnly
      Me.txt(0).SetFocus
      txt_GotFocus 0
      Exit Sub
   End If
   If Len(Me.txt(1).Text) <= 0 Then
      MsgBox "請輸入國外案號發文起日!!!", vbExclamation + vbOKOnly
      Me.txt(1).SetFocus
      txt_GotFocus 1
      Exit Sub
   End If
   If Val(Me.txt(0).Text) > Val(Me.txt(1).Text) Then
      MsgBox "國外案號發文日起日不可大於迄日!!!", vbExclamation + vbOKOnly
      Me.txt(0).SetFocus
      txt_GotFocus 0
      Exit Sub
   End If
   If MsgBox("您確定要執行 " & ChangeTStringToTDateString(Me.txt(0).Text) & " 至 " & ChangeTStringToTDateString(Me.txt(1).Text) & Chr(10) & Chr(13) & "國內外案件刪除作業嗎???", vbExclamation + vbOKCancel + vbDefaultButton2) = vbOK Then
      Me.Enabled = False
      '處理刪除作業
      Process
      Me.Enabled = True
   End If
Case 1 '結束
   Unload Me
End Select
End Sub
Private Sub Process()
Dim strSql As String
Dim rs As New ADODB.Recordset

'911105 nick transation
cnnConnection.BeginTrans

On Error GoTo ErrorHandler
Screen.MousePointer = vbHourglass
strSql = " AND CM10='0' "
strSql = strSql & " AND (CP10='101' OR CP10='102' OR CP10='103' OR CP10='104' OR CP10='105' OR CP10='125' ) "
strSql = strSql & " AND (( CP27>=" & Val(Me.txt(0).Text) + 19110000 & " AND CP27<=" & Val(Me.txt(1).Text) + 19110000 & ") OR (CP57>=" & Val(Me.txt(0).Text) + 19110000 & " AND CP57<=" & Val(Me.txt(1).Text + 19110000) & ")) "
strSql = "SELECT CM01,CM02,CM03,CM04,CM05,CM06,CM07,CM08,CM10,CP10,CP27,CP57 FROM CASEMAP, CASEPROGRESS WHERE CM01=CP01(+) AND CM02=CP02(+) AND CM03=CP03(+) AND CM04=CP04(+) " & strSql
If rs.State <> adStateClosed Then rs.Close
Set rs = Nothing
rs.CursorLocation = adUseClient
rs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If rs.RecordCount > 0 Then
   rs.MoveFirst
   While Not rs.EOF
      cnnConnection.Execute "Delete From CaseMap Where CM01='" & rs("CM01").Value & "'" & _
                                                 " AND CM02='" & rs("CM02").Value & "'" & _
                                                 " AND CM03='" & rs("CM03").Value & "'" & _
                                                 " AND CM04='" & rs("CM04").Value & "'" & _
                                                 " AND CM05='" & rs("CM05").Value & "'" & _
                                                 " AND CM06='" & rs("CM06").Value & "'" & _
                                                 " AND CM07='" & rs("CM07").Value & "'" & _
                                                 " AND CM08='" & rs("CM08").Value & "'" & _
                                                 " AND CM10='" & rs("CM10").Value & "'"
      rs.MoveNext
   Wend
   MsgBox "資料刪除完畢!!!", vbInformation + vbOKOnly
Else
   MsgBox "搜尋範圍內無資料!!!", vbExclamation + vbOKOnly
End If
If rs.State <> adStateClosed Then rs.Close
Set rs = Nothing

'911105 nick transation
cnnConnection.CommitTrans

Screen.MousePointer = vbDefault
Exit Sub

ErrorHandler:

'911105 nick transation
MsgBox (Err.Description)
cnnConnection.RollbackTrans

'MsgBox Err.Description
Err.Clear
If rs.State <> adStateClosed Then rs.Close
Set rs = Nothing
Screen.MousePointer = vbDefault

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyEscape
   Unload Me
End Select
End Sub

Private Sub Form_Load()
Me.KeyPreview = True
MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm050108 = Nothing
End Sub

Private Sub txt_GotFocus(Index As Integer)
txt(Index).SelStart = 0
txt(Index).SelLength = Len(txt(Index))
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
If CheckKeyIn(Index) = False Then
   Cancel = True
   txt_GotFocus Index
End If
End Sub

Private Function CheckKeyIn(Index As Integer) As Boolean
CheckKeyIn = True
Select Case Index
Case 0
   If PUB_CheckKeyInDate(Me.txt(Index)) <> 0 Then
      CheckKeyIn = False
   End If
Case 1
   If PUB_CheckKeyInDate(Me.txt(Index)) <> 0 Then
      CheckKeyIn = False
   End If
End Select
End Function
