VERSION 5.00
Begin VB.Form frm090110 
   BorderStyle     =   1  '單線固定
   Caption         =   "委查資料刪除作業"
   ClientHeight    =   2460
   ClientLeft      =   1650
   ClientTop       =   1050
   ClientWidth     =   5520
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   5520
   Begin VB.TextBox Txtdata 
      Height          =   300
      Index           =   1
      Left            =   1212
      MaxLength       =   7
      TabIndex        =   0
      Top             =   1128
      Width           =   825
   End
   Begin VB.TextBox Txtdata 
      Height          =   300
      Index           =   2
      Left            =   2376
      MaxLength       =   7
      TabIndex        =   1
      Top             =   1128
      Width           =   825
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   3612
      TabIndex        =   2
      Top             =   120
      Width           =   800
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   4440
      TabIndex        =   3
      Top             =   120
      Width           =   800
   End
   Begin VB.Label Label7 
      Caption         =   "（請輸入民國年）"
      Height          =   336
      Left            =   3384
      TabIndex        =   6
      Top             =   1152
      Width           =   1692
   End
   Begin VB.Label Label2 
      Caption         =   "查覆日期："
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   288
      TabIndex        =   5
      Top             =   1152
      Width           =   912
   End
   Begin VB.Label Label3 
      Caption         =   "－"
      Height          =   288
      Left            =   2124
      TabIndex        =   4
      Top             =   1188
      Width           =   276
   End
End
Attribute VB_Name = "frm090110"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2021/12/21 Form2.0已檢查 (無需修改的物件)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
Option Explicit
Dim s As Integer, i As Integer, j As Integer
Dim strSql As String
Dim rs As New ADODB.Recordset

Private Sub cmdOK_Click()
   If Txtdata(1) = Empty And Txtdata(2) = Empty Then
      s = MsgBox("請輸入查覆日期條件", , "使用者輸入錯誤")
      Txtdata(1).SetFocus
      Exit Sub
   End If
   'Add By Cheng 2002/03/21
   If PUB_CheckKeyInDate(Me.Txtdata(1)) = -1 Then
      Me.Txtdata(1).SetFocus
      Txtdata_GotFocus 1
      Exit Sub
   End If
   If PUB_CheckKeyInDate(Me.Txtdata(2)) = -1 Then
      Me.Txtdata(2).SetFocus
      Txtdata_GotFocus 2
      Exit Sub
   End If
   'Modify by Morgan 2010/8/16 百年蟲
   'If Txtdata(1) > Txtdata(2) Then
   If Val(Txtdata(1)) > Val(Txtdata(2)) Then
      s = MsgBox("查覆日期範圍錯誤", , "使用者輸入錯誤")
      Txtdata(1).SetFocus
      TextInverse Txtdata(1)
      Exit Sub
   End If
   Me.Enabled = False
   Me.Hide
   Process
   Me.Enabled = True
   Me.Show
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Me.Height = 3780
   Me.Width = 5625
   MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/18
   Set frm090110 = Nothing
End Sub

Private Sub Txtdata_GotFocus(Index As Integer)
   TextInverse Txtdata(Index)
End Sub

Private Sub txtdata_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 13 Then
      cmdOK_Click
   Else
      KeyAscii = UpperCase(KeyAscii)
   End If
End Sub

Private Sub Process()
Dim subSQL1 As String, subSQL2 As String, strMsg As String
   Screen.MousePointer = vbHourglass
   If rs.State <> adStateClosed Then
      rs.Close
   End If
   subSQL1 = "": subSQL2 = ""
   If Txtdata(1) <> Empty Then
      subSQL1 = " TMQ11>=" & Val(ChangeTStringToWString(Txtdata(1))) & ""
      subSQL2 = " AND TMQ06>=" & Val(ChangeTStringToWString(Txtdata(1))) & ""
   Else
      subSQL1 = " TMQ11 IS NOT NULL"
   End If
   If Txtdata(2) <> Empty Then
      If Len(subSQL1) <> 0 Then
         subSQL1 = subSQL1 + " AND "
         subSQL2 = subSQL2 + " AND "
      End If
      subSQL1 = subSQL1 + " TMQ11<=" & Val(ChangeTStringToWString(Txtdata(2))) & ""
      subSQL2 = subSQL2 + " TMQ06<=" & Val(ChangeTStringToWString(Txtdata(2))) & ""
   End If
   If Len(subSQL1) <> 0 Then
      subSQL1 = " WHERE (" & subSQL1 & ")"
   End If
   'Modified by Lydia 2016/03/28 扣除電子化查名單
   'strSql = "SELECT * FROM TRADEMARKQUERY" & subSQL1 & " OR (TMQ11 IS NULL " & subSQL2 & ") "
   'Modified by Lydia 2017/09/28 舊查名單TMQ18=>0 (tmq18 is null=>tmq18='0'
   strSql = "SELECT * FROM TRADEMARKQUERY" & subSQL1 & " OR (TMQ11 IS NULL " & subSQL2 & ") and tmq18='0'"
   rs.CursorLocation = adUseClient
   rs.Open strSql, cnnConnection, adOpenDynamic, adLockBatchOptimistic
   If rs.RecordCount <> 0 And rs.RecordCount > 0 Then
      'Modified by Lydia 2016/03/28
      'strMsg = "符合條件的資料共" + CStr(rs.RecordCount) + "筆資料，是否確定刪除資料？"
      strMsg = "符合條件的未電子化委查資料共" + CStr(rs.RecordCount) + "筆資料，是否確定刪除資料？"
      s = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton1, "詢問")
      If s = vbYes Then
         strSql = "DELETE TRADEMARKQUERY" & subSQL1 & " OR (TMQ11 IS NULL " & subSQL2 & ") "
         cnnConnection.Execute strSql
         DataErrorMessage (14), "" '刪除成功
      End If
   Else
      s = MsgBox("資料庫中沒有符合的資料!!", , "沒有資料")
   End If
   If rs.State <> adStateClosed Then
      rs.Close
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Sub Txtdata_Validate(Index As Integer, Cancel As Boolean)
Select Case Index
Case 1, 2 '查覆日期起, 迄
   If PUB_CheckKeyInDate(Me.Txtdata(Index)) = -1 Then
      Cancel = True
      Me.Txtdata(Index).SetFocus
      Txtdata_GotFocus Index
   End If
End Select
End Sub
