Attribute VB_Name = "basUpdSys"
Option Explicit
'變數宣告
Public Pub_str_UpdSysName As String '欲更新的專案名稱
Public Pub_str_SourcePath As String '各更新檔的來源路徑
Public Pub_str_WinSysPath As String '本機的WinSysPath
Public Pub_str_UpdPathFileName As String 'Server上Upd.ini的路徑與檔名
Public pub_str_VerPathFileName As String 'Server上Ver.ini的路徑與檔名

'API宣告用
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const HWND_BOTTOM = 1
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1


Sub main()

If App.PrevInstance Then
   MsgBox "更新程式重覆啟動!!!", vbExclamation
   End
End If

frmUpdSys.Show: DoEvents

End Sub

Public Sub Pub_WriteSysLog(strMsg As String)

On Error Resume Next
   
Open App.Path & "\Log_" & App.EXEName & ".txt" For Append As #10
Print #10, Now & "　" & strMsg
Close #10
If Err.Number <> 0 Then Err.Clear

End Sub

Public Sub Pub_DeleteLogFile()

On Error Resume Next

Kill App.Path & "\Log_" & App.EXEName & ".txt"
If Err.Number <> 0 Then Err.Clear

End Sub

