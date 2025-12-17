Attribute VB_Name = "basPublicClass"
Option Explicit
'Login之Server Name
Global Const ServerName = "M51CON"
'Login之User Name/Password
Global Const UserName = "PGMID"
Global Const Password = "PGMPWD"
Public Declare Function WNetGetUser Lib "mpr.dll" Alias "WNetGetUserA" (ByVal lpName As String, ByVal lpUserName As String, lpnLength As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags_ As Long) As Long
Public Const HWND_TOPMOST = -1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_SHOWWINDOW = &H40
'Connection宣告
Global cnnConnection As ADODB.Connection
'轉換字串以塞入SQL語法
Public Function CNULL(ByRef strNULL As String) As String
If strNULL = "" Then
   CNULL = "NULL"
Else
   CNULL = "'" + strNULL + "'"
End If
End Function
'顯示錯誤訊息視窗
Public Sub ErrorLog()
frm990003.Show vbModal
End Sub
'去除字串空白
Public Function MyTrim(ByRef strTemp As String) As String
Dim i As Long

For i = 1 To Len(strTemp)
       If Asc(Mid(strTemp, i, 1)) = 0 Then
          MyTrim = Mid(strTemp, 1, i - 1)
          Exit For
       End If
Next
End Function
'傳回現在應該是繳第幾年之費用
Public Function GetMoneyYears(ByRef strTemp As String) As Integer
Dim varTemp As Variant

If strTemp = "" Then
   GetMoneyYears = 1
Else
   varTemp = Split(strTemp, ",")
   GetMoneyYears = UBound(varTemp) + 1
End If
End Function
'傳回起算日之日期
Public Function GetStartDate(ByRef strTemp As String, cp() As String, field() As String, Optional bShowMsg As Boolean = True) As String
Select Case strTemp
             Case 收文日
                        GetStartDate = cp(5)
             Case 申請日
                        GetStartDate = field(10)
             Case 公開日
                        GetStartDate = cp(27)
             Case 准駁日
                        GetStartDate = cp(25)
             Case 公告日
                        GetStartDate = field(14)
             Case 發證日
                        GetStartDate = field(21)
End Select
If GetStartDate = "" Then
   If bShowMsg Then
      ShowMsg MsgText(9168)
   End If
End If
End Function
'將From移至畫面之中心
Public Sub MoveFormToCenter(ByRef frmTemp As Form)
Dim intX  As Integer, intY As Integer

intX = (Screen.Width - frmTemp.Width) / 2
intY = (Screen.Height - frmTemp.Height) / 2
frmTemp.Move intX, intY
End Sub
'本所案號後補零
Public Function ChangeCaseNumL(ByRef strTemp As String) As String
   If strTemp <> "" Then ChangeCaseNumL = strTemp + String(11 - Len(strTemp), "0")
End Function
'傳回長的客戶代號
Public Function ChangeCustomerL(ByRef strTemp As String, Optional IsAgent As Boolean) As String
'If IsAgent Then
'   If strTemp <> "" Then ChangeCustomerL = strTemp + String(8 - Len(strTemp), "0")
'Else
   If strTemp <> "" Then ChangeCustomerL = strTemp + String(9 - Len(strTemp), "0")
'End If
End Function
'傳回短的客戶代號
Public Function ChangeCustomerS(ByRef strTemp As String) As String
If strTemp <> "" Then ChangeCustomerS = IIf(Right(strTemp, 3) = "000", Mid(strTemp, 1, 6), IIf(Right(strTemp, 1) = "0", Mid(strTemp, 1, 8), strTemp))
End Function
