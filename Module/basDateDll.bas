Attribute VB_Name = "basDateDll"
Option Explicit
'轉換民國日期至西元之格式
Public Function ChangeTStringToWString(ByRef strTString As String) As String
If strTString = "" Then Exit Function
ChangeTStringToWString = Format(Val(strTString) + 19110000)
End Function
'轉換西元日期至民國之格式
Public Function ChangeWStringToTString(ByRef strWString As String) As String
If strWString = "" Then Exit Function
ChangeWStringToTString = Format(Val(strWString) - 19110000)
End Function
Public Function GetTaiwanThisYear() As String
GetTaiwanThisYear = Right(Format(Val(Format(Date, "YYYY")) - 1911), 2)
End Function

