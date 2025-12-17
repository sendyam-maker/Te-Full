VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Function SaveTrademarkData(ByRef cp() As String, ByRef tm() As String, ByRef sp() As String) As Boolean
Dim objPublicData As Object, intCaseKind As Integer, bolSuccuss As Boolean

'Add By Cheng 2002/11/14
Dim BolTransOk As Boolean
BolTransOk = True

On Error GoTo Err
Set objPublicData = CreateObject("prjTaieDll.clsPublicData")
cnnConnection.BeginTrans
If objPublicData.GetSystemKind(cp(1), intCaseKind) Then
   If objPublicData.SaveCaseProgressDatabase(cp()) Then
      If intCaseKind = ±M§Q Then
         If objPublicData.SaveTrademarkDatabase(tm()) Then bolSuccuss = True
      Else
         If objPublicData.SaveServicePracticeDatabase(sp()) Then bolSuccuss = True
      End If
   End If
End If
If bolSuccuss Then
    'Modify By Cheng 2002/11/14
    If BolTransOk Then
        cnnConnection.CommitTrans
    End If
Else
    'Modify By Cheng 2002/11/14
    If BolTransOk Then
        cnnConnection.RollbackTrans
    End If
End If
Set objPublicData = Nothing
SaveTrademarkData = True
Exit Function
Err:
    'Add By Cheng 2002/11/14
    If Err.Number = -2147168237 Then
       BolTransOk = False
       Resume Next
    End If

cnnConnection.RollbackTrans
ErrorLog
End Function
Public Property Set Connection(ByVal vNewValue As ADODB.Connection)
Set cnnConnection = vNewValue
End Property
