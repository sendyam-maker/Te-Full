VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class0704"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'檢查是否為專利商標種類代號
Public Function ChkSystemKind(StrSysCode As String) As Boolean
Dim obj01 As Object
Dim Rs As ADODB.Recordset
    Set obj01 = CreateObject("Prjtaiedll.clspublicdata")
    strSQL = "select * from SystemKind where sk02=" + CNULL(StrSysCode)
    Set Rs = obj01.ReadRst(strSQL)
    If Rs.EOF Then
        ChkSystemKind = False
        Exit Function
    End If
        ChkSystemKind = True
End Function

'新增資料到inventor
Public Function AddData0706(ce() As String) As Boolean
Dim i As Integer
'Add By Cheng 2002/11/14
Dim BolTransOk As Boolean
BolTransOk = True
    
    strSQL = "insert into changeevent("
    strSQL = strSQL + "ce01,ce02,ce03,ce04,ce05,ce06,ce07,ce08,ce09,ce10,"
    strSQL = strSQL + "ce11,ce12,ce13,ce14,ce15,ce16,ce17,ce18,ce19,ce20,"
    strSQL = strSQL + "ce21,ce22,ce23,ce24,ce25,ce26,ce27,ce28,ce29,ce30,"
    strSQL = strSQL + "ce31,ce32,ce33,ce34,ce35,ce36,ce37,ce38,ce39,ce40,"
    strSQL = strSQL + "ce41,ce42,ce43,ce44,ce45,ce46,ce47,ce48,ce49,ce50,"
    strSQL = strSQL + "ce51,ce52,ce53,ce54,ce55,ce56,ce57,ce58,ce59,ce60,"
    strSQL = strSQL + "ce61,ce62,ce63,ce64,ce65) values( "
    For i = 0 To UBound(ce)
        strSQL = strSQL + CNULL(ce(i)) + ","
    Next
    strSQL = Mid$(strSQL, 1, Len(strSQL) - 1)
    strSQL = strSQL + ") "
    Debug.Print strSQL
On Error GoTo Err
    cnnConnection.BeginTrans
    cnnConnection.Execute strSQL
    'Modify By Cheng 2002/11/14
    If BolTransOk Then
        cnnConnection.CommitTrans
    End If
    AddData0706 = True
    Exit Function
Err:
    'Add By Cheng 2002/11/14
    If Err.Number = -2147168237 Then
       BolTransOk = False
       Resume Next
    End If
    
    cnnConnection.RollbackTrans
    AddData0706 = False
    Exit Function
'    ErrorLog
End Function

Public Function EraseData0706(ce() As String) As Boolean
'Add By Cheng 2002/11/14
Dim BolTransOk As Boolean
BolTransOk = True

strSQL = "delete from changeevent where ce01=" + CNULL(ce(0))
    Debug.Print strSQL
On Error GoTo Err
    cnnConnection.BeginTrans
    cnnConnection.Execute strSQL
    'Modify By Cheng 2002/11/14
    If BolTransOk Then
        cnnConnection.CommitTrans
    End If
    EraseData0706 = True
    Exit Function
Err:
    'Add By Cheng 2002/11/14
    If Err.Number = -2147168237 Then
       BolTransOk = False
       Resume Next
    End If

    cnnConnection.RollbackTrans
    EraseData0706 = False
    ErrorLog
End Function

Public Function ModifyData0706(ce() As String) As Integer
Dim i As Integer, j As Integer
'Add By Cheng 2002/11/14
Dim BolTransOk As Boolean
BolTransOk = True
     
     strSQL = "update changeevent set "
    For i = 0 To 6
        For j = 0 To 9
                If i = 6 And j > 6 Then
                    Exit For
                Else
                    If (i * 10 + j) > 0 And (i * 10 + j) < 66 Then
                        strSQL = strSQL + "ce" + Trim$(Str(i)) + Trim$(Str(j)) + "=" + CNULL(ce(i * 10 + j - 1)) + ","
                    End If
                End If
        Next
    Next
    Debug.Print strSQL
    strSQL = Mid$(strSQL, 1, Len(strSQL) - 1)
    strSQL = strSQL + " where ce01=" + CNULL(ce(0))
    Debug.Print strSQL
On Error GoTo Err
    cnnConnection.BeginTrans
    cnnConnection.Execute strSQL
    'Modify By Cheng 2002/11/14
    If BolTransOk Then
        cnnConnection.CommitTrans
    End If
    ModifyData0706 = True
    Exit Function
Err:
    'Add By Cheng 2002/11/14
    If Err.Number = -2147168237 Then
       BolTransOk = False
       Resume Next
    End If
    
    cnnConnection.RollbackTrans
    ModifyData0706 = False
    ErrorLog
End Function

'檢查案件收文號是否存在
Public Function ChkIsExistCaseNo(strCaseNo As String, ByRef strCaseName As String) As Boolean
Dim CaseNo As String
Dim obj01 As Object
Dim Rs As ADODB.Recordset
    strSQL = "select  cp01  from caseprogress where cp09=" + CNULL(strCaseNo)
    Set obj01 = CreateObject("prjtaiedll.clspublicdata")
    Set Rs = obj01.ReadRst(strSQL)
    If Rs.EOF Then
        Rs.Close
        MsgBox "本收文號資料未存在於案件進度檔."
        ChkIsExistCaseNo = False
        Set obj01 = Nothing
        Exit Function
    Else
        strCaseName = IIf(IsNull(Rs(0)), "", Rs(0))
    End If
    Rs.Close
    Set obj01 = Nothing
    ChkIsExistCaseNo = True
End Function

Public Function AddData0705(fa() As String) As Boolean
Dim i As Integer
'Add By Cheng 2002/11/14
Dim BolTransOk As Boolean
BolTransOk = True
    
    strSQL = "insert into fagent("
    strSQL = strSQL + "FA01,FA02,FA03,FA04,FA05,FA06,FA07,FA08,FA09,FA10,"
    strSQL = strSQL + "FA11,FA12,FA13,FA14,FA15,FA16,FA17,FA18,FA19,FA20,"
    strSQL = strSQL + "FA21,FA22,FA23,FA24,FA25,FA26,FA27,FA28,FA29,FA30,"
    strSQL = strSQL + "FA31,FA32,FA33,FA34,FA35,FA36,FA37,FA38,FA39,FA40,"
    strSQL = strSQL + "FA41,FA42,FA43,FA44,FA45,FA52,FA53,FA54,fa55,fa56,"
    strSQL = strSQL + "fa57,fa58,fa59,fa60,fa61,fa62,FA63,FA64,FA65,FA66,FA67,FA68" + ") values( "
    For i = 0 To UBound(fa)
        strSQL = strSQL + CNULL(fa(i)) + ","
    Next
    strSQL = Mid$(strSQL, 1, Len(strSQL) - 1)
    strSQL = strSQL + ") "
    Debug.Print strSQL
On Error GoTo Err
    cnnConnection.BeginTrans
    cnnConnection.Execute strSQL
    'Modify By Cheng 2002/11/14
    If BolTransOk Then
        cnnConnection.CommitTrans
    End If
    AddData0705 = True
    Exit Function
Err:
    'Add By Cheng 2002/11/14
    If Err.Number = -2147168237 Then
       BolTransOk = False
       Resume Next
    End If
    cnnConnection.RollbackTrans
    AddData0705 = False
    ErrorLog
End Function

Public Function EraseData0705(fa() As String) As Boolean
'Add By Cheng 2002/11/14
Dim BolTransOk As Boolean
BolTransOk = True
    
    strSQL = "delete from fagent where fa01=" + CNULL(fa(0)) + " and fa02=" + CNULL(fa(1))
    Debug.Print strSQL
On Error GoTo Err
    cnnConnection.BeginTrans
    cnnConnection.Execute strSQL
    'Modify By Cheng 2002/11/14
    If BolTransOk Then
        cnnConnection.CommitTrans
    End If
    EraseData0705 = True
    Exit Function
Err:
    'Add By Cheng 2002/11/14
    If Err.Number = -2147168237 Then
       BolTransOk = False
       Resume Next
    End If

    cnnConnection.RollbackTrans
    EraseData0705 = False
    ErrorLog
End Function

Public Function ModifyData0705(fa() As String) As Integer
Dim i As Integer, j As Integer
'Add By Cheng 2002/11/14
Dim BolTransOk As Boolean
BolTransOk = True
     
     strSQL = "update fagent set "
    strSQL = strSQL + "FA03=" + CNULL(fa(2)) + ","
    strSQL = strSQL + "FA04 =" + CNULL(fa(3)) + ","
    strSQL = strSQL + "FA05=" + CNULL(fa(4)) + ","
    strSQL = strSQL + "FA06 =" + CNULL(fa(5)) + ","
    strSQL = strSQL + "FA07 =" + CNULL(fa(6)) + ","
    strSQL = strSQL + "FA08 =" + CNULL(fa(7)) + ","
    strSQL = strSQL + "FA09 = " + CNULL(fa(8)) + ","
    strSQL = strSQL + "FA10 =" + CNULL(fa(9)) + ","
    strSQL = strSQL + "FA11 =" + CNULL(fa(10)) + ","
    strSQL = strSQL + "FA12 =" + CNULL(fa(11)) + ","
    strSQL = strSQL + "FA13 =" + CNULL(fa(12)) + ","
    strSQL = strSQL + "FA14 =" + CNULL(fa(13)) + ","
    strSQL = strSQL + "FA15 =" + CNULL(fa(14)) + ","
    strSQL = strSQL + "FA16 =" + CNULL(fa(15)) + ","
    strSQL = strSQL + "FA17 =" + CNULL(fa(16)) + ","
    strSQL = strSQL + "FA18 =" + CNULL(fa(17)) + ","
    strSQL = strSQL + "FA19 =" + CNULL(fa(18)) + ","
    strSQL = strSQL + "FA20 =" + CNULL(fa(19)) + ","
    strSQL = strSQL + "FA21 =" + CNULL(fa(20)) + ","
    strSQL = strSQL + "FA22 =" + CNULL(fa(21)) + ","
    strSQL = strSQL + "FA23 =" + CNULL(fa(22)) + ","
    strSQL = strSQL + "FA24 =" + CNULL(fa(23)) + ","
    strSQL = strSQL + "FA25 =" + CNULL(fa(24)) + ","
    strSQL = strSQL + "FA26 =" + CNULL(fa(25)) + ","
    strSQL = strSQL + "FA27 =" + CNULL(fa(26)) + ","
    strSQL = strSQL + "FA28 =" + CNULL(fa(27)) + ","
    strSQL = strSQL + "FA29 =" + CNULL(fa(28)) + ","
    strSQL = strSQL + "FA30 =" + CNULL(fa(29)) + ","
    strSQL = strSQL + "FA31 =" + CNULL(fa(30)) + ","
    strSQL = strSQL + "FA32 =" + CNULL(fa(31)) + ","
    strSQL = strSQL + "FA33 =" + CNULL(fa(32)) + ","
    strSQL = strSQL + "FA34 =" + CNULL(fa(33)) + ","
    strSQL = strSQL + "FA35 =" + CNULL(fa(34)) + ","
    strSQL = strSQL + "FA36 =" + CNULL(fa(35)) + ","
    strSQL = strSQL + "FA37 =" + CNULL(fa(36)) + ","
    strSQL = strSQL + "FA38 =" + CNULL(fa(37)) + ","
    strSQL = strSQL + "FA39 =" + CNULL(fa(38)) + ","
    strSQL = strSQL + "FA40 =" + CNULL(fa(39)) + ","
    strSQL = strSQL + "FA41 =" + CNULL(fa(40)) + ","
    strSQL = strSQL + "FA42 =" + CNULL(fa(41)) + ","
    strSQL = strSQL + "FA43 =" + CNULL(fa(42)) + ","
    strSQL = strSQL + "FA44 =" + CNULL(fa(43)) + ","
    strSQL = strSQL + "FA45 =" + CNULL(fa(44)) + ","
    strSQL = strSQL + "FA52 =" + CNULL(fa(45)) + ","
    strSQL = strSQL + "FA53 =" + CNULL(fa(46)) + ","
    strSQL = strSQL + "FA54 =" + CNULL(fa(47)) + ","
    strSQL = strSQL + "FA55 =" + CNULL(fa(48)) + ","
    strSQL = strSQL + "FA56 =" + CNULL(fa(49)) + ","
    strSQL = strSQL + "FA57 =" + CNULL(fa(50)) + ","
    strSQL = strSQL + "FA58 =" + CNULL(fa(51)) + ","
    strSQL = strSQL + "FA59 =" + CNULL(fa(52)) + ","
    strSQL = strSQL + "FA60 =" + CNULL(fa(53)) + ","
    strSQL = strSQL + "FA61 =" + CNULL(fa(54)) + ","
    strSQL = strSQL + "FA62 =" + CNULL(fa(55)) + ","
    strSQL = strSQL + "FA63 =" + CNULL(fa(56)) + ","
    strSQL = strSQL + "FA64 =" + CNULL(fa(57)) + ","
    strSQL = strSQL + "FA65 =" + CNULL(fa(58)) + ","
    strSQL = strSQL + "FA66 =" + CNULL(fa(59)) + ","
    strSQL = strSQL + "FA67 =" + CNULL(fa(60)) + ","
    strSQL = strSQL + "FA68 =" + CNULL(fa(61))
    strSQL = strSQL + " where fa01=" + CNULL(fa(0)) + " and  fa02=" + CNULL(fa(1))
    Debug.Print strSQL
On Error GoTo Err
    cnnConnection.BeginTrans
    cnnConnection.Execute strSQL
    'Modify By Cheng 2002/11/14
    If BolTransOk Then
        cnnConnection.CommitTrans
    End If
    ModifyData0705 = True
    Exit Function
Err:
    'Add By Cheng 2002/11/14
    If Err.Number = -2147168237 Then
       BolTransOk = False
       Resume Next
    End If

    cnnConnection.RollbackTrans
    ModifyData0705 = False
    ErrorLog
End Function

