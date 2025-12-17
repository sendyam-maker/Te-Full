VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class0701"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'新增資料到inventor
Public Function AddData0709(in0() As String) As Boolean
Dim i As Integer
'Add By Cheng 2002/11/14
Dim BolTransOk As Boolean
BolTransOk = True

    strSQL = "insert into inventor(in01,in02,in03,in04,in05,in06,in07,in08,in09,in10,in11,in12) values("
    For i = 0 To UBound(in0()) - 1
        'Modify By Cheng 2002/12/11
'        strSQL = strSQL + CNULL(in0(i)) + ","
        strSQL = strSQL + CNULL(ChgSQL(in0(i))) + ","
    Next
    strSQL = Left(strSQL, Len(strSQL) - 1)
    strSQL = strSQL + ")"
    Debug.Print strSQL
On Error GoTo Err
    cnnConnection.BeginTrans
    cnnConnection.Execute strSQL
    'Modify By Cheng 2002/11/14
    If BolTransOk Then
        cnnConnection.CommitTrans
    End If
    AddData0709 = True
    Exit Function
Err:
    'Add By Cheng 2002/11/14
    If Err.Number = -2147168237 Then
       BolTransOk = False
       Resume Next
    End If
    
    cnnConnection.RollbackTrans
    AddData0709 = False
    ErrorLog
End Function
Public Function EraseData0709(in0() As String) As Boolean
'Add By Cheng 2002/11/14
Dim BolTransOk As Boolean
BolTransOk = True
    
    strSQL = "delete from inventor where in01='" + in0(0) + "' and in02='" + in0(1) + "' "
On Error GoTo Err
    cnnConnection.BeginTrans
    cnnConnection.Execute strSQL
    'Modify By Cheng 2002/11/14
    If BolTransOk Then
        cnnConnection.CommitTrans
    End If
    EraseData0709 = True
    Exit Function
Err:
    'Add By Cheng 2002/11/14
    If Err.Number = -2147168237 Then
       BolTransOk = False
       Resume Next
    End If
    
    cnnConnection.RollbackTrans
    EraseData0709 = False
    ErrorLog
End Function

Public Function ModifyData0709(in0() As String) As Integer
Dim i As Integer
'Add By Cheng 2002/11/14
Dim BolTransOk As Boolean
BolTransOk = True
    
    strSQL = "update inventor set "
    'Modify By Cheng 2002/12/11
'    strSQL = strSQL + "in03=" + CNULL(in0(2)) + ", in04=" + CNULL(in0(3)) + ",in05=" + CNULL(in0(4)) + _
'      ",in06=" + CNULL(in0(5)) + ",in07=" + CNULL(in0(6)) + ",in08=" + CNULL(in0(7)) + ",in09=" + CNULL(in0(8)) + _
'      ",in10=" + CNULL(in0(9)) + ",in11=" + CNULL(in0(10)) + ",in12=" + CNULL(in0(11)) + " where in01='" + in0(0) + "' and in02='" + in0(1) + "'"
    strSQL = strSQL + "in03=" + CNULL(ChgSQL(in0(2))) + ", in04=" + CNULL(ChgSQL(in0(3))) + ",in05=" + CNULL(ChgSQL(in0(4))) + _
      ",in06=" + CNULL(ChgSQL(in0(5))) + ",in07=" + CNULL(ChgSQL(in0(6))) + ",in08=" + CNULL(ChgSQL(in0(7))) + ",in09=" + CNULL(ChgSQL(in0(8))) + _
      ",in10=" + CNULL(ChgSQL(in0(9))) + ",in11=" + CNULL(ChgSQL(in0(10))) + ",in12=" + CNULL(ChgSQL(in0(11))) + " where in01='" + in0(0) + "' and in02='" + in0(1) + "'"
On Error GoTo Err
    cnnConnection.BeginTrans
    cnnConnection.Execute strSQL
    'Modify By Cheng 2002/11/14
    If BolTransOk Then
        cnnConnection.CommitTrans
    End If
    ModifyData0709 = True
    Exit Function
Err:
    'Add By Cheng 2002/11/14
    If Err.Number = -2147168237 Then
       BolTransOk = False
       Resume Next
    End If
    
    cnnConnection.RollbackTrans
    ModifyData0709 = False
    ErrorLog
End Function
'新增資料
Public Function AddData0707(dl() As String) As Boolean
Dim i As Integer
Dim j As Integer
Dim Date1 As String
Dim Date2 As String
Dim Date3 As String
'Add By Cheng 2002/11/14
Dim BolTransOk As Boolean
BolTransOk = True
    
    strSQL = "insert into datelimit (dl01,dl02,dl03,dl04) values("
    strSQL = strSQL + CNULL(dl(0)) + "," + CNULL(ChangeTStringToWString(dl(1))) + "," + CNULL(ChangeTStringToWString(dl(2))) + "," + CNULL(ChangeTStringToWString(dl(3))) + ")"
    Debug.Print strSQL
On Error GoTo Err
    cnnConnection.BeginTrans
    cnnConnection.Execute strSQL
    'Modify By Cheng 2002/11/14
    If BolTransOk Then
        cnnConnection.CommitTrans
    End If
    AddData0707 = True
    Exit Function
Err:
    'Add By Cheng 2002/11/14
    If Err.Number = -2147168237 Then
       BolTransOk = False
       Resume Next
    End If

    cnnConnection.RollbackTrans
    AddData0707 = False
    ErrorLog
End Function

Public Function EraseData0707(dl() As String) As Boolean
'Add By Cheng 2002/11/14
Dim BolTransOk As Boolean
BolTransOk = True
    
    strSQL = "delete from datelimit where dl01=" + CNULL(dl(0)) + " and dl02=" + CNULL(ChangeTStringToWString(dl(1)))
    Debug.Print strSQL
On Error GoTo Err
    cnnConnection.BeginTrans
    cnnConnection.Execute strSQL
    'Modify By Cheng 2002/11/14
    If BolTransOk Then
        cnnConnection.CommitTrans
    End If
    EraseData0707 = True
    Exit Function
Err:
    'Add By Cheng 2002/11/14
    If Err.Number = -2147168237 Then
       BolTransOk = False
       Resume Next
    End If
    
    cnnConnection.RollbackTrans
    EraseData0707 = False
    ErrorLog
End Function
'新增資料
Public Function ModifyData0707(dl() As String) As Boolean
Dim i As Integer
Dim j As Integer
Dim Date1 As String
Dim Date2 As String
Dim Date3 As String
'Add By Cheng 2002/11/14
Dim BolTransOk As Boolean
BolTransOk = True

    strSQL = "update datelimit set dl03=" + CNULL(ChangeTStringToWString(dl(2))) + ","
    strSQL = strSQL + " dl04=" + CNULL(ChangeTStringToWString(dl(3)))
    strSQL = strSQL + " where dl01=" + CNULL(dl(0)) + " and dl02=" + CNULL(ChangeTStringToWString(dl(1)))
    Debug.Print strSQL
On Error GoTo Err
    cnnConnection.BeginTrans
    cnnConnection.Execute strSQL
    'Modify By Cheng 2002/11/14
    If BolTransOk Then
        cnnConnection.CommitTrans
    End If
    ModifyData0707 = True
    Exit Function
Err:
    'Add By Cheng 2002/11/14
    If Err.Number = -2147168237 Then
       BolTransOk = False
       Resume Next
    End If
    
    cnnConnection.RollbackTrans
    ModifyData0707 = False
    ErrorLog
End Function
'新增資料
Public Function AddData0706(ce() As String) As Boolean
Dim i As Integer
Dim j As Integer
'Add By Cheng 2002/11/14
Dim BolTransOk As Boolean
BolTransOk = True

    strSQL = "insert into changeevent("
    strSQL = strSQL + "ce01,ce02,ce03,ce04,ce05,ce06,ce07,ce08,ce09,ce10,"
    strSQL = strSQL + "ce11,ce12,ce13,ce14,ce15,ce16,ce17,ce18,ce19,ce20,"
    strSQL = strSQL + "ce21,ce22,ce23,ce24,ce25,ce26,ce27,ce28,ce29,ce30,"
    strSQL = strSQL + "ce31,ce32,ce33,ce34,ce35,ce36,ce37,ce38,ce39,ce40,"
    strSQL = strSQL + "ce41,ce42,ce43,ce44,ce45,ce46,ce47,ce48,ce49,ce50,"
    strSQL = strSQL + "ce51,ce52,ce53,ce54,ce55,ce56,ce57,ce58,ce59,ce60,ce61,ce62,ce63,ce64,ce65)  values('"
    strSQL = strSQL + ce(0) + "'," + ce(1) + ",'"
    For i = 2 To 65
        strSQL = strSQL + ce(i) + "','"
    Next
    strSQL = Mid$(strSQL, 1, Len(strSQL) - 2) + ")"
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
    
    strSQL = "delete from changeevent where ce01='" + ce(0) + "'"
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
Dim i As Integer
Dim j As Integer
'Add By Cheng 2002/11/14
Dim BolTransOk As Boolean
BolTransOk = True
    
    strSQL = "update changeevent set "
    strSQL = strSQL + "ce02=" + IIf(ce(1) <> "", ce(1), "0") + " , "
    For j = 0 To 5
    For i = 0 To 9
        If j = 0 And i > 2 Then
            strSQL = strSQL + "CE0" + Trim$(Str(i)) + "='" + ce(i - 1) + "',"
        End If
        If j = 5 And i > 7 Then
        Else
            If j > 0 Then
                strSQL = strSQL + "CE" + Trim$(Str(j)) + Trim$(Str(i)) + "='" + ce((j * 10 + i) - 1) + "',"
            End If
        End If
    Next
    Next
    
    strSQL = Mid$(strSQL, 1, Len(strSQL) - 1) + "  where CE01='" + ce(0) + "'"
    
    
    
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

'新增資料
Public Function AddData0710(fac() As String) As Boolean
Dim i As Integer
Dim j As Integer
'Add By Cheng 2002/11/14
Dim BolTransOk As Boolean
BolTransOk = True

    strSQL = "insert into facase(fac01,fac02,fac03,fac04) values("
    strSQL = strSQL + CNULL(fac(0)) + "," + CNULL(fac(1)) + "," + CNULL(fac(2)) + "," + CNULL(fac(3)) + ")"
    Debug.Print strSQL
On Error GoTo Err
    cnnConnection.BeginTrans
    cnnConnection.Execute strSQL
    'Modify By Cheng 2002/11/14
    If BolTransOk Then
        cnnConnection.CommitTrans
    End If
    AddData0710 = True
    Exit Function
Err:
    'Add By Cheng 2002/11/14
    If Err.Number = -2147168237 Then
       BolTransOk = False
       Resume Next
    End If

    cnnConnection.RollbackTrans
    AddData0710 = False
    ErrorLog
End Function

Public Function EraseData0710(fac() As String) As Boolean
'Add By Cheng 2002/11/14
Dim BolTransOk As Boolean
BolTransOk = True
    
    strSQL = "delete from facase where fac01='" + fac(0) + "' and fac02='" + fac(1) + "'"
    Debug.Print strSQL
On Error GoTo Err
    cnnConnection.BeginTrans
    cnnConnection.Execute strSQL
    'Modify By Cheng 2002/11/14
    If BolTransOk Then
        cnnConnection.CommitTrans
    End If
    EraseData0710 = True
    Exit Function
Err:
    'Add By Cheng 2002/11/14
    If Err.Number = -2147168237 Then
       BolTransOk = False
       Resume Next
    End If

    cnnConnection.RollbackTrans
    EraseData0710 = False
    ErrorLog
End Function

Public Function ModifyData0710(fac() As String) As Integer
Dim i As Integer
Dim j As Integer
'Add By Cheng 2002/11/14
Dim BolTransOk As Boolean
BolTransOk = True

    strSQL = "update facase set fac03=" + IIf(fac(2) = "", "0", fac(2)) + ","
    strSQL = strSQL + "fac04='" + fac(3) + "' "
    strSQL = strSQL + " where fac01='" + fac(0) + "' and fac02='" + fac(1) + "'"
    Debug.Print strSQL
On Error GoTo Err
    cnnConnection.BeginTrans
    cnnConnection.Execute strSQL
    'Modify By Cheng 2002/11/14
    If BolTransOk Then
        cnnConnection.CommitTrans
    End If
    ModifyData0710 = True
    Exit Function
Err:
    'Add By Cheng 2002/11/14
    If Err.Number = -2147168237 Then
       BolTransOk = False
       Resume Next
    End If
     
    cnnConnection.RollbackTrans
    ModifyData0710 = False
    ErrorLog
End Function

'新增資料
Public Function AddData0708(cf() As String) As Boolean
Dim i As Integer
Dim j As Integer
'Add By Cheng 2002/11/14
Dim BolTransOk As Boolean
BolTransOk = True

    strSQL = "insert into casefee(cf01,cf02,cf03,cf04,cf05,cf06,cf07,"
    strSQL = strSQL + "cf08,cf09,cf10,cf11,cf12,cf13,cf14,cf15) values('"
    strSQL = strSQL + cf(0) + "','" + cf(1) + "','" + cf(2) + "'," + IIf(cf(3) <> "", cf(3), "0") + "," + IIf(cf(4) <> "", cf(4), "0") + "," + IIf(cf(5) <> "", cf(5), "0") + ","
    strSQL = strSQL + IIf(cf(6) <> "", cf(6), "0") + "," + IIf(cf(7) <> "", cf(7), "0") + ",'" + cf(8) + "','" + cf(9) + "'," + IIf(cf(10) <> "", cf(10), "0") + "," + IIf(cf(11) <> "", cf(11), "0") + "," + IIf(cf(12) <> "", cf(12), "0") + ","
    strSQL = strSQL + IIf(cf(13) <> "", cf(13), "0") + ",'" + cf(14) + "')"
    Debug.Print strSQL
On Error GoTo Err
    cnnConnection.BeginTrans
    cnnConnection.Execute strSQL
    'Modify By Cheng 2002/11/14
    If BolTransOk Then
        cnnConnection.CommitTrans
    End If
    AddData0708 = True
    Exit Function
Err:
    'Add By Cheng 2002/11/14
    If Err.Number = -2147168237 Then
       BolTransOk = False
       Resume Next
    End If

    cnnConnection.RollbackTrans
    AddData0708 = False
    ErrorLog
End Function

Public Function EraseData0708(cf() As String) As Boolean
'Add By Cheng 2002/11/14
Dim BolTransOk As Boolean
BolTransOk = True
    
    strSQL = "delete from casefee where cf01='" + cf(0) + "' and cf02='" + cf(1) + "' and cf03='" + cf(2) + "'"
    Debug.Print strSQL
On Error GoTo Err
    cnnConnection.BeginTrans
    cnnConnection.Execute strSQL
    'Modify By Cheng 2002/11/14
    If BolTransOk Then
        cnnConnection.CommitTrans
    End If
    EraseData0708 = True
    Exit Function
Err:
    'Add By Cheng 2002/11/14
    If Err.Number = -2147168237 Then
       BolTransOk = False
       Resume Next
    End If

    cnnConnection.RollbackTrans
    EraseData0708 = False
    ErrorLog
End Function

Public Function ModifyData0708(cf() As String) As Integer
Dim i As Integer
Dim j As Integer
'Add By Cheng 2002/11/14
Dim BolTransOk As Boolean
BolTransOk = True
    
    strSQL = "update casefee set "
    strSQL = strSQL + "cf04=" + IIf(cf(3) <> "", cf(3), "0") + ","
    strSQL = strSQL + "cf05=" + IIf(cf(4) <> "", cf(4), "0") + ","
    strSQL = strSQL + "cf06=" + IIf(cf(5) <> "", cf(5), "0") + ","
    strSQL = strSQL + "cf07=" + IIf(cf(6) <> "", cf(6), "0") + ","
    strSQL = strSQL + "cf08=" + IIf(cf(7) <> "", cf(7), "0") + ","
    strSQL = strSQL + "cf09='" + cf(8) + "',"
    strSQL = strSQL + "cf10='" + cf(9) + "',"
    strSQL = strSQL + "cf11=" + IIf(cf(10) <> "", cf(10), "0") + ","
    strSQL = strSQL + "cf12=" + IIf(cf(11) <> "", cf(11), "0") + ","
    strSQL = strSQL + "cf13=" + IIf(cf(12) <> "", cf(12), "0") + ","
    strSQL = strSQL + "cf14=" + IIf(cf(13) <> "", cf(13), "0") + ","
    strSQL = strSQL + "cf15='" + cf(14) + "' "
    strSQL = strSQL + " where cf01='" + cf(0) + "' and cf02='" + cf(1) + "' and cf03='" + cf(2) + "'"
    Debug.Print strSQL
On Error GoTo Err
    cnnConnection.BeginTrans
    cnnConnection.Execute strSQL
    'Modify By Chenhg 2002/11/14
    If BolTransOk Then
        cnnConnection.CommitTrans
    End If
    ModifyData0708 = True
    Exit Function
Err:
    'Add By Cheng 2002/11/14
    If Err.Number = -2147168237 Then
       BolTransOk = False
       Resume Next
    End If

    cnnConnection.RollbackTrans
    ModifyData0708 = False
    ErrorLog
End Function

