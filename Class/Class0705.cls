VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class0705"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'新增資料到inventor
Public Function AddData1201(ce() As String) As Boolean
Dim i As Integer
'Add By Cheng 2002/11/14
Dim BolTransOk As Boolean
BolTransOk = True
    
    Debug.Print strSQL
On Error GoTo Err
    cnnConnection.BeginTrans
    cnnConnection.Execute strSQL
    'Modify By Cheng 2002/11/14
    If BolTransOk Then
        cnnConnection.CommitTrans
    End If
    AddData1201 = True
    Exit Function
Err:
    'Add By Cheng 2002/11/14
    If Err.Number = -2147168237 Then
       BolTransOk = False
       Resume Next
    End If

    cnnConnection.RollbackTrans
    AddData1201 = False
    Exit Function
'    ErrorLog
End Function

Public Function EraseData1201(ce() As String) As Boolean
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
    EraseData1201 = True
    Exit Function
Err:
    'Add By Cheng 2002/11/14
    If Err.Number = -2147168237 Then
       BolTransOk = False
       Resume Next
    End If

    cnnConnection.RollbackTrans
    EraseData1201 = False
    ErrorLog
End Function
Public Function ModifyData1201(ce() As String) As Integer
Dim i As Integer, j As Integer
'Add By Cheng 2002/11/14
Dim BolTransOk As Boolean
BolTransOk = True
     
     strSQL = "update changeevent set "
    For i = 0 To 6
        For j = 0 To 9
                If i = 6 And j > 2 Then
                    Exit For
                Else
                    If (i * 10 + j) > 0 And (i * 10 + j) < 65 Then
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
    ModifyData1201 = True
    Exit Function
Err:
    'Add By Cheng 2002/11/14
    If Err.Number = -2147168237 Then
       BolTransOk = False
       Resume Next
    End If

    cnnConnection.RollbackTrans
    ModifyData1201 = False
    ErrorLog
End Function
Public Function ModifyData0712(fa() As String) As Boolean
Dim i As Integer
Dim Rss As ADODB.Recordset
Dim obj01 As Object
Dim strNewFA02 As String
Dim A(54) As String
'Add By Cheng 2002/11/14
Dim BolTransOk As Boolean
BolTransOk = True
    
    strSQL = "select fa02,fa03,fa04,fa05,fa06,fa07,fa08,fa09,fa10,"
    strSQL = strSQL + " fa11,fa12,fa13,fa14,fa15,fa16,fa17,fa18,fa19,fa20,"
    strSQL = strSQL + " fa21,fa22,fa23,fa24,fa25,fa26,fa27,fa28,fa29,fa30,"
    strSQL = strSQL + " fa31,fa32,fa33,fa34,fa35,fa36,fa37,fa38,fa39,fa40,"
    strSQL = strSQL + " fa41,fa42,fa43,fa44,fa45,fa55,fa56,fa57,fa58,fa59,"
    strSQL = strSQL + " fa60,fa61,fa62,fa63,fa64,fa65 "
    strSQL = strSQL + "from fagent where fa01=" + CNULL(fa(0)) + "order by fa02 desc"
    Debug.Print strSQL
    Set obj01 = CreateObject("prjtaiedll.clspublicData")
    Set Rss = obj01.ReadRst(strSQL)
    For i = 0 To 54
        A(i) = IIf(IsNull(Rss(i)), "", Rss(i))
    Next
    Rss.Close
    strNewFA02 = Trim$(Str(Val(A(0)) + 1))
    strSQL = "insert into fagent(fa01,fa02) values(" + CNULL(fa(0)) + "," + CNULL(strNewFA02) + ")"
    Debug.Print strSQL
    cnnConnection.BeginTrans
    cnnConnection.Execute strSQL
    strSQL = "update fagent set "
    strSQL = strSQL + "fa03=" + CNULL(A(1)) + ",fa04=" + CNULL(A(2)) + ",fa05=" + CNULL(A(3)) + ",fa06=" + CNULL(A(4)) + ",fa07=" + CNULL(A(5))
    strSQL = strSQL + ",fa08=" + CNULL(A(6)) + ",fa09=" + CNULL(A(7)) + ",fa10=" + CNULL(A(8)) + ",fa11=" + CNULL(A(9)) + ",fa12=" + CNULL(A(10))
    strSQL = strSQL + ",fa13=" + CNULL(A(11)) + ",fa14=" + CNULL(A(12)) + ",fa15=" + CNULL(A(13)) + ",fa16=" + CNULL(A(14)) + ",fa17=" + CNULL(A(15))
    strSQL = strSQL + ",fa18=" + CNULL(A(16)) + ",fa19=" + CNULL(A(17)) + ",fa20=" + CNULL(A(18)) + ",fa21=" + CNULL(A(19)) + ",fa22=" + CNULL(A(20))
    strSQL = strSQL + ",fa23=" + CNULL(A(21)) + ",fa24=" + CNULL(A(22)) + ",fa25=" + CNULL(A(23)) + ",fa26=" + CNULL(A(24)) + ",fa27=" + CNULL(A(25))
    strSQL = strSQL + ",fa28=" + CNULL(A(26)) + ",fa29=" + CNULL(A(27)) + ",fa30=" + CNULL(A(28)) + ",fa31=" + CNULL(A(29)) + ",fa32=" + CNULL(A(30))
    strSQL = strSQL + ",fa33=" + CNULL(A(31)) + ",fa34=" + CNULL(A(32)) + ",fa35=" + CNULL(A(33)) + ",fa36=" + CNULL(A(34)) + ",fa37=" + CNULL(A(35))
    strSQL = strSQL + ",fa38=" + CNULL(A(36)) + ",fa39=" + CNULL(A(37)) + ",fa40=" + CNULL(A(38)) + ",fa41=" + CNULL(A(39)) + ",fa42=" + CNULL(A(40))
    strSQL = strSQL + ",fa43=" + CNULL(A(41)) + ",fa44=" + CNULL(A(42)) + ",fa45=" + CNULL(A(43)) + ",fa55=" + CNULL(A(44)) + ",fa56=" + CNULL(A(45))
    strSQL = strSQL + ",fa57=" + CNULL(A(46)) + ",fa58=" + CNULL(A(47)) + ",fa59=" + CNULL(A(48)) + ",fa60=" + CNULL(A(49)) + ",fa61=" + CNULL(A(50))
    strSQL = strSQL + ",fa62=" + CNULL(A(51)) + ",fa63=" + CNULL(A(52)) + ",fa64=" + CNULL(A(53)) + ",fa65=" + CNULL(A(54))
    strSQL = strSQL + " where fa01=" + CNULL(fa(0)) + " and fa02=" + CNULL(strNewFA02)
    Debug.Print strSQL
    cnnConnection.Execute strSQL
    strSQL = "update fagent set "
    strSQL = strSQL + "fa03=" + CNULL(A(1)) + ",fa04=" + CNULL(fa(1)) + ",fa05=" + CNULL(fa(2)) + ",fa06=" + CNULL(fa(6)) + ",fa07=" + CNULL(A(5))
    strSQL = strSQL + ",fa08=" + CNULL(A(6)) + ",fa09=" + CNULL(A(7)) + ",fa10=" + CNULL(A(8)) + ",fa11=" + CNULL(A(9)) + ",fa12=" + CNULL(A(10))
    strSQL = strSQL + ",fa13=" + CNULL(A(11)) + ",fa14=" + CNULL(A(12)) + ",fa15=" + CNULL(A(13)) + ",fa16=" + CNULL(A(14)) + ",fa17=" + CNULL(A(15))
    strSQL = strSQL + ",fa18=" + CNULL(A(16)) + ",fa19=" + CNULL(A(17)) + ",fa20=" + CNULL(A(18)) + ",fa21=" + CNULL(A(19)) + ",fa22=" + CNULL(A(20))
    strSQL = strSQL + ",fa23=" + CNULL(A(21)) + ",fa24=" + CNULL(A(22)) + ",fa25=" + CNULL(A(23)) + ",fa26=" + CNULL(A(24)) + ",fa27=" + CNULL(A(25))
    strSQL = strSQL + ",fa28=" + CNULL(A(26)) + ",fa29=" + CNULL(A(27)) + ",fa30=" + CNULL(A(28)) + ",fa31=" + CNULL(A(29)) + ",fa32=" + CNULL(A(30))
    strSQL = strSQL + ",fa33=" + CNULL(A(31)) + ",fa34=" + CNULL(A(32)) + ",fa35=" + CNULL(A(33)) + ",fa36=" + CNULL(A(34)) + ",fa37=" + CNULL(A(35))
    strSQL = strSQL + ",fa38=" + CNULL(A(36)) + ",fa39=" + CNULL(A(37)) + ",fa40=" + CNULL(A(38)) + ",fa41=" + CNULL(A(39)) + ",fa42=" + CNULL(A(40))
    strSQL = strSQL + ",fa43=" + CNULL(A(41)) + ",fa44=" + CNULL(A(42)) + ",fa45=" + CNULL(A(43)) + ",fa55=" + CNULL(A(44)) + ",fa56=" + CNULL(A(45))
    strSQL = strSQL + ",fa57=" + CNULL(A(46)) + ",fa58=" + CNULL(A(47)) + ",fa59=" + CNULL(A(48)) + ",fa60=" + CNULL(A(49)) + ",fa61=" + CNULL(A(50))
    strSQL = strSQL + ",fa62=" + CNULL(A(51)) + ",fa63=" + CNULL(fa(3)) + ",fa64=" + CNULL(fa(4)) + ",fa65=" + CNULL(fa(5))
    strSQL = strSQL + " where fa01=" + CNULL(fa(0)) + " and fa02='0'"
    Debug.Print strSQL
    cnnConnection.Execute strSQL
    'Modify By Cheng 2002/11/14
    If BolTransOk Then
        cnnConnection.CommitTrans
    End If
    ModifyData0712 = True
    Exit Function
Err:
    'Add By Cheng 2002/11/14
    If Err.Number = -2147168237 Then
       BolTransOk = False
       Resume Next
    End If

    cnnConnection.RollbackTrans
    ModifyData0712 = False
    ErrorLog
End Function
