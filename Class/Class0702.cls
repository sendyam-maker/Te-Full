VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class0702"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'取得解除期限原因
Public Function GetReasonOfRelief(ByRef StrRorNum As String, ByRef StrRorName As String) As Boolean
Dim obj01 As Object
Dim Rs As ADODB.Recordset
    strSQL = "select ROR02 from ReasonOfRelief where Ror01=" + CNULL(StrRorNum)
    Set obj01 = CreateObject("prjtaiedll.clspublicdata")
    Set Rs = obj01.ReadRst(strSQL)
    If Not Rs.EOF Then
        StrRorName = IIf(IsNull(Rs(0)), "", Rs(0))
        GetReasonOfRelief = True
        Rs.Close
        Set obj01 = Nothing
    Else
        MsgBox "解除期限原因代碼錯誤.", vbCritical
        GetReasonOfRelief = False
        Rs.Close
        Set obj01 = Nothing
        Exit Function
    End If
End Function

'取得客戶名稱
Public Function GetCustomer1(ByRef strCustomer As String, ByRef strCustomerName As String, Optional ByRef strZipCode As String) As Boolean
Dim strSQL As String, rsRecordset As New ADODB.Recordset, i As Integer

On Error GoTo Err
strCustomer = ChangeCustomerL(strCustomer)
If Left(strCustomer, 1) = 客戶編號 Then
   strSQL = "select cu01||cu02,nvl(cu05,nvl(cu04,nvl(cu06,''))),cu27 from customer " & _
      "where " & ChgCustomer(strCustomer) & " order by cu01||cu02"
ElseIf Left(strCustomer, 1) = 代理人編號 Then
   strSQL = "select fa01||fa02,nvl(cu05,nvl(cu04,nvl(cu06,''))),cu27,cu01||cu02 " & _
      "from customer,fagent where cu01=fa03 and " & ChgFagent(strCustomer) + " order by cu01||cu02"
Else
   MsgBox "客戶代碼錯誤!!", vbCritical
   Exit Function
End If
rsRecordset.CursorLocation = adUseClient
rsRecordset.Open strSQL, cnnConnection
If rsRecordset.RecordCount > 0 Then
   If rsRecordset.Fields(0) = strCustomer Then
      If Left(strCustomer, 1) = 代理人編號 Then
         strCustomer = rsRecordset.Fields(3)
      End If
      strCustomer = ChangeCustomerS(strCustomer)
      strCustomerName = rsRecordset.Fields(1)
      strZipCode = IIf(IsNull(rsRecordset.Fields(2)), "", rsRecordset.Fields(2))
   End If
   GetCustomer1 = True
Else
   MsgBox "申請人代碼錯誤!!", vbCritical
End If
rsRecordset.Close
Exit Function
Err:
ErrorLog
End Function

Public Function Adddata0702(sp() As String) As Boolean
Dim i As Integer
Dim j As Integer
'Add By Cheng 2002/11/14
Dim BolTransOk As Boolean
BolTransOk = True
    'Modify By Cheng 2003/09/24
'    strSQL = "insert into Servicepractice(sp01,sp02,sp03,sp04,sp05,sp06,sp07,sp08,sp09,sp10,sp15,sp16,sp17,sp34,"
'    strSQL = strSQL + "sp28,sp18,sp29,sp26,sp30,sp37,sp35,sp27,sp31,sp33,sp36"
'    strSQL = strSQL + ") values("
    strSQL = "insert into Servicepractice(sp01,sp02,sp03,sp04,sp05,sp06,sp07,sp08,sp09,sp10,sp15,sp16,sp17,sp34,"
    strSQL = strSQL + "sp28,sp18,sp29,sp26,sp30,sp37,sp35,sp27,sp31,sp33,sp36, SP67"
    strSQL = strSQL + ") values("
    For i = 0 To 25
        strSQL = strSQL + CNULL(sp(i)) + ","
    Next
    strSQL = Mid$(strSQL, 1, Len(strSQL) - 1)
    strSQL = strSQL + ")"
    Debug.Print strSQL
On Error GoTo Err
    cnnConnection.BeginTrans
    cnnConnection.Execute strSQL
    'Modify By Cheng 2002/11/14
    If BolTransOk Then
        cnnConnection.CommitTrans
    End If
    Adddata0702 = True
    Exit Function
Err:
    'Add By Cheng 2002/11/14
    If Err.Number = -2147168237 Then
       BolTransOk = False
       Resume Next
    End If
    
    cnnConnection.RollbackTrans
    Adddata0702 = False
    ErrorLog
End Function

Public Function EraseData0702(sp() As String) As Boolean
'Add By Cheng 2002/11/14
Dim BolTransOk As Boolean
BolTransOk = True
    
    strSQL = "delete from servicepractice where sp01='" + sp(0) + "' and sp02='" + sp(1) + "' and sp03='" + sp(2) + "' and sp04='" + sp(3) + "'"
    Debug.Print strSQL
On Error GoTo Err
    cnnConnection.BeginTrans
    cnnConnection.Execute strSQL
    'Modify By Cheng 2002/11/14
    If BolTransOk Then
        cnnConnection.CommitTrans
    End If
    EraseData0702 = True
    Exit Function
Err:
    'Add By Cheng 2002/11/14
    If Err.Number = -2147168237 Then
       BolTransOk = False
       Resume Next
    End If
    
    cnnConnection.RollbackTrans
    EraseData0702 = False
    ErrorLog
End Function

Public Function EraseData0701(strSQL1 As String, Strsql2 As String) As Boolean
'Add By Cheng 2002/11/14
Dim BolTransOk As Boolean
BolTransOk = True
    
    Debug.Print strSQL1
On Error GoTo Err
    cnnConnection.BeginTrans
    cnnConnection.Execute strSQL1
    cnnConnection.Execute Strsql2
    'Modify By Cheng 2002/11/14
    If BolTransOk Then
        cnnConnection.CommitTrans
    End If
    EraseData0701 = True
    Exit Function
Err:
    'Add By Cheng 2002/11/14
    If Err.Number = -2147168237 Then
       BolTransOk = False
       Resume Next
    End If
    
    cnnConnection.RollbackTrans
    EraseData0701 = False
    ErrorLog
End Function

Public Function ModifyData0702(sp() As String, Optional BolWDB As Boolean = False) As Integer
Dim i As Integer
Dim j As Integer
'Add By Cheng 2002/11/14
Dim BolTransOk As Boolean
BolTransOk = True
    
    strSQL = "update servicepractice set sp05=" + CNULL(sp(4)) + ","
    '910910  nick tigger
    '***** start
    If BolWDB = True Then
        strSQL = "begin user_data.user_enabled:=1; " & strSQL
    End If
    '***** end
    strSQL = strSQL + "sp06=" + CNULL(sp(5)) + ",sp07=" + CNULL(sp(6)) + ","
    strSQL = strSQL + "sp08=" + CNULL(sp(7)) + ",sp09=" + CNULL(sp(8)) + ","
    strSQL = strSQL + "sp10=" + CNULL(sp(9)) + ",sp15=" + CNULL(sp(10)) + ","
    strSQL = strSQL + "sp16=" + CNULL(sp(11)) + ",sp17=" + CNULL(sp(12)) + ","
    strSQL = strSQL + "sp34=" + CNULL(sp(13)) + ",sp28=" + CNULL(sp(14)) + ","
    strSQL = strSQL + "sp18=" + CNULL(sp(15)) + ",sp29=" + CNULL(sp(16)) + ","
    strSQL = strSQL + "sp26=" + CNULL(sp(17)) + ",sp30=" + CNULL(sp(18)) + ","
    strSQL = strSQL + "sp37=" + CNULL(sp(19)) + ",sp35=" + CNULL(sp(20)) + ","
    strSQL = strSQL + "sp27=" + CNULL(sp(21)) + ",sp31=" + CNULL(sp(22)) + ","
    strSQL = strSQL + "sp33=" + CNULL(sp(23)) + ",sp36=" + CNULL(sp(24)) + ",sp67=" + CNULL(sp(25))
    strSQL = strSQL + " where sp01=" + CNULL(sp(0)) + " and sp02=" + CNULL(sp(1)) + " and sp03=" + CNULL(sp(2)) + "  and sp04=" + CNULL(sp(3))
    '910910  nick tigger
    '***** start
    If BolWDB = True Then
        strSQL = strSQL & "; end;"
    End If
    '***** end
On Error GoTo Err
    cnnConnection.BeginTrans
    cnnConnection.Execute strSQL
    'Modify By Cheng 2002/11/14
    If BolTransOk Then
        cnnConnection.CommitTrans
    End If
    ModifyData0702 = True
    Exit Function
Err:
    'Add By Cheng 2002/11/14
    If Err.Number = -2147168237 Then
       BolTransOk = False
       Resume Next
    End If
    
    cnnConnection.RollbackTrans
    ModifyData0702 = False
    ErrorLog
End Function

'取得使用者可使用的系統名稱
Public Function GetUserUseSysKind(ByRef strGroup As String, ByRef StrSysName As String) As Boolean
Dim Rs As ADODB.Recordset
Dim obj01 As Object
        Set obj01 = CreateObject("prjtaiedll.clspublicdata")
        strSQL = "select sg02 from staff_group where sg01='" + strGroup + "'"
        Set Rs = obj01.ReadRst(strSQL, True)
        If Rs.EOF Then
            Rs.Close
            GetUserUseSysKind = False
        Else
            StrSysName = Rs(0)
            GetUserUseSysKind = True
            Rs.Close
        End If
        Set obj01 = Nothing
End Function

'檢查新增之本所案號是否大於現在AutoNum
Public Function ChkCaseNoWithAuNum0702(StrNo1 As String, StrNo2 As String, StrNo3 As String, StrNo4 As String, strSysKind As String) As Boolean
 Dim Rs As ADODB.Recordset, Nret As Boolean, StrName As String, obj01 As Object
   strSQL = "select * from servicepractice where sp01=" + CNULL(StrNo1) + " and sp02=" + CNULL(StrNo2) + " and sp03=" + CNULL(StrNo3) + " and sp04=" + CNULL(StrNo4)
   Set obj01 = CreateObject("prjtaiedll.clspublicdata")
   Set Rs = obj01.ReadRst(strSQL)
   If Not Rs.EOF Then
      MsgBox "本所案號重覆 !", vbCritical
      ChkCaseNoWithAuNum0702 = False
      Rs.Close
      Exit Function
   End If
   Rs.Close
   StrName = ""
   strSQL = "select * from autonumber where au01=" + CNULL(strSysKind) + " and au03 >" + CStr(Val(StrNo2))
   Set Rs = obj01.ReadRst(strSQL)
   Set obj01 = Nothing
   If Rs.EOF Then
       Rs.Close
       MsgBox "本所案號不合規則 !", vbCritical
       ChkCaseNoWithAuNum0702 = False
   Else
       ChkCaseNoWithAuNum0702 = True  '可以新增資料
       Rs.Close
   End If
End Function

'檢查新增之本所案號是否大於現在AutoNum
Public Function ChkCaseNoWithAuNum0701(StrNo1 As String, StrNo2 As String, StrNo3 As String, StrNo4 As String, strSysKind As String) As Boolean
Dim Rs As ADODB.Recordset
Dim Nret As Boolean
Dim StrName As String
Dim obj01 As Object
        strSQL = "select * from patent where pa01=" + CNULL(StrNo1) + " and pa02=" + CNULL(StrNo2) + " and pa03=" + CNULL(StrNo3) + " and pa04=" + CNULL(StrNo4)
        Set obj01 = CreateObject("prjtaiedll.clspublicdata")
        Set Rs = obj01.ReadRst(strSQL)
        If Not Rs.EOF Then
             MsgBox "本所案號重覆"
             ChkCaseNoWithAuNum0701 = False
             Rs.Close
             Exit Function
        End If
        Rs.Close
        Select Case strSysKind
        Case "CFP"
            strSQL = "select  *  from patent where pa01='CFP' and pa02 >" + Format(StrNo2, "#####0")
            Debug.Print strSQL
        Case "FCP"
            strSQL = "select  *  from patent where pa01='FCP' and pa02 >" + Format(StrNo2, "#####0")
            Debug.Print strSQL
        Case "P"
            strSQL = "select  *  from patent where pa01='P' and pa02 >" + Format(StrNo2, "#####0")
            Debug.Print strSQL
        End Select
        Set Rs = obj01.ReadRst(strSQL)
        Set obj01 = Nothing
        If Rs.EOF Then
            Rs.Close
            ChkCaseNoWithAuNum0701 = False  '不可以新增資料
            MsgBox "本所案號不合規則"
            Exit Function
        Else
            ChkCaseNoWithAuNum0701 = True  '可以新增資料
            Rs.Close
        End If
End Function
'取得代理人備註
Public Function GetCustomerNote0702(StrNo As String, ByRef StrName As String) As Boolean
Dim Rs As ADODB.Recordset
Dim obj01 As Object
    Set obj01 = CreateObject("prjtaiedll.clspublicdata")
    strSQL = "select  cu79  from customer where cu01=" + CNULL(StrNo)
    Debug.Print strSQL
    Set Rs = obj01.ReadRst(strSQL)
    If Rs.EOF Then
        GetCustomerNote0702 = False
        Rs.Close
        Exit Function
    Else
        GetCustomerNote0702 = True
        StrName = IIf(IsNull(Rs(0)), "", Rs(0))
    End If
    Rs.Close
    Set obj01 = Nothing
End Function

'取得代理人備註
Public Function GetAgentNote0702(StrNo As String, ByRef StrName As String) As Boolean
Dim Rs As ADODB.Recordset
Dim obj01 As Object
    Set obj01 = CreateObject("prjtaiedll.clspublicdata")
    strSQL = "select  fa29  from fagent where fa01=" + CNULL(Mid$(StrNo, 1, 8)) + "and fa02=" + CNULL(Mid$(StrNo, 9, 1))
    Debug.Print strSQL
    Set Rs = obj01.ReadRst(strSQL)
    If Rs.EOF Then
        GetAgentNote0702 = False
        Rs.Close
        Exit Function
    Else
        GetAgentNote0702 = True
        StrName = IIf(IsNull(Rs(0)), "", Rs(0))
    End If
    Rs.Close
    Set obj01 = Nothing
End Function

'取得PriDate資料
Public Function GetPriDateData0701(pd() As String) As String()
Dim PD1() As String
Dim Rs As ADODB.Recordset
Dim i As Integer
Dim objPublicData As Object
    strSQL = "select * from Pridate where pd01=" + CNULL(pd(0)) + " and pd02=" + CNULL(pd(1)) + " and pd03=" + CNULL(pd(2)) + " and pd04=" + CNULL(pd(3))
    Set objPublicData = CreateObject("Prjtaiedll.clspublicdata")
    Set Rs = objPublicData.ReadRst(strSQL)
    If Rs.EOF Then
        GetPriDateData0701 = PD1()
        Rs.Close
    Else
        Rs.MoveLast
        Rs.MoveFirst
        ReDim PD1(Rs.RecordCount, 3)
        i = 0
        While Not Rs.EOF
            PD1(i, 0) = Rs(0)
            PD1(i, 1) = Rs(1)
            PD1(i, 2) = Rs(2)
            Rs.MoveNext
            i = i + 1
        Wend
        GetPriDateData0701 = PD1()
    End If
    Set objPublicData = Nothing
End Function
'取得年費通知人
Public Function GetAnnualAgent(ByRef StrNo As String, ByRef StrName As String, ByRef strTemp As String) As Boolean
Dim obj01 As Object
Dim obj02 As Object
Dim Rs As ADODB.Recordset
    Set obj01 = CreateObject("prjtaiedll.clspublicdata")
    Select Case Mid$(StrNo, 1, 1)
    Case "X"
        strSQL = "select cu03 from customer where cu01=" + CNULL(StrNo)
        Set Rs = obj01.ReadRst(strSQL)
        If Rs.EOF Then
            Rs.Close
            Set obj01 = Nothing
            GetAnnualAgent = False
            Exit Function
        Else
            strTemp = IIf(IsNull(Rs(0)), "", Rs(0))
            Rs.Close
        End If
    Case "Y"
        strTemp = StrNo
    End Select
    strSQL = "select  nvl(fa06,nvl(fa04,fa05))  from fagent where fa01=" + CNULL(Mid$(strTemp, 1, 8)) + " and fa02=" + CNULL(Mid$(strTemp, 9, 1))
    Set Rs = obj01.ReadRst(strSQL)
    If Rs.EOF Then
        Rs.Close
        Set obj01 = Nothing
        GetAnnualAgent = False
        Exit Function
    Else
        StrName = IIf(IsNull(Rs(0)), "", Rs(0))
        Rs.Close
        Set obj01 = Nothing
        GetAnnualAgent = True
    End If
End Function

'取得副本收受人,固定請款對象
Public Function GetCopyReceiver(ByRef StrNo As String, ByRef StrName As String) As Boolean
Dim obj01 As Object
Dim Rs As ADODB.Recordset
    Set obj01 = CreateObject("prjtaiedll.clspublicdata")
    Select Case Mid$(StrNo, 1, 1)
    Case "X"
        strSQL = "select nvl(cu04,nvl(cu05,cu06))  from customer where cu01=" + CNULL(Mid$(StrNo, 1, 8)) + "  and cu02='0'"
        Set Rs = obj01.ReadRst(strSQL)
    Case "Y"
        strSQL = "select nvl(fa05,nvl(fa04,fa06))  from customer where fa01=" + CNULL(Mid$(StrNo, 1, 8)) + "  and fa02='0'"
        Set Rs = obj01.ReadRst(strSQL)
    End Select
    If Rs.EOF Then
        Rs.Close
        GetCopyReceiver = False
        Exit Function
     Else
        StrName = IIf(IsNull(Rs(0)), "", Rs(0))
        GetCopyReceiver = True
        Rs.Close
    End If
    Set obj01 = Nothing
End Function

'取得下次繳費日,利用array傳值,array(0-3)為本所案號,
Public Function GetNextPayDate(pa() As String, ByRef StrName As String) As Boolean
Dim obj01 As Object
Dim Rs As ADODB.Recordset
Dim strTemp As String
 Dim i As Integer
    Set obj01 = CreateObject("prjtaiedll.clspublicdata")
    strTemp = "select np09 from nextprogress where np02=" + CNULL(pa(0)) + " and np03=" + CNULL(pa(1)) + "and np04=" + CNULL(pa(2)) + " and np05=" + CNULL(pa(3))
    strSQL = strTemp + " and np07='21' "
    Debug.Print strSQL
    Set Rs = obj01.ReadRst(strSQL)
    If Rs.EOF Then
        Rs.Close
        strSQL = strTemp + " and np07='" & 年費 & "'"
    Else
        StrName = IIf(IsNull(Rs(0)), "", Rs(0))
        GetNextPayDate = True
        Exit Function
    End If
    
    Debug.Print strSQL
    Set Rs = obj01.ReadRst(strSQL)
    
    If Rs.EOF Then
        Rs.Close
        strSQL = strTemp + " and np07='49'"
    Else
        StrName = IIf(IsNull(Rs(0)), "", Rs(0))
        GetNextPayDate = True
        Exit Function
    End If
    
    Debug.Print strSQL
    Set Rs = obj01.ReadRst(strSQL)
    
    If Rs.EOF Then
        Rs.Close
        GetNextPayDate = False
        StrName = ""
        Exit Function
    Else
        StrName = IIf(IsNull(Rs(0)), "", Rs(0))
        GetNextPayDate = True
        Exit Function
    End If
    Set obj01 = Nothing
End Function

'取得收款後辦案,代理人備註
Public Function GetFagentData0701(StrNo As String, ByRef StrName As String, ByRef strTemp As String) As Boolean
Dim obj01 As Object
Dim Rs As ADODB.Recordset
    Set obj01 = CreateObject("prjtaiedll.clspublicdata")
    strSQL = "select fa39,fa29 from fagent where fa01=" + CNULL(Mid$(StrNo, 1, 8)) + " and fa02=" + CNULL(Mid$(StrNo, 9, 1))
    Debug.Print strSQL
    Set Rs = obj01.ReadRst(strSQL)
    If Rs.EOF Then
        Rs.Close
        GetFagentData0701 = False
        Set obj01 = Nothing
        Exit Function
    Else
        StrName = IIf(IsNull(Rs(0)), "", Rs(0))
        strTemp = IIf(IsNull(Rs(1)), "", Rs(1))
        Rs.Close
        GetFagentData0701 = True
    End If
    Set obj01 = Nothing
End Function

'取得收款後辦案,客戶備註
Public Function GetCustomer0701(StrNo As String, ByRef StrName As String, ByRef strTemp As String) As Boolean
Dim obj01 As Object
Dim Rs As ADODB.Recordset
    Set obj01 = CreateObject("prjtaiedll.clspublicdata")
    StrNo = ChangeCustomerL(StrNo)
    strSQL = "select cu72,cu79 from customer where " + ChgCustomer(StrNo)
    Set Rs = obj01.ReadRst(strSQL)
    If Rs.EOF Then
        Rs.Close
        GetCustomer0701 = False
        Set obj01 = Nothing
        Exit Function
    Else
        StrName = IIf(IsNull(Rs(0)), "", Rs(0))
        strTemp = IIf(IsNull(Rs(1)), "", Rs(1))
        Rs.Close
        GetCustomer0701 = True
    End If
    StrNo = ChangeCustomerS(StrNo)
    Set obj01 = Nothing
End Function
Public Function GetInventor0701(strInventorNo As String, Optional strInID As String, Optional strInName1 As String, Optional strInName2 As String, Optional strInName3 As String) As Boolean
Dim obj01 As Object
Dim Rs As ADODB.Recordset
    strSQL = "select in03,in04,in05,in06 from inventor where in01||in02=" + CNULL(strInventorNo)
    Set obj01 = CreateObject("prjtaiedll.clsPublicData")
    Debug.Print strSQL
    Set Rs = obj01.ReadRst(strSQL)
    If Rs.EOF Then
        Rs.Close
        Set obj01 = Nothing
        GetInventor0701 = False
        Exit Function
    Else
        strInID = IIf(IsNull(Rs(0)), "", Rs(0))
        strInName1 = IIf(IsNull(Rs(1)), "", Rs(1))
        strInName2 = IIf(IsNull(Rs(2)), "", Rs(2))
        strInName3 = IIf(IsNull(Rs(3)), "", Rs(3))
        Rs.Close
        Set obj01 = Nothing
        GetInventor0701 = True
    End If
End Function
