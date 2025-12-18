VERSION 5.00
Begin VB.Form frm04010704 
   BorderStyle     =   1  '單線固定
   Caption         =   "重新委任批次收文(未回覆客戶)"
   ClientHeight    =   1620
   ClientLeft      =   255
   ClientTop       =   990
   ClientWidth     =   5565
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   5565
   Begin VB.CommandButton cmdOK 
      Caption         =   "補收追加聯合案(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   2
      Left            =   3465
      TabIndex        =   9
      Top             =   1140
      Width           =   1965
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "申請書補印(&A)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   5
      Left            =   2250
      TabIndex        =   8
      Top             =   90
      Width           =   1335
   End
   Begin VB.TextBox txtCaseQty 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Left            =   1260
      MaxLength       =   7
      TabIndex        =   0
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   0
      Left            =   4590
      TabIndex        =   3
      Top             =   90
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   400
      Index           =   1
      Left            =   3780
      TabIndex        =   2
      Top             =   90
      Width           =   800
   End
   Begin VB.TextBox txtAppDate 
      Height          =   270
      Left            =   1260
      MaxLength       =   7
      TabIndex        =   1
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label lblActCaseQty 
      Alignment       =   1  '靠右對齊
      Caption         =   "0"
      Height          =   180
      Left            =   1260
      TabIndex        =   7
      Top             =   555
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "預定案件數:                       ( 不可大於 500 )"
      Height          =   180
      Left            =   225
      TabIndex        =   6
      Top             =   870
      Width           =   3225
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "實際案件數:"
      Height          =   180
      Left            =   225
      TabIndex        =   5
      Top             =   555
      Width           =   945
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "申請書日期:"
      Height          =   180
      Left            =   225
      TabIndex        =   4
      Top             =   1230
      Width           =   945
   End
End
Attribute VB_Name = "frm04010704"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan2010/8/11 日期欄已修改
'Create by Morgan 2007/7/3
Option Explicit

Dim CP09() As String

Private Sub cmdok_Click(Index As Integer)
   Dim strAppDate As String
   Screen.MousePointer = vbHourglass
   Select Case Index
      Case 0 '結束
         Unload Me
         
      Case 1 '確定
         If TxtValidate = True Then
            Process
         End If
      'Add by Morgan 2007/12/13
      Case 2 '補收追加聯合案
         If FormSave1 = True Then
            AppPrintX "9612  "
         End If
         
      'Add by Morgan 2007/7/4 申請書補印
      Case 5
         strAppDate = InputBox("申請書日期：", "請輸入申請書日期", strSrvDate(2), Me.Left, Me.Top + Me.Height + 1000)
         If strAppDate <> "" Then
            If ChkDate(strAppDate) = True Then
               AppPrintX strAppDate
            End If
         End If
   End Select
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   Me.Top = Me.Top - 2000
   txtAppDate = strSrvDate(2)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm04010704 = Nothing
End Sub

Private Sub Process()

   Dim stCustNo1 As String '起始客戶編號
   Dim stCustNo2 As String '終止客戶編號
   Dim iRecs As Integer, iTRec As Integer
   
   '先排除個案發文通知的客戶，若確定要時在將lr11設'A'就可以
   'Modify by Morgan 2007/7/18 改沒印過的都要
   'strExc(0) = "select lr01 from linreasignrec where lr11<'B' and lr08 is null and lr12 is null order by 1"
   strExc(0) = "select lr01 from linreasignrec where lr08 is null and lr12 is null order by 1"
   intI = 1
   Set adoRecordset = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With adoRecordset
         stCustNo1 = .Fields(0)
         stCustNo2 = .Fields(0)
         Do While Not .EOF
            iRecs = Get928Case(stCustNo1, .Fields(0))
            If iRecs < Val(txtCaseQty) Then
               stCustNo2 = .Fields(0)
               iTRec = iRecs
            Else
               If iTRec = 0 Or iRecs = Val(txtCaseQty) Then
                  stCustNo2 = .Fields(0)
                  iTRec = iRecs
               End If
               Exit Do
            End If
            .MoveNext
         Loop
      End With
   End If
   
   If MsgBox("本次將收文 " & iTRec & " 筆，是否確定要繼續？", vbYesNo + vbDefaultButton1) = vbNo Then
      txtCaseQty.SetFocus
      txtCaseQty_GotFocus
      Exit Sub
   End If
   
   If FormSave(stCustNo1, stCustNo2) = True Then
      AppPrint
      txtCaseQty = ""
      txtCaseQty.SetFocus
   End If
   
End Sub

Private Sub AppPrint()

   Dim ii As Integer
   
   For ii = 1 To UBound(CP09)
      EndLetter "01", CP09(ii), "01", strUserNum
      strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
             "VALUES ('01','" & CP09(ii) & "','01','" & strUserNum & _
             "','勾選1','■')"
      cnnConnection.Execute strSql, intI
      
      strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
             "VALUES ('01','" & CP09(ii) & "','01','" & strUserNum & _
             "','其他日期','" & txtAppDate & "')"
      cnnConnection.Execute strSql, intI
      
      NowPrint CP09(ii), "01", "01", False, strUserNum
   Next
   PUB_BatchPrint "5"

End Sub

Private Function FormSave(p_CustNo1 As String, p_CustNo2 As String) As Boolean
   
   Dim cp(1 To 110) As String
   Dim iCol As Integer
   Dim sCols As String, sValues As String
   Dim iCnt As Integer, iRecs As Integer
   
   cnnConnection.BeginTrans
   
On Error GoTo ErrHnd

   '更新列印日期為0
   strSql = "update linreasignrec set lr12=0 where lr01>='" & p_CustNo1 & "' and lr01<='" & p_CustNo2 & "' and lr08 is null and  lr12 is null"
   
   cnnConnection.Execute strSql, intI
   
   iRecs = Get928Case(p_CustNo1, p_CustNo2, adoRecordset, True)
   If iRecs > 0 Then
      With adoRecordset
         Do While Not .EOF
            iCnt = iCnt + 1
            cp(1) = .Fields("lc01")
            cp(2) = .Fields("lc02")
            cp(3) = .Fields("lc03")
            cp(4) = .Fields("lc04")
            cp(5) = strSrvDate(1)
            cp(9) = AutoNo("B", 6)
            cp(10) = "928"
            cp(13) = PUB_GetAKindSalesNo(cp(1), cp(2), cp(3), cp(4))
            cp(12) = GetSalesArea(cp(13))
            cp(14) = strUserNum
            cp(20) = "N"
            cp(26) = "N"
            cp(27) = "NULL"
            cp(28) = Format(iCnt, "000000")
            cp(84) = 0
            cp(110) = "76012,81040"
            strSql = "insert into caseprogress(CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP27,CP28,CP84,CP110)" & _
               " VALUES(" & CNULL(cp(1)) & "," & CNULL(cp(2)) & "," & CNULL(cp(3)) & "," & CNULL(cp(4)) & "," & cp(5) & _
               "," & CNULL(cp(9)) & "," & CNULL(cp(10)) & "," & CNULL(cp(12)) & "," & CNULL(cp(13)) & "," & CNULL(cp(14)) & "," & CNULL(cp(20)) & "" & _
               "," & CNULL(cp(26)) & "," & cp(27) & "," & CNULL(cp(28)) & "," & cp(84) & "," & CNULL(cp(110)) & "" & _
               ")"
            cnnConnection.Execute strSql, intI
            
            ReDim Preserve CP09(iCnt)
            CP09(iCnt) = cp(9)
            .MoveNext
         Loop
      End With
   End If
   
   '更新列印日期
   strSql = "update linreasignrec set lr12=" & strSrvDate(1) & " where lr01>='" & p_CustNo1 & "' and lr01<='" & p_CustNo2 & "' and lr08 is null and lr12=0"
   cnnConnection.Execute strSql, intI
   cnnConnection.CommitTrans
   lblActCaseQty = iCnt
   FormSave = True
   Exit Function
   
ErrHnd:
   If Err.NUMBER <> 0 Then
      cnnConnection.RollbackTrans
      MsgBox Err.Description, vbCritical
   End If
   
End Function
'讀取待收文案件
Private Function Get928Case(p_CustNo1 As String, Optional p_CustNo2 As String, Optional p_Rst As ADODB.Recordset, Optional p_Add As Boolean = False) As Integer
   Dim stCon As String
   
   stCon = ""
   If p_CustNo2 = "" Then
      stCon = " and lr01='" & p_CustNo1 & "'"
   Else
      stCon = " and lr01>='" & p_CustNo1 & "' and lr01<='" & p_CustNo2 & "'"
   End If
   
   If p_Add = True Then
      stCon = stCon & " and lr12=0"
   Else
      stCon = stCon & " and lr12 is null"
   End If
   
   '郭--96/6/22
   '1.發明案76年以後,新型設計84年以後
   '2.案件已專利權消滅,或閉卷者不處理重新委任
   '3.案件進度檔中有其他來函,備註欄有"變更代理人"(模糊比對)者,不處理重新委任
   '4.案件程序中僅有:專利調查,調卷,鑑定報告,翻譯之程序,則刪除該案,即若還有其他程序者,則該案要保留.--96/6/28
   '玲玲 --96/7/5
   '5.有收文1606專利權己公告作廢,1907 通知退證註銷,1911 通知暫不續行審查,發文429放棄專利權 不處理重新委任
   strExc(0) = "select X.* from (" & _
      " select lc01,lc02,lc03,lc04 from linreasignrec,lincase" & _
      " where lr08 is null and lc07(+)=lr01 and lc06 in ('1','2')" & stCon & _
      " Union select lc01,lc02,lc03,lc04 from linreasignrec,lincase" & _
      " where lr08 is null and lc08(+)=lr01 and lc06 in ('1','2')" & stCon & _
      " Union select lc01,lc02,lc03,lc04 from linreasignrec,lincase" & _
      " where lr08 is null and lc09(+)=lr01 and lc06 in ('1','2')" & stCon & _
      " Union select lc01,lc02,lc03,lc04 from linreasignrec,lincase" & _
      " where lr08 is null and lc10(+)=lr01 and lc06 in ('1','2')" & stCon & _
      " Union select lc01,lc02,lc03,lc04 from linreasignrec,lincase" & _
      " where lr08 is null and lc11(+)=lr01 and lc06 in ('1','2')" & stCon & _
      ") X,patent where lc01='P' and pa01(+)=lc01 and pa02(+)=lc02 and pa03(+)=lc03 and pa04(+)=lc04" & _
      " and pa01='P' and (pa17='Y' or pa17 is null)" & _
      " and (pa57 is null or (pa57='Y' and pa25>=" & strSrvDate(1) & "))" & _
      " and not exists(select * from caseprogress where cp01=lc01 and cp02=lc02 and cp03=lc03 and cp04=lc04" & _
      " and ( cp10 in ('1606','1907','1911','1604','928','429') OR instr(cp64,'專利權消滅')>0 ))" & _
      " and not exists(select * from caseprogress where cp01=lc01 and cp02=lc02 and cp03=lc03 and cp04=lc04" & _
      " and cp10='1902' AND instr(cp64,'變更代理人')>0)" & _
      " and not exists(select * from caseprogress where cp01=lc01 and cp02=lc02" & _
      " and cp03=lc03 and cp04=lc04 and cp09<'C' and cp27>0 and cp110<>'65002')" & _
      " and exists(select * from caseprogress where cp01=lc01 and cp02=lc02" & _
      " and cp03=lc03 and cp04=lc04 and cp10 NOT IN ('201','927','903','904','906'))" & _
      " and length(pa11)<9"
      
   '追加聯合案
   strExc(0) = strExc(0) & " union select X.* from (" & _
      " select lc01,lc02,lc03,lc04 from linreasignrec,lincase" & _
      " where lr08 is null and lc07(+)=lr01 and lc06 in ('1','2')" & stCon & _
      " Union select lc01,lc02,lc03,lc04 from linreasignrec,lincase" & _
      " where lr08 is null and lc08(+)=lr01 and lc06 in ('1','2')" & stCon & _
      " Union select lc01,lc02,lc03,lc04 from linreasignrec,lincase" & _
      " where lr08 is null and lc09(+)=lr01 and lc06 in ('1','2')" & stCon & _
      " Union select lc01,lc02,lc03,lc04 from linreasignrec,lincase" & _
      " where lr08 is null and lc10(+)=lr01 and lc06 in ('1','2')" & stCon & _
      " Union select lc01,lc02,lc03,lc04 from linreasignrec,lincase" & _
      " where lr08 is null and lc11(+)=lr01 and lc06 in ('1','2')" & stCon & _
      ") X,patent p1 where lc01='P' and pa01(+)=lc01 and pa02(+)=lc02 and pa03(+)=lc03 and pa04(+)=lc04" & _
      " and not exists(select * from caseprogress where cp01=lc01 and cp02=lc02 and cp03=lc03 and cp04=lc04" & _
      " and ( cp10 in ('1606','1907','1911','1604','928','429') OR instr(cp64,'專利權消滅')>0 ))" & _
      " and not exists(select * from caseprogress where cp01=lc01 and cp02=lc02 and cp03=lc03 and cp04=lc04" & _
      " and cp10='1902' AND instr(cp64,'變更代理人')>0)" & _
      " and not exists(select * from caseprogress where cp01=lc01 and cp02=lc02" & _
      " and cp03=lc03 and cp04=lc04 and cp09<'C' and cp27>0 and cp110<>'65002')" & _
      " and exists(select * from caseprogress where cp01=lc01 and cp02=lc02" & _
      " and cp03=lc03 and cp04=lc04 and cp10 NOT IN ('201','927','903','904','906'))" & _
      " and length(pa11)>=9"
      
   'Modify by Morgan 2007/12/11
   'strExc(0) = strExc(0) & _
      " and pa01='P' and exists(select * from patent p2 where p2.pa11=p1.pa11 and p2.pa01=p1.pa02" & _
      " and (p2.pa17='Y' or p2.pa17 is null) and (p2.pa57 is null or (p2.pa57='Y' and p2.pa25>=" & strSrvDate(1) & "))" & _
      " and not exists(select * from caseprogress where cp01=p2.pa01 and cp02=p2.pa02 and cp03=p2.pa03 and cp04=p2.pa04" & _
      " and ( cp10='1604' OR instr(cp64,'專利權消滅')>0 )))"
      
   strExc(0) = strExc(0) & _
      " and pa01='P' and exists(select * from patent p2 where substr(p2.pa11,1,8)=substr(p1.pa11,1,8) and p2.pa02=p1.pa02" & _
      " and p2.pa03='0' and (p2.pa17='Y' or p2.pa17 is null) and (p2.pa57 is null or (p2.pa57='Y' and p2.pa25>=" & strSrvDate(1) & "))" & _
      " and not exists(select * from caseprogress where cp01=p2.pa01 and cp02=p2.pa02 and cp03=p2.pa03 and cp04=p2.pa04" & _
      " and ( cp10='1604' OR instr(cp64,'專利權消滅')>0 )))"
   'end 2007/12/11
      
   strExc(0) = strExc(0) & " order by 1,2,3,4"
   
   intI = 1
   Set p_Rst = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      Get928Case = p_Rst.RecordCount
   Else
      Get928Case = 0
   End If
End Function


Private Sub txtAppDate_GotFocus()
   TextInverse txtAppDate
End Sub

Private Sub txtAppDate_Validate(Cancel As Boolean)
   If txtAppDate <> "" Then
      If ChkDate(txtAppDate) = False Then
         Cancel = True
      ElseIf Val(txtAppDate) < Val(strSrvDate(2)) Then
         Cancel = True
         MsgBox "申請書日期不可小於系統日，請重新輸入 !", vbCritical
      End If
   End If
End Sub


Private Function TxtValidate() As Boolean

   Dim bolResult As Boolean
      
   If Val(txtCaseQty) = 0 Or Val(txtCaseQty) > 500 Then
      MsgBox "預定案件數輸入錯誤!", vbExclamation
      txtCaseQty.SetFocus
      txtCaseQty_GotFocus
      Exit Function
   End If
   
   If txtAppDate = "" Then
      MsgBox "請輸入申請書日期!", vbExclamation
      txtAppDate.SetFocus
      Exit Function
   End If
      
   txtAppDate_Validate bolResult
   If bolResult = True Then
      txtAppDate.SetFocus
      txtAppDate_GotFocus
      Exit Function
   End If
   TxtValidate = True
End Function

Private Sub txtCaseQty_GotFocus()
   TextInverse txtCaseQty
End Sub


Private Sub AppPrintX(p_AppDate As String)
     
   Dim rsPrint As New ADODB.Recordset
   
   strExc(0) = "select cp09 from caseprogress" & _
      " where cp01='P' and cp05=" & strSrvDate(1) & " and cp10='928' and cp09>'B' and cp27 IS NULL " & _
      " AND NOT EXISTS(SELECT * FROM LETTERDEMAND WHERE LD01=CP65 AND LD04=CP09 AND LD09=CP10" & _
      " AND LD10='01' AND LD11='01' AND LD12='5') order by cp01,cp02,cp03,cp04"
      
   intI = 1
   Set rsPrint = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If MsgBox("共有 " & rsPrint.RecordCount & "筆申請書需補印，是否確定要繼續？", vbYesNo + vbDefaultButton1) = vbYes Then
         With rsPrint
         Do While Not .EOF
            
            EndLetter "01", .Fields(0), "01", strUserNum
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   "VALUES ('01','" & .Fields(0) & "','01','" & strUserNum & _
                   "','勾選1','■')"
            cnnConnection.Execute strSql, intI
            
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   "VALUES ('01','" & .Fields(0) & "','01','" & strUserNum & _
                   "','其他日期','" & p_AppDate & "')"
            cnnConnection.Execute strSql, intI
            
            NowPrint .Fields(0), "01", "01", False, strUserNum
         
            .MoveNext
         Loop
         End With
         PUB_BatchPrint "5"
      End If
   End If
   Set rsPrint = Nothing
End Sub

'讀取待收文案件
Private Function Get928CaseX(p_Rst As ADODB.Recordset) As Integer
   Dim stCon As String
   
   stCon = stCon & " and lr12 is not null"
         
   '追加聯合案
   strExc(0) = "select X.* from (" & _
      " select lc01,lc02,lc03,lc04 from linreasignrec,lincase" & _
      " where lr08 is null and lc07(+)=lr01 and lc06 in ('1','2')" & stCon & _
      " Union select lc01,lc02,lc03,lc04 from linreasignrec,lincase" & _
      " where lr08 is null and lc08(+)=lr01 and lc06 in ('1','2')" & stCon & _
      " Union select lc01,lc02,lc03,lc04 from linreasignrec,lincase" & _
      " where lr08 is null and lc09(+)=lr01 and lc06 in ('1','2')" & stCon & _
      " Union select lc01,lc02,lc03,lc04 from linreasignrec,lincase" & _
      " where lr08 is null and lc10(+)=lr01 and lc06 in ('1','2')" & stCon & _
      " Union select lc01,lc02,lc03,lc04 from linreasignrec,lincase" & _
      " where lr08 is null and lc11(+)=lr01 and lc06 in ('1','2')" & stCon & _
      ") X,patent p1 where lc01='P' and pa01(+)=lc01 and pa02(+)=lc02 and pa03(+)=lc03 and pa04(+)=lc04" & _
      " and not exists(select * from caseprogress where cp01=lc01 and cp02=lc02 and cp03=lc03 and cp04=lc04" & _
      " and ( cp10 in ('1606','1907','1911','1604','928','429') OR instr(cp64,'專利權消滅')>0 ))" & _
      " and not exists(select * from caseprogress where cp01=lc01 and cp02=lc02 and cp03=lc03 and cp04=lc04" & _
      " and cp10='1902' AND instr(cp64,'變更代理人')>0)" & _
      " and not exists(select * from caseprogress where cp01=lc01 and cp02=lc02" & _
      " and cp03=lc03 and cp04=lc04 and cp09<'C' and cp27>0 and cp110<>'65002')" & _
      " and exists(select * from caseprogress where cp01=lc01 and cp02=lc02" & _
      " and cp03=lc03 and cp04=lc04 and cp10 NOT IN ('201','927','903','904','906'))" & _
      " and length(pa11)>=9"
  
   '有母案還存活或沒母案的
   '母案:申請號前8碼相同
   strExc(0) = strExc(0) & _
      " and pa01='P' and( exists(select * from patent p2 where p2.pa11=substr(p1.pa11,1,8)" & _
      " and (p2.pa17='Y' or p2.pa17 is null) and (p2.pa57 is null or (p2.pa57='Y' and p2.pa25>=" & strSrvDate(1) & "))" & _
      " and not exists(select * from caseprogress where cp01=p2.pa01 and cp02=p2.pa02 and cp03=p2.pa03 and cp04=p2.pa04" & _
      " and ( cp10='1604' OR instr(cp64,'專利權消滅')>0 )))" & _
      " or not exists(select * from patent p2 where p2.pa11=substr(p1.pa11,1,8)))"
   
   strExc(0) = strExc(0) & " order by 1,2,3,4"
   
   intI = 1
   Set p_Rst = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      Get928CaseX = p_Rst.RecordCount
   Else
      Get928CaseX = 0
   End If
End Function

Private Function FormSave1() As Boolean
   
   Dim cp(1 To 110) As String
   Dim iCol As Integer
   Dim sCols As String, sValues As String
   Dim iCnt As Integer, iRecs As Integer
   
   cnnConnection.BeginTrans
   
On Error GoTo ErrHnd
  
   iRecs = Get928CaseX(adoRecordset)
   If iRecs > 0 Then
      With adoRecordset
         Do While Not .EOF
            iCnt = iCnt + 1
            cp(1) = .Fields("lc01")
            cp(2) = .Fields("lc02")
            cp(3) = .Fields("lc03")
            cp(4) = .Fields("lc04")
            cp(5) = strSrvDate(1)
            cp(9) = AutoNo("B", 6)
            cp(10) = "928"
            cp(13) = PUB_GetAKindSalesNo(cp(1), cp(2), cp(3), cp(4))
            cp(12) = GetSalesArea(cp(13))
            cp(14) = strUserNum
            cp(20) = "N"
            cp(26) = "N"
            cp(27) = "NULL"
            cp(28) = Format(iCnt, "000000")
            cp(84) = 0
            cp(110) = "76012,81040"
            strSql = "insert into caseprogress(CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP27,CP28,CP84,CP110)" & _
               " VALUES(" & CNULL(cp(1)) & "," & CNULL(cp(2)) & "," & CNULL(cp(3)) & "," & CNULL(cp(4)) & "," & cp(5) & _
               "," & CNULL(cp(9)) & "," & CNULL(cp(10)) & "," & CNULL(cp(12)) & "," & CNULL(cp(13)) & "," & CNULL(cp(14)) & "," & CNULL(cp(20)) & "" & _
               "," & CNULL(cp(26)) & "," & cp(27) & "," & CNULL(cp(28)) & "," & cp(84) & "," & CNULL(cp(110)) & "" & _
               ")"
            cnnConnection.Execute strSql, intI
            .MoveNext
         Loop
      End With
   End If
   
   cnnConnection.CommitTrans
   lblActCaseQty = iCnt
   FormSave1 = True
   Exit Function
   
ErrHnd:
   If Err.NUMBER <> 0 Then
      cnnConnection.RollbackTrans
      MsgBox Err.Description, vbCritical
   End If
   
End Function

