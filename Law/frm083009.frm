VERSION 5.00
Begin VB.Form frm083009 
   BorderStyle     =   1  '單線固定
   Caption         =   "業績分析表"
   ClientHeight    =   2010
   ClientLeft      =   2715
   ClientTop       =   2790
   ClientWidth     =   4275
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   4275
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   0
      Left            =   1428
      MaxLength       =   3
      TabIndex        =   0
      Top             =   864
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   1
      Left            =   2388
      MaxLength       =   2
      TabIndex        =   1
      Top             =   864
      Width           =   375
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   400
      Left            =   2376
      TabIndex        =   2
      Top             =   120
      Width           =   800
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   3204
      TabIndex        =   3
      Top             =   120
      Width           =   800
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收文年月：                年               月"
      Height          =   180
      Index           =   1
      Left            =   468
      TabIndex        =   4
      Top             =   876
      Width           =   2832
   End
End
Attribute VB_Name = "frm083009"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/3 日期欄已修改
Option Explicit

Dim SDay As String, EDay As String, PLeft(0 To 17) As Integer
Dim PLeft1(0 To 17) As Integer, StS(1 To 8) As String
Dim m_print As Integer

Private Sub cmdBack_Click()
   Unload Me
End Sub

Private Sub cmdPrint_Click()
   m_print = 0
   If Text1(0) = "" Or Text1(1) = "" Then
      MsgBox "年度、月份必須有值 !", vbCritical
      Exit Sub
   Else
      If Val(CInt(Text1(0) & Format(Text1(1), "00"))) > Val(CInt((Year(Date) - 1911) & Format(Date, "MM"))) Then
         MsgBox "收文年月不得大於系統日 !", vbCritical
         Exit Sub
      End If
   End If
   Screen.MousePointer = 11
   GetPrintLeft
   PrintCase
   Screen.MousePointer = 0
   If m_print = 0 Then
      MsgBox "列印結束!", vbInformation
   End If
End Sub

Private Sub PrintCase()
 Dim i As Integer, j As Integer, Page As Integer, iPrint As Integer
 Dim TmpArea As String, Qty As String, StNum(1 To 2) As String
 Dim StDay(0 To 2) As String
 Dim StRng(0 To 25) As String, stVal(1 To 16) As Single
  Dim strDeptName(0 To 25) As String
 
 Dim strValTotal As Single '總計
 
 Dim bolChk As Boolean
 Dim iTotal As Integer
 Dim Wo As DAO.Workspace, Db As DAO.Database, Rc As DAO.Recordset, Rc1 As DAO.Recordset
'On Error GoTo err
   If CreateDatabase = False Then
      MsgBox "無法建立暫存區，列印失敗 !", vbInformation
      m_print = 1
      Exit Sub
   End If
   Set Wo = DBEngine.Workspaces(0)
   Set Db = Wo.OpenDatabase(App.Path & "\Case.mdb", False, False, ";PWD=taie")
   Qty = "DELETE FROM TEMP"
   Db.Execute Qty
   
   If Me.Tag = 0 Then
      StNum(1) = "'410102','411102','414102','418102'"
      StNum(2) = "'414101','418101','4181','4182','4183','4184','4141'"
      StS(1) = "顧問": StS(2) = "法務"
   Else
      StNum(1) = "'416101'"
      StNum(2) = "'416102'"
      StS(1) = "FCL": StS(2) = "CFL"
   End If
   StS(3) = "本月達成": StS(4) = "上月達成": StS(5) = "去年同期"
   
   '北所
   strExc(0) = "select a0901,a0902 from acc090 where a0901 like 'S1%' order by a0901"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))    'edit by nickc 2007/02/07 不用 dll 了 Set rstemp = objLawDll.ReadRstMsg(intI, strExc(0))
   iTotal = 1
   bolChk = True
   Do While Not RsTemp.EOF
      StRng(iTotal) = " AND ST03='" & RsTemp.Fields("a0901") & "'"
      If Not IsNull(RsTemp.Fields(1)) Then
         strDeptName(iTotal) = RsTemp.Fields("a0902")
      Else
         strDeptName(iTotal) = ""
      End If
      iTotal = iTotal + 1
      RsTemp.MoveNext
   Loop
   StRng(iTotal) = " AND ST03 like 'S1%'"
   strDeptName(iTotal) = "北所業務"
   iTotal = iTotal + 1
   
   '中所
   strExc(0) = "select a0901,a0902 from acc090 where a0901 like 'S2%' order by a0901"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))    'edit by nickc 2007/02/07 不用 dll 了 Set rstemp = objLawDll.ReadRstMsg(intI, strExc(0))
   bolChk = True
   Do While Not RsTemp.EOF
      StRng(iTotal) = " AND ST03='" & RsTemp.Fields("a0901") & "'"
      If Not IsNull(RsTemp.Fields(1)) Then
         strDeptName(iTotal) = RsTemp.Fields("a0902")
      Else
         strDeptName(iTotal) = ""
      End If
      iTotal = iTotal + 1
      RsTemp.MoveNext
   Loop
   StRng(iTotal) = " AND ST03 like 'S2%'"
   strDeptName(iTotal) = "中所業務"
   iTotal = iTotal + 1
   
   '分所
   strExc(0) = "select a0901,a0902 from acc090 where a0901 between 'S3' and 'S99' order by a0901"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))    'edit by nickc 2007/02/07 不用 dll 了 Set rstemp = objLawDll.ReadRstMsg(intI, strExc(0))
   bolChk = True
   Do While Not RsTemp.EOF
      StRng(iTotal) = " AND ST03='" & RsTemp.Fields("a0901") & "'"
      If Not IsNull(RsTemp.Fields(1)) Then
         strDeptName(iTotal) = RsTemp.Fields("a0902")
      Else
         strDeptName(iTotal) = ""
      End If
      iTotal = iTotal + 1
      RsTemp.MoveNext
   Loop
   StRng(iTotal) = " AND ST03 between 'S3' and 'S99'"
   strDeptName(iTotal) = "分所業務"
   iTotal = iTotal + 1
   
   '國外
   StRng(iTotal) = " AND SUBSTR(ST03,1,1)='F'"
   strDeptName(iTotal) = "國外"
   iTotal = iTotal + 1
   
   '其他
   StRng(iTotal) = " AND (SUBSTR(ST03,1,1)<>'F' AND SUBSTR(ST03,1,1)<>'S' AND SUBSTR(ST03,1,1)<>'L')"
   strDeptName(iTotal) = "其他"
   iTotal = iTotal + 1
   
   '律師
   StRng(iTotal) = " AND SUBSTR(ST03,1,1)='L'"
   strDeptName(iTotal) = "律師"
   
   For i = 1 To iTotal
   
      For j = 1 To 5
         Select Case j
            Case 1 '416101 顧問
               StRng(0) = "(A0205 BETWEEN " & Text1(0) & Format(Text1(1), "00") & "01 AND " & _
                  Text1(0) & Format(Text1(1), "00") & "31) AND AX205 IN (" & StNum(1) & ")" & StRng(i)
            Case 2 '416102 法務
               StRng(0) = "(A0205 BETWEEN " & Text1(0) & Format(Text1(1), "00") & "01 AND " & _
                  Text1(0) & Format(Text1(1), "00") & "31) AND AX205 IN (" & StNum(2) & ")" & StRng(i)
            Case 3 '416101,416102 本月達成
               StRng(0) = "(A0205 BETWEEN " & Text1(0) & Format(Text1(1), "00") & "01 AND " & _
                  Text1(0) & Format(Text1(1), "00") & "31) AND AX205 IN (" & StNum(1) & "," & StNum(2) & ")" & StRng(i)
            Case 4 '416101,416102 上月達成
               If Text1(1) = 1 Then
                  StDay(1) = Val(Text1(0)) - 1
                  StDay(2) = "12"
               Else
                  StDay(1) = Text1(0)
                  StDay(2) = Format(Text1(1) - 1, "00")
               End If
               StRng(0) = "(A0205 BETWEEN " & StDay(1) & StDay(2) & "01 AND " & _
                  StDay(1) & StDay(2) & "31) AND AX205 IN (" & StNum(1) & "," & StNum(2) & ")" & StRng(i)
            Case 5 '去年同期
               StDay(1) = Text1(0) - 1
               StDay(2) = Format(Text1(1), "00")
               StRng(0) = "(A0205 BETWEEN " & StDay(1) & StDay(2) & "01 AND " & _
                  StDay(1) & StDay(2) & "31) AND AX205 IN (" & StNum(1) & "," & StNum(2) & ")" & StRng(i)
         End Select
         
         stVal(j) = "0"
         
         strExc(0) = "SELECT (SUM(AX207)-SUM(AX206))/1000 FROM ACC021,ACC020,ACC090,STAFF " & _
            "WHERE " & StRng(0) & " AND A0201=AX201(+) and A0202=AX202(+) AND AX209=ST01(+) AND ST03=A0901(+)"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))    'edit by nickc 2007/02/07 不用 dll 了 Set rstemp = objLawDll.ReadRstMsg(intI, strExc(0))
         If Not IsNull(RsTemp.Fields(0)) = True Then
            stVal(j) = RsTemp.Fields(0)
         End If
         
         RsTemp.Close
      Next
      
      strExc(0) = "INSERT INTO TEMP (TMP01,TMP02,TMP03,TMP04,TMP05,TMP06) VALUES " & _
         "('" & strDeptName(i) & "','" & stVal(1) & "','" & stVal(2) & "','" & _
         stVal(3) & "','" & stVal(4) & "','" & stVal(5) & "')"
      Db.Execute strExc(0)
      
   Next
   
   '總計
   strExc(0) = "SELECT SUM(VAL(TMP02)),SUM(VAL(TMP03)),SUM(VAL(TMP04)),SUM(VAL(TMP05))," & _
      "SUM(VAL(TMP06)) FROM TEMP WHERE TMP01<>'北所業務' AND TMP01<>'中所業務' AND TMP01<>'分所業務'"
   Set Rc = Db.OpenRecordset(strExc(0))
   
   strExc(0) = "INSERT INTO TEMP (TMP01,TMP02,TMP03,TMP04,TMP05,TMP06) VALUES " & _
      "('總計','" & Rc.Fields(0) & "','" & Rc.Fields(1) & "','" & _
      Rc.Fields(2) & "','" & Rc.Fields(3) & "','" & Rc.Fields(4) & "')"
   Db.Execute strExc(0)
   
   Page = 1
   CaseTitle TmpArea, 1
   iPrint = 2800
   
   Qty = "SELECT TMP01,TMP02,TMP03,TMP04,TMP05,TMP06 FROM TEMP"
   Set Rc = Db.OpenRecordset(Qty)
   With Rc
      Do While Not .EOF
         Printer.CurrentX = PLeft(0):      Printer.CurrentY = iPrint
         Printer.Print .Fields(0)
         Printer.CurrentX = PLeft(2) + 700 - (Printer.TextWidth(Format(.Fields(1), "0.00"))) - 100
         Printer.CurrentY = iPrint
         Printer.Print Format(.Fields(1), "0.00")
         Printer.CurrentX = PLeft(4) + 700 - (Printer.TextWidth(Format(.Fields(2), "0.00"))) - 100
         Printer.CurrentY = iPrint
         Printer.Print Format(.Fields(2), "0.00")
         Printer.CurrentX = PLeft(6) + 700 - (Printer.TextWidth(Format(.Fields(3), "0.00"))) - 100
         Printer.CurrentY = iPrint
         Printer.Print Format(.Fields(3), "0.00")
         Printer.CurrentX = PLeft(8) + 700 - (Printer.TextWidth(Format(.Fields(4), "0.00"))) - 100
         Printer.CurrentY = iPrint
         Printer.Print Format(.Fields(4), "0.00")
         Printer.CurrentX = PLeft(10) + 700 - (Printer.TextWidth(Format(.Fields(5), "0.00"))) - 100
         Printer.CurrentY = iPrint
         Printer.Print Format(.Fields(5), "0.00")
         iPrint = iPrint + 300
         .MoveNext
      Loop
   End With
   Printer.CurrentX = PLeft1(0)
   Printer.CurrentY = iPrint
   Printer.Print String(215, "-")
   
   Rc.Close
   Printer.EndDoc
   
'國外法務
'本月目標
   If Me.Tag = 0 Then
      StS(6) = "法務處"
   ElseIf Me.Tag = 1 Then
      StS(6) = "國外法務"
   End If
   StS(7) = "全所": StS(8) = "占全所"
   strExc(0) = "SELECT SUM(A0409)/1000 FROM ACC040 WHERE A0401='" & Text1(0) & _
      "' AND A0402=" & Text1(1) & " AND A0403||A0404='1TOT' AND A0405 IN (" & _
      StNum(1) & "," & StNum(2) & ")"
   RsTemp.Open strExc(0), cnnConnection
   If IsNull(RsTemp.Fields(0)) = True Then
      stVal(1) = "0"
   Else
      stVal(1) = RsTemp.Fields(0)
   End If
   RsTemp.Close
   Qty = "SELECT TMP04 FROM TEMP WHERE TMP01='總計'"
   Set Rc = Db.OpenRecordset(Qty)
   stVal(2) = Rc.Fields(0)
   If stVal(2) = 0 Or stVal(1) = 0 Then
      'StVal(3) = 100
      stVal(3) = 0
   ElseIf stVal(2) <> 0 And stVal(1) <> 0 Then
'      StVal(3) = StVal(1) * 100 / StVal(2)
       stVal(3) = (stVal(2) / stVal(1)) * 100
   End If
'上月達成
   Qty = "SELECT TMP05 FROM TEMP WHERE TMP01='總計'"
   Set Rc = Db.OpenRecordset(Qty)
   stVal(4) = Rc.Fields(0)
   stVal(5) = stVal(2) - stVal(4)
   If stVal(4) = 0 Or stVal(5) = 0 Then
      'StVal(6) = 100
      stVal(6) = 0
   ElseIf stVal(4) <> 0 And stVal(5) <> 0 Then
      stVal(6) = (stVal(5) / stVal(4)) * 100
   End If
'去年同期
   Qty = "SELECT TMP06 FROM TEMP WHERE TMP01='總計'"
   Set Rc = Db.OpenRecordset(Qty)
   stVal(7) = Rc.Fields(0)
   stVal(8) = stVal(2) - stVal(7)
   If stVal(7) = 0 Or stVal(8) = 0 Then
      stVal(9) = 0
   ElseIf stVal(7) <> 0 And stVal(8) <> 0 Then
      stVal(9) = stVal(8) * 100 / stVal(7)
   End If
'本季目標
   Select Case CInt(Text1(1))
      Case 1, 2, 3
         StDay(1) = "1": StDay(2) = "3"
      Case 4, 5, 6
         StDay(1) = "4": StDay(2) = "6"
      Case 7, 8, 9
         StDay(1) = "7": StDay(2) = "9"
      Case 10, 11, 12
         StDay(1) = "10": StDay(2) = "12"
   End Select
   strExc(0) = "SELECT SUM(A0409)/1000 FROM ACC040 WHERE A0401=" & Text1(0) & _
      " AND (A0402 BETWEEN " & StDay(1) & " AND " & StDay(2) & ") AND " & _
      "A0403||A0404='1TOT' AND A0405 IN (" & StNum(1) & "," & StNum(2) & ")"
   RsTemp.Open strExc(0), cnnConnection
   If IsNull(RsTemp.Fields(0)) = True Then
      stVal(10) = "0"
   Else
      stVal(10) = RsTemp.Fields(0)
   End If
   RsTemp.Close
   StRng(0) = " AND (A0205 BETWEEN " & Text1(0) & Format(StDay(1), "00") & _
      "00 AND " & Text1(0) & Format(StDay(2), "00") & "31)"
'   Qty = "SELECT (SUM(AX207)-SUM(AX206))/1000 FROM ACC021,ACC020 WHERE " & _
      "AX201||AX202=A0201||A0202 AND AX205 IN (" & StNum(1) & "," & StNum(2) & ")" & StRng(0)
   Qty = "SELECT (SUM(AX207)-SUM(AX206))/1000 FROM ACC021,ACC020 WHERE " & _
      "A0201=AX201(+) and A0202=AX202(+) AND AX205 IN (" & StNum(1) & "," & StNum(2) & ")" & StRng(0)
   
   RsTemp.Open Qty, cnnConnection
   If IsNull(RsTemp.Fields(0)) = True Then
      stVal(11) = "0.00"
   Else
      stVal(11) = Format(RsTemp.Fields(0), "0.00")
   End If
   RsTemp.Close
   If stVal(10) = 0 Or stVal(11) = 0 Then
      stVal(12) = 0
   ElseIf stVal(10) <> 0 And stVal(11) <> 0 Then
      stVal(12) = stVal(11) * 100 / stVal(10)
   End If
'上季目標
   Select Case CInt(Text1(1))
      Case 1, 2, 3
         StDay(1) = "10": StDay(2) = "12"
      Case 4, 5, 6
         StDay(1) = "1": StDay(2) = "3"
      Case 7, 8, 9
         StDay(1) = "4": StDay(2) = "6"
      Case 10, 11, 12
         StDay(1) = "7": StDay(2) = "9"
   End Select
   If StDay(1) = "10" Then
      StDay(0) = Text1(0) - 1
   Else
      StDay(0) = Text1(0)
   End If
   
   strExc(0) = "SELECT SUM(A0409)/1000 FROM ACC040 WHERE A0401=" & StDay(0) & _
      " AND (A0402 BETWEEN " & StDay(1) & " AND " & StDay(2) & ") AND " & _
      "A0403||A0404='1TOT' AND A0405 IN (" & StNum(1) & "," & StNum(2) & ")"
   RsTemp.Open strExc(0), cnnConnection
   If IsNull(RsTemp.Fields(0)) = True Then
      stVal(13) = "0"
   Else
      stVal(13) = RsTemp.Fields(0)
   End If
   RsTemp.Close
   StRng(0) = " AND (A0205 BETWEEN " & StDay(0) & Format(StDay(1), "00") & _
      "00 AND " & StDay(0) & Format(StDay(2), "00") & "31)"
   Qty = "SELECT (SUM(AX207)-SUM(AX206))/1000 FROM ACC021,ACC020 WHERE " & _
      "A0201=AX201(+) and A0202=AX202(+) AND AX205 IN (" & StNum(1) & "," & StNum(2) & ")" & StRng(0)
   RsTemp.Open Qty, cnnConnection
   If IsNull(RsTemp.Fields(0)) = True Then
      stVal(14) = "0.00"
   Else
      stVal(14) = Format(RsTemp.Fields(0), "0.00")
   End If
   RsTemp.Close
   If stVal(13) = 0 Or stVal(14) = 0 Then
      stVal(15) = 0
   ElseIf stVal(13) <> 0 And stVal(14) <> 0 Then
      stVal(15) = stVal(14) * 100 / stVal(13)
   End If
   Qty = "DELETE FROM TEMP"
   Db.Execute Qty
   Qty = "INSERT INTO TEMP (TMP01,TMP02,TMP03,TMP04,TMP05,TMP06,TMP07,TMP08,TMP09," & _
      "TMP10,TMP11,TMP12,TMP13,TMP14,TMP15,TMP16) VALUES ('" & StS(6) & "','" & _
      stVal(1) & "','" & stVal(2) & "','" & stVal(3) & "','" & stVal(4) & "','" & _
      stVal(5) & "','" & stVal(6) & "','" & stVal(7) & "','" & stVal(8) & "','" & _
      stVal(9) & "','" & stVal(10) & "','" & stVal(11) & "','" & stVal(12) & "','" & _
      stVal(13) & "','" & stVal(14) & "','" & stVal(15) & "')"
   Db.Execute Qty
   
'全所
'本月目標
'   strExc(0) = "SELECT SUM(A0409)/1000 FROM ACC040 WHERE A0401=" & Text1(0) & _
'      " AND A0402=" & Text1(1) & " AND A0403||A0404='1TOT' AND SUBSTR(A0405,1,1)='4'"
   strExc(0) = "SELECT SUM(A0409)/1000 FROM ACC040 WHERE A0401=" & Text1(0) & _
      " AND A0402=" & Text1(1) & " AND A0403 ='1' AND A0404='TOT'" & _
      " AND A0405 BETWEEN '4' AND '49999'"
      
   RsTemp.Open strExc(0), cnnConnection
   If IsNull(RsTemp.Fields(0)) = True Then
      stVal(1) = "0.00"
   Else
      stVal(1) = RsTemp.Fields(0)
   End If
   RsTemp.Close
   
   StDay(1) = Text1(0)
   StDay(2) = Format(Text1(1), "00")
   
'   StRng(0) = " AND (A0205 BETWEEN " & StDay(1) & StDay(2) & "01 AND " & _
'      StDay(1) & StDay(2) & "31) AND SUBSTR(AX205,1,1)='4'"
   StRng(0) = " AND (A0205 BETWEEN " & StDay(1) & StDay(2) & "01 AND " & _
      StDay(1) & StDay(2) & "31) AND AX205 BETWEEN '4' AND '4999'"
            
   strExc(0) = "SELECT (SUM(AX207)-SUM(AX206))/1000 FROM ACC021,ACC020 WHERE " & _
      "A0201=AX201(+) and A0202=AX202(+)" & StRng(0)
   RsTemp.Open strExc(0), cnnConnection
   If IsNull(RsTemp.Fields(0)) = True Then
      stVal(2) = "0.00"
   Else
      stVal(2) = RsTemp.Fields(0)
   End If
   RsTemp.Close
   If stVal(2) = 0 Or stVal(1) = 0 Then
      stVal(3) = 0
   ElseIf stVal(2) <> 0 And stVal(1) <> 0 Then
      stVal(3) = stVal(2) * 100 / stVal(1)
   End If
   
'上月達成
   If CInt(Text1(1)) = "1" Then
      StDay(1) = Text1(0) - 1
      StDay(2) = "12"
   Else
      StDay(1) = Text1(0)
      StDay(2) = Format(CInt(Text1(1)) - 1, "00")
   End If
   
'   StRng(0) = " AND (A0205 BETWEEN " & StDay(1) & StDay(2) & "01 AND " & _
'      StDay(1) & StDay(2) & "31) AND SUBSTR(AX205,1,1)='4'"
   StRng(0) = " AND (A0205 BETWEEN " & StDay(1) & StDay(2) & "01 AND " & _
      StDay(1) & StDay(2) & "31) AND AX205 BETWEEN '4' AND '4999'"

   strExc(0) = "SELECT (SUM(AX207)-SUM(AX206))/1000 FROM ACC021,ACC020 WHERE " & _
      "A0201=AX201(+) and A0202=AX202(+)" & StRng(0)
   RsTemp.Open strExc(0), cnnConnection
   If IsNull(RsTemp.Fields(0)) = True Then
      stVal(4) = "0.00"
   Else
      stVal(4) = RsTemp.Fields(0)
   End If
   RsTemp.Close
   stVal(5) = stVal(2) - stVal(4)
   stVal(6) = stVal(5) * 100 / stVal(4)
   
'去年同期
   StDay(1) = CInt(Text1(0)) - 1
   StDay(2) = Format(Text1(1), "00")

'   StRng(0) = " AND (A0205 BETWEEN " & StDay(1) & StDay(2) & "01 AND " & _
'      StDay(1) & StDay(2) & "31) AND SUBSTR(AX205,1,1)='4'"
   StRng(0) = " AND (A0205 BETWEEN " & StDay(1) & StDay(2) & "01 AND " & _
      StDay(1) & StDay(2) & "31) AND AX205 BETWEEN '4' AND '4999'"

   strExc(0) = "SELECT (SUM(AX207)-SUM(AX206))/1000 FROM ACC021,ACC020 WHERE " & _
      "A0201=AX201(+) and A0202=AX202(+)" & StRng(0)
   RsTemp.Open strExc(0), cnnConnection
   If IsNull(RsTemp.Fields(0)) = True Then
      stVal(7) = "0.00"
   Else
      stVal(7) = RsTemp.Fields(0)
   End If
   RsTemp.Close
   stVal(8) = stVal(2) - stVal(7)
   If stVal(7) = 0 Or stVal(8) = 0 Then
      stVal(9) = 0
   ElseIf stVal(7) <> 0 And stVal(8) <> 0 Then
      stVal(9) = stVal(8) * 100 / stVal(7)
   End If
'本季目標
   Select Case CInt(Text1(1))
      Case 1, 2, 3
         StDay(1) = "1": StDay(2) = "3"
      Case 4, 5, 6
         StDay(1) = "4": StDay(2) = "6"
      Case 7, 8, 9
         StDay(1) = "7": StDay(2) = "9"
      Case 10, 11, 12
         StDay(1) = "10": StDay(2) = "12"
   End Select
'   strExc(0) = "SELECT SUM(A0409)/1000 FROM ACC040 WHERE A0401=" & Text1(0) & _
'      " AND (A0402 BETWEEN " & StDay(1) & " AND " & StDay(2) & ") AND " & _
'      "A0403||A0404='1TOT' AND SUBSTR(A0405,1,1)='4'"
   strExc(0) = "SELECT SUM(A0409)/1000 FROM ACC040 WHERE A0401=" & Text1(0) & _
      " AND (A0402 BETWEEN " & StDay(1) & " AND " & StDay(2) & ") AND " & _
      "A0403='1' AND A0404='TOT' AND A0405 BETWEEN '4' AND '4999' "
      
      
   RsTemp.Open strExc(0), cnnConnection
   If IsNull(RsTemp.Fields(0)) = True Then
      stVal(10) = "0"
   Else
      stVal(10) = RsTemp.Fields(0)
   End If
   RsTemp.Close
   StRng(0) = " AND (A0205 BETWEEN " & Text1(0) & Format(StDay(1), "00") & _
      "00 AND " & Text1(0) & Format(StDay(2), "00") & "31)"
'   Qty = "SELECT (SUM(AX207)-SUM(AX206))/1000 FROM ACC021,ACC020 WHERE " & _
'      "A0201=AX201(+) and A0202=AX202(+) AND SUBSTR(AX205,1,1)='4'" & StRng(0)
   Qty = "SELECT (SUM(AX207)-SUM(AX206))/1000 FROM ACC021,ACC020 WHERE " & _
      "A0201=AX201(+) and A0202=AX202(+) AND AX205 BETWEEN '4' AND '4999' " & StRng(0)
   
   RsTemp.Open Qty, cnnConnection
   If IsNull(RsTemp.Fields(0)) = True Then
      stVal(11) = "0.00"
   Else
      stVal(11) = Format(RsTemp.Fields(0), "0.00")
   End If
   RsTemp.Close
   If stVal(10) <> 0 Then
      stVal(12) = stVal(11) * 100 / stVal(10)
   Else
      stVal(12) = 0
   End If
   
'上季目標
   Select Case CInt(Text1(1).Text)
      Case 1, 2, 3
         StDay(1) = "10": StDay(2) = "12"
      Case 4, 5, 6
         StDay(1) = "1": StDay(2) = "3"
      Case 7, 8, 9
         StDay(1) = "4": StDay(2) = "6"
      Case 10, 11, 12
         StDay(1) = "7": StDay(2) = "9"
   End Select
   If StDay(1) = "10" Then
      StDay(0) = Text1(0) - 1
   Else
      StDay(0) = Text1(0)
   End If
   
   strExc(0) = "SELECT SUM(A0409)/1000 FROM ACC040 WHERE A0401=" & StDay(0) & _
      " AND (A0402 BETWEEN " & StDay(1) & " AND " & StDay(2) & ") AND " & _
      "A0403 = '1' AND A0404='TOT' AND A0405 BETWEEN '4' AND '4999'"
   RsTemp.Open strExc(0), cnnConnection
   If IsNull(RsTemp.Fields(0)) = True Then
      stVal(13) = "0"
   Else
      stVal(13) = RsTemp.Fields(0)
   End If
   RsTemp.Close
   StRng(0) = " AND (A0205 BETWEEN " & StDay(0) & Format(StDay(1), "00") & _
      "00 AND " & StDay(0) & Format(StDay(2), "00") & "31)"
'   Qty = "SELECT (SUM(AX207)-SUM(AX206))/1000 FROM ACC021,ACC020 WHERE " & _
'      "A0201=AX201(+) and A0202=AX202(+) AND SUBSTR(AX205,1,1)='4'" & StRng(0)
   Qty = "SELECT (SUM(AX207)-SUM(AX206))/1000 FROM ACC021,ACC020 WHERE " & _
      "A0201=AX201(+) and A0202=AX202(+) AND AX205 BETWEEN '4' AND '4999' " & StRng(0)
   
   
   RsTemp.Open Qty, cnnConnection
   If IsNull(RsTemp.Fields(0)) = True Then
      stVal(14) = "0.00"
   Else
      stVal(14) = Format(RsTemp.Fields(0), "0.00")
   End If
   RsTemp.Close
   If stVal(13) <> 0 Then
      stVal(15) = stVal(14) * 100 / stVal(13)
   Else
      stVal(15) = 0
   End If
   Qty = "INSERT INTO TEMP (TMP01,TMP02,TMP03,TMP04,TMP05,TMP06,TMP07,TMP08,TMP09," & _
      "TMP10,TMP11,TMP12,TMP13,TMP14,TMP15,TMP16) VALUES ('" & StS(7) & "','" & _
      stVal(1) & "','" & stVal(2) & "','" & stVal(3) & "','" & stVal(4) & "','" & _
      stVal(5) & "','" & stVal(6) & "','" & stVal(7) & "','" & stVal(8) & "','" & _
      stVal(9) & "','" & stVal(10) & "','" & stVal(11) & "','" & stVal(12) & "','" & _
      stVal(13) & "','" & stVal(14) & "','" & stVal(15) & "')"
   Db.Execute Qty
   Qty = "SELECT TMP02,TMP03,TMP04,TMP05,TMP06,TMP07,TMP08,TMP09,TMP10,TMP11,TMP12," & _
      "TMP13,TMP14,TMP15,TMP16 FROM TEMP WHERE TMP01='" & StS(6) & "'"
   Set Rc = Db.OpenRecordset(Qty)
   Qty = "SELECT TMP02,TMP03,TMP04,TMP05,TMP06,TMP07,TMP08,TMP09,TMP10,TMP11,TMP12," & _
      "TMP13,TMP14,TMP15,TMP16 FROM TEMP WHERE TMP01='" & StS(7) & "'"
   Set Rc1 = Db.OpenRecordset(Qty)
   For i = 0 To 14
      If Rc1.Fields(i) = 0 Then
         stVal(i + 1) = 0
      Else
         stVal(i + 1) = Rc.Fields(i) * 100 / Rc1.Fields(i)
      End If
   Next
   With Rc
      Qty = "INSERT INTO TEMP (TMP01,TMP02,TMP03,TMP04,TMP05,TMP06,TMP07,TMP08," & _
      "TMP09,TMP10,TMP11,TMP12,TMP13,TMP14,TMP15,TMP16) VALUES ('" & StS(8) & "','" & _
      stVal(1) & "','" & stVal(2) & "','" & stVal(3) & "','" & stVal(4) & "','" & _
      stVal(5) & "','" & stVal(6) & "','" & stVal(7) & "','" & stVal(8) & "','" & _
      stVal(9) & "','" & stVal(10) & "','" & stVal(11) & "','" & stVal(12) & "','" & _
      stVal(13) & "','" & stVal(14) & "','" & stVal(15) & "')"
      Db.Execute Qty
   End With
   GetPrintLeft
   CaseTitle1 1
   iPrint = 3100
   Qty = "SELECT TMP01,TMP02,TMP03,TMP04,TMP05,TMP06,TMP07,TMP08,TMP09,TMP10," & _
      "TMP11,TMP12,TMP13,TMP14,TMP15,TMP16,TMP17 FROM TEMP"
   Set Rc = Db.OpenRecordset(Qty)
   With Rc
      Do While Not .EOF
         Printer.CurrentX = PLeft1(0):      Printer.CurrentY = iPrint
         Printer.Print .Fields(0)
         Printer.CurrentX = PLeft1(1) + 600 - (Printer.TextWidth(CheckStr(Format(.Fields(1), "0.00"))))
         Printer.CurrentY = iPrint
         Printer.Print Format(.Fields(1), "0.00")
         Printer.CurrentX = PLeft1(2) + 600 - (Printer.TextWidth(CheckStr(Format(.Fields(2), "0.00"))))
         Printer.CurrentY = iPrint
         Printer.Print Format(.Fields(2), "0.00")
         Printer.CurrentX = PLeft1(3) + 600 - (Printer.TextWidth(CheckStr(Format(.Fields(3), "0.00"))))
         Printer.CurrentY = iPrint
         Printer.Print Format(.Fields(3), "0.00")
         Printer.CurrentX = PLeft1(4) + 600 - (Printer.TextWidth(CheckStr(Format(.Fields(4), "0.00"))))
         Printer.CurrentY = iPrint
         Printer.Print Format(.Fields(4), "0.00")
         Printer.CurrentX = PLeft1(5) + 600 - (Printer.TextWidth(CheckStr(Format(.Fields(5), "0.00"))))
         Printer.CurrentY = iPrint
         Printer.Print Format(.Fields(5), "0.00")
         Printer.CurrentX = PLeft1(6) + 600 - (Printer.TextWidth(CheckStr(Format(.Fields(6), "0.00"))))
         Printer.CurrentY = iPrint
         Printer.Print Format(.Fields(6), "0.00")
         Printer.CurrentX = PLeft1(7) + 600 - (Printer.TextWidth(CheckStr(Format(.Fields(7), "0.00"))))
         Printer.CurrentY = iPrint
         Printer.Print Format(.Fields(7), "0.00")
         Printer.CurrentX = PLeft1(8) + 600 - (Printer.TextWidth(CheckStr(Format(.Fields(8), "0.00"))))
         Printer.CurrentY = iPrint
         Printer.Print Format(.Fields(8), "0.00")
         Printer.CurrentX = PLeft1(9) + 600 - (Printer.TextWidth(CheckStr(Format(.Fields(9), "0.00"))))
         Printer.CurrentY = iPrint
         Printer.Print Format(.Fields(9), "0.00")
         Printer.CurrentX = PLeft1(10) + 600 - (Printer.TextWidth(CheckStr(Format(.Fields(10), "0.00"))))
         Printer.CurrentY = iPrint
         Printer.Print Format(.Fields(10), "0.00")
         Printer.CurrentX = PLeft1(11) + 600 - (Printer.TextWidth(CheckStr(Format(.Fields(11), "0.00"))))
         Printer.CurrentY = iPrint
         Printer.Print Format(.Fields(11), "0.00")
         Printer.CurrentX = PLeft1(12) + 600 - (Printer.TextWidth(CheckStr(Format(.Fields(12), "0.00"))))
         Printer.CurrentY = iPrint
         Printer.Print Format(.Fields(12), "0.00")
         Printer.CurrentX = PLeft1(13) + 600 - (Printer.TextWidth(CheckStr(Format(.Fields(13), "0.00"))))
         Printer.CurrentY = iPrint
         Printer.Print Format(.Fields(13), "0.00")
         Printer.CurrentX = PLeft1(14) + 600 - (Printer.TextWidth(CheckStr(Format(.Fields(14), "0.00"))))
         Printer.CurrentY = iPrint
         Printer.Print Format(.Fields(14), "0.00")
         Printer.CurrentX = PLeft1(15) + 600 - (Printer.TextWidth(CheckStr(Format(.Fields(15), "0.00"))))
         Printer.CurrentY = iPrint
         Printer.Print Format(.Fields(15), "0.00")
         Printer.CurrentX = PLeft1(16) + 600 - (Printer.TextWidth(CheckStr(Format(.Fields(16), "0.00"))))
         Printer.CurrentY = iPrint
         Printer.Print Format(.Fields(16), "0.00")
         iPrint = iPrint + 300
         .MoveNext
      Loop
   End With
   Printer.CurrentX = PLeft1(0)
   Printer.CurrentY = iPrint
   Printer.Print String(225, "-")
   Rc.Close
   Printer.EndDoc
   Exit Sub
ErrHand:
   MsgBox Err.Description
End Sub

Private Sub GetPrintLeft()
   PLeft1(0) = 500:       PLeft1(1) = 1500
   PLeft1(2) = 2500:      PLeft1(3) = 3500
   PLeft1(4) = 4500:      PLeft1(5) = 5500
   PLeft1(6) = 6500:      PLeft1(7) = 7500
   PLeft1(8) = 8500:      PLeft1(9) = 9500
   PLeft1(10) = 10500:    PLeft1(11) = 11500
   PLeft1(12) = 12500:    PLeft1(13) = 13500
   PLeft1(14) = 14500:    PLeft1(15) = 15500
   PLeft1(16) = 16500:    PLeft1(17) = 17500
  
   PLeft(0) = 500
   PLeft(1) = 1500         '北一
   PLeft(2) = 2400         '北三
   PLeft(3) = 3300         '北四
   PLeft(4) = 4200         '北五
   PLeft(5) = 5100         '北所
   PLeft(6) = 6200         '中一
   PLeft(7) = 7100         '中二
   PLeft(8) = 8000         '中三
   PLeft(9) = 8900         '中所
   PLeft(10) = 10000       '台南
   PLeft(11) = 10900       '高雄
   PLeft(12) = 11800       '分所
   PLeft(13) = 12900       '國外
   PLeft(14) = 13800       '其他
   PLeft(15) = 14700       '總計
   PLeft(16) = 15600       '律師
  ' PLeft(17) = 15200
  
End Sub

Private Sub CaseTitle(ByVal Area As String, ByVal Page As String)
 Dim i As Integer, St As String
   i = 500
   Printer.Orientation = 1
   Printer.Font.Size = 22
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.CurrentX = 4000:         Printer.CurrentY = i
   If Me.Tag = 0 Then
      Printer.Print "法務處業績分析表1"
   ElseIf Me.Top = 1 Then
      Printer.Print "國外法務業績分析表1"
   End If
   Printer.Font.Underline = False
   Printer.Font.Size = 11
   Printer.Font.Bold = False
   Printer.CurrentX = 5000:         Printer.CurrentY = i + 500
   Printer.Print "收文年月 : " & Text1(0) & " / " & Format(Text1(1), "00")
   Printer.CurrentX = 500:              Printer.CurrentY = i + 800
   Printer.Print "列印人 : " & strUserName
   Printer.CurrentX = 9000:            Printer.CurrentY = i + 800
   Printer.Print "列印日期 : " & ChangeTStringToTDateString(GetTaiwanTodayDate)
   Printer.CurrentX = 9000:            Printer.CurrentY = i + 1100
   Printer.Print "頁次 : " & Page
   Printer.CurrentX = 500:              Printer.CurrentY = i + 1400
   Printer.Print String(225, "-")
   
   Printer.CurrentX = PLeft(2):         Printer.CurrentY = i + 1700
   Printer.Print StS(1)
   Printer.CurrentX = PLeft(4):         Printer.CurrentY = i + 1700
   Printer.Print StS(2)
   Printer.CurrentX = PLeft(6):         Printer.CurrentY = i + 1700
   Printer.Print StS(3)
   Printer.CurrentX = PLeft(8):         Printer.CurrentY = i + 1700
   Printer.Print StS(4)
   Printer.CurrentX = PLeft(10):         Printer.CurrentY = i + 1700
   Printer.Print StS(5)
  ' Printer.CurrentX = PLeft(17):         Printer.CurrentY = i + 2000
  ' Printer.Print String(205, "-")
End Sub

Private Sub CaseTitle1(ByVal Page As String)
 Dim i As Integer, St As String
   i = 500
   Printer.Orientation = 2
   Printer.Font.Size = 22
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.CurrentX = 6000:         Printer.CurrentY = i
   If Me.Tag = 0 Then
      Printer.Print "法務處業績分析表2"
   ElseIf Me.Top = 1 Then
      Printer.Print "國外法務業績分析表2"
   End If
   Printer.Font.Underline = False
   Printer.Font.Size = 11
   Printer.Font.Bold = False
   Printer.CurrentX = 6900:         Printer.CurrentY = i + 500
   Printer.Print "收文年月 : " & Text1(0) & " / " & Format(Text1(1), "00")
   Printer.CurrentX = 500:              Printer.CurrentY = i + 800
   Printer.Print "列印人 : " & strUserName
   Printer.CurrentX = 13000:            Printer.CurrentY = i + 800
   Printer.Print "列印日期 : " & ChangeTStringToTDateString(GetTaiwanTodayDate)
   Printer.CurrentX = 13000:            Printer.CurrentY = i + 1100
   Printer.Print "頁次 : " & Page
   Printer.CurrentX = 500:              Printer.CurrentY = i + 1700
   Printer.Print String(225, "-")
   
   
   Printer.CurrentX = PLeft1(1):         Printer.CurrentY = i + 2000
   Printer.Print "本月目標"
   Printer.CurrentX = PLeft1(2):         Printer.CurrentY = i + 2000
   Printer.Print "實際達成"
   Printer.CurrentX = PLeft1(3):         Printer.CurrentY = i + 2000
   Printer.Print "達成率"
   Printer.CurrentX = PLeft1(4):         Printer.CurrentY = i + 2000
   Printer.Print "上月達成"
   Printer.CurrentX = PLeft1(5):         Printer.CurrentY = i + 2000
   Printer.Print "成長點數"
   Printer.CurrentX = PLeft1(6):         Printer.CurrentY = i + 2000
   Printer.Print "成長率"
   Printer.CurrentX = PLeft1(7):         Printer.CurrentY = i + 2000
   Printer.Print "去年同期"
   Printer.CurrentX = PLeft1(8):         Printer.CurrentY = i + 2000
   Printer.Print "成長點數"
   Printer.CurrentX = PLeft1(9):         Printer.CurrentY = i + 2000
   Printer.Print "成長率"
   Printer.CurrentX = PLeft1(10):         Printer.CurrentY = i + 2000
   Printer.Print "本季目標"
   Printer.CurrentX = PLeft1(11):         Printer.CurrentY = i + 2000
   Printer.Print "達成點數"
   Printer.CurrentX = PLeft1(12):         Printer.CurrentY = i + 2000
   Printer.Print "達成率"
   Printer.CurrentX = PLeft1(13):         Printer.CurrentY = i + 2000
   Printer.Print "上季目標"
   Printer.CurrentX = PLeft1(14):         Printer.CurrentY = i + 2000
   Printer.Print "達成點數"
   Printer.CurrentX = PLeft1(15):         Printer.CurrentY = i + 2000
   Printer.Print "達成率"
End Sub

Private Sub Form_Activate()
  Text1(0).SetFocus
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm083009 = Nothing
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   If (KeyAscii > 57 Or KeyAscii < 48) And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub
