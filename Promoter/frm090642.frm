VERSION 5.00
Begin VB.Form frm090642 
   BorderStyle     =   1  '單線固定
   Caption         =   "每周承辦會議統計查詢"
   ClientHeight    =   2484
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4284
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2484
   ScaleWidth      =   4284
   Begin VB.CommandButton cmdExcel 
      BackColor       =   &H00C0FFC0&
      Caption         =   "產生 EXCEL (&O)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   1056
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   1824
      Width           =   2292
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   0
      Left            =   3384
      TabIndex        =   5
      Top             =   96
      Width           =   852
   End
   Begin VB.TextBox txtDate 
      Height          =   345
      Index           =   0
      Left            =   1536
      MaxLength       =   7
      TabIndex        =   0
      Top             =   732
      Width           =   795
   End
   Begin VB.TextBox txtDate 
      Height          =   345
      Index           =   1
      Left            =   2532
      MaxLength       =   7
      TabIndex        =   1
      Top             =   732
      Width           =   795
   End
   Begin VB.TextBox txtDate 
      Height          =   345
      Index           =   2
      Left            =   1536
      MaxLength       =   7
      TabIndex        =   2
      Top             =   1164
      Width           =   795
   End
   Begin VB.TextBox txtDate 
      Height          =   345
      Index           =   3
      Left            =   2532
      MaxLength       =   7
      TabIndex        =   3
      Top             =   1164
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "當期統計區間："
      Height          =   216
      Left            =   216
      TabIndex        =   7
      Top             =   816
      Width           =   1368
   End
   Begin VB.Line Line1 
      X1              =   2112
      X2              =   2802
      Y1              =   912
      Y2              =   912
   End
   Begin VB.Line Line2 
      X1              =   2136
      X2              =   2826
      Y1              =   1332
      Y2              =   1332
   End
   Begin VB.Label Label2 
      Caption         =   "上期統計區間："
      Height          =   216
      Left            =   216
      TabIndex        =   6
      Top             =   1236
      Width           =   1416
   End
End
Attribute VB_Name = "frm090642"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create by Lydia 2024/12/02
Option Explicit
Dim oObj As Control
Dim m_ESeq As String
Dim strFileName As String

Private Sub cmdExit_Click(Index As Integer)
   Unload Me
End Sub

Private Sub cmdExcel_Click()
Dim intErr As Integer
Dim bolTmp As Boolean
   
   intErr = -1
   For Each oObj In txtDate
      If Trim(oObj.Text) = "" Then
         intErr = oObj.Index
         MsgBox "請輸入" & IIf(intErr < 2, "當期", "上期") & "統計區間！", vbCritical
         GoTo JumpToExit
      Else
         Call txtDate_Validate(oObj.Index, bolTmp)
         If bolTmp = True Then
            GoTo JumpToExit
         End If
      End If
   Next
   
   If txtDate(0) > strSrvDate(2) Then
      intErr = 0
      MsgBox "當期統計區間起值不可大於系統日！", vbCritical
      GoTo JumpToExit
   End If
   If txtDate(2) > strSrvDate(2) Then
      intErr = 2
      MsgBox "上期統計區間起值不可大於系統日！", vbCritical
      GoTo JumpToExit
   End If
   If Val(txtDate(0)) >= Val(txtDate(1)) Then
      intErr = 0
      MsgBox "當期統計區間起值不可大於等於迄值！", vbCritical
      GoTo JumpToExit
   End If
   If Val(txtDate(2)) >= Val(txtDate(3)) Then
      intErr = 2
      MsgBox "上期統計區間起值不可大於等於迄值！", vbCritical
      GoTo JumpToExit
   End If
   If Val(txtDate(2)) >= Val(txtDate(1)) Then
      intErr = 2
      MsgBox "上期統計區間起值不可大於等於當期統計區間！", vbCritical
      GoTo JumpToExit
   End If
   If DBDATE(txtDate(2)) <> CompDate(1, -1, DBDATE(txtDate(0))) Then
      If MsgBox("上期統計區間起值不等於當期統計區間起值減一個月，是否重新輸入？", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
         intErr = 2
         GoTo JumpToExit
      End If
   End If
   If GetLastDay(DBDATE(Mid(txtDate(1), 5) & "01")) = DBDATE(txtDate(1)) Then
      If GetLastDay(DBDATE(Mid(txtDate(3), 5) & "01")) <> DBDATE(txtDate(3)) Then
         If MsgBox("上期統計區間迄值不等於該月最後一天，是否重新輸入？", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
            intErr = 3
            GoTo JumpToExit
         End If
      End If
   Else
      If DBDATE(txtDate(3)) <> CompDate(1, -1, DBDATE(txtDate(1))) Then
         If MsgBox("上期統計區間迄值不等於當期統計區間迄值減一個月，是否重新輸入？", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
            intErr = 3
            GoTo JumpToExit
         End If
      End If
   End If
   
   strFileName = strExcelPath & txtDate(0) & Me.Caption & MsgText(43)
   RidFile strFileName
   
   m_ESeq = "1"
   cnnConnection.Execute "delete from rdatafactory where formname='" & Me.Name & "' and id = '" & strUserNum & "' and seqno ='" & m_ESeq & "' "
   
   Screen.MousePointer = vbHourglass
   cmdExcel.Enabled = False
   If Process = False Then
      strFileName = ""
   End If
   cmdExcel.Enabled = True
   Screen.MousePointer = vbDefault
   
   If strFileName <> "" Then
      MsgBox "Excel檔案產生完成！檔案位置：" & strExcelPathN
   End If
   
   Exit Sub
   
JumpToExit:
   If intErr >= 1 Then
      txtDate(intErr).SetFocus
      txtDate_GotFocus intErr
   End If
End Sub

Private Sub Form_Load()

   MoveFormToCenter Me
   
   If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
       MkDir strExcelPath
   End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   Set frm090642 = Nothing
End Sub

Private Sub txtDate_GotFocus(Index As Integer)
   TextInverse txtDate(Index)
End Sub

Private Sub txtDate_Validate(Index As Integer, Cancel As Boolean)
Dim strTemp1 As String

   If txtDate(Index) <> "" Then
      strTemp1 = txtDate(Index)
      If CheckIsTaiwanDate(strTemp1) = False Then
         MsgBox "請輸入民國日期!", vbCritical
         txtDate(Index).SetFocus
         txtDate_GotFocus Index
         Cancel = True
         Exit Sub
      End If
   End If
End Sub

Private Function Process() As Boolean
Dim strQ1 As String, intQ As Integer
Dim strCon As String, intPos As Integer
Dim strDateN(0 To 1) As String, strDateP(0 To 1) As String
   
   ClearQueryLog (Me.Name)
   strDateN(0) = DBDATE(txtDate(0))
   strDateN(1) = DBDATE(txtDate(1))
   strDateP(0) = DBDATE(txtDate(2))
   strDateP(1) = DBDATE(txtDate(3))
   pub_QL05 = pub_QL05 & ";當期統計區間：" & txtDate(0) & "-" & txtDate(1)
   pub_QL05 = pub_QL05 & ";上期統計區間：" & txtDate(2) & "-" & txtDate(3)
   
On Error GoTo ErrHandle

'*****系統別R001
   strSql = "insert into rdatafactory (formname,id,seqno,rowseq,r001) values ('" & Me.Name & "', '" & strUserNum & "','" & m_ESeq & "','1','P') "
   cnnConnection.Execute strSql
   strSql = "insert into rdatafactory (formname,id,seqno,rowseq,r001) values ('" & Me.Name & "', '" & strUserNum & "','" & m_ESeq & "','2','CFP') "
   cnnConnection.Execute strSql
   
'*****當期收文R002,上期R003
   'P案含改請(排除設計)
   intPos = 2
   strCon = ""
   strSql = "select LISTAGG(CPM02,',') WITHIN GROUP (ORDER BY CPM02) AS SYSLIST from casepropertymap where cpm01='P' and cpm03 like '%設計%' and length(cpm02)=3"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      strCon = " and cp10 not in (" & GetAddStr("" & RsTemp.Fields("syslist")) & ") "
   End If
   
   For intQ = 0 To 1
      '範圍條件來自frm100105_2.DoTemp
      strQ1 = "select cp01,count(cp01||cp02||cp03||cp04) cnt from caseprogress, patent where cp05>=" & IIf(intQ = 0, strDateN(0), strDateP(0)) & " and cp05<= " & IIf(intQ = 0, strDateN(1), strDateP(1)) & _
              " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) AND CP159=0 AND cp26 IS NULL AND cp21 IS NULL and cp09< 'B' and cp01||cp02<>'TT999999' and pa09>='000' and pa09<='000' " & _
              " and cp01= 'P' and (instr('" & NewCasePtyList & "',CP10)>0 or substr(CP10,1,1)='3') " & strCon & _
              " group by cp01 "
      intI = 1
      strExc(0) = "0"
      Set RsTemp = ClsLawReadRstMsg(intI, strQ1)
      If intI = 1 Then
         strExc(0) = Val("" & RsTemp.Fields("cnt"))
      End If
      strSql = "Update rdatafactory set " & "R" & Format(intPos + intQ, "000") & "='" & strExc(0) & "' where formname='" & Me.Name & "' and id='" & strUserNum & "' and seqno = '" & m_ESeq & "' and r001='P' "
      cnnConnection.Execute strSql, intI
   Next intQ
   '當期發文R012,上期R013
   intPos = 12
   For intQ = 0 To 1
      '範圍條件來自frm100105_2.DoTemp
      strQ1 = "select cp01,count(cp01||cp02||cp03||cp04) cnt from caseprogress, patent where cp27>=" & IIf(intQ = 0, strDateN(0), strDateP(0)) & " and cp27<= " & IIf(intQ = 0, strDateN(1), strDateP(1)) & _
              " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) AND CP159=0 AND cp26 IS NULL AND cp21 IS NULL and cp09< 'B' and cp01||cp02<>'TT999999' and pa09>='000' and pa09<='000' " & _
              " and cp01= 'P' and (instr('" & NewCasePtyList & "',CP10)>0 or substr(CP10,1,1)='3') " & strCon & _
              " group by cp01 "
      intI = 1
      strExc(0) = "0"
      Set RsTemp = ClsLawReadRstMsg(intI, strQ1)
      If intI = 1 Then
         strExc(0) = Val("" & RsTemp.Fields("cnt"))
      End If
      strSql = "Update rdatafactory set " & "R" & Format(intPos + intQ, "000") & "='" & strExc(0) & "' where formname='" & Me.Name & "' and id='" & strUserNum & "' and seqno = '" & m_ESeq & "' and r001='P' "
      cnnConnection.Execute strSql, intI
   Next intQ
   
   'CFP案含改請(排除設計)
   intPos = 2
   strCon = ""
   strSql = "select LISTAGG(CPM02,',') WITHIN GROUP (ORDER BY CPM02) AS SYSLIST from casepropertymap where cpm01='CFP' and cpm03 like '%設計%' and length(cpm02)=3"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      strCon = " and cp10 not in (" & GetAddStr("" & RsTemp.Fields("syslist")) & ") "
   End If
   
   For intQ = 0 To 1
      '範圍條件來自frm100105_2.DoTemp---不記件(CP26)也列入統計
      strQ1 = "select cp01,count(cp01||cp02||cp03||cp04) cnt from caseprogress, patent where cp05>=" & IIf(intQ = 0, strDateN(0), strDateP(0)) & " and cp05<= " & IIf(intQ = 0, strDateN(1), strDateP(1)) & _
              " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) AND CP159=0 AND cp21 IS NULL and cp09< 'B' and cp01||cp02<>'TT999999' " & _
              " and cp01= 'CFP' and (instr('" & NewCasePtyList & "',CP10)>0 or substr(CP10,1,1)='3') " & strCon & _
              " group by cp01 "
      intI = 1
      strExc(0) = "0"
      Set RsTemp = ClsLawReadRstMsg(intI, strQ1)
      If intI = 1 Then
         strExc(0) = Val("" & RsTemp.Fields("cnt"))
      End If
      strSql = "Update rdatafactory set " & "R" & Format(intPos + intQ, "000") & "='" & strExc(0) & "' where formname='" & Me.Name & "' and id='" & strUserNum & "' and seqno = '" & m_ESeq & "' and r001='CFP' "
      cnnConnection.Execute strSql, intI
   Next intQ
   '當期發文R012,上期R013
   intPos = 12
   For intQ = 0 To 1
      '範圍條件來自frm100105_2.DoTemp---不記件(CP26)也列入統計
      strQ1 = "select cp01,count(cp01||cp02||cp03||cp04) cnt from caseprogress, patent where cp27>=" & IIf(intQ = 0, strDateN(0), strDateP(0)) & " and cp27<= " & IIf(intQ = 0, strDateN(1), strDateP(1)) & _
              " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) AND CP159=0 AND cp21 IS NULL and cp09< 'B' and cp01||cp02<>'TT999999' " & _
              " and cp01= 'CFP' and (instr('" & NewCasePtyList & "',CP10)>0 or substr(CP10,1,1)='3') " & strCon & _
              " group by cp01 "
      intI = 1
      strExc(0) = "0"
      Set RsTemp = ClsLawReadRstMsg(intI, strQ1)
      If intI = 1 Then
         strExc(0) = Val("" & RsTemp.Fields("cnt"))
      End If
      strSql = "Update rdatafactory set " & "R" & Format(intPos + intQ, "000") & "='" & strExc(0) & "' where formname='" & Me.Name & "' and id='" & strUserNum & "' and seqno = '" & m_ESeq & "' and r001='CFP' "
      cnnConnection.Execute strSql, intI
   Next intQ
   
'*****當期齊備R004,上期R005
   'P案：案件性質分別統計101-102、301-302、307
   intPos = 4
   For intQ = 0 To 1
      '範圍條件來自frm090613.Process
      strQ1 = "select cp01,count(ep02) cnt from engineerprogress,caseprogress, patent where ep06>=" & IIf(intQ = 0, strDateN(0), strDateP(0)) & " and ep06<= " & IIf(intQ = 0, strDateN(1), strDateP(1)) & _
              " and ep02=cp09(+) and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) AND cp26 IS NULL and pa09>='000' and pa09<='000' " & _
              " and cp01= 'P' and cp10 in ('101','102','301','302','307') group by cp01 "
      intI = 1
      strExc(0) = "0"
      Set RsTemp = ClsLawReadRstMsg(intI, strQ1)
      If intI = 1 Then
         strExc(0) = Val("" & RsTemp.Fields("cnt"))
      End If
      strSql = "Update rdatafactory set " & "R" & Format(intPos + intQ, "000") & "='" & strExc(0) & "' where formname='" & Me.Name & "' and id='" & strUserNum & "' and seqno = '" & m_ESeq & "' and r001='P' "
      cnnConnection.Execute strSql, intI
   Next intQ
   
   'CFP案：案件性質分別統計101-102、301-302、307-307、113-113、118-118
   For intQ = 0 To 1
      '範圍條件來自frm090613.Process
      strQ1 = "select cp01,count(ep02) cnt from engineerprogress,caseprogress, patent where ep06>=" & IIf(intQ = 0, strDateN(0), strDateP(0)) & " and ep06<= " & IIf(intQ = 0, strDateN(1), strDateP(1)) & _
              " and ep02=cp09(+) and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) AND cp26 IS NULL " & _
              " and cp01= 'CFP' and cp10 in ('101','102','301','302','307','113','118') group by cp01 "
      intI = 1
      strExc(0) = "0"
      Set RsTemp = ClsLawReadRstMsg(intI, strQ1)
      If intI = 1 Then
         strExc(0) = Val("" & RsTemp.Fields("cnt"))
      End If
      strSql = "Update rdatafactory set " & "R" & Format(intPos + intQ, "000") & "='" & strExc(0) & "' where formname='" & Me.Name & "' and id='" & strUserNum & "' and seqno = '" & m_ESeq & "' and r001='CFP' "
      cnnConnection.Execute strSql, intI
   Next intQ
   
'*****當期會稿R006,上期R007
   'P案：案件性質分別統計101-102、301-302、307
   intPos = 6
   For intQ = 0 To 1
      '範圍條件來自frm090613.Process  'Modified by Morgan 2017/12/18 會稿日或會完日都要剔除不會稿案件(ep34)--王副總
      strQ1 = "select cp01,count(ep02) cnt from engineerprogress,caseprogress, patent where ep07>=" & IIf(intQ = 0, strDateN(0), strDateP(0)) & " and ep07<= " & IIf(intQ = 0, strDateN(1), strDateP(1)) & _
              " and ep02=cp09(+) and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and NVL(ep34,'Y')='Y' AND cp26 IS NULL and pa09>='000' and pa09<='000' " & _
              " and cp01= 'P' and cp10 in ('101','102','301','302','307') group by cp01 "
      intI = 1
      strExc(0) = "0"
      Set RsTemp = ClsLawReadRstMsg(intI, strQ1)
      If intI = 1 Then
         strExc(0) = Val("" & RsTemp.Fields("cnt"))
      End If
      strSql = "Update rdatafactory set " & "R" & Format(intPos + intQ, "000") & "='" & strExc(0) & "' where formname='" & Me.Name & "' and id='" & strUserNum & "' and seqno = '" & m_ESeq & "' and r001='P' "
      cnnConnection.Execute strSql, intI
   Next intQ
   
   'CFP案：案件性質分別統計101-102、301-302、307-307、113-113、118-118
   For intQ = 0 To 1
      '範圍條件來自frm090613.Process
      strQ1 = "select cp01,count(ep02) cnt from engineerprogress,caseprogress, patent where ep07>=" & IIf(intQ = 0, strDateN(0), strDateP(0)) & " and ep07<= " & IIf(intQ = 0, strDateN(1), strDateP(1)) & _
              " and ep02=cp09(+) and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and NVL(ep34,'Y')='Y' AND cp26 IS NULL " & _
              " and cp01= 'CFP' and cp10 in ('101','102','301','302','307','113','118') group by cp01 "
      intI = 1
      strExc(0) = "0"
      Set RsTemp = ClsLawReadRstMsg(intI, strQ1)
      If intI = 1 Then
         strExc(0) = Val("" & RsTemp.Fields("cnt"))
      End If
      strSql = "Update rdatafactory set " & "R" & Format(intPos + intQ, "000") & "='" & strExc(0) & "' where formname='" & Me.Name & "' and id='" & strUserNum & "' and seqno = '" & m_ESeq & "' and r001='CFP' "
      cnnConnection.Execute strSql, intI
   Next intQ
   
'*****待辦-新案R008,上月平均由人員補上R009,未會稿-非新案R010,上月平均由人員補上R011
   strSql = "Update rdatafactory set r008=0, r009=0, r010=0, r011=0 where formname='" & Me.Name & "' and id='" & strUserNum & "' "
   cnnConnection.Execute strSql, intI
   
   strQ1 = " Select cp01,type,''||count(*) as cnt From (Select cp01,ep02,Decode(Substr(cp10,1,1),'3','新申請案',Decode(Instr('" & NewCasePtyList & "',cp10),0, '非新申請案','新申請案')) Type " & _
           " From CaseProgress, EngineerProgress,staff Where cp158=0 And cp159=0 and cp14=st01(+) And st03>='P1' And st03<='P11' And cp09=ep02(+) And cp26 is null " & _
           " And Nvl(ep06,0) >0 and nvl(ep07,0) =0 and cp01 in ('P','CFP') ) group by cp01,type "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strQ1)
   If intI = 1 Then
      RsTemp.MoveFirst
      Do While Not RsTemp.EOF
         If InStr("" & RsTemp.Fields("type"), "非新") > 0 Then
            strCon = "R010"
         Else
            strCon = "R008"
         End If
         strSql = "Update rdatafactory set " & strCon & "=" & RsTemp.Fields("cnt") & " where formname='" & Me.Name & "' and id='" & strUserNum & "' and seqno = '" & m_ESeq & "' and r001='" & RsTemp.Fields("cp01") & "' "
         cnnConnection.Execute strSql, intI
         RsTemp.MoveNext
      Loop
   End If
   
   If ProcSaveExcel = True Then
      Process = True
   End If
   
   Exit Function
   
ErrHandle:
   If Err.Number <> 0 Then
      MsgBox "執行失敗：" & Err.Description, vbCritical
   End If
End Function

Private Function ProcSaveExcel() As Boolean
Dim xlsPoint As New Excel.Application
Dim WksPoint As New Worksheet
Dim bolOpenxlsPoint As Boolean
Dim intB1 As Integer, intB2 As Integer
Dim xRow As Integer, tmpArr As Variant

   strSql = "select R001,R002,R003,R004,R005,R006,R007,R008,R009,R010,R011,R012,R013 from rdatafactory where formname='" & Me.Name & "' and id='" & strUserNum & "' and seqno = '" & m_ESeq & "' order by rowseq "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      InsertQueryLog (RsTemp.RecordCount)
      '-------預設Excel
      xlsPoint.SheetsInNewWorkbook = 1
      xlsPoint.Workbooks.add
      xlsPoint.Application.Visible = False
      xlsPoint.Worksheets(1).Name = Mid(txtDate(0), 4, 4) & "-" & Mid(txtDate(1), 4, 4)
      Set WksPoint = xlsPoint.Worksheets(1)
      bolOpenxlsPoint = True
      
      ReDim tmpArr(1 To 13)
      xRow = xRow + 1
      For intI = 0 To 12
         If intI = 0 Then
            WksPoint.Range(Chr(65 + intI) & ":" & Chr(65 + intI)).ColumnWidth = 10
         Else
            WksPoint.Range(Chr(65 + intI) & ":" & Chr(65 + intI)).ColumnWidth = 13
         End If
         WksPoint.Range(Chr(65 + intI) & ":" & Chr(65 + intI)).Font.Size = 12
         WksPoint.Range(Chr(65 + intI) & ":" & Chr(65 + intI)).Font.Name = "標楷體"
      Next intI
      WksPoint.Range("B" & xRow).Value = "收文"
      WksPoint.Range("B" & xRow & ":C" & xRow).MergeCells = True
      WksPoint.Range("B" & xRow & ":C" & xRow).HorizontalAlignment = xlCenter
      WksPoint.Range("B" & xRow & ":C" & xRow).VerticalAlignment = xlCenter
      WksPoint.Range("B" & xRow & ":C" & xRow).Interior.ColorIndex = 22 '底色=粉紅
      
      WksPoint.Range("D" & xRow).Value = "齊備"
      WksPoint.Range("D" & xRow & ":E" & xRow).MergeCells = True
      WksPoint.Range("D" & xRow & ":E" & xRow).HorizontalAlignment = xlCenter
      WksPoint.Range("D" & xRow & ":E" & xRow).VerticalAlignment = xlCenter
      WksPoint.Range("D" & xRow & ":E" & xRow).Interior.ColorIndex = 36 '底色=黃色
      
      WksPoint.Range("F" & xRow).Value = "會稿"
      WksPoint.Range("F" & xRow & ":G" & xRow).MergeCells = True
      WksPoint.Range("F" & xRow & ":G" & xRow).HorizontalAlignment = xlCenter
      WksPoint.Range("F" & xRow & ":G" & xRow).VerticalAlignment = xlCenter
      WksPoint.Range("F" & xRow & ":G" & xRow).Interior.ColorIndex = 22 '底色=粉紅
      
      WksPoint.Range("H" & xRow).Value = "待辦(" & Mid(strSrvDate(2), 4, 4) & ")"
      WksPoint.Range("H" & xRow & ":K" & xRow).MergeCells = True
      WksPoint.Range("H" & xRow & ":K" & xRow).HorizontalAlignment = xlCenter
      WksPoint.Range("H" & xRow & ":K" & xRow).VerticalAlignment = xlCenter
      WksPoint.Range("H" & xRow & ":K" & xRow).Interior.ColorIndex = 36 '底色=黃色
      
      WksPoint.Range("L" & xRow).Value = "發文"
      WksPoint.Range("L" & xRow & ":M" & xRow).MergeCells = True
      WksPoint.Range("L" & xRow & ":M" & xRow).HorizontalAlignment = xlCenter
      WksPoint.Range("L" & xRow & ":M" & xRow).VerticalAlignment = xlCenter
      WksPoint.Range("L" & xRow & ":M" & xRow).Interior.ColorIndex = 22 '底色=粉紅
      
      xRow = xRow + 1
      WksPoint.Range(xRow & ":" & xRow).RowHeight = 60
      WksPoint.Range(xRow & ":" & xRow).HorizontalAlignment = xlCenter
      WksPoint.Range(xRow & ":" & xRow).VerticalAlignment = xlCenter
      For intI = 1 To 12
         If intI >= 7 And intI <= 10 Then
            Select Case intI
               Case 7
                  WksPoint.Range(Chr(65 + intI) & xRow).Value = "新案" & vbCrLf & "(僅工程師)"
               Case 8
                  WksPoint.Range(Chr(65 + intI) & xRow).Value = "新案" & vbCrLf & "(" & Mid(txtDate(3), 4, 2) & "月平均)" & vbCrLf & "(僅工程師)"
               Case 9
                  WksPoint.Range(Chr(65 + intI) & xRow).Value = "非新案" & vbCrLf & "(僅工程師)"
               Case 10
                  WksPoint.Range(Chr(65 + intI) & xRow).Value = "非新案" & vbCrLf & "(" & Mid(txtDate(3), 4, 2) & "月平均)" & vbCrLf & "(僅工程師)"
            End Select
         Else
            If intI Mod 2 = 1 Then
               WksPoint.Range(Chr(65 + intI) & xRow).Value = Mid(txtDate(0), 4, 2) & "/" & Mid(txtDate(0), 6, 2) & "-" & Mid(txtDate(1), 4, 2) & "/" & Mid(txtDate(1), 6, 2)
            Else
               WksPoint.Range(Chr(65 + intI) & xRow).Value = Mid(txtDate(2), 4, 2) & "/" & Mid(txtDate(2), 6, 2) & "-" & Mid(txtDate(3), 4, 2) & "/" & Mid(txtDate(3), 6, 2)
            End If
         End If
      Next intI
      
      xRow = xRow + 1
      intB1 = xRow
      '-------預設Excel
      RsTemp.MoveFirst
      Do While Not RsTemp.EOF
         For intI = 1 To UBound(tmpArr)
            tmpArr(intI) = "" & RsTemp.Fields("R" & Format(intI, "000"))
         Next intI
         WksPoint.Range(Chr(65) & xRow & ":" & Chr(65 + UBound(tmpArr) - 1) & xRow).Value = tmpArr
         WksPoint.Range(Chr(65 + 1) & xRow & ":" & Chr(65 + UBound(tmpArr) - 1) & xRow).HorizontalAlignment = xlRight
         xRow = xRow + 1
         RsTemp.MoveNext
      Loop
      
      intB2 = xRow - 1
      WksPoint.Range(Chr(65) & xRow).Value = "總數"
      For intI = 2 To UBound(tmpArr)
         WksPoint.Range(Chr(65 + intI - 1) & xRow).Formula = "=SUM(" & Chr(65 + intI - 1) & intB1 & ":" & Chr(65 + intI - 1) & intB2 & ")"
         WksPoint.Range(Chr(65 + intI - 1) & xRow).HorizontalAlignment = xlRight
      Next intI
      
      xRow = xRow + 8
      WksPoint.Range(Chr(65) & xRow).Value = "統計標準："
      xRow = xRow + 1
      WksPoint.Range(Chr(65) & xRow).Value = "*收文發文新案Y，含分割改請"
      xRow = xRow + 1
      WksPoint.Range(Chr(65) & xRow).Value = "*P僅查台灣案000，"
      xRow = xRow + 1
      WksPoint.Range(Chr(65) & xRow).Value = "*收發文看(專利-->內專-->統計報表-->收文統計表/發文統計表)，"
      xRow = xRow + 1
      WksPoint.Range(Chr(65) & xRow).Value = "*齊備/會稿之承辦人部門下P10-P11(暫不下)，新案Y(101-102、301-302、307、CFP-113、118)"
      xRow = xRow + 1
      WksPoint.Range(Chr(65) & xRow).Value = "*核駁已收文未發文：P：1002、1202、1227、1221、1810、1802、1807，CFP：1002、1006、1206、1209"
      
      xlsPoint.Sheets(1).Select '選擇工作表
   
      '判斷版本
      If Val(xlsPoint.Version) < 12 Then
           xlsPoint.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=-4143
      Else
           xlsPoint.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=56
      End If
      xlsPoint.Workbooks.Close
      xlsPoint.Quit
   Else
      InsertQueryLog (0)
   End If
   ProcSaveExcel = True
   Exit Function
   
ErrHnd1:
    If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
    If bolOpenxlsPoint = True Then
        xlsPoint.Workbooks(1).Close xlDoNotSaveChanges
        xlsPoint.Quit
    End If
End Function
