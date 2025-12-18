VERSION 5.00
Begin VB.Form frm083007 
   BorderStyle     =   1  '單線固定
   Caption         =   "人員工作分析表"
   ClientHeight    =   2490
   ClientLeft      =   3045
   ClientTop       =   1680
   ClientWidth     =   4140
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   4140
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   5
      Left            =   2760
      MaxLength       =   7
      TabIndex        =   7
      Top             =   1575
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   4
      Left            =   1560
      MaxLength       =   7
      TabIndex        =   6
      Top             =   1575
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   3
      Left            =   2760
      MaxLength       =   6
      TabIndex        =   5
      Top             =   1185
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   2
      Left            =   1560
      MaxLength       =   6
      TabIndex        =   4
      Top             =   1185
      Width           =   735
   End
   Begin VB.OptionButton Opt1 
      Caption         =   "協辦人員："
      Height          =   300
      Index           =   1
      Left            =   360
      TabIndex        =   3
      Top             =   1185
      Width           =   1215
   End
   Begin VB.OptionButton Opt1 
      Caption         =   "承  辦  人："
      Height          =   300
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   792
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   0
      Left            =   1560
      MaxLength       =   6
      TabIndex        =   1
      Top             =   792
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   1
      Left            =   2760
      MaxLength       =   6
      TabIndex        =   2
      Top             =   792
      Width           =   735
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   400
      Left            =   2256
      TabIndex        =   8
      Top             =   120
      Width           =   800
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   3084
      TabIndex        =   9
      Top             =   120
      Width           =   800
   End
   Begin VB.Label lblNote 
      Caption         =   "此報表尚需法律所重新調整！"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   90
      TabIndex        =   11
      Top             =   150
      Width           =   1845
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   2400
      X2              =   2640
      Y1              =   1695
      Y2              =   1695
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   2400
      X2              =   2640
      Y1              =   1305
      Y2              =   1305
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收文日期："
      Height          =   180
      Index           =   1
      Left            =   600
      TabIndex        =   10
      Top             =   1575
      Width           =   900
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   2400
      X2              =   2640
      Y1              =   912
      Y2              =   912
   End
End
Attribute VB_Name = "frm083007"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/08/31 法務系統的工作點數分配功能先上線(110/9/1)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/3 日期欄已修改
'Memo by Lydia 2015/06/12 隱藏
Option Explicit

Dim SDay As String, EDay As String, PLeft(0 To 10) As Integer
Dim m_print As Integer
Dim blnClkSure As Boolean '判斷是否按下確定按鈕
Dim strPointSql As String 'Add By Sindy 2010/5/4

Private Sub cmdBack_Click()
   Unload Me
End Sub

Private Sub cmdPrint_Click()
   blnClkSure = False
   
   If Me.Text1(4).Text = "" Or Me.Text1(5).Text = "" Then
      MsgBox "收文日期不可空白!!!", vbExclamation + vbOKOnly
      blnClkSure = True
      If Me.Text1(4).Text = "" Then
         Me.Text1(4).SetFocus
         Text1_GotFocus 4
      Else
         Me.Text1(5).SetFocus
         Text1_GotFocus 5
      End If
      Exit Sub
   End If
   
   If Me.Opt1(0).Value Then
      If Me.Text1(0).Text <> "" And Me.Text1(1).Text <> "" Then
         If Me.Text1(0).Text > Me.Text1(1).Text Then
            MsgBox "承辦人範圍輸入錯誤!!!", vbExclamation + vbOKOnly
            blnClkSure = True
            Me.Text1(0).SetFocus
            Text1_GotFocus 0
            Exit Sub
         End If
      End If
   Else
      If Me.Text1(2).Text <> "" And Me.Text1(3).Text <> "" Then
         If Me.Text1(2).Text > Me.Text1(3).Text Then
           'Modified by Lydia 2015/10/05
           'MsgBox "法務人員範圍輸入錯誤!!!", vbExclamation + vbOKOnly
            MsgBox "協辦人員範圍輸入錯誤!!!", vbExclamation + vbOKOnly
            blnClkSure = True
            Me.Text1(2).SetFocus
            Text1_GotFocus 2
            Exit Sub
         End If
      End If
   End If
   
   If PUB_CheckKeyInDate(Me.Text1(4)) = -1 Then
      Me.Text1(4).SetFocus
      Text1_GotFocus 4
      Exit Sub
   End If
   If PUB_CheckKeyInDate(Me.Text1(5)) = -1 Then
      Me.Text1(5).SetFocus
      Text1_GotFocus 5
      Exit Sub
   End If
   If Me.Text1(4).Text <> "" And Me.Text1(5).Text <> "" Then
      If Val(Me.Text1(4).Text) > Val(Me.Text1(5).Text) Then
         MsgBox "收文日期範圍輸入錯誤!!!", vbExclamation + vbOKOnly
         blnClkSure = True
         Me.Text1(4).SetFocus
         Text1_GotFocus 4
         Exit Sub
      End If
   End If
         
   m_print = 0
   Screen.MousePointer = 11
   GetPrintLeft
   PrintCase
   Screen.MousePointer = 0
   If m_print = 0 Then
      MsgBox "列印結束!", vbInformation
   End If
End Sub

Private Sub PrintCase()
Dim i As Integer, St As String, Page As Integer, iPrint As Integer
Dim TmpArea As String, Qty As String
'Dim Wo As DAO.Workspace, Db As DAO.Database, Rc As DAO.Recordset 'Remove by Lydia 2020/04/10
Dim strSql As String, strCP10 As String
'Added by Lydia 2020/04/10 暫存檔的序號
Dim mESeqNo As String
Dim xRows As Integer
Dim rsQuery As New ADODB.Recordset
 
 If Me.Tag = 0 Then
    strSql = "'L','LA'"
 ElseIf Me.Tag = 1 Then
    'Modify By Sindy 2009/07/24 增加LIN系統類別
    strSql = "'FCL','CFL','LIN'"
 End If
'On Error GoTo ErrHand
   If Opt1(0).Value = True Then
      St = "CP14"
   Else
      St = "CP29"
   End If
   'Modified by Lydia 2020/04/10 改用暫存檔Rdatafactory
   'If CreateDatabase = False Then
   '   MsgBox "無法建立暫存區，列印失敗 !", vbInformation
   '   m_print = 1
   '   Exit Sub
   'End If
   intI = 1
   Qty = "select '合計','0','0','0','0','合計' from dual"
   Set RsTemp = ClsLawReadRstMsg(intI, Qty)
   Set rsQuery = PUB_CreateRecordset(RsTemp, , , , Me.Name, mESeqNo)
   xRows = xRows + 1
   rsQuery.Close
   'end 2020/04/10
   
   'Modified by Lydia 2015/06/02 一致顯示-工作分配點數的人員
   'strExc(0) = "SELECT DISTINCT " & St & ",DECODE(LENGTH(ST02),NULL,ST01,ST02) FROM CASEPROGRESS,STAFF WHERE " & _
      "CP01 IN (" & strSql & ") AND (CP57 IS NULL AND CP26 IS NULL) AND CP09<'C' AND " & St & "=ST01(+)" & strGetcdnSQL & " ORDER BY " & St
      '去掉同部門判斷"AND (substr(s1.st15,1,2)=substr(s2.st15,1,2) or s2.st15 is null) "
     strExc(0) = " select nvl(a1n04,CP14) A01,nvl(s1.st02,s2.st02) st02 FROM CASEPROGRESS,acc1n0,STAFF s1,staff s2 " & _
               "WHERE CP01 IN (" & strSql & ") AND (CP57 IS NULL AND CP26 IS NULL) AND CP09<'C' " & strGetcdnSQL & strPointSql & _
               "and cp09=a1n03(+) and a1n02(+)<>'1' and a1n05(+)>0 and a1n04=s1.st01(+) and CP14=s2.st01(+) " & _
               "group by nvl(a1n04,CP14),nvl(s1.st02,s2.st02) order by 1 "
   If Opt1(0).Value = False Then strExc(0) = Replace(strExc(0), "CP14", "CP29")
     
   If RsTemp.State = adStateOpen Then RsTemp.Close
   RsTemp.Open strExc(0), cnnConnection
   If RsTemp.EOF And RsTemp.BOF Then
      MsgBox "資料庫內無資料 !", vbInformation
      m_print = 1
      Exit Sub
   End If
   'Modified by Lydia 2020/04/10 改用暫存檔Rdatafactory
   'Set Wo = DBEngine.Workspaces(0)
   'Set Db = Wo.OpenDatabase(App.path & "\Case.mdb", False, False, ";PWD=taie")
   'Qty = "DELETE FROM TEMP"
   'Db.Execute Qty
   'With RsTemp
   '   .MoveFirst
   '   Do While Not .EOF
   '      If IsNull(.Fields(0)) = True Then
   '         Qty = "INSERT INTO TEMP (TMP01,TMP02,TMP03,TMP04,TMP05,TMP06) VALUES ('空白','0','0','0','0',NULL)"
   '      Else
   '         Qty = "INSERT INTO TEMP (TMP01,TMP02,TMP03,TMP04,TMP05,TMP06) VALUES ('" & .Fields(0) & "','0','0','0','0','" & .Fields(1) & "')"
   '      End If
   '      Db.Execute Qty
   '      .MoveNext
   '   Loop
   '   Qty = "INSERT INTO TEMP (TMP01,TMP02,TMP03,TMP04,TMP05,TMP06) VALUES ('合計','0','0','0','0','合計')"
   '   Db.Execute Qty
   '   .Close
   'End With
   Qty = "update rdatafactory set rowseq='1' where formname='" & Me.Name & "' and id='" & strUserNum & "' and seqno='" & mESeqNo & "' "
   cnnConnection.Execute Qty
   With RsTemp
        .MoveFirst
        Do While Not .EOF
              xRows = xRows + 1
              If "" & .Fields("A01") = "" Then
                 Qty = "insert into rdatafactory(formname,id,seqno,rowseq,r001,r002,r003,r004,r005,r006) " & _
                           "values ('" & Me.Name & "', '" & strUserNum & "', '" & mESeqNo & "', " & xRows & ",'空白','0','0','0','0',NULL)"
              Else
                 '分配人員非法律所人員請姓名後加*
                 strExc(1) = "" & .Fields("st02")
                 If PUB_ChkLCompStaff(.Fields("A01")) = False Then
                     strExc(1) = strExc(1) & "*"
                 End If
                 Qty = "insert into rdatafactory(formname,id,seqno,rowseq,r001,r002,r003,r004,r005,r006) " & _
                           "values ('" & Me.Name & "', '" & strUserNum & "', '" & mESeqNo & "', " & xRows & ",'" & .Fields("A01") & "','0','0','0','0','" & strExc(1) & "')"
              End If
              cnnConnection.Execute Qty
              .MoveNext
        Loop
   End With
   If RsTemp.State = adStateOpen Then RsTemp.Close
   'end 2020/04/10
   
   strCP10 = "'2101','2102','2103','2104'," & _
         "'2105','2106','2108','2121','2122','2123','2124','2201','2202','2203','2204'," & _
         "'2205','2206','2207','2109','2137','2139','2140','1111'," & _
         "'1112','1113','1115','1122','1124','1132','1133','1134','1301','1311','1312'," & _
         "'1313','1314','1304'"
'委任訴訟------------------------
'Modified by Lydia 2015/06/02 一致使用-工作分配點數
   'Modify By Sindy 2010/5/11
'   If opt1(1).Value = True Then '協辦人員-點數改抓請款單
'      strExc(0) = "SELECT CP29,COUNT(*),SUM((nvl(a1k11,0)-nvl(a1k09,0))/1000) FROM CASEPROGRESS,acc1k0 WHERE " & _
'         "CP01 IN (" & strSql & ") AND CP10 IN (" & strCP10 & ") AND (CP57 IS NULL AND CP26 IS NULL) AND CP09<'C' AND a1k01(+)=cp60 " & strGetcdnSQL & _
'         " GROUP BY CP29"
'   Else
      '承辦人-點數改抓點數分配檔
      'Modify by Morgan 2011/6/1 若有建點數分配資料時點數改分配點數(目前L會有分配) cp18->nvl(a0n03/1000,cp18)
      'Modified by Lydia 2015/06/02 +工作分配點數
      'strExc(0) = "SELECT CP14,COUNT(*),SUM(a1n05) FROM ( " & _
         "SELECT CP14,decode(substr(cp60,1,1),'X',decode(a1k25,null,a1n05,''),nvl(a0n03/1000,cp18)) as a1n05 FROM CASEPROGRESS,acc1n0,acc1k0 ,acc0n0 where a0n02(+)=cp09 and " & _
         "CP01 IN (" & strSql & ") AND CP10 IN (" & strCP10 & ") AND (CP57 IS NULL AND CP26 IS NULL) AND CP09<'C' " & _
         "AND a1n01(+)=cp60 AND a1n02(+)='2' AND a1n03(+)=cp09 AND a1n04(+)=CP14 AND a1k01(+)=cp60 " & strGetcdnSQL & _
         " Union All " & _
         "SELECT a1n04 as CP14,decode(substr(cp60,1,1),'X',decode(a1k25,null,a1n05,''),nvl(a0n03/1000,cp18)) as a1n05 FROM CASEPROGRESS,acc1n0,acc1k0,staff a,staff b ,acc0n0 where a0n02(+)=cp09 and " & _
         "CP01 IN (" & strSql & ") AND CP10 IN (" & strCP10 & ") AND (CP57 IS NULL AND CP26 IS NULL) AND CP09<'C' " & _
         "AND a1n01(+)=cp60 AND a1n02(+)='2' AND a1n03(+)=cp09 AND a1n04(+)<>CP14 AND a1n05>0 AND a1k01(+)=cp60 AND a1n04=a.st01(+) AND CP14=b.st01(+) AND substr(a.st15,1,2)=substr(b.st15,1,2) " & strPointSql & _
         ") GROUP BY CP14 ORDER BY CP14 "
'   End If
      'Modified by Lydia 2015/06/09 承辦人為null,卻有分配點數(ex.AA4002302),分2階段抓資料
      '累計件數
        strExc(4) = "select cp09,1 as or1,CP14 as stno, 0.000 as n05 " & _
                    "FROM CASEPROGRESS where CP01 IN (" & strSql & ") AND CP10 IN (" & strCP10 & ") " & _
                    "AND (CP57 IS NULL AND CP26 IS NULL) AND CP09<'C' " & strGetcdnSQL & _
                    " group by cp09,CP14 "
       
      '累計點數
        'Memo by Lydia 2015/06/11 因為法務部還處於整合狀態,有部門判斷會造成不同結果
        '去掉同部門判斷"AND (substr(c.st15,1,2)=substr(b.st15,1,2) or b.st15 is null) and (substr(c.st15,1,2)=substr(a.st15,1,2) or a.st15 is null) "
        strExc(5) = "SELECT cp09,0 as or1,decode(p3.a1n04,null,decode(p2.a1n04,null,CP14, p2.a1n04), p3.a1n04) as stno," & _
                    "decode(p3.a1n01, Null, decode(p2.a1n01, Null, cp18, p2.a1n05), p3.a1n05) As n05 " & _
                    "FROM CASEPROGRESS,acc1n0 p2,acc1n0 p3,staff a,staff b,staff c where " & _
                    "CP01 IN (" & strSql & ") AND CP10 IN (" & strCP10 & ") AND (CP57 IS NULL AND CP26 IS NULL) AND CP09<'C' " & strGetcdnSQL & _
                    "AND p3.a1n02(+)='3' AND p3.a1n03(+)=cp09  AND p2.a1n02(+)='2' AND p2.a1n03(+)=cp09 " & _
                    "AND p3.a1n04=a.st01(+) AND p2.a1n04=b.st01(+) AND CP14=c.st01(+) "

        strExc(6) = strExc(4) + " UNION ALL " + strExc(5)
        strExc(0) = "select stno,sum(or1) cnt,sum(n05) s05 from (" & strExc(6) & ") where 1=1 " & IIf(InStr(strPointSql, "CP14") > 0, Replace(strPointSql, "CP14", "stno"), Replace(strPointSql, "CP29", "stno")) & " group by stno order by stno "
   If Opt1(0).Value = False Then strExc(0) = Replace(strExc(0), "CP14", "CP29")
'end 2015/06/02

   RsTemp.Open strExc(0), cnnConnection
   If Not RsTemp.EOF Then
      With RsTemp
         .MoveFirst
         Do While Not .EOF
            'Modified by Lydia 2020/04/10 改用暫存檔Rdatafactory
'            If IsNull(.Fields(0)) = True Then
'               Qty = "UPDATE TEMP SET TMP04="
'                      If IsNull(.Fields(1)) = False Then
'                         Qty = Qty & "'" & .Fields(1) & "'"
'                      Else
'                         Qty = Qty & "'0'"
'                      End If
'
'                      Qty = Qty & ",TMP05="
'                      If IsNull(.Fields(2)) = False Then
'                         Qty = Qty & "'" & .Fields(2) & "'"
'                      Else
'                         Qty = Qty & "'0'"
'                      End If
'                      Qty = Qty & " WHERE TMP01='空白'"
'            Else
'               Qty = "UPDATE TEMP SET TMP04="
'                      If IsNull(.Fields(1)) = False Then
'                         Qty = Qty & "'" & .Fields(1) & "'"
'                      Else
'                         Qty = Qty & "'0'"
'                      End If
'                      Qty = Qty & ",TMP05="
'                      If IsNull(.Fields(2)) = False Then
'                         Qty = Qty & "'" & .Fields(2) & "'"
'                      Else
'                         Qty = Qty & "'0'"
'                      End If
'
'                     Qty = Qty & " WHERE TMP01='" & .Fields(0) & "'"
'            End If
'            Db.Execute Qty
            Qty = "update rdatafactory set R004=" & CNULL(.Fields("cnt"), True) & ", R005=" & CNULL(.Fields("s05"), True) & _
                      " where formname='" & Me.Name & "' and id='" & strUserNum & "' and seqno='" & mESeqNo & "' "
            If "" & .Fields("stno") = "" Then
                 Qty = Qty & " and r001='空白' "
            Else
                 Qty = Qty & " and r001='" & .Fields("stno") & "' "
            End If
            cnnConnection.Execute Qty
            'end 2020/04/10
            .MoveNext
         Loop
      End With
   End If
   RsTemp.Close
   
'雜文------------------------
'Modified by Lydia 2015/06/02 一致使用-工作分配點數
'   'Modify By Sindy 2010/5/11
'   If opt1(1).Value = True Then '協辦人員-點數改抓請款單
'      strExc(0) = "SELECT CP29,COUNT(*),SUM((nvl(a1k11,0)-nvl(a1k09,0))/1000) FROM CASEPROGRESS,acc1k0 WHERE " & _
'         "CP01 IN (" & strSql & ") AND NOT (CP10 IN (" & strCP10 & ")) AND (CP57 IS NULL AND CP26 IS NULL) AND CP09<'C' AND a1k01(+)=cp60 " & strGetcdnSQL & _
'         " GROUP BY CP29"
'   Else
      '承辦人-點數改抓點數分配檔
      'Modify by Morgan 2011/6/1 若有建點數分配資料時點數改分配點數(目前L會有分配) cp18->nvl(a0n03/1000,cp18)
      'Modified by Lydia 2015/06/02 +工作分配點數
      'strExc(0) = "SELECT CP14,COUNT(*),SUM(a1n05) FROM ( " & _
         "SELECT CP14,decode(substr(cp60,1,1),'X',decode(a1k25,null,a1n05,''),nvl(a0n03/1000,cp18)) as a1n05 FROM CASEPROGRESS,acc1n0,acc1k0 ,acc0n0 where a0n02(+)=cp09 and " & _
         "CP01 IN (" & strSql & ") AND NOT (CP10 IN (" & strCP10 & ")) AND (CP57 IS NULL AND CP26 IS NULL) AND CP09<'C' " & _
         "AND a1n01(+)=cp60 AND a1n02(+)='2' AND a1n03(+)=cp09 AND a1n04(+)=CP14 AND a1k01(+)=cp60 " & strGetcdnSQL & _
         " Union All " & _
         "SELECT a1n04 as CP14,decode(substr(cp60,1,1),'X',decode(a1k25,null,a1n05,''),nvl(a0n03/1000,cp18)) as a1n05 FROM CASEPROGRESS,acc1n0,acc1k0,staff a,staff b ,acc0n0 where a0n02(+)=cp09 and " & _
         "CP01 IN (" & strSql & ") AND NOT (CP10 IN (" & strCP10 & ")) AND (CP57 IS NULL AND CP26 IS NULL) AND CP09<'C' " & _
         "AND a1n01(+)=cp60 AND a1n02(+)='2' AND a1n03(+)=cp09 AND a1n04(+)<>CP14 AND a1n05>0 AND a1k01(+)=cp60 AND a1n04=a.st01(+) AND CP14=b.st01(+) AND substr(a.st15,1,2)=substr(b.st15,1,2) " & strPointSql & _
         ") GROUP BY CP14 ORDER BY CP14 "
              
'       '承辦人分配到的點數CP14=a1n04
'        strExc(4) = "SELECT cp09,1 as or1,decode(p1.a1n04,null,decode(p3.a1n04,null,CP14,p3.a1n04),p1.a1n04) as CP14," & _
'                    "decode(p1.a1n01,null,decode(p3.a1n01,null,cp18,p3.a1n05) ,p1.a1n05) as n05 " & _
'                    "FROM CASEPROGRESS,acc1n0 p1,acc1n0 p3,acc1k0 ,acc0n0 where a0n02(+)=cp09 and " & _
'                    "CP01 IN (" & strSql & ") AND NOT (CP10 IN (" & strCP10 & ")) AND (CP57 IS NULL AND CP26 IS NULL) AND CP09<'C' " & _
'                    "AND p1.a1n02(+)='3' AND p1.a1n03(+)=cp09 AND p1.a1n04(+)=CP14 " & _
'                    "AND p3.a1n01(+)=cp60 AND p3.a1n02(+)='2' AND p3.a1n03(+)=cp09 AND p3.a1n04(+)=CP14 AND a1k01(+)=cp60 " & strGetcdnSQL
'       If opt1(1).Value = True Then '協辦人員只抓有分配
'         strExc(6) = strExc(4)
'       Else
'        '非承辦人分配到的點數CP14<>a1n04
'         strExc(5) = "SELECT cp09,0 as or1,nvl(p2.a1n04,p4.a1n04) as CP14,decode(p2.a1n01,null,decode(p4.a1n01,null,cp18,p4.a1n05) ,p2.a1n05) as n05 " & _
'                     "FROM CASEPROGRESS,acc1n0 p2,acc1n0 p4,acc1k0 ,acc0n0,staff a,staff b,staff c where a0n02(+)=cp09 and " & _
'                     "CP01 IN (" & strSql & ") AND NOT (CP10 IN (" & strCP10 & ")) AND (CP57 IS NULL AND CP26 IS NULL) AND CP09<'C' " & _
'                     "AND p2.a1n02(+)='3' AND p2.a1n03(+)=cp09 AND p2.a1n04(+)<>CP14 AND p2.a1n05(+)>0 " & _
'                     "AND p4.a1n01(+)=cp60 AND p4.a1n02(+)='2' AND p4.a1n03(+)=cp09 AND p4.a1n04(+)<>CP14 AND p4.a1n05(+)>0 AND a1k01(+)=cp60 " & _
'                     "AND p2.a1n04=a.st01(+) AND p4.a1n04=b.st01(+) AND CP14=c.st01(+) " & _
'                     "AND (substr(c.st15,1,2)=substr(b.st15,1,2) or b.st15 is null) and (substr(c.st15,1,2)=substr(a.st15,1,2) or a.st15 is null) " & strPointSql
'         strExc(6) = strExc(4) + " UNION ALL " + strExc(5)
'       End If
      'Modified by Lydia 2015/06/09 承辦人為null,卻有分配點數(ex.AA4002302),分2階段抓資料
      '累計件數
        strExc(4) = "select cp09,1 as or1,CP14 as stno, 0.000 as n05 " & _
                    "FROM CASEPROGRESS where CP01 IN (" & strSql & ") AND NOT(CP10 IN (" & strCP10 & ")) " & _
                    "AND (CP57 IS NULL AND CP26 IS NULL) AND CP09<'C' " & strGetcdnSQL & _
                    " group by cp09,CP14 "
       
      '累計點數
        'Memo by Lydia 2015/06/11 因為法務部還處於整合狀態,有部門判斷會造成不同結果
        '去掉同部門判斷"AND (substr(c.st15,1,2)=substr(b.st15,1,2) or b.st15 is null) and (substr(c.st15,1,2)=substr(a.st15,1,2) or a.st15 is null) "
        strExc(5) = "SELECT cp09,0 as or1,decode(p3.a1n04,null,decode(p2.a1n04,null,CP14, p2.a1n04), p3.a1n04) as stno," & _
                    "decode(p3.a1n01, Null, decode(p2.a1n01, Null, cp18, p2.a1n05), p3.a1n05) As n05 " & _
                    "FROM CASEPROGRESS,acc1n0 p2,acc1n0 p3,staff a,staff b,staff c where " & _
                    "CP01 IN (" & strSql & ") AND NOT(CP10 IN (" & strCP10 & ")) AND (CP57 IS NULL AND CP26 IS NULL) AND CP09<'C' " & strGetcdnSQL & _
                    "AND p3.a1n02(+)='3' AND p3.a1n03(+)=cp09  AND p2.a1n02(+)='2' AND p2.a1n03(+)=cp09 " & _
                    "AND p3.a1n04=a.st01(+) AND p2.a1n04=b.st01(+) AND CP14=c.st01(+) "
        strExc(6) = strExc(4) + " UNION ALL " + strExc(5)
        strExc(0) = "select stno,sum(or1) cnt,sum(n05) s05 from (" & strExc(6) & ") where 1=1 " & IIf(InStr(strPointSql, "CP14") > 0, Replace(strPointSql, "CP14", "stno"), Replace(strPointSql, "CP29", "stno")) & " group by stno order by stno "
'   End If
   If Opt1(0).Value = False Then strExc(0) = Replace(strExc(0), "CP14", "CP29")
'end 2015/06/02
   RsTemp.Open strExc(0), cnnConnection
   If Not RsTemp.EOF Then
      With RsTemp
         .MoveFirst
         Do While Not .EOF
            'Modified by Lydia 2020/04/10 改用暫存檔Rdatafactory
'            If IsNull(.Fields(0)) = True Then
'               Qty = "UPDATE TEMP SET TMP02="
'                     If Not IsNull(.Fields(1)) Then
'                        Qty = Qty & "'" & .Fields(1) & "'"
'                     Else
'                        Qty = Qty & "'0'"
'                     End If
'                     Qty = Qty & ",TMP03="
'                     If Not IsNull(.Fields(2)) Then
'                        Qty = Qty & "'" & .Fields(2) & "'"
'                     Else
'                        Qty = Qty & "'0'"
'                     End If
'                     Qty = Qty & " WHERE TMP01='空白'"
'            Else
'               Qty = "UPDATE TEMP SET TMP02="
'                     If Not IsNull(.Fields(1)) Then
'                        Qty = Qty & "'" & .Fields(1) & "'"
'                     Else
'                        Qty = Qty & "'0'"
'                     End If
'                     Qty = Qty & ",TMP03="
'
'                     If Not IsNull(.Fields(2)) Then
'                        Qty = Qty & "'" & .Fields(2) & "'"
'                     Else
'                        Qty = Qty & "'0'"
'                     End If
'                     Qty = Qty & " WHERE TMP01='" & .Fields(0) & "'"
'            End If
'            Db.Execute Qty
            Qty = "update rdatafactory set R002=" & CNULL(.Fields("cnt"), True) & ", R003=" & CNULL(.Fields("s05"), True) & _
                      " where formname='" & Me.Name & "' and id='" & strUserNum & "' and seqno='" & mESeqNo & "' "
            If "" & .Fields("stno") = "" Then
                 Qty = Qty & " and r001='空白' "
            Else
                 Qty = Qty & " and r001='" & .Fields("stno") & "' "
            End If
            cnnConnection.Execute Qty
            'end 2020/04/10
            .MoveNext
         Loop
      End With
   End If
   RsTemp.Close
   
   'Modified by Lydia 2020/04/10 改用暫存檔Rdatafactory
   'Qty = "SELECT SUM(VAL(iif(isnull(TMP02),'0',TMP02))),SUM(VAL(IIF(isnull(TMP03),'0',TMP03))),SUM(VAL(IIF(isnull(TMP04),'0',TMP04))),SUM(VAL(IIF(isnull(TMP05),'0',TMP05))) FROM TEMP WHERE TMP01<>'合計'"
   'Set Rc = Db.OpenRecordset(Qty)
   'Qty = "UPDATE TEMP SET TMP02='" & Rc.Fields(0) & "',TMP03='" & Rc.Fields(1) & _
   '   "',TMP04='" & Rc.Fields(2) & "',TMP05='" & Rc.Fields(3) & "' WHERE TMP01='合計'"
   'Db.Execute Qty
   'Rc.Close
   cnnConnection.Execute "delete from rdatafactory where r001='合計' and formname='" & Me.Name & "' and id='" & strUserNum & "' and seqno='" & mESeqNo & "' "
   xRows = xRows + 1
   Qty = "insert into rdatafactory(formname,id,seqno,rowseq,r001,r002,r003,r004,r005,r006) " & _
             "SELECT '" & Me.Name & "', '" & strUserNum & "', '" & mESeqNo & "', " & xRows & ",'合計',sum(r002) as r002, sum(r003) as r003, sum(r004) as r004, sum(r005) as r005, '合計' " & _
             " FROM rdatafactory where formname='" & Me.Name & "' and id='" & strUserNum & "' and seqno='" & mESeqNo & "' "
   'end 2020/04/10
   
   i = 1
   Page = 1
   CaseTitle TmpArea, 1
   iPrint = 3100
   'Modified by Lydia 2020/04/10 改用暫存檔Rdatafactory
   'Qty = "SELECT TMP01,TMP02,TMP03,TMP04,TMP05,TMP06 FROM TEMP"
   'Set Rc = Db.OpenRecordset(Qty)
   'If Not Rc.EOF Then
   '   With Rc
   Qty = "select r001,r002,r003,r004,r005,r006 from rdatafactory WHERE  formname='" & Me.Name & "' and id='" & strUserNum & "' and seqno='" & mESeqNo & "' order by rowseq "
   intI = 1
   Set rsQuery = ClsLawReadRstMsg(intI, Qty)
   If intI = 1 Then
      With rsQuery
   'end 2020/04/10
            .MoveFirst
            Do While Not .EOF
               Printer.CurrentX = PLeft(0):      Printer.CurrentY = iPrint
               If IsNull(.Fields(5)) = False Then
                  Printer.Print .Fields(5)
               Else
                  Printer.Print ""
               End If
            
               Printer.CurrentX = CInt(PLeft(1)) + 1000 - (Printer.TextWidth(Format(CheckStr(.Fields(1)), "0.0")))
               Printer.CurrentY = iPrint
               Printer.Print Format(CheckStr(.Fields(1)), "0")
               Printer.CurrentX = CInt(PLeft(2)) + 1000 - (Printer.TextWidth(Format(CheckStr(.Fields(2)), "0.0")))
               Printer.CurrentY = iPrint
               Printer.Print Format(CheckStr(.Fields(2)), "0.0")
               Printer.CurrentX = CInt(PLeft(3)) + 1000 - (Printer.TextWidth(Format(CheckStr(.Fields(3)), "0.0")))
               Printer.CurrentY = iPrint
               Printer.Print Format(CheckStr(.Fields(3)), "0")
               Printer.CurrentX = CInt(PLeft(4)) + 1000 - (Printer.TextWidth(Format(CheckStr(.Fields(4)), "0.0")))
               Printer.CurrentY = iPrint
               Printer.Print Format(CheckStr(.Fields(4)), "0.0")
               .MoveNext
               If Not .EOF Then
                  If (i Mod 35 = 0) Then
                     Printer.NewPage
                     Page = Page + 1
                     CaseTitle St, Page
                     iPrint = 3100
                     i = 0
                  End If
                  i = i + 1
                  iPrint = iPrint + 300
                  If .Fields(0) = "合計" Then
                     Printer.CurrentX = PLeft(0):      Printer.CurrentY = iPrint
                     Printer.Print String(130, "-")
                     i = i + 1
                     iPrint = iPrint + 300
                  End If
               End If
            Loop
      End With
   End If 'Added by Lydia 2020/04/10
   'Remove by Lydia 2020/04/10
   'End If
   'Rc.Close
   'end 2020/04/10
   
   'Modify By Sindy 2010/5/11
   iPrint = iPrint + 900
   Printer.CurrentX = PLeft(0):      Printer.CurrentY = iPrint
   If Opt1(1).Value = True Then
      Printer.Print "PS.點數為請款單點數，未扣除跨部門合作點數"
   Else
      Printer.Print "PS.不含非個人點數，有扣除專利處配合開庭分配點數。"
   End If
   iPrint = iPrint + 300
   '2010/5/11 End
   Printer.EndDoc
   Exit Sub
ErrHand:
   MsgBox Err.Description
End Sub

Private Sub GetPrintLeft()
   PLeft(0) = 500:    PLeft(1) = 2000
   PLeft(2) = 3300:   PLeft(3) = 5600
   PLeft(4) = 7000
End Sub

Private Sub CaseTitle(ByVal Area As String, ByVal Page As String)
 Dim i As Integer, St As String
   i = 500
   Printer.Orientation = 1
   Printer.Font.Size = 22
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.CurrentX = 4000:         Printer.CurrentY = i
   Printer.Print "人員工作分析表"
   Printer.Font.Underline = False
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   Printer.CurrentX = 3900:         Printer.CurrentY = i + 500
   Printer.Print "收文日期 : " & ChangeTStringToTDateString(Text1(4)) & _
      " - " & ChangeTStringToTDateString(Text1(5))
   Printer.Font.Bold = False
   Printer.CurrentX = 500:              Printer.CurrentY = i + 800
   Printer.Print "列印人 : " & strUserName
   Printer.CurrentX = 9000:            Printer.CurrentY = i + 800
   Printer.Print "列印日期 : " & ChangeTStringToTDateString(GetTaiwanTodayDate)
   Printer.CurrentX = 9000:            Printer.CurrentY = i + 1100
   Printer.Print "頁次 : " & Page
   
   'Printer.Font.Underline = True
   Printer.CurrentX = 2500:              Printer.CurrentY = i + 1400
   Printer.Print "雜          文"
   Printer.CurrentX = 6100:              Printer.CurrentY = i + 1400
   Printer.Print "委  任  訴  訟"
   'Printer.Font.Underline = False
   Printer.Line (2500, i + 1700)-(4200, i + 1700)
   Printer.Line (6100, i + 1700)-(7800, i + 1700)
   
   Printer.CurrentX = 500:              Printer.CurrentY = i + 1700
   Printer.Print String(130, "-")
   Printer.CurrentX = PLeft(0):         Printer.CurrentY = i + 2000
   If Opt1(0).Value = True Then
      St = "承辦人"
   Else
      'Modified by Lydia 2015/10/05
     ' St = "法務人員"
      St = "協辦人員"
   End If
   Printer.Print St
   Printer.CurrentX = PLeft(1):         Printer.CurrentY = i + 2000
   Printer.Print "承辦件數"
   Printer.CurrentX = PLeft(2):         Printer.CurrentY = i + 2000
   Printer.Print "合計點數"
   Printer.CurrentX = PLeft(3):         Printer.CurrentY = i + 2000
   Printer.Print "承辦件數"
   Printer.CurrentX = PLeft(4):         Printer.CurrentY = i + 2000
   Printer.Print "合計點數"
   Printer.CurrentX = 500:         Printer.CurrentY = i + 2300
   Printer.Print String(130, "-")
End Sub

Private Sub Form_Activate()
  Text1(0).SetFocus
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   Opt1(0).Value = True
End Sub
Private Function strGetcdnSQL() As String
   strExc(1) = ""
   strPointSql = "" 'Add By Sindy 2010/5/11
   If Opt1(0).Value = True Then
      If Text1(0).Text = "" And Text1(1).Text <> "" Then
         'Modified by Lydia 2015/06/09
         'strPointSql = strPointSql & " and a1n04<='" + Text1(1) + "' " 'Add By Sindy 2010/5/11
          strExc(1) = " and ((CP14<='" + Text1(1) + "') OR (CP14 IS NULL)) "
      ElseIf Text1(0).Text <> "" And Text1(1).Text <> "" Then
         'Modified by Lydia 2015/06/09
         'strPointSql = strPointSql & " and (a1n04 BETWEEN '" + Text1(0) + "' AND '" + Text1(1) + "') " 'Add By Sindy 2010/5/11
         strExc(1) = " and ((CP14 BETWEEN '" + Text1(0) + "' AND '" + Text1(1) + "') OR (CP14 IS NULL)) "
      End If
   Else
      If Text1(2).Text = "" And Text1(3).Text <> "" Then
         strExc(1) = " and ((CP29<='" + Text1(3) + "') OR (CP29 IS NULL)) "
      ElseIf Text1(2).Text <> "" And Text1(3).Text <> "" Then
         strExc(1) = " and ((CP29 BETWEEN '" + Text1(2) + "' AND '" + Text1(3) + "') OR (CP29 IS NULL)) "
      End If
   End If
   'Modified by Lydia 2015/06/09 分成２段SQL
   strPointSql = strExc(1)
   strExc(1) = ""
   If Text1(4).Text = "" And Text1(5).Text <> "" Then
      'Modified by Lydia 2015/06/09
      'strPointSql = strPointSql & " and CP05<='" + ChangeTStringToWString(Text1(5)) + "' " 'Add By Sindy 2010/5/11
      strExc(1) = strExc(1) + " and CP05<='" + ChangeTStringToWString(Text1(5)) + "' "
   ElseIf Text1(4).Text <> "" And Text1(5).Text <> "" Then
      'Modified by Lydia 2015/06/09
      'strPointSql = strPointSql & " and (CP05 BETWEEN '" + ChangeTStringToWString(Text1(4)) + "' AND '" + ChangeTStringToWString(Text1(5)) + "') " 'Add By Sindy 2010/5/11
      strExc(1) = strExc(1) + " and (CP05 BETWEEN '" + ChangeTStringToWString(Text1(4)) + "' AND '" + ChangeTStringToWString(Text1(5)) + "') "
   End If
   strGetcdnSQL = strExc(1)
End Function

Private Sub Form_Unload(Cancel As Integer)
   Set frm083007 = Nothing
End Sub

Private Sub Opt1_Click(Index As Integer)
 Dim i As Integer
   For i = 0 To 3
      Text1(i).Enabled = False
   Next
   If Index = 0 Then
      Text1(0).Enabled = True
      Text1(1).Enabled = True
   Else
      Text1(2).Enabled = True
      Text1(3).Enabled = True
   End If
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0, 1, 2, 3
         KeyAscii = UpperCase(KeyAscii)
      Case 4, 5
         KeyAscii = Pub_NumAscii(KeyAscii)
   End Select
End Sub

Private Sub Text1_LostFocus(Index As Integer)
   Select Case Index
      Case 1 '承辦人
            'If Len(Text1(Index - 1)) <> 0 Then
            '   If Left(Text1(Index - 1), 6) <> Left(Text1(Index), 6) Then
            '       MsgBox "承辦人前 6 碼必須相同", , "USER 輸入錯誤"
            '       Text1(Index - 1).SetFocus
            '       Exit Sub
            '   End If
            'End If
            'Add/Modify By Cheng 2002/09/09
            If blnClkSure = False Then
               If RunNick(Text1(Index - 1), Text1(Index)) Then
                  Text1(Index - 1).SetFocus
               End If
            Else
               blnClkSure = False
            End If
      Case 3 '協辦人員
            'If Len(Text1(Index - 1)) <> 0 Then
            '   If Left(Text1(Index - 1), 6) <> Left(Text1(Index), 6) Then
            '       MsgBox "協辦人員前 6 碼必須相同", , "USER 輸入錯誤"
            '       Text1(Index - 1).SetFocus
            '       Exit Sub
            '   End If
            'End If
            'Add/Modify By Cheng 2002/09/09
            If blnClkSure = False Then
               If RunNick(Text1(Index - 1), Text1(Index)) Then
                  Text1(Index - 1).SetFocus
               End If
            Else
               blnClkSure = False
            End If
      Case 5 '收文日
            'Add/Modify By Cheng 2002/09/09
            If blnClkSure = False Then
               If RunNick(Text1(Index - 1), Text1(Index)) Then
                  Text1(Index - 1).SetFocus
                  Exit Sub
               End If
            Else
               blnClkSure = False
            End If
      End Select
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
 Dim strTempName As String, i As Integer, t As Integer
   If Text1(Index) = "" Then Exit Sub
   Select Case Index
      Case 4, 5
         If CheckIsTaiwanDate(Text1(Index)) = False Then
            Cancel = True
            TextInverse Text1(Index)
            Exit Sub
         End If
   End Select
   If Cancel Then TextInverse Text1(Index)
End Sub
