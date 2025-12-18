VERSION 5.00
Begin VB.Form frm050307 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "收文簿"
   ClientHeight    =   2460
   ClientLeft      =   4185
   ClientTop       =   2970
   ClientWidth     =   3480
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   3480
   Begin VB.TextBox txtPA46 
      Height          =   285
      Left            =   1620
      TabIndex        =   5
      Top             =   1560
      Width           =   285
   End
   Begin VB.CheckBox Check1 
      Caption         =   "列印新案件承辦人明細表"
      Height          =   345
      Left            =   270
      TabIndex        =   6
      Top             =   1950
      Width           =   2355
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   996
      TabIndex        =   0
      Top             =   516
      Width           =   1905
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   996
      MaxLength       =   7
      TabIndex        =   1
      Top             =   840
      Width           =   870
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   2016
      MaxLength       =   7
      TabIndex        =   2
      Top             =   840
      Width           =   870
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   996
      MaxLength       =   4
      TabIndex        =   3
      Top             =   1200
      Width           =   870
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   2028
      MaxLength       =   4
      TabIndex        =   4
      Top             =   1200
      Width           =   870
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   1290
      TabIndex        =   7
      Top             =   20
      Width           =   800
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   2112
      TabIndex        =   8
      Top             =   20
      Width           =   800
   End
   Begin VB.Label Label1 
      Caption         =   "PCT進入國家階段：　（Y：國家階段）"
      Height          =   180
      Index           =   3
      Left            =   90
      TabIndex        =   12
      Top             =   1620
      Width           =   3180
   End
   Begin VB.Line Line2 
      X1              =   1512
      X2              =   2376
      Y1              =   1332
      Y2              =   1332
   End
   Begin VB.Line Line1 
      X1              =   1548
      X2              =   2340
      Y1              =   972
      Y2              =   972
   End
   Begin VB.Label Label1 
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   96
      TabIndex        =   11
      Top             =   552
      Width           =   936
   End
   Begin VB.Label Label1 
      Caption         =   "收文日："
      Height          =   180
      Index           =   1
      Left            =   96
      TabIndex        =   10
      Top             =   876
      Width           =   732
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家："
      Height          =   180
      Index           =   2
      Left            =   96
      TabIndex        =   9
      Top             =   1212
      Width           =   936
   End
End
Attribute VB_Name = "frm050307"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Morgan 2012/12/11 智權人員欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
Option Explicit
Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, s As Integer, strTemp1 As Variant, strTemp2 As Variant
Dim iPrint As Integer, Page As Integer, k As Integer, PLeft(0 To 11) As Integer, strTemp(0 To 11) As String, StrTest As String
Dim StrTest2 As String, StrTest99 As String, iY As Integer, strSQL2 As String
'Add By Cheng 2002/09/11
Dim blnClkSure As Boolean '判斷是否按下確定按鈕
Dim m_strSaleZoneCode '業務區代碼
Dim m_strSaleZone '業務區名稱
Dim m_strReceiveDate '收文日

Private Sub cmdok_Click(Index As Integer)
Select Case Index
Case 0 '確定
      'Add By Cheng 2002/09/11
      blnClkSure = False
     If Len(txt1(0)) = 0 Then
        s = MsgBox("系統類別不可空白!!", , "USER 輸入錯誤")
        txt1(0).SetFocus
        Exit Sub
     Else
         'Add By Cheng 2002/03/19
         If PUB_CheckKeyInDate(Me.txt1(1)) = -1 Then
            Me.txt1(1).SetFocus
            txt1_GotFocus 1
            Exit Sub
         End If
         If PUB_CheckKeyInDate(Me.txt1(2)) = -1 Then
            Me.txt1(2).SetFocus
            txt1_GotFocus 2
            Exit Sub
         End If
         'Add By Cheng 2002/09/11
         If Me.txt1(1).Text <> "" And Me.txt1(2).Text <> "" Then
            If Val(Me.txt1(1).Text) > Val(Me.txt1(2).Text) Then
               MsgBox "收文日範圍輸入錯誤!!!", vbExclamation + vbOKOnly
               blnClkSure = True
               Me.txt1(1).SetFocus
               txt1_GotFocus 1
               Exit Sub
            End If
         End If
         
        If Len(txt1(2)) = 0 Then
            s = MsgBox("收文日區間不可空白!!", , "USER 輸入錯誤")
            txt1(1).SetFocus
            txt1_GotFocus (1)
            Exit Sub
        Else
            'Add By Cheng 2002/09/11
            If Me.txt1(3).Text <> "" And Me.txt1(4).Text <> "" Then
               If Me.txt1(3).Text > Me.txt1(4).Text Then
                  MsgBox "申請國家範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                  blnClkSure = True
                  Me.txt1(3).SetFocus
                  txt1_GotFocus 3
                  Exit Sub
               End If
            End If
            Screen.MousePointer = vbHourglass
            Me.Enabled = False
'            StrTest = StrTest2
            StrTest = Me.txt1(0).Text
            strTemp1 = Split(UCase(StrTest), ",")
            strTemp2 = Split(UCase(txt1(0)), ",")
'            For i = 0 To UBound(strTemp1)
'                s = 0
'                For j = 0 To UBound(strTemp2)
'                    If strTemp2(j) = strTemp1(i) Then
'                        s = 1
'                        Exit For
'                    End If
'                Next j
'                If s = 0 Then
'                    StrTest = Replace(StrTest, strTemp1(i), "")
'                End If
'            Next i
            For i = 0 To UBound(strTemp2)
                s = 0
                For j = 0 To UBound(strTemp1)
                    If strTemp2(i) = strTemp1(j) Then
                        s = 1
                        Exit For
                    End If
                Next j
                If s = 0 Then
                    StrTest = Replace(StrTest, strTemp2(i), "")
                End If
            Next i
            Process
            Me.Enabled = True
            Screen.MousePointer = vbDefault
        End If
     End If
Case 1 '結束
     Unload Me
Case Else
End Select
End Sub

Private Sub Process()
'Add By Cheng 2002/10/28
Dim blnNoData2Print As Boolean
'Add By Cheng 2003/04/15
Dim strCP01 As String '系統類別

   ClearQueryLog (Me.Name) '2009/12/21 ADD BY SONIA 清除查詢印表記錄檔欄位
   
   '預設無資料可列印
   blnNoData2Print = True
   Screen.MousePointer = vbHourglass
   'cnnConnection.Execute "DELETE FROM R050307 "
   strSQL1 = "AND CP09 < 'C' AND CP10<>'907' AND CP10<>'913'"
   strSQL2 = strSQL1
   'Add By Cheng 2001/12/28
   '若為CFP案
   If StrStartSystemByNick = "CFP" Or StrStartSystemByNick = "CPS" Then
      strSQL1 = ""
      strSQL2 = ""
      strSQL1 = "AND CP10<>'907' AND CP10<>'913'"
      strSQL2 = strSQL1
   End If
   If Len(StrTest) <> 0 Then
      strSQL1 = strSQL1 + " AND CP01 IN (" & SQLGrpStr(StrTest, 1) & ") "
      strSQL2 = strSQL2 + " AND CP01 IN (" & SQLGrpStr(StrTest, 5) & ") "
      pub_QL05 = pub_QL05 & ";" & Label1(0) & StrTest                  '2009/12/17 add by sonia
   End If
   If Len(Trim(txt1(1))) <> 0 Then
      strSQL1 = strSQL1 + " AND CP05>=" & Val(ChangeTStringToWString(txt1(1))) & " "
      strSQL2 = strSQL2 + " AND CP05>=" & Val(ChangeTStringToWString(txt1(1))) & " "
      pub_QL05 = pub_QL05 & ";" & Label1(1) & txt1(1) & "-" & txt1(2)        '2009/12/21 add by sonia
   End If
   If Len(Trim(txt1(2))) <> 0 Then
      strSQL1 = strSQL1 & " AND CP05<=" & Val(ChangeTStringToWString(txt1(2))) & " "
      strSQL2 = strSQL2 & " AND CP05<=" & Val(ChangeTStringToWString(txt1(2))) & " "
   End If
   If Len(txt1(3)) <> 0 Then
      strSQL1 = strSQL1 + " AND SUBSTR(PA09,1,3)>='" & txt1(3) & "' "
      strSQL2 = strSQL2 + " AND SUBSTR(SP09,1,3)>='" & txt1(3) & "' "
      pub_QL05 = pub_QL05 & ";" & Label1(2) & txt1(3) & "-" & txt1(4)        '2009/12/21 add by sonia
   End If
   If Len(txt1(4)) <> 0 Then
       strSQL1 = strSQL1 + " AND SUBSTR(PA09,1,3)<='" & txt1(4) & "' "
       strSQL2 = strSQL2 + " AND SUBSTR(SP09,1,3)<='" & txt1(4) & "' "
   End If
   'Add by Morgan 2005/2/4
   If txtPA46 = "Y" Then
      strSQL1 = strSQL1 & " And PA09<>'056' AND PA46='Y' "
      pub_QL05 = pub_QL05 & ";" & Left(Label1(3), 9)                 '2009/12/17 add by sonia
   End If
   CheckOC
   
   'MODIFY BY SONIA 91.10.24 指定國家的子號案號資料不印
   'Modify By Cheng 2001/12/28
   '多選擇部門別(S2.ST03)欄位作為列印時判斷用
   '多選擇發文日(CP27)欄位作為列印時判斷用
   'strSQL = "SELECT CP09,NA03,CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(PA05,NVL(PA06,PA07)),decode(pa09,'000',cpm03,cpm04)," & SQLDate("CP06") & ",CP18,NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),S1.ST02,S2.ST02,CP26 FROM CASEPROGRESS,PATENT,STAFF S1,STAFF S2,NATION,CASEPROPERTYMAP,CUSTOMER WHERE CP57 IS NULL AND cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND cp13=S1.ST01(+) AND cp14=S2.ST01(+) AND cp01=cpm01(+) AND cp10=cpm02(+) AND SUBSTR(PA26,1,8)=cu01(+) AND decode(SUBSTR(Pa26,9,1),null,'0',substr(pa26,9,1))=cu02(+) AND pa09=na01(+) " & strSQL1
   'strSQL = "SELECT CP09,NA03,CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(PA05,NVL(PA06,PA07)),decode(pa09,'000',cpm03,cpm04)," & SQLDate("CP06") & ",CP18,NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),S1.ST02,S2.ST02,CP26,S2.ST03 FROM CASEPROGRESS,PATENT,STAFF S1,STAFF S2,NATION,CASEPROPERTYMAP,CUSTOMER WHERE CP57 IS NULL AND CP04='00' AND cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND cp13=S1.ST01(+) AND cp14=S2.ST01(+) AND cp01=cpm01(+) AND cp10=cpm02(+) AND SUBSTR(PA26,1,8)=cu01(+) AND decode(SUBSTR(Pa26,9,1),null,'0',substr(pa26,9,1))=cu02(+) AND pa09=na01(+) " & strSQL1
   'Modify By Cheng 2003/02/21
   '加多國案件欄位
   'strSQL = "SELECT CP09,NA03,CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(PA05,NVL(PA06,PA07)),decode(pa09,'000',cpm03,cpm04)," & SQLDate("CP06") & ",CP18,NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),S1.ST02,S2.ST02,CP26,S2.ST03,CP27 FROM CASEPROGRESS,PATENT,STAFF S1,STAFF S2,NATION,CASEPROPERTYMAP,CUSTOMER WHERE CP57 IS NULL AND CP04='00' AND cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND cp13=S1.ST01(+) AND cp14=S2.ST01(+) AND cp01=cpm01(+) AND cp10=cpm02(+) AND SUBSTR(PA26,1,8)=cu01(+) AND decode(SUBSTR(Pa26,9,1),null,'0',substr(pa26,9,1))=cu02(+) AND pa09=na01(+) " & strSQL1
   'strSQL = "SELECT CP09,NA03,CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(PA05,NVL(PA06,PA07)),decode(pa09,'000',cpm03,cpm04)," & SQLDate("CP06") & ",CP18,NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),S1.ST02,S2.ST02,CP26,S2.ST03,CP27,CP21 FROM CASEPROGRESS,PATENT,STAFF S1,STAFF S2,NATION,CASEPROPERTYMAP,CUSTOMER WHERE CP57 IS NULL AND CP04='00' AND cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND cp13=S1.ST01(+) AND cp14=S2.ST01(+) AND cp01=cpm01(+) AND cp10=cpm02(+) AND SUBSTR(PA26,1,8)=cu01(+) AND decode(SUBSTR(Pa26,9,1),null,'0',substr(pa26,9,1))=cu02(+) AND pa09=na01(+) " & strSQL1
   'strSQL = strSQL & " union all select CP09,NA03,CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),decode(sp09,'000',cpm03,cpm04)," & SQLDate("CP06") & ",CP18,NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)),S1.ST02,S2.ST02,CP26 FROM CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,NATION,CASEPROPERTYMAP,CUSTOMER WHERE CP57 IS NULL AND cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND cp13=S1.ST01(+) AND cp14=S2.ST01(+) AND cp01=cpm01(+) AND cp10=cpm02(+) AND SUBSTR(sP08,1,8)=cu01(+) AND decode(SUBSTR(sP08,9,1),null,'0',substr(sp08,9,1))=cu02(+) AND sp09=na01(+) " & strSQL2
   'strSQL = strSQL & " union all select CP09,NA03,CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),decode(sp09,'000',cpm03,cpm04)," & SQLDate("CP06") & ",CP18,NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)),S1.ST02,S2.ST02,CP26,S2.ST03 FROM CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,NATION,CASEPROPERTYMAP,CUSTOMER WHERE CP57 IS NULL AND CP04='00' AND cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND cp13=S1.ST01(+) AND cp14=S2.ST01(+) AND cp01=cpm01(+) AND cp10=cpm02(+) AND SUBSTR(sP08,1,8)=cu01(+) AND decode(SUBSTR(sP08,9,1),null,'0',substr(sp08,9,1))=cu02(+) AND sp09=na01(+) " & strSQL2
   'Modify By Cheng 2003/02/21
   'strSQL = strSQL & " union all select CP09,NA03,CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),decode(sp09,'000',cpm03,cpm04)," & SQLDate("CP06") & ",CP18,NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)),S1.ST02,S2.ST02,CP26,S2.ST03,CP27 FROM CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,NATION,CASEPROPERTYMAP,CUSTOMER WHERE CP57 IS NULL AND CP04='00' AND cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND cp13=S1.ST01(+) AND cp14=S2.ST01(+) AND cp01=cpm01(+) AND cp10=cpm02(+) AND SUBSTR(sP08,1,8)=cu01(+) AND decode(SUBSTR(sP08,9,1),null,'0',substr(sp08,9,1))=cu02(+) AND sp09=na01(+) " & strSQL2
   '加多國案件欄位
   'strSQL = strSQL & " union all select CP09,NA03,CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),decode(sp09,'000',cpm03,cpm04)," & SQLDate("CP06") & ",CP18,NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)),S1.ST02,S2.ST02,CP26,S2.ST03,CP27,CP21 FROM CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,NATION,CASEPROPERTYMAP,CUSTOMER WHERE CP57 IS NULL AND CP04='00' AND cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND cp13=S1.ST01(+) AND cp14=S2.ST01(+) AND cp01=cpm01(+) AND cp10=cpm02(+) AND SUBSTR(sP08,1,8)=cu01(+) AND decode(SUBSTR(sP08,9,1),null,'0',substr(sp08,9,1))=cu02(+) AND sp09=na01(+) " & strSQL2
   strSql = "SELECT CP09,NA03,CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(PA05,NVL(PA06,PA07)),decode(pa09,'000',cpm03,cpm04)," & SQLDate("CP06") & ",CP18,NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),S1.ST02,S2.ST02,CP26,S2.ST03,CP27,CP21,CP01 FROM CASEPROGRESS,PATENT,STAFF S1,STAFF S2,NATION,CASEPROPERTYMAP,CUSTOMER WHERE CP57 IS NULL AND CP04='00' AND cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND cp13=S1.ST01(+) AND cp14=S2.ST01(+) AND cp01=cpm01(+) AND cp10=cpm02(+) AND SUBSTR(PA26,1,8)=cu01(+) AND decode(SUBSTR(Pa26,9,1),null,'0',substr(pa26,9,1))=cu02(+) AND pa09=na01(+) " & strSQL1
   strSql = strSql & " union all select CP09,NA03,CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),decode(sp09,'000',cpm03,cpm04)," & SQLDate("CP06") & ",CP18,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),S1.ST02,S2.ST02,CP26,S2.ST03,CP27,CP21, CP01 FROM CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,NATION,CASEPROPERTYMAP,CUSTOMER WHERE CP57 IS NULL AND CP04='00' AND cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND cp13=S1.ST01(+) AND cp14=S2.ST01(+) AND cp01=cpm01(+) AND cp10=cpm02(+) AND SUBSTR(sP08,1,8)=cu01(+) AND decode(SUBSTR(sP08,9,1),null,'0',substr(sp08,9,1))=cu02(+) AND sp09=na01(+) " & strSQL2
   'strSQL = strSQL + " ORDER BY CP09 "
   strSql = strSql + " ORDER BY CP01, CP09 "
   
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
       With adoRecordset
           .MoveFirst
           'Add By Cheng 2003/04/15
           '記錄系統類別
           strCP01 = "" & .Fields("CP01").Value
           k = 0
           Page = 1
           ' PrintTitle   '2009/12/21 CANCEL BY SONIA 否則若所有資料都GoTo NextRecord則仍會印表頭
           '2009/12/21 END
           Do While .EOF = False
            'Add By Cheng 2001/12/28
            '若為CFP案
            If StrStartSystemByNick = "CFP" Or StrStartSystemByNick = "CPS" Then
               '若總收文號為C開頭者, 其部門別不可為P12
               If Left("" & .Fields(0).Value, 1) = "C" And .Fields(11).Value = "P12" Then
                  GoTo NextRecord
               End If
            End If
               'Add By Cheng 2002/10/28
            If StrStartSystemByNick = "P" Or StrStartSystemByNick = "PS" Then
               '若總收文號為B開頭者, 已發文資料不印
               If Left("" & .Fields(0).Value, 1) = "B" And IsNull(.Fields("CP27").Value) = False Then
                  GoTo NextRecord
               End If
            End If
            '2009/12/21 ADD BY SONIA
            If k = 0 Then
               PrintTitle
            End If
            '2009/12/21 END
               'Add By Cheng 2003/04/15
               '若系統類別不同則跳頁
               If strCP01 <> "" & .Fields("CP01").Value Then
                   strCP01 = "" & .Fields("CP01").Value
                   Page = Page + 1
                   Printer.CurrentX = 500
                   Printer.CurrentY = iPrint
                   Printer.Print String(200, "-")
                   Printer.NewPage
                   PrintTitle
               End If
               blnNoData2Print = False
               For i = 0 To 10
                   strTemp(i) = CheckStr(.Fields(i))
               Next i
               strTemp(1) = StrConv(MidB(StrConv(strTemp(1), vbFromUnicode), 1, 8), vbUnicode)
               'strTemp(5) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(5)))
               strTemp(3) = StrConv(MidB(StrConv(strTemp(3), vbFromUnicode), 1, 20), vbUnicode)
               strTemp(4) = StrConv(MidB(StrConv(strTemp(4), vbFromUnicode), 1, 10), vbUnicode)
               strTemp(7) = StrConv(MidB(StrConv(strTemp(7), vbFromUnicode), 1, 12), vbUnicode)
               strTemp(8) = StrConv(MidB(StrConv(strTemp(8), vbFromUnicode), 1, 8), vbUnicode)
               strTemp(9) = StrConv(MidB(StrConv(strTemp(9), vbFromUnicode), 1, 8), vbUnicode)
               '是否多國案
               strTemp(11) = "" & .Fields(13).Value
               If iPrint > 10000 Then
                   Page = Page + 1
                   Printer.CurrentX = 500
                   Printer.CurrentY = iPrint
                   Printer.Print String(200, "-")
                   Printer.NewPage
                   PrintTitle
               End If
               PrintDatil
               k = k + 1
               DoEvents
   'Add By Cheng 2001/12/28
NextRecord:
               .MoveNext
           Loop
           Printer.CurrentX = 500
           Printer.CurrentY = iPrint
           If k > 0 Then Printer.Print String(200, "-")
           Printer.EndDoc
       End With
       'Modify By Cheng 2003/04/21
       'Add By Cheng 2003/04/15
       '若從CFP進入, 列印新案承辦人明細表
   '    If StrStartSystemByNick = "CFP" Then PrintNewCaseList
       '2009/12/21 MODIFY BY SONIA
       'If StrStartSystemByNick = "CFP" And Me.Check1.Value = vbChecked Then PrintNewCaseList
       If k > 0 Then
         If StrStartSystemByNick = "CFP" And Me.Check1.Value = vbChecked Then
            pub_QL05 = pub_QL05 & ";同時" & Me.Check1.Caption
            InsertQueryLog (k)
            PrintNewCaseList
         Else
            InsertQueryLog (k)
         End If
       Else
          InsertQueryLog (0)  '2010/1/7 ADD BY SONIA
          ShowNoData
          Screen.MousePointer = vbDefault
          Exit Sub
       End If
       '2009/12/21 END
   Else
       InsertQueryLog (0)  '2010/1/7 ADD BY SONIA
       ShowNoData
       Screen.MousePointer = vbDefault
       Exit Sub
   End If
   CheckOC
   'Modify By Cheng 2002/10/28
   'ShowPrintOk
   If blnNoData2Print = False Then ShowPrintOk
   Screen.MousePointer = vbDefault
End Sub

Private Sub PrintTitle()
GetPleft
iPrint = 500
Printer.Orientation = 2
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 6500
Printer.CurrentY = iPrint
Printer.Print Trim(Me.Tag) & " 收文簿"
iPrint = iPrint + 500
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
Printer.CurrentX = 6200
Printer.CurrentY = iPrint
Printer.Print "收文日：" & Format(ChangeTStringToTDateString(txt1(1)) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(2))
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
iPrint = iPrint + 300
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "頁　　次：" & str(Page)
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
Printer.Font.Underline = True
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "總收文號"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "申請國家"
'Add By Cheng 2003/02/21
'加多國案件標題
Printer.CurrentX = PLeft(11)
Printer.CurrentY = iPrint
Printer.Print "多國"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "本所案號"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "案件名稱"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "案件性質"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "本所期限"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "點數"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "申請人"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print "智權人員"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
Printer.Print "承辦人"
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iPrint
Printer.Print "是否算"
iPrint = iPrint + 300
'Add By Cheng 2003/02/21
'加多國案件標題
Printer.CurrentX = PLeft(11)
Printer.CurrentY = iPrint
Printer.Print "案件"
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iPrint
Printer.Print "案件數"
Printer.Font.Underline = False
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
End Sub

Private Sub PrintDatil()
'Modify By Cheng 2003/02/21
'For i = 0 To 5
'    Printer.CurrentX = PLeft(i)
'    Printer.CurrentY = iPrint
'    Printer.Print strTemp(i)
'Next i
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print strTemp(0)
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print strTemp(1)
Printer.CurrentX = PLeft(11)
Printer.CurrentY = iPrint
Printer.Print strTemp(11)
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print strTemp(2)
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print strTemp(3)
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print strTemp(4)
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print strTemp(5)
Printer.CurrentX = PLeft(6) + 500 - Printer.TextWidth(strTemp(6))
Printer.CurrentY = iPrint
Printer.Print strTemp(6)
For i = 7 To 10
    Printer.CurrentX = PLeft(i)
    Printer.CurrentY = iPrint
    Printer.Print strTemp(i)
Next i

iPrint = iPrint + 300
End Sub

Private Sub GetPleft()
Erase PLeft
PLeft(0) = 500
PLeft(1) = 1700
'Add By Cheng 2003/02/21
'是否多國案
PLeft(11) = 2800
PLeft(2) = 2800 + 750
PLeft(3) = 4600 + 750
PLeft(4) = 7400 + 750
PLeft(5) = 8700 + 750
PLeft(6) = 10000 + 750
PLeft(7) = 10700 + 750
PLeft(8) = 12300 + 750
PLeft(9) = 13400 + 750
PLeft(10) = 14500 + 750
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
Select Case StrStartSystemByNick
Case "P", "PS"
     StrTest2 = "P,PS"
Case "CFP", "CPS"
     StrTest2 = "CFP,CPS"
Case Else
     s = MsgBox("程式錯誤!!傳入值錯誤  " & Me.Name, , "ERROR")
     Unload Me
     Exit Sub
End Select
Me.Tag = StrStartSystemByNick
StrTest = StrTest2
strTemp1 = Split(UCase(StrTest), ",")
strTemp2 = Split(UCase(GetSystemKindByNick), ",")
For i = 0 To UBound(strTemp1)
    s = 0
    For j = 0 To UBound(strTemp2)
        If strTemp2(j) = strTemp1(i) Then
            s = 1
            Exit For
        End If
    Next j
    If s = 0 Then
        StrTest = Replace(StrTest, strTemp1(i), "")
    End If
Next i
txt1(0) = StrTest
txt1(1) = strSrvDate(2)
txt1(2) = strSrvDate(2)
'Add By Cheng 2003/04/21
'若非從CFP進入
If StrStartSystemByNick <> "CFP" Then
   Me.Check1.Enabled = False
Else
   Me.Check1.Value = vbChecked
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm050307 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txt1_LostFocus(Index As Integer)
Select Case Index
Case 0
     strTemp1 = Split(UCase(GetSystemKindByNick), ",")
     strTemp2 = Split(UCase(txt1(0)), ",")
     For i = 0 To UBound(strTemp2)
        s = 0
        For j = 0 To UBound(strTemp1)
            If strTemp1(j) = strTemp2(i) Then
                s = 1
                Exit For
            End If
        Next j
        If s = 0 Then
            s = MsgBox(strUserName & " 沒有 " & strTemp2(i) & " 的權限 ", , "權限問題")
            txt1(0).SetFocus
        End If
    Next i
Case 2, 4 '收文日, 申請國家
   'Modify By Cheng 2002/09/11
   If blnClkSure = False Then
      If RunNick(txt1(Index - 1), txt1(Index)) Then
         txt1(Index - 1).SetFocus
      End If
   Else
      blnClkSure = False
   End If
End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
Select Case Index
Case 1, 2 '收文日起, 迄
   If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
      Cancel = True
      Me.txt1(Index).SetFocus
      txt1_GotFocus Index
   End If
End Select
End Sub

'Add By Cheng 2003/04/15
Private Sub PrintNewCaseList()
strSQL1 = ""
strSQL2 = ""
'組字串
'系統類別
If Len(txt1(0)) <> 0 Then
   strSQL1 = strSQL1 & " and cp01 in (" & SQLGrpStr(txt1(0), 1) & ") "
   strSQL2 = strSQL2 & " and cp01 in (" & SQLGrpStr(txt1(0), 5) & ") "
End If
'收文日
If Len(Trim(txt1(1))) <> 0 Then
    strSQL1 = strSQL1 + " AND CP05>=" & Val(ChangeTStringToWString(txt1(1))) & " "
    strSQL2 = strSQL2 + " AND CP05>=" & Val(ChangeTStringToWString(txt1(1))) & " "
End If
If Len(Trim(txt1(2))) <> 0 Then
    strSQL1 = strSQL1 + " AND CP05<=" & Val(ChangeTStringToWString(txt1(2))) & " "
    strSQL2 = strSQL2 + " AND CP05<=" & Val(ChangeTStringToWString(txt1(2))) & " "
End If
'申請國家
If Len(Trim(txt1(3))) <> 0 Then
    strSQL1 = strSQL1 + " AND SUBSTR(PA09,1,3)>='" & txt1(3) & "' "
    strSQL2 = strSQL2 + " AND SUBSTR(SP09,1,3)>='" & txt1(3) & "' "
End If
If Len(Trim(txt1(4))) <> 0 Then
    strSQL1 = strSQL1 + " AND SUBSTR(PA09,1,3)<='" & txt1(4) & "' "
    strSQL2 = strSQL2 + " AND SUBSTR(SP09,1,3)<='" & txt1(4) & "' "
End If
strSQL1 = strSQL1 + " And CP04 ='00' "
strSQL2 = strSQL2 + " And CP04 ='00' "
'組合
'Modify by Morgan 2010/7/15 +改抓所有新案的案件性質
'strSql = "SELECT S1.ST03,A0902," & SQLDate("CP05") & ",CP13,S1.ST02,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,NVL(CU04,DECODE(CU05,NULL,CU06,CU05|| ' '||CU88||' '||CU89||' '||CU90)),NVL(PA05,NVL(PA06,PA07)),CP18,CP14,S2.ST02, CP21 FROM CASEPROGRESS,STAFF S1,STAFF S2,NATION,PATENT,CUSTOMER,ACC090 WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND PA09=NA01(+) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND CP10>='101' AND CP10<='105' AND CP09<'B' AND (CP14<>'72006' AND S2.ST03<>'P12') And SUBSTR(PA26,1,8)=CU01(+) And SUBSTR(PA26,9,1)=CU02(+) AND S1.ST03=A0901(+) " & strSQL1
'strSql = strSql + " UNION ALL SELECT S1.ST03,A0902," & SQLDate("CP05") & ",CP13,S1.ST02,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,NVL(CU04,DECODE(CU05,NULL,CU06,CU05|| ' '||CU88||' '||CU89||' '||CU90)),NVL(SP05,NVL(SP06,SP07)),CP18,CP14,S2.ST02, CP21 FROM CASEPROGRESS,STAFF S1,STAFF S2,NATION,SERVICEPRACTICE,CUSTOMER,ACC090 WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND SP09=NA01(+) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND CP10>='101' AND CP10<='105' AND CP09<'B' AND (CP14<>'72006' AND S2.ST03<>'P12') AND SUBSTR(SP08,1,8)=CU01(+) AND SUBSTR(SP08,9,1)=CU02(+) AND S1.ST03=A0901(+) " & strSQL2
'Modify by Morgan 2010/8/12 百年蟲
'strSql = "SELECT S1.ST03,A0902," & SQLDate("CP05") & ",CP13,S1.ST02,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,NVL(CU04,DECODE(CU05,NULL,CU06,CU05|| ' '||CU88||' '||CU89||' '||CU90)),NVL(PA05,NVL(PA06,PA07)),CP18,CP14,S2.ST02, CP21 FROM CASEPROGRESS,STAFF S1,STAFF S2,NATION,PATENT,CUSTOMER,ACC090 WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND PA09=NA01(+) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND CP10 in(" & NewCasePtyList & ") AND CP09<'B' AND (CP14<>'72006' AND S2.ST03<>'P12') And SUBSTR(PA26,1,8)=CU01(+) And SUBSTR(PA26,9,1)=CU02(+) AND S1.ST03=A0901(+) " & strSQL1
'modify by sonia 2016/3/3 +CP14<>'87025'
strSql = "SELECT S1.ST03,A0902,substrb(' '||sqldatet(CP05),-9),CP13,S1.ST02,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,NVL(CU04,DECODE(CU05,NULL,CU06,CU05|| ' '||CU88||' '||CU89||' '||CU90)),NVL(PA05,NVL(PA06,PA07)),CP18,CP14,S2.ST02, CP21 FROM CASEPROGRESS,STAFF S1,STAFF S2,NATION,PATENT,CUSTOMER,ACC090 WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND PA09=NA01(+) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND CP10 in(" & NewCasePtyList & ") AND CP09<'B' AND (CP14<>'72006' AND CP14<>'87025' AND S2.ST03<>'P12') And SUBSTR(PA26,1,8)=CU01(+) And SUBSTR(PA26,9,1)=CU02(+) AND S1.ST03=A0901(+) " & strSQL1
strSql = strSql + " UNION ALL SELECT S1.ST03,A0902," & SQLDate("CP05") & ",CP13,S1.ST02,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,NVL(CU04,DECODE(CU05,NULL,CU06,CU05|| ' '||CU88||' '||CU89||' '||CU90)),NVL(SP05,NVL(SP06,SP07)),CP18,CP14,S2.ST02, CP21 FROM CASEPROGRESS,STAFF S1,STAFF S2,NATION,SERVICEPRACTICE,CUSTOMER,ACC090 WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND SP09=NA01(+) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND CP10 in(" & NewCasePtyList & ") AND CP09<'B' AND (CP14<>'72006' AND CP14<>'87025' AND S2.ST03<>'P12') AND SUBSTR(SP08,1,8)=CU01(+) AND SUBSTR(SP08,9,1)=CU02(+) AND S1.ST03=A0901(+) " & strSQL2
strSql = strSql & " ORDER BY 1, 3, 4, 6"
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
Else
'    ShowNoData
    CheckOC
    Screen.MousePointer = vbDefault
    Exit Sub
End If
'列印資料
PrintData
End Sub

'Add By Cheng 2003/04/15
Sub PrintData()
Page = 1
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    adoRecordset.MoveFirst
    m_strSaleZoneCode = "" & adoRecordset.Fields(0).Value
    m_strSaleZone = "" & adoRecordset.Fields(1).Value
    m_strReceiveDate = "" & adoRecordset.Fields(2).Value
    PrintTitleA
    Do While adoRecordset.EOF = False
'        For i = 0 To 10
        For i = 0 To 11
            strTemp(i) = CheckStr(adoRecordset.Fields(i))
        Next i
        strTemp(6) = StrConv(MidB(StrConv(strTemp(6), vbFromUnicode), 1, 28), vbUnicode)
        strTemp(7) = StrConv(MidB(StrConv(strTemp(7), vbFromUnicode), 1, 28), vbUnicode)
        '若業務區不同
        If m_strSaleZoneCode <> strTemp(0) Then
            Printer.NewPage
            Page = Page + 1
            m_strSaleZoneCode = "" & adoRecordset.Fields(0).Value
            m_strSaleZone = "" & adoRecordset.Fields(1).Value
            m_strReceiveDate = "" & adoRecordset.Fields(2).Value
            PrintTitleA
        End If
        If iPrint > 10000 Then
            Printer.NewPage
            Page = Page + 1
            m_strSaleZoneCode = "" & adoRecordset.Fields(0).Value
            m_strSaleZone = "" & adoRecordset.Fields(1).Value
            m_strReceiveDate = "" & adoRecordset.Fields(2).Value
            PrintTitleA
        End If
        PrintDatilA
        adoRecordset.MoveNext
    Loop
End If
CheckOC
Printer.EndDoc
End Sub

Sub PrintTitleA()     '印抬頭
    GetPleftA
    iPrint = 500
    Printer.Orientation = 2
    Printer.Font.Name = "細明體"
    Printer.Font.Size = 22
    Printer.Font.Bold = True
    Printer.Font.Underline = True
    Printer.CurrentX = 6300
    Printer.CurrentY = iPrint
    Printer.Print "新案承辦人明細表"
    iPrint = iPrint + 500
    Printer.Font.Size = 12
    Printer.Font.Bold = False
    Printer.Font.Underline = False
    Printer.CurrentX = 500
    Printer.CurrentY = iPrint
    Printer.Print "列印人：" & strUserName
    Printer.CurrentX = 6500
    Printer.CurrentY = iPrint
    Printer.Print "收文日：" & Format(ChangeTStringToTDateString(txt1(1)) & " ", "@@@@@@@@@@")
    Printer.CurrentX = 13000
    Printer.CurrentY = iPrint
    Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
    iPrint = iPrint + 300
    Printer.CurrentX = 500
    Printer.CurrentY = iPrint
    Printer.Print "業務區：" & m_strSaleZoneCode & "　　" & m_strSaleZone
    Printer.CurrentX = 13000
    Printer.CurrentY = iPrint
    Printer.Print "頁　　次：" & str(Page)
    iPrint = iPrint + 300
    Printer.CurrentX = 500
    Printer.CurrentY = iPrint
    Printer.Print String(200, "-")
    iPrint = iPrint + 300
    Printer.CurrentX = PLeft(0)
    Printer.CurrentY = iPrint
    Printer.Print "智權人員"
    Printer.CurrentX = PLeft(1)
    Printer.CurrentY = iPrint
    Printer.Print "本所案號"
    Printer.CurrentX = PLeft(2)
    Printer.CurrentY = iPrint
    Printer.Print "申請人"
    Printer.CurrentX = PLeft(3)
    Printer.CurrentY = iPrint
    Printer.Print "案件名稱"
    Printer.CurrentX = PLeft(4)
    Printer.CurrentY = iPrint
    Printer.Print "點數"
    Printer.CurrentX = PLeft(5)
    Printer.CurrentY = iPrint
    Printer.Print "承辦人"
    'Add By Cheng 2003/03/2
    Printer.CurrentX = PLeft(6)
    Printer.CurrentY = iPrint
    Printer.Print "多國案"
    iPrint = iPrint + 300
    Printer.CurrentX = 500
    Printer.CurrentY = iPrint
    Printer.Print String(200, "-")
    iPrint = iPrint + 300
End Sub

Sub GetPleftA()
    Erase PLeft
    PLeft(0) = 500
    PLeft(1) = 2000
    PLeft(2) = 4500
    PLeft(3) = 8500
    PLeft(4) = 13000
    PLeft(5) = 14000
    'Add By Cheng 2003/03/28
    PLeft(6) = 15250 '多國
End Sub

Sub PrintDatilA()           '印內容
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print strTemp(4)
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print strTemp(5)
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print strTemp(6)
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print strTemp(7)
Printer.CurrentX = PLeft(4) + 500 - Printer.TextWidth(Format(strTemp(8), "###.00"))
Printer.CurrentY = iPrint
Printer.Print Format(strTemp(8), "###.00")
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print strTemp(10)
'Add By Cheng 2003/04/15
'是否多國案
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print strTemp(11)
iPrint = iPrint + 300
End Sub
'Add by Morgan 2005/2/4 加PCT進入國家階段條件
Private Sub txtPA46_GotFocus()
   'edit by nickc 2007/07/11 切換輸入法改用API
   'txtPA46.IMEMode = 2
   CloseIme
   TextInverse txtPA46
End Sub

Private Sub txtPA46_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And Chr(KeyAscii) <> "Y" Then
      KeyAscii = 0
      Beep
   End If
End Sub
