VERSION 5.00
Begin VB.Form frm060311 
   BorderStyle     =   1  '單線固定
   Caption         =   "智權人員收文明細表"
   ClientHeight    =   3540
   ClientLeft      =   2070
   ClientTop       =   3810
   ClientWidth     =   5040
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   5040
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   13
      Left            =   1185
      MaxLength       =   3
      TabIndex        =   1
      Top             =   1170
      Width           =   990
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   14
      Left            =   2295
      MaxLength       =   3
      TabIndex        =   2
      Top             =   1170
      Width           =   990
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   4080
      TabIndex        =   14
      Top             =   165
      Width           =   800
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   3285
      TabIndex        =   13
      Top             =   165
      Width           =   756
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   12
      Left            =   1590
      MaxLength       =   1
      TabIndex        =   12
      Top             =   3000
      Width           =   240
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   10
      Left            =   2295
      MaxLength       =   9
      TabIndex        =   11
      Top             =   2670
      Width           =   990
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   9
      Left            =   1185
      MaxLength       =   9
      TabIndex        =   10
      Top             =   2670
      Width           =   990
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   8
      Left            =   2295
      MaxLength       =   99
      TabIndex        =   9
      Top             =   2370
      Width           =   990
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   7
      Left            =   1185
      MaxLength       =   9
      TabIndex        =   8
      Top             =   2370
      Width           =   990
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   6
      Left            =   1185
      MaxLength       =   6
      TabIndex        =   7
      Top             =   2070
      Width           =   990
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   5
      Left            =   2295
      MaxLength       =   4
      TabIndex        =   6
      Top             =   1755
      Width           =   990
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   1185
      MaxLength       =   4
      TabIndex        =   5
      Top             =   1755
      Width           =   990
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   2295
      MaxLength       =   7
      TabIndex        =   4
      Top             =   1470
      Width           =   990
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   1185
      MaxLength       =   7
      TabIndex        =   3
      Top             =   1470
      Width           =   990
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1185
      TabIndex        =   0
      Top             =   840
      Width           =   2115
   End
   Begin VB.Label Label1 
      Caption         =   "( 空白表全部 )"
      Height          =   180
      Index           =   7
      Left            =   3345
      TabIndex        =   25
      Top             =   900
      Width           =   1110
   End
   Begin VB.Line Line5 
      X1              =   2190
      X2              =   2280
      Y1              =   1290
      Y2              =   1290
   End
   Begin VB.Line Line4 
      X1              =   2190
      X2              =   2280
      Y1              =   2835
      Y2              =   2835
   End
   Begin VB.Line Line3 
      X1              =   2190
      X2              =   2280
      Y1              =   2490
      Y2              =   2490
   End
   Begin VB.Line Line2 
      X1              =   2190
      X2              =   2280
      Y1              =   1875
      Y2              =   1875
   End
   Begin VB.Line Line1 
      X1              =   2190
      X2              =   2280
      Y1              =   1590
      Y2              =   1590
   End
   Begin VB.Label lbl1 
      Height          =   180
      Left            =   2295
      TabIndex        =   24
      Top             =   2130
      Width           =   1170
   End
   Begin VB.Label Label1 
      Caption         =   "(Y:印)"
      Height          =   180
      Index           =   11
      Left            =   1905
      TabIndex        =   23
      Top             =   3060
      Width           =   600
   End
   Begin VB.Label Label1 
      Caption         =   "是否列印明細："
      Height          =   180
      Index           =   8
      Left            =   285
      TabIndex        =   22
      Top             =   3060
      Width           =   1260
   End
   Begin VB.Label Label1 
      Caption         =   "代理人："
      Height          =   180
      Index           =   6
      Left            =   285
      TabIndex        =   21
      Top             =   2730
      Width           =   720
   End
   Begin VB.Label Label1 
      Caption         =   "申請人："
      Height          =   180
      Index           =   5
      Left            =   285
      TabIndex        =   20
      Top             =   2415
      Width           =   720
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員："
      Height          =   180
      Index           =   4
      Left            =   285
      TabIndex        =   19
      Top             =   2130
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質："
      Height          =   180
      Index           =   3
      Left            =   285
      TabIndex        =   18
      Top             =   1815
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "日期："
      Height          =   180
      Index           =   2
      Left            =   285
      TabIndex        =   17
      Top             =   1530
      Width           =   555
   End
   Begin VB.Label Label1 
      Caption         =   "業務區別："
      Height          =   180
      Index           =   1
      Left            =   285
      TabIndex        =   16
      Top             =   1215
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   285
      TabIndex        =   15
      Top             =   885
      Width           =   900
   End
End
Attribute VB_Name = "frm060311"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
Option Explicit

Dim i As Integer, j As Integer, s As Integer, strTemp3(0 To 9) As String
Dim SavDay1 As String, SavDay2 As String, SavDay3 As String, SavDay4 As String, SavDay5 As String
Dim iPrint As Integer, Page As Integer, strTemp(0 To 9) As String
Dim PLeft(0 To 9) As Integer, strTemp1 As Variant, strTemp2 As Variant
Dim StrTemp4(0 To 4) As String, StrTemp5(0 To 4) As String
Dim m_bFirstLine As Boolean
Dim m_strSharePointVTB As String '分配點數語法
Dim m_intR As Integer

Private Sub cmdok_Click(Index As Integer)
   Select Case Index
      Case 0
         Screen.MousePointer = vbHourglass
         If TxtValidate = True Then
            Me.Enabled = False
            m_strSharePointVTB = ""
            'Modify by Morgan 2010/5/20 改抓分配點數資料
            'Process
            ProcessNew
            Me.Enabled = True
         End If
         Screen.MousePointer = vbDefault
      
      Case 1
         Unload Me
   End Select
End Sub

Sub Process()

   Dim strSQL1 As String, strSQL2 As String, StrSQL3 As String
   Dim strVTB As String
   
   cnnConnection.Execute "DELETE FROM R060311 WHERE ID='" & strUserNum & "' "
   strSQL1 = ""
   strSQL2 = ""
   'Modify by Morgan 2007/11/27 只抓A類收文
   StrSQL3 = " AND CP57 IS NULL AND CP09<'B'"
   
   '系統別
   If Len(txt1(0)) <> 0 Then
      strSQL1 = strSQL1 + " AND CP01 IN (" & SQLGrpStr(txt1(0), 1) & ") "
      strSQL2 = strSQL2 + " AND CP01 IN (" & SQLGrpStr(txt1(0), 5) & ") "
   End If
   
   '業務區
   If Len(txt1(13)) <> 0 Then
       StrSQL3 = StrSQL3 + " AND CP12>='" & txt1(13) & "' "
   End If
   If Len(txt1(14)) <> 0 Then
       StrSQL3 = StrSQL3 + " AND CP12<='" & txt1(14) & "' "
   End If
   
   '收文日
   If Len(txt1(2)) <> 0 Then
     StrSQL3 = StrSQL3 + " AND CP05>=" & Val(ChangeTStringToWString(txt1(2))) & " "
   End If
   StrSQL3 = StrSQL3 + " AND CP05<=" & Val(ChangeTStringToWString(txt1(3)))
   
   '案件性質
   If Len(txt1(4)) <> 0 Then
       StrSQL3 = StrSQL3 + " AND CP10>='" & txt1(4) & "' "
   End If
   If Len(txt1(5)) <> 0 Then
       StrSQL3 = StrSQL3 + " AND CP10<='" & txt1(5) & "' "
   End If
   
   '智權人員
   If Len(txt1(6)) <> 0 Then
       StrSQL3 = StrSQL3 + " AND CP13='" & txt1(6) & "' "
   End If
   
   '申請人
   If Len(Trim(txt1(7))) <> 0 And Len(Trim(txt1(8))) <> 0 Then
       strSQL1 = strSQL1 + " AND ((PA26>='" & GetNewFagent(txt1(7)) & "' AND PA26<='" & GetNewFagent(txt1(8)) & "') OR (PA27>='" & GetNewFagent(txt1(7)) & "' AND PA27<='" & GetNewFagent(txt1(8)) & "') OR (PA28>='" & GetNewFagent(txt1(7)) & "' AND PA28<='" & GetNewFagent(txt1(8)) & "') OR (PA29>='" & GetNewFagent(txt1(7)) & "' AND PA29<='" & GetNewFagent(txt1(8)) & "') OR (PA30>='" & GetNewFagent(txt1(7)) & "' AND PA30<='" & GetNewFagent(txt1(8)) & "')) "
       strSQL2 = strSQL2 + " AND ((SP08>='" & GetNewFagent(txt1(7)) & "' AND SP08<='" & GetNewFagent(txt1(8)) & "') OR (SP58<='" & GetNewFagent(txt1(7)) & "' AND SP58<='" & GetNewFagent(txt1(8)) & "') OR (SP59>='" & GetNewFagent(txt1(7)) & "' AND SP59<='" & GetNewFagent(txt1(8)) & "')) "
   Else
       If Len(Trim(txt1(7))) <> 0 And Len(Trim(txt1(8))) = 0 Then
           strSQL1 = strSQL1 + " AND (PA26>='" & GetNewFagent(txt1(7)) & "' OR PA27>='" & GetNewFagent(txt1(7)) & "' OR PA28>='" & GetNewFagent(txt1(7)) & "' OR PA29>='" & GetNewFagent(txt1(7)) & "' OR PA30>='" & GetNewFagent(txt1(7)) & "') "
           strSQL2 = strSQL2 + " AND (SP08>='" & GetNewFagent(txt1(7)) & "' OR SP58>='" & GetNewFagent(txt1(7)) & "' OR SP59>='" & GetNewFagent(txt1(7)) & "') "
       Else
           If Len(Trim(txt1(7))) = 0 And Len(Trim(txt1(8))) <> 0 Then
               strSQL1 = strSQL1 + " AND (PA26<='" & GetNewFagent(txt1(8)) & "' OR PA27<='" & GetNewFagent(txt1(8)) & "' OR PA28<='" & GetNewFagent(txt1(8)) & "' OR PA29<='" & GetNewFagent(txt1(8)) & "' OR PA30<='" & GetNewFagent(txt1(8)) & "') "
               strSQL2 = strSQL2 + " AND (SP08<='" & GetNewFagent(txt1(8)) & "' OR SP58<='" & GetNewFagent(txt1(8)) & "' OR SP59<='" & GetNewFagent(txt1(8)) & "') "
           End If
       End If
   End If
   
   '代理人
   If Len(Trim(txt1(9))) <> 0 And Len(Trim(txt1(10))) <> 0 Then
       strSQL1 = strSQL1 + " AND PA75>='" & GetNewFagent(txt1(9)) & "' AND PA75<='" & GetNewFagent(txt1(10)) & "' "
       strSQL2 = strSQL2 + " AND SP26>='" & GetNewFagent(txt1(9)) & "' AND SP26<='" & GetNewFagent(txt1(10)) & "' "
   Else
       If Len(Trim(txt1(9))) <> 0 And Len(Trim(txt1(10))) = 0 Then
           strSQL1 = strSQL1 + " AND PA75>='" & GetNewFagent(txt1(9)) & "' "
           strSQL2 = strSQL2 + " AND SP26>='" & GetNewFagent(txt1(9)) & "' "
       Else
           If Len(Trim(txt1(9))) = 0 And Len(Trim(txt1(10))) <> 0 Then
               strSQL1 = strSQL1 + " AND PA75<='" & GetNewFagent(txt1(10)) & "' "
               strSQL2 = strSQL2 + " AND SP26<='" & GetNewFagent(txt1(10)) & "' "
           End If
       End If
   End If
   'Modify by Morgan 2008/1/23 改抓服務費/1000,不必考慮收據,但排除銷帳及作廢的請款單
   'strExc(0) = "select nvl(a0902,a0903),ST02," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(PA05,NVL(PA06,PA07)),RPAD(CP10,4,' ')||NVL(CPM03,CPM04),NVL(NA03,NA04)," & SQLDate("CP06") & "," & SQLDate("CP27") & ",DECODE(CP20,NULL,DECODE(SUBSTR(CP60,1,1),'E',(A0K06+A0K07)/1000,'X',A1K11/1000,NULL,0)) FROM CASEPROGRESS,PATENT,STAFF,ACC1K0,ACC090,ACC0K0,NATION,CASEPROPERTYMAP,FAGENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND PA01 IS NOT NULL AND CP01=CPM01(+) AND CP10=CPM02(+) AND FA10=NA01(+) AND SUBSTR(PA75,1,8) = FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) AND CP13=ST01(+) AND CP60=A1K01(+) AND CP60=A0K01(+) AND CP12=A0901(+) " & strSQL1 & StrSQL3 & _
   '   " UNION ALL select nvl(a0902,a0903),ST02," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),RPAD(CP10,4,' ')||NVL(CPM03,CPM04),NVL(NA03,NA04)," & SQLDate("CP06") & "," & SQLDate("CP27") & ",DECODE(CP20,NULL,DECODE(SUBSTR(CP60,1,1),'E',(A0K06+A0K07)/1000,'X',A1K11/1000,NULL,0)) FROM CASEPROGRESS,SERVICEPRACTICE,STAFF,ACC1K0,ACC090,ACC0K0,NATION,CASEPROPERTYMAP,FAGENT WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND SP01 IS NOT NULL AND CP01=CPM01(+) AND CP10=CPM02(+) AND FA10=NA01(+) AND SUBSTR(SP26,1,8) = FA01(+) AND SUBSTR(SP26,9,1)=FA02(+) AND CP13=ST01(+) AND CP60=A1K01(+) AND CP60=A0K01(+) AND CP12=A0901(+) " & strSQL2 & StrSQL3
   'Modify By Sindy 2013/1/15
'   strExc(0) = "select nvl(a0902,a0903),ST02," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(PA05,NVL(PA06,PA07)),RPAD(CP10,4,' ')||NVL(CPM03,CPM04),NVL(NA03,NA04)," & SQLDate("CP06") & "," & SQLDate("CP27") & ",DECODE(CP20,NULL,(A1K11-nvl(A1K09,0)-nvl(A1K06,0)*A1K10)/1000),CP09,CP60,CP13 FROM CASEPROGRESS,PATENT,STAFF,ACC1K0,ACC090,NATION,CASEPROPERTYMAP,FAGENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND PA01 IS NOT NULL AND CP01=CPM01(+) AND CP10=CPM02(+) AND FA10=NA01(+) AND SUBSTR(PA75,1,8) = FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) AND CP13=ST01(+) AND CP60=A1K01(+) and a1k12 is null and a1k25 is null AND CP12=A0901(+) " & strSQL1 & StrSQL3 & _
'    " UNION ALL select nvl(a0902,a0903),ST02," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),RPAD(CP10,4,' ')||NVL(CPM03,CPM04),NVL(NA03,NA04)," & SQLDate("CP06") & "," & SQLDate("CP27") & ",DECODE(CP20,NULL,(A1K11-nvl(A1K09,0)-nvl(A1K06,0)*A1K10)/1000),CP09,CP60,CP13 FROM CASEPROGRESS,SERVICEPRACTICE,STAFF,ACC1K0,ACC090,NATION,CASEPROPERTYMAP,FAGENT WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND SP01 IS NOT NULL AND CP01=CPM01(+) AND CP10=CPM02(+) AND FA10=NA01(+) AND SUBSTR(SP26,1,8) = FA01(+) AND SUBSTR(SP26,9,1)=FA02(+) AND CP13=ST01(+) AND CP60=A1K01(+) and a1k12 is null and a1k25 is null AND CP12=A0901(+) " & strSQL2 & StrSQL3
   strExc(0) = "select nvl(a0902,a0903),ST02," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(PA05,NVL(PA06,PA07)),RPAD(CP10,4,' ')||NVL(CPM03,CPM04),NVL(NA03,NA04)," & SQLDate("CP06") & "," & SQLDate("CP27") & ",DECODE(CP20,NULL,(A1K11-nvl(A1K09,0)-nvl(A1K06,0))/1000),CP09,CP60,CP13 FROM CASEPROGRESS,PATENT,STAFF,ACC1K0,ACC090,NATION,CASEPROPERTYMAP,FAGENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND PA01 IS NOT NULL AND CP01=CPM01(+) AND CP10=CPM02(+) AND FA10=NA01(+) AND SUBSTR(PA75,1,8) = FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) AND CP13=ST01(+) AND CP60=A1K01(+) and a1k12 is null and a1k25 is null AND CP12=A0901(+) " & strSQL1 & StrSQL3 & _
    " UNION ALL select nvl(a0902,a0903),ST02," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),RPAD(CP10,4,' ')||NVL(CPM03,CPM04),NVL(NA03,NA04)," & SQLDate("CP06") & "," & SQLDate("CP27") & ",DECODE(CP20,NULL,(A1K11-nvl(A1K09,0)-nvl(A1K06,0))/1000),CP09,CP60,CP13 FROM CASEPROGRESS,SERVICEPRACTICE,STAFF,ACC1K0,ACC090,NATION,CASEPROPERTYMAP,FAGENT WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND SP01 IS NOT NULL AND CP01=CPM01(+) AND CP10=CPM02(+) AND FA10=NA01(+) AND SUBSTR(SP26,1,8) = FA01(+) AND SUBSTR(SP26,9,1)=FA02(+) AND CP13=ST01(+) AND CP60=A1K01(+) and a1k12 is null and a1k25 is null AND CP12=A0901(+) " & strSQL2 & StrSQL3
   '2013/1/15 End
   intI = 1
   Set adoRecordset = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With adoRecordset
          .MoveFirst
          Do While .EOF = False
              For i = 0 To 9
                  strTemp(i) = CheckStr(.Fields(i))
              Next i
               'Add by Morgan 2008/1/23 請款點數算在最後收文的智權人員上
               If Not IsNull(.Fields("CP60")) And Val(strTemp(9)) > 0 Then
                  strExc(1) = "SELECT SUBSTR(MAX(CP05||CP09),9) FROM CASEPROGRESS WHERE CP60='" & .Fields("CP60") & "'"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(1))
                  If intI = 1 Then
                     If .Fields("CP09") <> RsTemp.Fields(0) Then
                        strTemp(9) = Empty
                     End If
                  End If
               End If
              strSql = "INSERT INTO R060311(R044001,R044002,R044003,R044004,R044005,R044006,R044007,R044008,R044009,R044010,id,R044011) VALUES('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & ChgSQL(strTemp(8)) & "','" & ChgSQL(strTemp(9)) & "','" & strUserNum & "','" & .Fields("cp13") & "') "
              cnnConnection.Execute strSql
              .MoveNext
          Loop
          
      End With
      PrintData
      ShowPrintOk
   Else
      ShowNoData
   End If
   
End Sub

Sub PrintData()
   
   Dim rsQuery1 As ADODB.Recordset
   Dim strKeyNow As String, strKeyLast As String
   
   strExc(0) = "SELECT * FROM R060311 WHERE ID='" & strUserNum & "' ORDER BY R044001,R044011,R044003,r044004,R044006 "
   intI = 1
   Set adoRecordset = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      Printer.Orientation = 2
      DoEvents
      Page = 0
      SavDay1 = " "
      SavDay2 = " "
      SavDay3 = "   "
      SavDay4 = "   "
      SavDay5 = "" 'Add by Morgan 2010/4/20
      With adoRecordset
         .MoveFirst
         Do While .EOF = False
            For i = 0 To 9
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            'Add by Morgan 2007/11/27 案件性質前面加代碼
            strTemp(5) = Mid(strTemp(5), 5)
            If strTemp(0) <> SavDay1 Or strTemp(1) <> SavDay2 Then
            
               If Page > 0 Then
                  PrintEnd
                  If strTemp(0) <> SavDay1 And txt1(12) = "" Then
                     SavDay2 = ""
                     SavDay5 = ""
                     PrintEnd
                  End If
               End If
                              
               'Add by Morgan 2010/4/20 檢查是否有無資料但有分配點數的情形
               If m_intR > 0 Then
                  strKeyNow = "" & .Fields("R044001") & .Fields("R044011")
                  strExc(0) = "select distinct A0902 R044001,st02 R044002,st01 R044011 from R060311_1,staff,acc090 where ID='" & strUserNum & "' and st01(+)=R01 and a0901(+)=ST15 and A0902||st01>'" & strKeyLast & " ' and A0902||st01<'" & strKeyNow & "' order by R044001,R044011"
                  intI = 1
                  Set rsQuery1 = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     With rsQuery1
                        Do While Not .EOF
                           SavDay1 = "" & .Fields("R044001")
                           SavDay2 = "" & .Fields("R044002")
                           SavDay3 = "   "
                           SavDay4 = "   "
                           SavDay5 = "" & .Fields("R044011")
                           
                           If txt1(12) = "Y" Or Page = 0 Then
                              PrintTitle
                           End If
                           
                           '明細
                           If txt1(12) = "Y" Then
                              iPrint = iPrint + 300
                              PrintSharePoint SavDay5
                           '統計
                           Else
                              PrintEnd
                           End If
                           .MoveNext
                        Loop
                     End With
                  End If
               End If
               'end 2010/4/20
               
               SavDay1 = "" & .Fields("R044001")
               SavDay2 = "" & .Fields("R044002")
               SavDay3 = "  "
               SavDay4 = "  "
               SavDay5 = "" & .Fields("R044011") 'Add by Morgan 2010/4/20
               
               If txt1(12) = "Y" Or Page = 0 Then
                  PrintTitle
               End If
               strKeyLast = strKeyNow
            End If
            'Add by Morgan 2007/9/26 加可不印明細只印統計
            If txt1(12) = "Y" Then
               If SavDay3 = strTemp(2) Then
                   strTemp(2) = ""
                   If SavDay4 = strTemp(3) Then
                     strTemp(3) = ""
                     strTemp(4) = ""
                   Else
                      SavDay4 = strTemp(3)
                   End If
               Else
                  SavDay3 = strTemp(2)
                  SavDay4 = strTemp(3)
               End If
               strTemp(4) = StrToStr(strTemp(4), 14)
               strTemp(5) = StrToStr(strTemp(5), 4)
               strTemp(6) = StrToStr(strTemp(6), 4)
               strTemp(7) = StrToStr(strTemp(7), 8)
               PrintDatil
            End If
            .MoveNext
         Loop
      End With
      PrintEnd
      
      'Add by Morgan 2010/4/20 檢查是否有無資料但有分配點數的情形
      If m_intR > 0 And txt1(12) = "Y" Then
         strExc(0) = "select distinct A0902 R044001,st02 R044002,st01 R044011 from R060311_1,staff,acc090 where ID='" & strUserNum & "' and st01(+)=R01 and a0901(+)=ST15 and A0902||st01>'" & strKeyLast & " ' order by R044001,R044011"
         intI = 1
         Set rsQuery1 = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            With rsQuery1
               Do While Not .EOF
                  SavDay1 = "" & .Fields("R044001")
                  SavDay2 = "" & .Fields("R044002")
                  SavDay3 = "   "
                  SavDay4 = "   "
                  SavDay5 = "" & .Fields("R044011")
                  PrintTitle
                  iPrint = iPrint + 300
                  PrintSharePoint SavDay5
                  .MoveNext
               Loop
            End With
         End If
      End If
      'end 2010/4/20
                  
      'Add by Morgan 2007/12/4 不印明細時加印合計
      If txt1(12) = "" Then
         SavDay1 = "合計"
         SavDay2 = "合計"
         SavDay5 = "" 'Add by Morgan 2010/4/20
         PrintEnd 1
      End If
      'end 2007/12/4
      
      Printer.EndDoc
   End If
   Set rsQuery1 = Nothing
End Sub

Sub PrintDatil()
   NewLine
   For i = 2 To 8
       Printer.CurrentX = PLeft(i)
       Printer.CurrentY = iPrint
       Printer.Print strTemp(i)
   Next i
   Printer.CurrentX = PLeft(9) + 1000 - (Printer.TextWidth(Format(strTemp(9), "#,##0.00")))
   Printer.CurrentY = iPrint
   Printer.Print Format(strTemp(9), "#,##0.00")
   m_bFirstLine = False
End Sub

Sub GetPleft()
   Erase PLeft
   PLeft(0) = 500
   PLeft(1) = 500
   PLeft(2) = 500
   PLeft(3) = 2500
   PLeft(4) = 4500
   PLeft(5) = 8200
   PLeft(6) = 10000
   PLeft(7) = 12200
   PLeft(8) = 13400
   PLeft(9) = 14500
End Sub

Sub PrintTitle()

   Dim strTitle As String
   
   Page = Page + 1
   If Page > 1 Then
      Printer.NewPage
   End If
   
   If txt1(12) = "Y" Then
      strTitle = "國外部專利處業務組收文明細表"
   Else
      strTitle = "國外部專利處業務組收文統計表"
   End If
   
   GetPleft
   iPrint = 500
   With Printer
      .Font.Size = 22
      .Font.Bold = True
      .Font.Underline = True
      .CurrentX = .ScaleWidth / 2 - (.TextWidth(strTitle) / 2)
      .CurrentY = iPrint
      Printer.Print strTitle
      iPrint = iPrint + 500
      .Font.Size = 12
      .Font.Bold = False
      .Font.Underline = False
      strExc(0) = "收文日：" & Format(ChangeTStringToTDateString(txt1(2)), "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(3))
      .CurrentX = .ScaleWidth / 2 - (.TextWidth(strExc(0)) / 2)
      .CurrentY = iPrint
      Printer.Print strExc(0)
      
      iPrint = iPrint + 300
      .CurrentX = 500
      .CurrentY = iPrint
      Printer.Print "列印人：" & strUserName
      .CurrentX = 13000
      .CurrentY = iPrint
      Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
      iPrint = iPrint + 300
      .CurrentX = 13000
      .CurrentY = iPrint
      Printer.Print "頁    次：" & str(Page)
      
      iPrint = iPrint + 300
      .CurrentX = 500
      .CurrentY = iPrint
      Printer.Print "業務區：" & SavDay1
         
      If txt1(12) = "Y" Then
         .CurrentX = 3000
         .CurrentY = iPrint
         Printer.Print "智權人員：" & SavDay2
      End If
      
      'Add by Morgan 2008/1/23
      .CurrentX = 6000
      .CurrentY = iPrint
      Printer.Print "PS：每一筆請款單點數計入該請款單之最後收文承辦智權同仁"
   End With
      
   PrintLine 1
   If txt1(12) = "Y" Then 'Add by Morgan 2007/9/26 判斷是否要印明細
      PrintTitle3
   Else
      PrintTitle4
   End If
   PrintLine 1
   m_bFirstLine = True
End Sub

Sub PrintEnd(Optional p_iOpt As Integer = 0)
   Dim stCon As String, stCon1 As String
   Dim iRows As Integer
   
   If m_bFirstLine = True Then
      m_bFirstLine = False
   Else
      PrintLine
   End If
   
   stCon = ""
   If p_iOpt = 0 Then 'Add by Morgan 2007/12/4 加可印合計
      If Len(SavDay2) <> 0 Then
         stCon = stCon & " AND R044002='" & SavDay2 & "'"
      End If
   End If
   
   strSql = "SELECT SUM(R044010) FROM R060311 WHERE ID='" & strUserNum & "'" & stCon
   
   CheckOC2
   With adoRecordset1
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount <> 0 Then
         strExc(1) = "" & .Fields(0)
      Else
         strExc(1) = "0"
      End If
      
   End With
   
   NewLine
   If txt1(12) = "" Then
      Printer.CurrentX = PLeft(0)
      Printer.CurrentY = iPrint
      If SavDay2 = "" Then
         Printer.Print SavDay1
      Else
         Printer.Print SavDay2
      End If
   End If
   Printer.CurrentX = PLeft(9) + 1200 - (Printer.TextWidth(Format(strExc(1), "#,##0.00")))
   Printer.CurrentY = iPrint
   Printer.Print Format(strExc(1), "#,##0.00")
   iPrint = iPrint - 300
   iRows = 0
   
   strSql = "SELECT R044006,COUNT(*) FROM R060311 WHERE ID='" & strUserNum & "'" & stCon & " GROUP BY R044006 "
   intI = 1
   Set adoRecordset1 = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      With adoRecordset1
      .MoveFirst
      Do While .EOF = False
         For i = 0 To 4
            StrTemp4(i) = CheckStr(.Fields(0))
            StrTemp4(i) = Mid(StrTemp4(i), 5)
            StrTemp5(i) = CheckStr(.Fields(1))
            .MoveNext
            If .EOF = True Then
                For j = i + 1 To 4
                    StrTemp4(j) = ""
                    StrTemp5(j) = ""
                Next j
                Exit For
            End If
         Next i
         PrintSubTot1
         'Add by Morgan 2010/5/10 列印分配點數小計
         iRows = iRows + 1
         If iRows = 2 Then
            If txt1(12) = "" Then
               PrintShareTot SavDay5
            End If
         End If
      Loop
      End With
   End If
   'Add by Morgan 2010/5/10 列印分配點數小計
   If iRows < 2 Then
      If txt1(12) = "" Then
         PrintShareTot SavDay5, True
      End If
   End If
   
   'Add by Morgan 2010/4/19
   If m_strSharePointVTB <> "" Then
      If txt1(12) = "Y" Then
         PrintSharePoint SavDay5
      End If
   End If
End Sub
'Add by Morgan 2010/5/10
'列印分配點數合計
Private Sub PrintShareTot(Optional p_ID As String, Optional bolNewLine As Boolean)
   Dim stSQL As String, intR As Integer, stCon As String
   Dim adoRst As ADODB.Recordset
   
   If p_ID <> "" Then
      stCon = " and R01='" & p_ID & "'"
   End If
   
   stSQL = "select NVL(sum(R03),0) C1 from R060311_1 where ID='" & strUserNum & "'" & stCon
   intR = 1
   Set adoRst = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      If bolNewLine Then NewLine 300
      With adoRst
      If .Fields(0) > 0 Then
         strExc(1) = "分配點數："
         Printer.CurrentX = PLeft(9) - Printer.TextWidth(strExc(1))
         Printer.CurrentY = iPrint
         Printer.Print strExc(1)
         
         strExc(1) = Format(.Fields(0), "0.00")
         Printer.CurrentX = PLeft(9) + 1200 - Printer.TextWidth(strExc(1))
         Printer.CurrentY = iPrint
         Printer.Print strExc(1)
      End If
      End With
   End If
   
   Set adoRst = Nothing
End Sub

Sub PrintSubTot1()
   Dim lngX0 As Long
   
   NewLine 300
   
   If txt1(12) = "Y" Then
      lngX0 = PLeft(1)
   Else
      lngX0 = 2000
   End If
   
   For j = 0 To 4
      With Printer
         .CurrentX = lngX0 + (j * 2300)
         .CurrentY = iPrint
         Printer.Print StrConv(MidB(StrConv(StrTemp4(j), vbFromUnicode), 1, 14), vbUnicode)
         .CurrentX = lngX0 + ((j + 1) * 2300) - 400 - .TextWidth(StrTemp5(j))
         .CurrentY = iPrint
         Printer.Print StrTemp5(j)
      End With
   Next j
   m_bFirstLine = False
End Sub

Private Sub NewLine(Optional iHeight As Integer = 400)
   iPrint = iPrint + iHeight
   If iPrint > Printer.ScaleHeight - 800 Then
      PrintTitle
      iPrint = iPrint + iHeight
   End If
End Sub

Private Sub PrintTitle3()
   iPrint = iPrint + 300
   With Printer
      .CurrentX = PLeft(2)
      .CurrentY = iPrint
      Printer.Print "收文日"
      .CurrentX = PLeft(3)
      .CurrentY = iPrint
      Printer.Print "本所案號"
      .CurrentX = PLeft(4)
      .CurrentY = iPrint
      Printer.Print "案件名稱"
      .CurrentX = PLeft(5)
      .CurrentY = iPrint
      Printer.Print "案件性質"
      .CurrentX = PLeft(6)
      .CurrentY = iPrint
      Printer.Print "國    籍"
      .CurrentX = PLeft(7)
      .CurrentY = iPrint
      Printer.Print "本所期限"
      .CurrentX = PLeft(8)
      .CurrentY = iPrint
      Printer.Print "發文日"
      .CurrentX = PLeft(9)
      .CurrentY = iPrint
      Printer.Print "請款點數"
   End With
End Sub

Private Sub PrintTitle4()
   iPrint = iPrint + 300
   With Printer
      .CurrentX = PLeft(1)
      .CurrentY = iPrint
      Printer.Print "智權人員"
      .CurrentX = 2000
      .CurrentY = iPrint
      Printer.Print "各項收文案件總數"
      .CurrentX = PLeft(9)
      .CurrentY = iPrint
      Printer.Print "請款總點數"
   End With
End Sub

Private Sub PrintLine(Optional iType As Integer = 0)
   iPrint = iPrint + 300
   With Printer
      .CurrentX = PLeft(1)
      .CurrentY = iPrint
      If iType = 1 Then
         Printer.Line (PLeft(1), iPrint + 150)-(PLeft(1) + .TextWidth(String(200, "-")), iPrint + 150)
      Else
         Printer.Print String(200, "-")
      End If
   End With
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   'txt1(0) = GetSystemKindByNick 'Remove by Morgan 2007/11/27 不再預設
   txt1(13) = "F20"
   txt1(14) = "F29"
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm060311 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   CloseIme
   TextInverse txt1(Index)
End Sub

Private Sub txt1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
       cmdok(0).SetFocus
   End If
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   Select Case Index
      Case 1
         If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 8 Then
            KeyAscii = 0
         End If
      Case 12
         If KeyAscii <> 89 And KeyAscii <> 8 Then
            KeyAscii = 0
         End If
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   If txt1(Index) = "" Then Exit Sub
   Select Case Index
      Case 2, 3
         Cancel = Not ChkDate(txt1(Index))
               
      Case 6
         lbl1.Caption = GetPrjSales(txt1(Index), "智權人員")
         If Me.txt1(Index).Text <> "" Then
            If Me.txt1(Index).Text = Me.lbl1.Caption Then
               Me.lbl1.Caption = ""
               Cancel = True
            End If
         End If

   End Select
End Sub

Private Function TxtValidate() As Boolean
   
   Dim oText As TextBox, bCancel As Boolean
   For Each oText In txt1
      txt1_Validate oText.Index, bCancel
      If bCancel = True Then
         txt1(oText.Index).SetFocus
         txt1_GotFocus oText.Index
         Exit Function
      End If
   Next
   
'Remove by Morgan 2007/11/27 不再控管權限
'   If Len(txt1(0)) = 0 Then
'      s = MsgBox("系統類別不可空白!!", , "USER 輸入錯誤")
'      txt1(0).SetFocus
'      Exit Sub
'   End If
         
'   strTemp1 = Split(Replace(UCase(GetSystemKindByNick), ",,", ""), ",")
'   strTemp2 = Split(Replace(UCase(txt1(0)), ",,", ""), ",")
'   For i = 0 To UBound(strTemp2)
'      s = 0
'      For j = 0 To UBound(strTemp1)
'         If strTemp1(j) = strTemp2(i) Then
'            s = 1
'            Exit For
'          End If
'      Next j
'      If s = 0 Then
'         s = MsgBox(strUserName & " 沒有 " & strTemp2(i) & " 的權限 ", , "權限問題")
'         txt1(0).SetFocus
'         txt1_GotFocus 0
'         Exit Function
'      End If
'   Next i
   
   If Me.txt1(13).Text <> "" And Me.txt1(14).Text <> "" Then
      If Me.txt1(13).Text > Me.txt1(14).Text Then
         MsgBox "業務區別範圍輸入錯誤!!!", vbExclamation + vbOKOnly
         Me.txt1(13).SetFocus
         txt1_GotFocus 13
         Exit Function
      End If
   End If
   
   If Len(txt1(3)) = 0 Then
       s = MsgBox("日期區間不可空白!!", , "USER 輸入錯誤")
       If Len(txt1(2)) = 0 Then txt1(2).SetFocus
       Exit Function
   End If
   
   If Me.txt1(2).Text <> "" And Me.txt1(3).Text <> "" Then
      If Val(Me.txt1(2).Text) > Val(Me.txt1(3).Text) Then
         MsgBox "日期範圍輸入錯誤!!!", vbExclamation + vbOKOnly
         Me.txt1(2).SetFocus
         txt1_GotFocus 2
         Exit Function
      End If
   End If
            
   If Me.txt1(4).Text <> "" And Me.txt1(5).Text <> "" Then
      If Me.txt1(4).Text > Me.txt1(5).Text Then
         MsgBox "案件性質範圍輸入錯誤!!!", vbExclamation + vbOKOnly
         Me.txt1(4).SetFocus
         txt1_GotFocus 4
         Exit Function
      End If
   End If
   
   '申請人
   If Len(txt1(7)) <> 0 Then
      If Left(txt1(7), 6) <> Left(txt1(8), 6) Then
          s = MsgBox("申請人前 6 碼必須相同", , "USER 輸入錯誤")
          txt1(7).SetFocus
          txt1_GotFocus 7
          Exit Function
      End If
   End If
   If Me.txt1(7).Text <> "" And Me.txt1(8).Text <> "" Then
      If Me.txt1(7).Text > Me.txt1(8).Text Then
         MsgBox "申請人範圍輸入錯誤!!!", vbExclamation + vbOKOnly
         Me.txt1(7).SetFocus
         txt1_GotFocus 7
         Exit Function
      End If
   End If
            
   '代理人
   If Len(txt1(9)) <> 0 Then
      If Left(txt1(9), 6) <> Left(txt1(10), 6) Then
          s = MsgBox("代理人前 6 碼必須相同", , "USER 輸入錯誤")
          txt1(9).SetFocus
          txt1_GotFocus 9
          Exit Function
      End If
   End If
   If Me.txt1(9).Text <> "" And Me.txt1(10).Text <> "" Then
      If Me.txt1(9).Text > Me.txt1(10).Text Then
         MsgBox "代理人範圍輸入錯誤!!!", vbExclamation + vbOKOnly
         Me.txt1(9).SetFocus
         txt1_GotFocus 9
         Exit Function
      End If
   End If
   
   TxtValidate = True
End Function

'Add by Morgan 2010/4/19
Private Sub ProcessNew()

   Dim strSQL1 As String, strSQL2 As String, StrSQL3 As String, stCon1N0 As String
   Dim strVTB As String
   Dim stConPA As String, stConSP As String
   Dim rsQuery1 As ADODB.Recordset
   Dim strKeyNow As String, strKeyLast As String
   
   ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/10 清除查詢印表記錄檔欄位
   cnnConnection.Execute "DELETE FROM R060311 WHERE ID='" & strUserNum & "' "
   strSQL1 = ""
   strSQL2 = ""
   'Modify by Morgan 2007/11/27 只抓A類收文
   StrSQL3 = " AND CP57 IS NULL AND CP09<'B'"
   stCon1N0 = ""
   stConPA = ""
   stConSP = ""
   
   '系統別
   If Len(txt1(0)) <> 0 Then
      strSQL1 = strSQL1 + " AND CP01 IN (" & SQLGrpStr(txt1(0), 1) & ") "
      strSQL2 = strSQL2 + " AND CP01 IN (" & SQLGrpStr(txt1(0), 5) & ") "
      stCon1N0 = stCon1N0 & " AND CP01 IN (" & SQLGrpStr(txt1(0), 1) & "," & SQLGrpStr(txt1(0), 5) & ")"
      pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(0) 'Add By Sindy 2010/12/10
   End If
   
   '業務區
   If Len(txt1(13)) <> 0 Then
       StrSQL3 = StrSQL3 + " AND CP12>='" & txt1(13) & "' "
   End If
   If Len(txt1(14)) <> 0 Then
       StrSQL3 = StrSQL3 + " AND CP12<='" & txt1(14) & "' "
   End If
   If Len(txt1(13)) <> 0 Or Len(txt1(14)) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label1(1) & txt1(13) & "-" & txt1(14) 'Add By Sindy 2010/12/10
   End If
   
   '收文日
   If Len(txt1(2)) <> 0 Then
     StrSQL3 = StrSQL3 + " AND CP05>=" & Val(ChangeTStringToWString(txt1(2))) & " "
   End If
   StrSQL3 = StrSQL3 + " AND CP05<=" & Val(ChangeTStringToWString(txt1(3)))
   pub_QL05 = pub_QL05 & ";" & Label1(2) & txt1(2) & "-" & txt1(3) 'Add By Sindy 2010/12/10
   
   '案件性質
   If Len(txt1(4)) <> 0 Then
       StrSQL3 = StrSQL3 + " AND CP10>='" & txt1(4) & "' "
   End If
   If Len(txt1(5)) <> 0 Then
       StrSQL3 = StrSQL3 + " AND CP10<='" & txt1(5) & "' "
   End If
   If Len(txt1(4)) <> 0 Or Len(txt1(5)) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label1(3) & txt1(4) & "-" & txt1(5) 'Add By Sindy 2010/12/10
   End If
   
   '智權人員
   If Len(txt1(6)) <> 0 Then
       StrSQL3 = StrSQL3 + " AND CP13='" & txt1(6) & "' "
       pub_QL05 = pub_QL05 & ";" & Label1(4) & txt1(6) & lbl1 'Add By Sindy 2010/12/10
   End If
         
   '申請人
   If Len(Trim(txt1(7))) <> 0 And Len(Trim(txt1(8))) <> 0 Then
       stConPA = stConPA + " AND ((PA26>='" & GetNewFagent(txt1(7)) & "' AND PA26<='" & GetNewFagent(txt1(8)) & "') OR (PA27>='" & GetNewFagent(txt1(7)) & "' AND PA27<='" & GetNewFagent(txt1(8)) & "') OR (PA28>='" & GetNewFagent(txt1(7)) & "' AND PA28<='" & GetNewFagent(txt1(8)) & "') OR (PA29>='" & GetNewFagent(txt1(7)) & "' AND PA29<='" & GetNewFagent(txt1(8)) & "') OR (PA30>='" & GetNewFagent(txt1(7)) & "' AND PA30<='" & GetNewFagent(txt1(8)) & "')) "
       stConSP = stConSP + " AND ((SP08>='" & GetNewFagent(txt1(7)) & "' AND SP08<='" & GetNewFagent(txt1(8)) & "') OR (SP58<='" & GetNewFagent(txt1(7)) & "' AND SP58<='" & GetNewFagent(txt1(8)) & "') OR (SP59>='" & GetNewFagent(txt1(7)) & "' AND SP59<='" & GetNewFagent(txt1(8)) & "')) "
   Else
       If Len(Trim(txt1(7))) <> 0 And Len(Trim(txt1(8))) = 0 Then
           stConPA = stConPA + " AND (PA26>='" & GetNewFagent(txt1(7)) & "' OR PA27>='" & GetNewFagent(txt1(7)) & "' OR PA28>='" & GetNewFagent(txt1(7)) & "' OR PA29>='" & GetNewFagent(txt1(7)) & "' OR PA30>='" & GetNewFagent(txt1(7)) & "') "
           stConSP = stConSP + " AND (SP08>='" & GetNewFagent(txt1(7)) & "' OR SP58>='" & GetNewFagent(txt1(7)) & "' OR SP59>='" & GetNewFagent(txt1(7)) & "') "
       Else
           If Len(Trim(txt1(7))) = 0 And Len(Trim(txt1(8))) <> 0 Then
               stConPA = stConPA + " AND (PA26<='" & GetNewFagent(txt1(8)) & "' OR PA27<='" & GetNewFagent(txt1(8)) & "' OR PA28<='" & GetNewFagent(txt1(8)) & "' OR PA29<='" & GetNewFagent(txt1(8)) & "' OR PA30<='" & GetNewFagent(txt1(8)) & "') "
               stConSP = stConSP + " AND (SP08<='" & GetNewFagent(txt1(8)) & "' OR SP58<='" & GetNewFagent(txt1(8)) & "' OR SP59<='" & GetNewFagent(txt1(8)) & "') "
           End If
       End If
   End If
   If Len(Trim(txt1(7))) <> 0 Or Len(Trim(txt1(8))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label1(5) & txt1(7) & "-" & txt1(8) 'Add By Sindy 2010/12/10
   End If
   
   '代理人
   If Len(Trim(txt1(9))) <> 0 And Len(Trim(txt1(10))) <> 0 Then
       stConPA = stConPA + " AND PA75>='" & GetNewFagent(txt1(9)) & "' AND PA75<='" & GetNewFagent(txt1(10)) & "' "
       stConSP = stConSP + " AND SP26>='" & GetNewFagent(txt1(9)) & "' AND SP26<='" & GetNewFagent(txt1(10)) & "' "
   Else
       If Len(Trim(txt1(9))) <> 0 And Len(Trim(txt1(10))) = 0 Then
           stConPA = stConPA + " AND PA75>='" & GetNewFagent(txt1(9)) & "' "
           stConSP = stConSP + " AND SP26>='" & GetNewFagent(txt1(9)) & "' "
       Else
           If Len(Trim(txt1(9))) = 0 And Len(Trim(txt1(10))) <> 0 Then
               stConPA = stConPA + " AND PA75<='" & GetNewFagent(txt1(10)) & "' "
               stConSP = stConSP + " AND SP26<='" & GetNewFagent(txt1(10)) & "' "
           End If
       End If
   End If
   If Len(Trim(txt1(9))) <> 0 Or Len(Trim(txt1(10))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label1(6) & txt1(9) & "-" & txt1(10) 'Add By Sindy 2010/12/10
   End If
   
   If txt1(12) = "Y" Then
      pub_QL05 = pub_QL05 & ";" & Label1(8) & txt1(12) 'Add By Sindy 2010/12/10
   End If
   
   stCon1N0 = stCon1N0 & StrSQL3
   
   If stConPA <> "" Then
      stCon1N0 = stCon1N0 & " AND ( EXISTS(SELECT * FROM PATENT WHERE PA01=CP01 AND PA02=CP02 AND PA03=CP03 AND PA04=CP04 " & stConPA & ")"
      stCon1N0 = stCon1N0 & " OR EXISTS(SELECT * FROM SERVICEPRACTICE WHERE SP01=CP01 AND SP02=CP02 AND SP03=CP03 AND SP04=CP04 " & stConSP & ")"
      stCon1N0 = stCon1N0 & ")"
   End If

   'Add by Morgan 20104/20 分配點數語法
   m_strSharePointVTB = "select '" & strUserNum & "',A1N04,A1N03,A1N05,cp01,cp02,cp03,cp04,cp10,cp05" & _
      " From caseprogress,acc1n0,acc1k0" & _
      " Where cp60>'X'" & stCon1N0 & _
      " and a1k01(+)=cp60 and a1k25 is null and a1n01(+)=cp60 and a1n02(+)='1' and a1n03(+)=cp09 and a1n05>0" & _
      " and a1n04<>cp13"
   
   cnnConnection.Execute "DELETE FROM R060311_1 WHERE ID='" & strUserNum & "'"
   strSql = "INSERT INTO R060311_1(ID,R01,R02,R03,R04,R05,R06,R07,R08,R09)" & m_strSharePointVTB
   cnnConnection.Execute strSql, m_intR
   
   '排除銷帳及作廢的請款單
   'Modified by Lydia 2018/06/05 修改顯示案件性質 '020',CPM04,CPM03 => '000',CPM03,CPM04
   strExc(0) = "select nvl(a0902,a0903),ST02," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(PA05,NVL(PA06,PA07)),RPAD(CP10,4,' ')||decode(pa09,'000',CPM03,CPM04),NVL(NA03,NA04)," & SQLDate("CP06") & "," & SQLDate("CP27") & ",a1n05,CP09,CP60,cp13 FROM CASEPROGRESS,PATENT,STAFF,ACC1K0,ACC090,NATION,CASEPROPERTYMAP,FAGENT,acc1n0 WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND PA01 IS NOT NULL AND CP01=CPM01(+) AND CP10=CPM02(+) AND FA10=NA01(+) AND SUBSTR(PA75,1,8) = FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) AND CP13=ST01(+) AND CP60=A1K01(+) and a1k12 is null and a1k25 is null AND CP12=A0901(+) and a1n01(+)=cp60 and a1n02(+)='1' and a1n03(+)=cp09 and a1n04(+)=cp13 " & strSQL1 & StrSQL3 & stConPA & _
      " UNION ALL select nvl(a0902,a0903),ST02," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),RPAD(CP10,4,' ')||decode(sp09,'000',CPM03,CPM04),NVL(NA03,NA04)," & SQLDate("CP06") & "," & SQLDate("CP27") & ",a1n05,CP09,CP60,cp13 FROM CASEPROGRESS,SERVICEPRACTICE,STAFF,ACC1K0,ACC090,NATION,CASEPROPERTYMAP,FAGENT,acc1n0 WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND SP01 IS NOT NULL AND CP01=CPM01(+) AND CP10=CPM02(+) AND FA10=NA01(+) AND SUBSTR(SP26,1,8) = FA01(+) AND SUBSTR(SP26,9,1)=FA02(+) AND CP13=ST01(+) AND CP60=A1K01(+) and a1k12 is null and a1k25 is null AND CP12=A0901(+) and a1n01(+)=cp60 and a1n02(+)='1' and a1n03(+)=cp09 and a1n04(+)=cp13 " & strSQL2 & StrSQL3 & stConSP
   'end 2018/06/05
   intI = 1
   Set adoRecordset = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With adoRecordset
          InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/10
          .MoveFirst
          Do While .EOF = False
              For i = 0 To 9
                  strTemp(i) = CheckStr(.Fields(i))
              Next i
              strSql = "INSERT INTO R060311(R044001,R044002,R044003,R044004,R044005,R044006,R044007,R044008,R044009,R044010,id,R044011)" & _
               " VALUES('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & ChgSQL(strTemp(8)) & "','" & ChgSQL(strTemp(9)) & "','" & strUserNum & "','" & .Fields("cp13") & "') "
              cnnConnection.Execute strSql
              .MoveNext
          Loop
      End With
      PrintData
      ShowPrintOk
   'Add by Morgan 2010/4/20
   Else
      '跑明細報表且有分配點數資料
      If txt1(12) = "Y" And m_intR > 0 Then
         strExc(0) = "select distinct A0902 R044001,st02 R044002,st01 R044011 from R060311_1,staff,acc090 where ID='" & strUserNum & "' and st01(+)=R01 and a0901(+)=ST15 order by R044001,R044011"
         intI = 1
         Set rsQuery1 = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            Printer.Orientation = 2
            DoEvents
            Page = 0
            SavDay1 = " "
            SavDay2 = " "
            SavDay5 = ""
            With rsQuery1
            InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/10
            .MoveFirst
            Do While Not .EOF
               SavDay1 = CheckStr(.Fields("R044001"))
               SavDay2 = CheckStr(.Fields("R044002"))
               SavDay3 = "   "
               SavDay4 = "   "
               SavDay5 = "" & .Fields("R044011")
               PrintTitle
               iPrint = iPrint + 300
               PrintSharePoint SavDay5
               .MoveNext
            Loop
            End With
            Printer.EndDoc
            ShowPrintOk
         End If
      Else
         InsertQueryLog (0) 'Add By Sindy 2010/12/10
         ShowNoData
      End If
   End If
   Set rsQuery1 = Nothing
End Sub

'Add by Morgan 2010/4/19
'列印分配點數
Private Sub PrintSharePoint(Optional p_ID As String)
   Dim lngX0 As Long, lngX As Long, dblTot As Double
   Dim stCon As String
   
   'Modified by Lydia 2018/06/05 修改顯示案件性質 '020',CPM04,CPM03 => '000',CPM03,CPM04
   strExc(0) = "select sqldatet(R09) C1" & _
      ",R04||'-'||R05||decode(R06||R07,'000','','-'||R06||'-'||R07) C2" & _
      ",substrb(decode(nvl(pa09,sp09),'000',cpm03,cpm04),1,8) C3" & _
      ",R03 C4" & _
      " from R060311_1,patent,servicepractice,casepropertymap" & _
      " WHERE ID='" & strUserNum & "' and R01='" & p_ID & "'" & _
      " AND cpm01(+)=R04 and cpm02(+)=R08" & _
      " and pa01(+)=R04 and pa02(+)=R05 and pa03(+)=R06 and pa04(+)=R07" & _
      " and sp01(+)=R04 and sp02(+)=R05 and sp03(+)=R06 and sp04(+)=R07"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      NewLine 600
      strExc(1) = Printer.FontName
      Printer.FontName = "細明體"
      
      lngX0 = PLeft(0)
      Printer.CurrentX = lngX0
      Printer.CurrentY = iPrint
      Printer.Print "請款單分配點數："
      Printer.FontName = strExc(1)
      
      NewLine

      strExc(1) = Printer.FontName
      Printer.FontName = "細明體"
      
      lngX = lngX0
      Printer.CurrentX = lngX
      Printer.CurrentY = iPrint
      Printer.Print "收文日"

      lngX = lngX + Printer.TextWidth(String(11, "@"))
      Printer.CurrentX = lngX
      Printer.CurrentY = iPrint
      Printer.Print "本所案號"

      lngX = lngX + Printer.TextWidth(String(17, "@"))
      Printer.CurrentX = lngX
      Printer.CurrentY = iPrint
      Printer.Print "案件性質"

      lngX = lngX + Printer.TextWidth(String(10, "@"))
      Printer.CurrentX = lngX
      Printer.CurrentY = iPrint
      Printer.Print "點數"
      
      Printer.FontName = strExc(1)

      NewLine
      
      strExc(1) = Printer.FontName
      Printer.FontName = "細明體"
      Printer.CurrentX = lngX0
      Printer.CurrentY = iPrint
      Printer.Print String(45, "-")
      Printer.FontName = strExc(1)

      iPrint = iPrint - 100

      With RsTemp
      Do While Not .EOF
         NewLine
         strExc(1) = Printer.FontName
         Printer.FontName = "細明體"
         '日期
         lngX = lngX0
         Printer.CurrentX = lngX
         Printer.CurrentY = iPrint
         Printer.Print .Fields("C1")
         '本所案號
         lngX = lngX + Printer.TextWidth(String(11, "@"))
         Printer.CurrentX = lngX
         Printer.CurrentY = iPrint
         Printer.Print .Fields("C2")
         '案件性質
         lngX = lngX + Printer.TextWidth(String(17, "@"))
         Printer.CurrentX = lngX
         Printer.CurrentY = iPrint
         Printer.Print .Fields("C3")
         '點數
         lngX = lngX + Printer.TextWidth(String(10, "@"))
         Printer.CurrentX = lngX + Printer.TextWidth(String(7, "@")) - Printer.TextWidth(Format(.Fields("C4"), "0.00"))
         Printer.CurrentY = iPrint
         Printer.Print Format(.Fields("C4"), "0.00")
         Printer.FontName = strExc(1)
         
         dblTot = dblTot + Val("" & .Fields("C4"))
         .MoveNext
      Loop
      End With

      NewLine 300
      
      strExc(1) = Printer.FontName
      Printer.FontName = "細明體"
      Printer.CurrentX = lngX0
      Printer.CurrentY = iPrint
      Printer.Print String(45, "-")
      Printer.FontName = strExc(1)

      NewLine 300
      
      strExc(1) = Printer.FontName
      Printer.FontName = "細明體"
      Printer.CurrentX = lngX + Printer.TextWidth(String(7, "@")) - Printer.TextWidth(Format(dblTot, "0.00"))
      Printer.CurrentY = iPrint
      Printer.Print Format(dblTot, "0.00")
      Printer.FontName = strExc(1)
   End If
End Sub

