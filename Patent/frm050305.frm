VERSION 5.00
Begin VB.Form frm050305 
   BorderStyle     =   1  '單線固定
   Caption         =   "催審表"
   ClientHeight    =   2280
   ClientLeft      =   2265
   ClientTop       =   1965
   ClientWidth     =   3510
   ControlBox      =   0   'False
   LinkTopic       =   "Form5"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   3510
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   8
      Left            =   2580
      MaxLength       =   4
      TabIndex        =   8
      Top             =   1890
      Width           =   800
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   7
      Left            =   1290
      MaxLength       =   4
      TabIndex        =   7
      Top             =   1890
      Width           =   800
   End
   Begin VB.TextBox TXT1 
      Enabled         =   0   'False
      Height          =   264
      Index           =   5
      Left            =   1290
      MaxLength       =   7
      TabIndex        =   3
      Top             =   1200
      Width           =   800
   End
   Begin VB.TextBox TXT1 
      Enabled         =   0   'False
      Height          =   264
      Index           =   6
      Left            =   2580
      MaxLength       =   7
      TabIndex        =   4
      Top             =   1200
      Width           =   800
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   1
      Left            =   1290
      MaxLength       =   7
      TabIndex        =   1
      Top             =   870
      Width           =   800
   End
   Begin VB.OptionButton opt 
      Caption         =   "發文日期："
      Height          =   225
      Index           =   1
      Left            =   120
      TabIndex        =   14
      Top             =   1230
      Width           =   1215
   End
   Begin VB.OptionButton opt 
      Caption         =   "催審期限："
      Height          =   225
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   900
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.CommandButton CMDOK 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   2625
      TabIndex        =   10
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton CMDOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   1800
      TabIndex        =   9
      Top             =   45
      Width           =   800
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   4
      Left            =   2580
      MaxLength       =   4
      TabIndex        =   6
      Top             =   1530
      Width           =   800
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   3
      Left            =   1290
      MaxLength       =   4
      TabIndex        =   5
      Top             =   1530
      Width           =   800
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   2
      Left            =   2580
      MaxLength       =   7
      TabIndex        =   2
      Top             =   870
      Width           =   800
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   0
      Left            =   1290
      TabIndex        =   0
      Top             =   540
      Width           =   2085
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "案件性質："
      Height          =   180
      Left            =   390
      TabIndex        =   15
      Top             =   1950
      Width           =   900
   End
   Begin VB.Line Line1 
      X1              =   1485
      X2              =   2889
      Y1              =   2010
      Y2              =   2010
   End
   Begin VB.Line Line2 
      X1              =   1710
      X2              =   2934
      Y1              =   1350
      Y2              =   1350
   End
   Begin VB.Line Line4 
      X1              =   1485
      X2              =   2889
      Y1              =   1650
      Y2              =   1650
   End
   Begin VB.Line Line3 
      X1              =   1755
      X2              =   2979
      Y1              =   1005
      Y2              =   1005
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "申請國家："
      Height          =   180
      Left            =   390
      TabIndex        =   12
      Top             =   1590
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "系統類別："
      Height          =   180
      Left            =   390
      TabIndex        =   11
      Top             =   585
      Width           =   900
   End
End
Attribute VB_Name = "frm050305"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/12/28 申請案號欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
Option Explicit
Dim s As Integer, strSql As String, strTemp1 As Variant, strTemp2 As Variant, StrTest1 As String, StrTest2 As String
Dim i As Integer, j As Integer, k As Integer, strTemp(0 To 8) As String
Dim PLeft(0 To 8) As Integer, iPrint As Integer, iLine As Integer, Page As Integer, TmpArea As String
Dim St As String, GetTaiwanTodayDate1 As String
'Add By Cheng 2002/08/07
Dim SavDay(0 To 1) As String
'Add By Cheng 2002/09/12
Dim blnClkSure As Boolean '判斷是否按下確定按鈕

Private Sub cmdok_Click(Index As Integer)
Select Case Index
Case 0 '確定
      'Add By Cheng 2002/09/12
      blnClkSure = False
     If Len(txt1(0)) = 0 Then
        s = MsgBox("系統類別不可空白!!", , "USER 輸入錯誤")
        txt1(0).SetFocus
        Exit Sub
     Else
         'Add By Cheng 2002/08/05
         '選擇催審期限
         If Me.opt(0).Value Then
            If CheckIsTaiwanDate(Me.txt1(1).Text) = False Then
               Me.txt1(1).SetFocus
               Exit Sub
            End If
            If CheckIsTaiwanDate(Me.txt1(2).Text) = False Then
               Me.txt1(2).SetFocus
               Exit Sub
            End If
            If Val(Me.txt1(1).Text) > Val(Me.txt1(2).Text) Then
               MsgBox "催審期限區間輸入錯誤!!!", vbExclamation + vbOKOnly
               blnClkSure = True
               Me.txt1(1).SetFocus
               Exit Sub
            End If
         '選擇發文日期
         Else
            If CheckIsTaiwanDate(Me.txt1(5).Text) = False Then
               Me.txt1(5).SetFocus
               Exit Sub
            End If
            If CheckIsTaiwanDate(Me.txt1(6).Text) = False Then
               Me.txt1(6).SetFocus
               Exit Sub
            End If
            If Val(Me.txt1(5).Text) > Val(Me.txt1(6).Text) Then
               MsgBox "催審期限區間輸入錯誤!!!", vbExclamation + vbOKOnly
               blnClkSure = True
               Me.txt1(5).SetFocus
               Exit Sub
            End If
         End If
               
         'Modify by Cheng 2002/08/05
'         'Add By Cheng 2002/03/20
'         If PUB_CheckKeyInDate(Me.TXT1(1)) = -1 Then
'            Me.TXT1(1).SetFocus
'            txt1_GotFocus 1
'            Exit Sub
'         End If
'         If PUB_CheckKeyInDate(Me.TXT1(2)) = -1 Then
'            Me.TXT1(2).SetFocus
'            txt1_GotFocus 2
'            Exit Sub
'         End If
'        If Len(TXT1(2)) = 0 Then
'            s = MsgBox("催審期限不可空白!!", , "USER 輸入錯誤")
'            TXT1(1).SetFocus
'            txt1_GotFocus (1)
'            Exit Sub
'        Else
            'Add By Cheng 2002/09/12
            If Me.txt1(3).Text <> "" And Me.txt1(4).Text <> "" Then
               If Me.txt1(3).Text > Me.txt1(4).Text Then
                  MsgBox "申請國家範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                  blnClkSure = True
                  Me.txt1(3).SetFocus
                  txt1_GotFocus 3
                  Exit Sub
               End If
            End If
                        
            Me.Enabled = False
            StrMenu
            Me.Enabled = True
'        End If
     End If
Case 1
     Unload Me
Case Else
End Select
End Sub

Sub StrMenu()
'Add By Cheng 2002/08/07
Dim blnUpdate As Boolean '是否要更新下一程序的期限

ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/3 清除查詢印表記錄檔欄位
Screen.MousePointer = vbHourglass
StrTest1 = ""
StrTest2 = ""
'系統類別
If Len(St) <> 0 Then
   StrTest1 = StrTest1 & " and cp01 in (" & SQLGrpStr(St, 1) & ") "
   StrTest2 = StrTest2 & " and cp01 in (" & SQLGrpStr(St, 5) & ") "
   pub_QL05 = pub_QL05 & ";" & Label1 & txt1(0) 'Add By Sindy 2010/12/3
End If
'Modify By Cheng 2002/08/05
'催審期限
If Me.opt(0).Value Then
   If Len(txt1(1)) <> 0 Then
       StrTest1 = StrTest1 + " AND NP09>=" & Val(ChangeTStringToWString(txt1(1))) & " "
       StrTest2 = StrTest2 + " AND NP09>=" & Val(ChangeTStringToWString(txt1(1))) & " "
   End If
   If Len(txt1(2)) <> 0 Then
       StrTest1 = StrTest1 + " AND NP09<=" & Val(ChangeTStringToWString(txt1(2))) & " "
       StrTest2 = StrTest2 + " AND NP09<=" & Val(ChangeTStringToWString(txt1(2))) & " "
   End If
   If Len(txt1(1)) <> 0 Or Len(txt1(2)) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & opt(0).Caption & txt1(1) & "-" & txt1(2) 'Add By Sindy 2010/12/3
   End If
End If
'Add By Cheng 2002/08/05
'發文日期
If Me.opt(1).Value Then
   If Len(txt1(5)) <> 0 Then
       StrTest1 = StrTest1 + " AND CP27>=" & Val(ChangeTStringToWString(txt1(5))) & " "
       StrTest2 = StrTest2 + " AND CP27>=" & Val(ChangeTStringToWString(txt1(5))) & " "
   End If
   If Len(txt1(6)) <> 0 Then
       StrTest1 = StrTest1 + " AND CP27<=" & Val(ChangeTStringToWString(txt1(6))) & " "
       StrTest2 = StrTest2 + " AND CP27<=" & Val(ChangeTStringToWString(txt1(6))) & " "
   End If
   StrTest1 = StrTest1 & " AND CP24 IS NULL AND CP57 IS NULL AND CP09 <'C' "
   StrTest2 = StrTest2 & " AND CP24 IS NULL AND CP57 IS NULL AND CP09 <'C' "
   If Len(txt1(5)) <> 0 Or Len(txt1(6)) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & opt(1).Caption & txt1(5) & "-" & txt1(6) 'Add By Sindy 2010/12/3
   End If
End If
'申請國家
If Len(txt1(3)) <> 0 Then
    StrTest1 = StrTest1 + " AND SUBSTR(PA09,1,3)>='" & txt1(3) & "' "
    StrTest2 = StrTest2 + " AND SUBSTR(SP09,1,3)>='" & txt1(3) & "' "
End If
If Len(txt1(4)) <> 0 Then
    StrTest1 = StrTest1 + " AND SUBSTR(PA09,1,3)<='" & txt1(4) & "' "
    StrTest2 = StrTest2 + " AND SUBSTR(SP09,1,3)<='" & txt1(4) & "' "
End If
If Len(txt1(3)) <> 0 Or Len(txt1(4)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label3 & txt1(3) & "-" & txt1(4) 'Add By Sindy 2010/12/3
End If
'Add By Cheng 2003/04/15
'案件性質
If Len(txt1(7)) <> 0 Then
    StrTest1 = StrTest1 + " AND CP10>='" & txt1(7) & "' "
    StrTest2 = StrTest2 + " AND CP10>='" & txt1(7) & "' "
End If
If Len(txt1(8)) <> 0 Then
    StrTest1 = StrTest1 + " AND CP10<='" & txt1(8) & "' "
    StrTest2 = StrTest2 + " AND CP10<='" & txt1(8) & "' "
End If
If Len(txt1(7)) <> 0 Or Len(txt1(8)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label2 & txt1(7) & "-" & txt1(8) 'Add By Sindy 2010/12/3
End If
'Modify By Cheng 2002/08/05
'若選擇催審期限
If Me.opt(0).Value Then
   'StrSQL = "SELECT CP27 AS A,PA11,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS B,NVL(PA05,NVL(PA06,PA07)),NA03,PTM03,CPM03,S1.ST02,S2.ST02 FROM NEXTPROGRESS,CASEPROGRESS,PATENT,NATION,CASEPROPERTYMAP,PATENTTRADEMARK  WHERE NP07=411 AND NP06 IS NULL AND CP09=NP01 AND " & StrTest1
   'Modify By Cheng 2002/08/07
   '多顯示下一程序的期限(NP08,NP09), 總收文號(NP01), 下一程序(NP07), 序號(NP22)
'   strSQL = "select CP27,PA11,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,nvl(PA05,NVL(PA06,PA07)),NVL(NA03,NA04),decode(pa09,'000',ptm03,ptm04),decode(pa09,'000',cpm03,cpm04),S1.ST02,S2.ST02 FROM NEXTPROGRESS,CASEPROGRESS,PATENT,NATION,CASEPROPERTYMAP,PATENTTRADEMARKMAP,STAFF S1,STAFF S2 WHERE cp01=cpm01(+) AND cp10=cpm02(+) AND CP09=NP01 AND NP06 IS NULL AND NP07=411 AND NP02=PA01(+) AND NP03=PA02(+) AND NP04=PA03(+) AND NP05=PA04(+) AND (PA57<>'Y' OR PA57 IS NULL) AND PA09=NA01(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) AND CP14=S1.ST01(+) AND NP10=S2.ST02(+) " & StrTest1
'   strSQL = strSQL + " union all select CP27,SP11,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,nvl(SP05,NVL(SP06,SP07)),NVL(NA03,NA04),'',decode(sp09,'000',cpm03,cpm04),S1.ST02,S2.ST02 FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,NATION,CASEPROPERTYMAP,STAFF S1,STAFF S2 WHERE cp01=cpm01(+) AND cp10=cpm02(+) AND CP09=NP01 AND NP06 IS NULL AND NP07=305 AND NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+)  AND  (SP15<>'Y' OR SP15 IS NULL) AND SP09=NA01(+) AND CP14=S1.ST01(+) AND NP10=S2.ST02(+) " & StrTest2
'92.03.27 nick add left join
'   strSQL = "select CP27,PA11,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,nvl(PA05,NVL(PA06,PA07)),NVL(NA03,NA04),ptm03,decode(pa09,'000',cpm03,cpm04),S1.ST02,S2.ST02,NP08,NP09,NP01,NP07,NP22 FROM NEXTPROGRESS,CASEPROGRESS,PATENT,NATION,CASEPROPERTYMAP,PATENTTRADEMARKMAP,STAFF S1,STAFF S2 WHERE cp01=cpm01(+) AND cp10=cpm02(+) AND CP09=NP01 AND NP06 IS NULL AND NP07=411 AND NP02=PA01(+) AND NP03=PA02(+) AND NP04=PA03(+) AND NP05=PA04(+) AND (PA57<>'Y' OR PA57 IS NULL) AND PA09=NA01(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) AND CP14=S1.ST01(+) AND NP10=S2.ST01(+) " & StrTest1
'   strSQL = strSQL + " union all select CP27,SP11,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,nvl(SP05,NVL(SP06,SP07)),NVL(NA03,NA04),'',decode(sp09,'000',cpm03,cpm04),S1.ST02,S2.ST02,NP08,NP09,NP01,NP07,NP22 FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,NATION,CASEPROPERTYMAP,STAFF S1,STAFF S2 WHERE cp01=cpm01(+) AND cp10=cpm02(+) AND CP09=NP01 AND NP06 IS NULL AND NP07=305 AND NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+)  AND  (SP15<>'Y' OR SP15 IS NULL) AND SP09=NA01(+) AND CP14=S1.ST01(+) AND NP10=S2.ST01(+) " & StrTest2
   strSql = "select CP27,PA11,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,nvl(PA05,NVL(PA06,PA07)),NVL(NA03,NA04),ptm03,decode(pa09,'000',cpm03,cpm04),S1.ST02,S2.ST02,NP08,NP09,NP01,NP07,NP22 FROM NEXTPROGRESS,CASEPROGRESS,PATENT,NATION,CASEPROPERTYMAP,PATENTTRADEMARKMAP,STAFF S1,STAFF S2 WHERE cp01=cpm01(+) AND cp10=cpm02(+) AND NP01=cp09(+) AND NP06 IS NULL AND NP07=411 AND NP02=PA01(+) AND NP03=PA02(+) AND NP04=PA03(+) AND NP05=PA04(+) AND (PA57<>'Y' OR PA57 IS NULL) AND PA09=NA01(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) AND CP14=S1.ST01(+) AND NP10=S2.ST01(+) " & StrTest1
   'Added by Morgan 2012/5/25 +1603(發文日抓核准日)
   strSql = strSql & " union all select PA20 CP27,PA11,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,nvl(PA05,NVL(PA06,PA07)),NVL(NA03,NA04),ptm03,decode(pa09,'000',cpm03,cpm04),S1.ST02,S2.ST02,NP08,NP09,NP01,NP07,NP22 FROM NEXTPROGRESS,CASEPROGRESS,PATENT,NATION,CASEPROPERTYMAP,PATENTTRADEMARKMAP,STAFF S1,STAFF S2 WHERE np02=cpm01(+) AND np07=cpm02(+) AND NP01=cp09(+) AND NP06 IS NULL AND NP07=1603 AND NP02=PA01(+) AND NP03=PA02(+) AND NP04=PA03(+) AND NP05=PA04(+) AND (PA57<>'Y' OR PA57 IS NULL) AND PA09=NA01(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) AND CP14=S1.ST01(+) AND NP10=S2.ST01(+) " & StrTest1
   'end 2012/5/25
   strSql = strSql + " union all select CP27,SP11,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,nvl(SP05,NVL(SP06,SP07)),NVL(NA03,NA04),'',decode(sp09,'000',cpm03,cpm04),S1.ST02,S2.ST02,NP08,NP09,NP01,NP07,NP22 FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,NATION,CASEPROPERTYMAP,STAFF S1,STAFF S2 WHERE cp01=cpm01(+) AND cp10=cpm02(+) AND NP01=cp09(+) AND NP06 IS NULL AND NP07=305 AND NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+)  AND  (SP15<>'Y' OR SP15 IS NULL) AND SP09=NA01(+) AND CP14=S1.ST01(+) AND NP10=S2.ST01(+) " & StrTest2
   strSql = strSql + " ORDER BY CP27,A "
'若選擇發文日期(若案件國家檔的實查時間"CF05"為NULL, 則不管制)
Else
   'Modify By Cheng 2002/08/21
'   strSQL = "select CP27,PA11,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,nvl(PA05,NVL(PA06,PA07)),NVL(NA03,NA04),ptm03,decode(pa09,'000',cpm03,cpm04),S1.ST02,S2.ST02 FROM NEXTPROGRESS,CASEPROGRESS,PATENT P1,NATION,CASEPROPERTYMAP,PATENTTRADEMARKMAP,STAFF S1,STAFF S2,CASEFEE WHERE cp01=cpm01(+) AND cp10=cpm02(+) AND CP09=NP01 AND NP06 IS NULL AND NP07=411 AND NP02=PA01(+) AND NP03=PA02(+) AND NP04=PA03(+) AND NP05=PA04(+) AND (PA57<>'Y' OR PA57 IS NULL) AND PA09=NA01(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) AND CP14=S1.ST01(+) AND NP10=S2.ST02(+) AND CP01=CF01 AND PA09=CF02 AND CP10=CF03 AND CF05 IS NOT NULL " & StrTest1
'   strSQL = strSQL + " union all select CP27,SP11,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,nvl(SP05,NVL(SP06,SP07)),NVL(NA03,NA04),'',decode(sp09,'000',cpm03,cpm04),S1.ST02,S2.ST02 FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,NATION,CASEPROPERTYMAP,STAFF S1,STAFF S2,CASEFEE WHERE cp01=cpm01(+) AND cp10=cpm02(+) AND CP09=NP01 AND NP06 IS NULL AND NP07=305 AND NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+)  AND  (SP15<>'Y' OR SP15 IS NULL) AND SP09=NA01(+) AND CP14=S1.ST01(+) AND NP10=S2.ST02(+) AND CP01=CF01 AND SP09=CF02 AND CP10=CF03 AND CF05 IS NOT NULL " & StrTest2
'   strSQL = strSQL + " ORDER BY CP27,A "
'92.03.27 nick add left join
'   strSQL = "select CP27,PA11,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,nvl(PA05,NVL(PA06,PA07)),NVL(NA03,NA04),ptm03,decode(pa09,'000',cpm03,cpm04),S1.ST02,S2.ST02 FROM CASEPROGRESS,PATENT P1,NATION,CASEPROPERTYMAP,PATENTTRADEMARKMAP,STAFF S1,STAFF S2,CASEFEE WHERE cp01=cpm01(+) AND cp10=cpm02(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND (PA57<>'Y' OR PA57 IS NULL) AND PA09=NA01(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CF01 AND PA09=CF02 AND CP10=CF03 AND CF05 IS NOT NULL " & StrTest1
'   strSQL = strSQL + " union all select CP27,SP11,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,nvl(SP05,NVL(SP06,SP07)),NVL(NA03,NA04),'',decode(sp09,'000',cpm03,cpm04),S1.ST02,S2.ST02 FROM CASEPROGRESS,SERVICEPRACTICE,NATION,CASEPROPERTYMAP,STAFF S1,STAFF S2,CASEFEE WHERE cp01=cpm01(+) AND cp10=cpm02(+) AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+)  AND  (SP15<>'Y' OR SP15 IS NULL) AND SP09=NA01(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CF01 AND SP09=CF02 AND CP10=CF03 AND CF05 IS NOT NULL " & StrTest2
   strSql = "select CP27,PA11,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,nvl(PA05,NVL(PA06,PA07)),NVL(NA03,NA04),ptm03,decode(pa09,'000',cpm03,cpm04),S1.ST02,S2.ST02 FROM CASEPROGRESS,PATENT P1,NATION,CASEPROPERTYMAP,PATENTTRADEMARKMAP,STAFF S1,STAFF S2,CASEFEE WHERE cp01=cpm01(+) AND cp10=cpm02(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND (PA57<>'Y' OR PA57 IS NULL) AND PA09=NA01(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CF01(+) AND PA09=CF02(+) AND CP10=CF03(+) AND CF05 IS NOT NULL " & StrTest1
   strSql = strSql + " union all select CP27,SP11,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,nvl(SP05,NVL(SP06,SP07)),NVL(NA03,NA04),'',decode(sp09,'000',cpm03,cpm04),S1.ST02,S2.ST02 FROM CASEPROGRESS,SERVICEPRACTICE,NATION,CASEPROPERTYMAP,STAFF S1,STAFF S2,CASEFEE WHERE cp01=cpm01(+) AND cp10=cpm02(+) AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+)  AND  (SP15<>'Y' OR SP15 IS NULL) AND SP09=NA01(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CF01(+) AND SP09=CF02(+) AND CP10=CF03(+) AND CF05 IS NOT NULL " & StrTest2
   strSql = strSql + " ORDER BY CP27,A "
End If

CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    With adoRecordset
         InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/3
         'Remove by Morgan 2009/9/16 都不再詢問,改由每日批次報表控制週期
         ''Add By Cheng 2002/08/07
         ''若選擇催審期限, 加詢問使用者是否要更新
         'If Me.opt(0).Value Then
         '   If MsgBox("是否再管制三個月?", vbExclamation + vbYesNo) = vbYes Then
         '      blnUpdate = True
         '   Else
         '      blnUpdate = False
         '   End If
         ''其他選項一律不更新
         'Else
         '   blnUpdate = False
         'End If
        
        .MoveFirst
        Page = 1
        PrintTitle Page
        Do While .EOF = False
            If IsNull(.Fields(0)) Then
                strTemp(0) = ""
            Else
                strTemp(0) = Trim(ChangeTStringToTDateString(ChangeWStringToTString(.Fields(0))))
            End If
            If IsNull(.Fields(1)) Then
                strTemp(1) = ""
            Else
                strTemp(1) = Trim(.Fields(1))
            End If
            If IsNull(.Fields(2)) Then
                strTemp(2) = ""
            Else
                strTemp(2) = Trim(.Fields(2))
            End If
            If IsNull(.Fields(3)) Then
                strTemp(3) = ""
            Else
                strTemp(3) = Trim(.Fields(3))
            End If
            If IsNull(.Fields(4)) Then
                strTemp(4) = ""
            Else
                strTemp(4) = Trim(.Fields(4))
            End If
            If IsNull(.Fields(5)) Then
                strTemp(5) = ""
            Else
                strTemp(5) = Trim(.Fields(5))
            End If
            If IsNull(.Fields(6)) Then
                strTemp(6) = ""
            Else
                strTemp(6) = Trim(.Fields(6))
            End If
            If IsNull(.Fields(7)) Then
                strTemp(7) = ""
            Else
                strTemp(7) = Trim(.Fields(7))
            End If
            If IsNull(.Fields(8)) Then
                strTemp(8) = ""
            Else
                strTemp(8) = Trim(.Fields(8))
            End If
            strTemp(3) = StrToStr(strTemp(3), 10)
            
            'Remove by Morgan 2009/9/16 改由每日批次報表控制週期
            ''Add By Cheng 2002/08/07
            'If blnUpdate Then
            '   SavDay(0) = Format(DateAdd("M", 3, Format(Format(CheckStr(.Fields(9)), "####/##/##"), "yyyy/mm/dd")), "yyyymmdd")
            '   SavDay(1) = Format(DateAdd("M", 3, Format(Format(CheckStr(.Fields(10)), "####/##/##"), "yyyy/mm/dd")), "yyyymmdd")
            '   cnnConnection.Execute "UPDATE NEXTPROGRESS SET NP08=" & Val(SavDay(0)) & ",NP09=" & Val(SavDay(1)) & " WHERE NP01='" & CheckStr(.Fields(11)) & "' AND NP07=" & Val(CheckStr(.Fields(12))) & " AND NP22=" & Val(CheckStr(.Fields(13)))
            'End If
            
            PrintDatil
            If iPrint > 10000 Then
                Printer.CurrentX = 500
                Printer.CurrentY = iPrint
                Printer.Print String(200, "-")
                Printer.NewPage
                Page = Page + 1
                PrintTitle Page
            End If
            .MoveNext
        Loop
    End With
Else
   InsertQueryLog (0) 'Add By Sindy 2010/12/3
   ShowNoData
   Screen.MousePointer = vbDefault
   Exit Sub
End If
'Add By Cheng 2002/08/22
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print String(250, "-")

Printer.EndDoc
ShowPrintOk
Screen.MousePointer = vbDefault
End Sub

'********************
'印內容
'********************
Sub PrintDatil()
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print strTemp(0)
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print strTemp(1)
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
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print strTemp(6)
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print strTemp(7)
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print strTemp(8)
iPrint = iPrint + 300
End Sub

'********************
'定義報表X軸
'********************
Sub GetPleft()
Erase PLeft
PLeft(0) = 500
PLeft(1) = 1700
PLeft(2) = 4200
PLeft(3) = 6100
PLeft(4) = 8800
PLeft(5) = 10300
PLeft(6) = 11600
PLeft(7) = 13000
PLeft(8) = 14500

End Sub
'********************
'印抬頭
'********************
Sub PrintTitle(ByVal Page As Integer)
GetPleft
iPrint = 500
Printer.Orientation = 2
DoEvents
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.Font.Name = "細明體"
Printer.CurrentX = 6700
Printer.CurrentY = iPrint
Printer.Print "催 審 表"
Printer.Font.Underline = False
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.CurrentX = 6200
Printer.CurrentY = iPrint + 500
'Modify By Cheng 2002/08/22
'Printer.Print "催審期限：" & Format(ChangeTStringToTDateString(TXT1(1)) & " ", "@@@@@@@@@") & "－" & ChangeTStringToTDateString(TXT1(2))
Printer.Print IIf(Me.opt(0).Value, "催審期限：" & Format(ChangeTStringToTDateString(txt1(1)) & " ", "@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(2)), _
               "發文日期：" & Format(ChangeTStringToTDateString(txt1(5)) & " ", "@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(6)))
Printer.CurrentX = 500
Printer.CurrentY = iPrint + 800
Printer.Print "列印人　：" & strUserName
Printer.CurrentX = 13000
Printer.CurrentY = iPrint + 800
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
Printer.CurrentX = 500
Printer.CurrentY = iPrint + 1100
Printer.Print "申請國家：" & Format(txt1(3) & " ", "@@@@") & "－" & txt1(4)
Printer.CurrentX = 13000
Printer.CurrentY = iPrint + 1100
Printer.Print "頁　　次：" & str(Page)
Printer.CurrentX = 500
Printer.CurrentY = iPrint + 1400
Printer.Print String(200, "-")
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint + 1700
Printer.Print "發文日"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint + 1700
Printer.Print "申請案號"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint + 1700
Printer.Print "本所案號"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint + 1700
Printer.Print "案件名稱"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint + 1700
Printer.Print "申請國家"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint + 1700
Printer.Print "專利種類"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint + 1700
Printer.Print "案件性質"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint + 1700
Printer.Print "承辦人"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint + 1700
Printer.Print "智權人員"
Printer.CurrentX = 500
Printer.CurrentY = iPrint + 2000
Printer.Print String(200, "-")

iPrint = iPrint + 2300

End Sub

Private Sub Form_Load()
MoveFormToCenter Me
GetTaiwanTodayDate1 = GetTaiwanTodayDate
txt1(0) = GetSystemKindByNick
strTemp1 = Split(UCase("CFP,CPS"), ",")
strTemp2 = Split(UCase(txt1(0)), ",")
s = 0
St = ""
For i = 0 To UBound(strTemp1)
    For j = 0 To UBound(strTemp2)
        If strTemp2(j) = strTemp1(i) Then
            s = 1
            'Modify By Cheng 2002/08/22
'            St = St + strTemp(i)
            St = St + strTemp1(i) & ","
            Exit For
        End If
    Next j
Next i
If s = 0 Then
    s = MsgBox(strUserName & " 沒有 CFP 與 CPS 的權限 ", , "權限問題")
    Unload Me
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm050305 = Nothing
End Sub

Private Sub opt_Click(Index As Integer)
Select Case Index
Case 0 '催審期限
   Me.opt(0).Value = True
   Me.txt1(1).Enabled = True
   Me.txt1(2).Enabled = True
   Me.txt1(1).SetFocus

   Me.opt(1).Value = False
   Me.txt1(5).Enabled = False
   Me.txt1(6).Enabled = False

Case 1 '發文日期
   Me.opt(0).Value = False
   Me.txt1(1).Enabled = False
   Me.txt1(2).Enabled = False

   Me.opt(1).Value = True
   Me.txt1(5).Enabled = True
   Me.txt1(6).Enabled = True
   Me.txt1(5).SetFocus

End Select
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
            txt1(0).SelStart = 0
            txt1(0).SelLength = Len(txt1(0))
            Exit Sub
        End If
     Next i
'Modify By Cheng 2002/08/05
'Case 2, 4
Case 2, 4, 6
   'Modify By Cheng 202/09/12
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
Case 1, 2 '催審期限
   If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
      Cancel = True
      Me.txt1(Index).SetFocus
      txt1_GotFocus Index
   End If
'Add By Cheng 2002/08/05
Case 5, 6 '發文日期
   If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
      Cancel = True
      Me.txt1(Index).SetFocus
      txt1_GotFocus Index
   End If
End Select
End Sub
