VERSION 5.00
Begin VB.Form frm020302 
   BorderStyle     =   1  '單線固定
   Caption         =   "承辦人期限管制表"
   ClientHeight    =   3010
   ClientLeft      =   3480
   ClientTop       =   1950
   ClientWidth     =   3720
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3010
   ScaleWidth      =   3720
   Begin VB.OptionButton Option1 
      Caption         =   "指定期限："
      Height          =   180
      Index           =   2
      Left            =   70
      TabIndex        =   21
      Top             =   1650
      Width           =   1200
   End
   Begin VB.TextBox TXT1 
      Height          =   270
      Index           =   9
      Left            =   1275
      MaxLength       =   7
      TabIndex        =   5
      Top             =   1620
      Width           =   990
   End
   Begin VB.TextBox TXT1 
      Height          =   270
      Index           =   10
      Left            =   2355
      MaxLength       =   7
      TabIndex        =   6
      Top             =   1620
      Width           =   990
   End
   Begin VB.CommandButton CMDOK 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   2808
      TabIndex        =   12
      Top             =   130
      Width           =   756
   End
   Begin VB.CommandButton CMDOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2016
      TabIndex        =   11
      Top             =   130
      Width           =   756
   End
   Begin VB.TextBox TXT1 
      Height          =   270
      Index           =   8
      Left            =   2355
      MaxLength       =   4
      TabIndex        =   10
      Top             =   2570
      Width           =   990
   End
   Begin VB.TextBox TXT1 
      Height          =   270
      Index           =   7
      Left            =   1275
      MaxLength       =   4
      TabIndex        =   9
      Top             =   2570
      Width           =   990
   End
   Begin VB.TextBox TXT1 
      Height          =   270
      Index           =   6
      Left            =   1275
      MaxLength       =   6
      TabIndex        =   8
      Top             =   2260
      Width           =   930
   End
   Begin VB.TextBox TXT1 
      Height          =   270
      Index           =   5
      Left            =   1275
      MaxLength       =   1
      TabIndex        =   7
      Top             =   1940
      Width           =   300
   End
   Begin VB.TextBox TXT1 
      Height          =   270
      Index           =   4
      Left            =   2355
      MaxLength       =   7
      TabIndex        =   4
      Top             =   1300
      Width           =   990
   End
   Begin VB.TextBox TXT1 
      Height          =   270
      Index           =   3
      Left            =   1275
      MaxLength       =   7
      TabIndex        =   3
      Top             =   1300
      Width           =   990
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   2
      Left            =   2355
      MaxLength       =   7
      TabIndex        =   2
      Top             =   980
      Width           =   990
   End
   Begin VB.TextBox TXT1 
      Height          =   270
      Index           =   1
      Left            =   1275
      MaxLength       =   7
      TabIndex        =   1
      Top             =   980
      Width           =   990
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   0
      Left            =   1275
      TabIndex        =   0
      Top             =   650
      Width           =   2055
   End
   Begin VB.OptionButton Option1 
      Caption         =   "法定期限："
      Height          =   180
      Index           =   1
      Left            =   70
      TabIndex        =   19
      Top             =   1330
      Width           =   1200
   End
   Begin VB.OptionButton Option1 
      Caption         =   "本所期限："
      Height          =   180
      Index           =   0
      Left            =   70
      TabIndex        =   18
      Top             =   1050
      Value           =   -1  'True
      Width           =   1200
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   1760
      X2              =   2960
      Y1              =   1740
      Y2              =   1740
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   1790
      X2              =   2990
      Y1              =   2750
      Y2              =   2750
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   1740
      X2              =   2940
      Y1              =   1420
      Y2              =   1420
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   1730
      X2              =   2930
      Y1              =   1110
      Y2              =   1110
   End
   Begin VB.Label LBL1 
      Height          =   180
      Left            =   2280
      TabIndex        =   20
      Top             =   2280
      Width           =   1080
   End
   Begin VB.Label Label1 
      Caption         =   "(1. 商申  2.商爭  3.全部)"
      Height          =   180
      Index           =   4
      Left            =   1620
      TabIndex        =   17
      Top             =   1990
      Width           =   1880
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家："
      Height          =   180
      Index           =   3
      Left            =   70
      TabIndex        =   16
      Top             =   2600
      Width           =   910
   End
   Begin VB.Label Label1 
      Caption         =   "承辦人："
      Height          =   180
      Index           =   2
      Left            =   70
      TabIndex        =   15
      Top             =   2300
      Width           =   760
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質："
      Height          =   180
      Index           =   1
      Left            =   70
      TabIndex        =   14
      Top             =   1980
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   70
      TabIndex        =   13
      Top             =   710
      Width           =   950
   End
End
Attribute VB_Name = "frm020302"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/9 日期欄已修改
Option Explicit

Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, s As Integer, SavDay1 As String, SavDay2 As String
Dim strSQL2 As String, iPrint As Integer, Page As Integer, strTemp(0 To 15) As String, strTemp3 As String
Dim PLeft(0 To 14) As Integer, strTemp1 As Variant, strTemp2 As Variant, StrSQL6 As String


Private Sub cmdOK_Click(Index As Integer)
Select Case Index
Case 0
     Printer.Orientation = 2
     DoEvents
     If Len(txt1(0)) = 0 Then
         s = MsgBox("系統類別不可空白!!", , "USER 輸入錯誤")
         txt1(0).SetFocus
         Exit Sub
     Else
         'Add By Cheng 2002/09/19
         strTemp1 = Split(UCase(GetSystemKindByNick), ",")
         strTemp2 = Split(UCase(txt1(0)), ",")
         For i = 0 To UBound(strTemp2)
            s = 0
            For j = 0 To UBound(strTemp1)
                If strTemp2(i) = strTemp1(j) Then
                    s = 1
                    Exit For
                End If
            Next j
            If s = 0 Then
                s = MsgBox(strUserName & " 沒有 " & strTemp2(i) & " 的權限!! ", , "USER 權限問題")
                txt1(0).SetFocus
                txt1(0).SelStart = 0
                txt1(0).SelLength = Len(txt1(0))
                Exit Sub
            End If
         Next i
     
         If Option1(0).Value = True Then
            'Add By Cheng 2002/03/21
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
            
            If Len(txt1(2)) = 0 Then
                 s = MsgBox("本所期限區間不可空白!!", , "USER 輸入錯誤")
                 txt1(1).SetFocus
                 txt1_GotFocus (1)
                 Exit Sub
            End If
         ElseIf Option1(1).Value = True Then
            'Add By Cheng 2002/03/21
            If PUB_CheckKeyInDate(Me.txt1(3)) = -1 Then
               Me.txt1(3).SetFocus
               txt1_GotFocus 3
               Exit Sub
            End If
            If PUB_CheckKeyInDate(Me.txt1(4)) = -1 Then
               Me.txt1(4).SetFocus
               txt1_GotFocus 4
               Exit Sub
            End If
            
            If Len(txt1(4)) = 0 Then
                 s = MsgBox("法定期限區間不可空白!!", , "USER 輸入錯誤")
                 txt1(3).SetFocus
                 txt1_GotFocus (3)
                 Exit Sub
            End If
         'Add By Sindy 2024/1/16 +指定期限
         Else
            If PUB_CheckKeyInDate(Me.txt1(9)) = -1 Then
               Me.txt1(9).SetFocus
               txt1_GotFocus 9
               Exit Sub
            End If
            If PUB_CheckKeyInDate(Me.txt1(10)) = -1 Then
               Me.txt1(10).SetFocus
               txt1_GotFocus 10
               Exit Sub
            End If
            
            If Len(txt1(10)) = 0 Then
                 s = MsgBox("指定期限區間不可空白!!", , "USER 輸入錯誤")
                 txt1(9).SetFocus
                 txt1_GotFocus (9)
                 Exit Sub
            End If
         '2024/1/16 END
         End If
         If Len(txt1(5)) = 0 Then
             s = MsgBox("案件性質不可空白!!", , "USER 輸入錯誤")
             txt1(5).SetFocus
             Exit Sub
         Else
             ClearQueryLog (Me.Name) 'Add By Sindy 2010/9/30 清除查詢印表記錄檔欄位
             Screen.MousePointer = vbHourglass
             Me.Enabled = False
             Process
             Me.Enabled = True
             Screen.MousePointer = vbDefault
         End If
     End If
Case 1
     Unload Me
Case Else
End Select
End Sub

Sub Process()
'add by sonia 2017/10/13
Dim Rs As New ADODB.Recordset
Dim SQL As String
'end 2017/10/13

   Screen.MousePointer = vbHourglass
   cnnConnection.Execute "DELETE FROM R020302 WHERE ID='" & strUserNum & "' "
   strSQL1 = ""
   strSQL2 = ""
   StrSQL6 = ""
   If Len(txt1(0)) <> 0 Then
      strSQL1 = strSQL1 + " AND CP01 IN (" & SQLGrpStr(txt1(0), 2) & ") "
      strSQL2 = strSQL2 + " AND CP01 in (" & SQLGrpStr(txt1(0), 5) & ") "
      pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(0) 'Add By Sindy 2010/9/30
   End If
   If Option1(0).Value = True Then
       If Len(txt1(1)) <> 0 Then
          strSQL1 = strSQL1 + " AND CP06>=" & Val(ChangeTStringToWString(txt1(1))) & " "
          strSQL2 = strSQL2 + " AND CP06>=" & Val(ChangeTStringToWString(txt1(1))) & " "
       End If
       If Len(Trim(txt1(2))) <> 0 Then
         strSQL1 = strSQL1 + " AND CP06<=" & Val(ChangeTStringToWString(txt1(2))) & " "
         strSQL2 = strSQL2 + " AND CP06<=" & Val(ChangeTStringToWString(txt1(2))) & " "
       End If
       If Len(txt1(1)) <> 0 Or Len(Trim(txt1(2))) <> 0 Then
         pub_QL05 = pub_QL05 & ";" & Option1(0).Caption & txt1(1) & "-" & txt1(2) 'Add By Sindy 2010/9/30
       End If
   ElseIf Option1(1).Value = True Then
       If Len(txt1(3)) <> 0 Then
            strSQL1 = strSQL1 + " AND CP07>=" & Val(ChangeTStringToWString(txt1(3))) & ""
            strSQL2 = strSQL2 + " AND CP07>=" & Val(ChangeTStringToWString(txt1(3))) & ""
       End If
       If Len(Trim(txt1(4))) <> 0 Then
         strSQL1 = strSQL1 + " AND CP07<=" & Val(ChangeTStringToWString(txt1(4))) & " "
         strSQL2 = strSQL2 + " AND CP07<=" & Val(ChangeTStringToWString(txt1(4))) & " "
       End If
       If Len(txt1(3)) <> 0 Or Len(Trim(txt1(4))) <> 0 Then
         pub_QL05 = pub_QL05 & ";" & Option1(1).Caption & txt1(3) & "-" & txt1(4) 'Add By Sindy 2010/9/30
       End If
   'Add By Sindy 2024/1/16 +指定期限
   Else
       If Len(txt1(9)) <> 0 Then
            strSQL1 = strSQL1 + " AND CP142>=" & Val(ChangeTStringToWString(txt1(9))) & ""
            strSQL2 = strSQL2 + " AND CP142>=" & Val(ChangeTStringToWString(txt1(9))) & ""
       End If
       If Len(Trim(txt1(10))) <> 0 Then
         strSQL1 = strSQL1 + " AND CP142<=" & Val(ChangeTStringToWString(txt1(10))) & " "
         strSQL2 = strSQL2 + " AND CP142<=" & Val(ChangeTStringToWString(txt1(10))) & " "
       End If
       If Len(txt1(9)) <> 0 Or Len(Trim(txt1(10))) <> 0 Then
         pub_QL05 = pub_QL05 & ";" & Option1(2).Caption & txt1(9) & "-" & txt1(10)
       End If
   '2024/1/16 END
   End If
   strSQL1 = strSQL1 & " AND (TM29 IS NULL OR TM29='') AND CP57 IS NULL AND CP27 IS NULL "
   strSQL2 = strSQL2 & " AND (SP15 IS NULL OR SP15='') AND CP57 IS NULL AND CP27 IS NULL "
   StrSQL6 = " AND CP10<>'1728' "   'add by sonia 2017/7/10 桂英說剔除收款寄證1728(T-208573)
   
   Select Case Val(txt1(5))
      'Modify By Cheng 2002/09/19
      'Case 1
      '     strsql6 = strsql6 + " AND S1.ST05='95' "
      'Case 2
      '     strsql6 = strsql6 + " AND S1.ST05='97' "
      Case 1 '商申
           StrSQL6 = StrSQL6 + " AND (CP14 IS NULL OR (S1.ST05>='91' AND S1.ST05<='99')) "
            '案件性質三碼不抓"4"及"6"字頭, 四碼不抓"14"及"16"字頭
           StrSQL6 = StrSQL6 + " AND ( SUBSTR(DECODE(LENGTH(CP10),3,'0'||CP10,CP10),1,2)<>'04' AND SUBSTR(DECODE(LENGTH(CP10),3,'0'||CP10,CP10),1,2)<>'06' AND SUBSTR(DECODE(LENGTH(CP10),3,'0'||CP10,CP10),1,2)<>'14' AND SUBSTR(DECODE(LENGTH(CP10),3,'0'||CP10,CP10),1,2)<>'16' ) "
           pub_QL05 = pub_QL05 & ";" & Label1(1) & "商申" 'Add By Sindy 2010/9/30
      Case 2 '商爭
           StrSQL6 = StrSQL6 + " AND (CP14 IS NULL OR (S1.ST05>='91' AND S1.ST05<='99')) "
            '案件性質三碼只抓"4"或"6"字頭, 四碼只抓"14"或"16"字頭
           StrSQL6 = StrSQL6 + " AND ( SUBSTR(DECODE(LENGTH(CP10),3,'0'||CP10,CP10),1,2)='04' OR SUBSTR(DECODE(LENGTH(CP10),3,'0'||CP10,CP10),1,2)='06' OR SUBSTR(DECODE(LENGTH(CP10),3,'0'||CP10,CP10),1,2)='14' OR SUBSTR(DECODE(LENGTH(CP10),3,'0'||CP10,CP10),1,2)='16' ) "
           pub_QL05 = pub_QL05 & ";" & Label1(1) & "商爭" 'Add By Sindy 2010/9/30
      Case 3 '全部
           StrSQL6 = StrSQL6 + " AND (CP14 IS NULL OR (S1.ST05>='91' AND S1.ST05<='99')) "
           pub_QL05 = pub_QL05 & ";" & Label1(1) & "全部" 'Add By Sindy 2010/9/30
      Case Else
      End Select
   
   If Len(txt1(6)) <> 0 Then
       StrSQL6 = StrSQL6 + " AND CP14='" & txt1(6) & "' "
       pub_QL05 = pub_QL05 & ";" & Label1(2) & txt1(6) & Lbl1 'Add By Sindy 2010/9/30
   End If
   If Len(txt1(7)) <> 0 Then
       strSQL1 = strSQL1 + " AND TM10>='" & txt1(7) & "' "
       strSQL2 = strSQL2 + " AND SP09>='" & txt1(7) & "' "
   End If
   If Len(txt1(8)) <> 0 Then
       strSQL1 = strSQL1 + " AND TM10<='" & txt1(8) & "' "
       strSQL2 = strSQL2 + " AND SP09<='" & txt1(8) & "' "
   End If
   If Len(txt1(7)) <> 0 Or Len(txt1(8)) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label1(3) & txt1(7) & "-" & txt1(8)  'Add By Sindy 2010/9/30
   End If
   'modify by sonia 2017/10/13 +cp01,cp10以判斷FCT是否為爭議案FCT-040300未分案時會列在此報表
   'Modify By Sindy 2024/1/16 + ,SQLDateT2(CP142)||decode(CP164,'1','當天','2','之前','3','之後',cp164) CP142
   strSql = "SELECT s1.st01,S1.ST02,CP06," & SQLDate("CP07") & "," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,TM09,NVL(TM15,TM12),NVL(TM05,NVL(TM06,TM07)),decode(tm10,'000',CPM03,CPM04),nvl(CP49,cp64),DECODE(CP22,'Y','是','N','否',NULL,'是'),S2.ST02,CP09,CP27,CP01,CP10,SQLDateT2(CP142)||decode(CP164,'1','當天','2','之前','3','之後',cp164) CP142 FROM CASEPROGRESS,TRADEMARK,CASEPROPERTYMAP,STAFF S1,STAFF S2 WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & StrSQL6 & strSQL1
   strSql = strSql + " union all select s1.st01,S1.ST02,CP06," & SQLDate("CP07") & "," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,'',SP11,NVL(SP05,NVL(SP06,SP07)),decode(sp09,'000',CPM03,CPM04),nvl(CP49,cp64),DECODE(CP22,'Y','是','N','否',NULL,'是'),S2.ST02,CP09,CP27,CP01,CP10,SQLDateT2(CP142)||decode(CP164,'1','當天','2','之前','3','之後',cp164) CP142 FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,STAFF S1,STAFF S2 WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & StrSQL6 & strSQL2
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   With adoRecordset
       If .RecordCount <> 0 And .RecordCount > 0 Then
           InsertQueryLog (.RecordCount) 'Add By Sindy 2010/9/30
           .MoveFirst
           DoEvents
           Do While .EOF = False
               For i = 0 To 14 '13
                   'Add By Sindy 2024/1/16
                   If i = 14 Then
                     strTemp(i) = CheckStr(.Fields("cp142"))
                   Else
                   '2024/1/16 END
                     strTemp(i) = CheckStr(.Fields(i + 1))
                   End If
               Next i
               If strTemp(1) < GetTodayDate Then
                  '逾所限
                  strTemp(1) = "*" & ChangeTStringToTDateString(ChangeWStringToTString(strTemp(1)))
               Else
                   '當天
                   If strTemp(1) = GetTodayDate Then
                       strTemp(1) = "V" & ChangeTStringToTDateString(ChangeWStringToTString(strTemp(1)))
                   Else
                     '未過所限
                     If UCase(Mid(strTemp(12), 1, 1)) = "C" And Len(strTemp(13)) = 0 Then
                         'C類未發文
                         strTemp(1) = "#" & ChangeTStringToTDateString(ChangeWStringToTString(strTemp(1)))
                     Else
                         strTemp(1) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(1)))
                     End If
                   End If
               End If
               'add by sonia 2017/10/13 未分案案件判斷是否屬於內商承辦FCT-040300
               If Rs.State <> adStateClosed Then Rs.Close
               SQL = "select * from STAFF_GROUP where sg02='" & .Fields("CP01") & "' and sg03='" & .Fields("CP10") & "' and sg01='C1'"
               Rs.Open SQL, cnnConnection, adOpenStatic, adLockReadOnly
               '若有資料
               If "" & .Fields(0) <> "" Or Not Rs.EOF Then
               'end 2017/10/13
                  'Modify By Sindy 2024/1/16 +,r055014
                  strSql = "INSERT INTO R020302(r055001,r055002,r055003,r055004,r055005,r055006,r055007,r055008,r055009,r055010,r055011,r055012,id,r055013,r055014)" & _
                           " VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & ChgSQL(strTemp(8)) & "','" & ChgSQL(strTemp(9)) & "','" & ChgSQL(strTemp(10)) & "','" & ChgSQL(strTemp(11)) & "','" & strUserNum & "','" & ChgSQL(CheckStr(.Fields(0))) & "','" & ChgSQL(strTemp(14)) & "') "
                  cnnConnection.Execute strSql
               'add by sonia 2017/10/13
               End If
               If Rs.State <> adStateClosed Then Rs.Close
               Set Rs = Nothing
               'end 2017/10/13
               .MoveNext
               DoEvents
           Loop
       Else
           InsertQueryLog (0) 'Add By Sindy 2010/9/30
           ShowNoData
           Screen.MousePointer = vbDefault
           Exit Sub
       End If
   End With
   CheckOC
   PrintData
   Screen.MousePointer = vbDefault
End Sub

Sub PrintTitle()
GetPleft
iPrint = 500
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 6000
Printer.CurrentY = iPrint
Printer.Print "內商承辦人期限管制表"
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 500
Printer.CurrentX = 6700
Printer.CurrentY = iPrint
If Option1(0).Value = True Then
    Printer.Print "本所期限：" & Format(ChangeTStringToTDateString(txt1(1)) & " ", "@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(2))
ElseIf Option1(1).Value = True Then
    Printer.Print "法定期限：" & Format(ChangeTStringToTDateString(txt1(3)) & " ", "@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(4))
'Add By Sindy 2024/1/16 +指定期限
Else
    Printer.Print "指定期限：" & Format(ChangeTStringToTDateString(txt1(9)) & " ", "@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(10))
'2024/1/16 END
End If
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "承辦人：" & strTemp3
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "頁　　次：" & str(Page)
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(135, "-")
iPrint = iPrint + 300
Printer.Font.Size = 10
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "本所期限"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "法定期限"
'Add By Sindy 2024/1/16
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "指定期限"
'2024/1/16 END
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "收文日"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "本所案號"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "商品類別"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "申請案號/審定號"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print "案件名稱"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
Printer.Print "案件性質"
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iPrint
Printer.Print "條款內容/備註"
'Printer.CurrentX = PLeft(10)
'Printer.CurrentY = iPrint
'Printer.Print "是否出名"
Printer.CurrentX = PLeft(12)
Printer.CurrentY = iPrint
Printer.Print "智權人員"
iPrint = iPrint + 300
Printer.Font.Size = 12
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(135, "-")
iPrint = iPrint + 300
Printer.Font.Size = 10
End Sub

Sub PrintDatil()
For i = 1 To 12 '11
    Printer.CurrentX = PLeft(i)
    Printer.CurrentY = iPrint
    Printer.Print strTemp(i)
Next i
iPrint = iPrint + 300
End Sub

Sub GetPleft()
Erase PLeft
PLeft(0) = 500
PLeft(1) = 500
PLeft(2) = 1500
PLeft(3) = 2500 '指定送件日
PLeft(4) = 4000
PLeft(5) = 5000
PLeft(6) = 6500
PLeft(7) = 7500
PLeft(8) = 9200
PLeft(9) = 12000
PLeft(10) = 13500
PLeft(11) = 14000
PLeft(12) = 15500
End Sub

Sub PrintData()

'Add By Cheng 2001/12/28
'宣告變數
Dim Rs As New ADODB.Recordset '資料錄
Dim arrayTemp9
Dim SQL As String
'91/03/12 日期排序不能用符號排序
'nick
'Modify By Sindy 2024/1/16 +,r055014
strSql = "SELECT r055001,r055002,r055003,r055014,r055004,r055005,r055006,r055007,r055008,r055009,r055010,r055011,r055012,id,r055013 FROM R020302 WHERE ID='" & strUserNum & "' order by r055013,r055001,decode(substr(r055002,1,1),'#',substr(r055002,2,10),'*',substr(r055002,2,10),'V',substr(r055002,2,10),r055002),r055003,r055004"
Page = 1
strTemp3 = ""
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        strTemp3 = CheckStr(.Fields(0))
        PrintTitle
        Do While .EOF = False
            For i = 0 To 12 '11
               strTemp(i) = CheckStr(.Fields(i))
            Next i
            If strTemp3 <> strTemp(0) Then
                strTemp3 = strTemp(0)
                Page = Page + 1
                Printer.NewPage
                PrintTitle
            End If
            strTemp(6) = StrToStr(strTemp(6), 7)
            strTemp(7) = StrToStr(strTemp(7), 8)
            strTemp(8) = StrToStr(strTemp(8), 15)
            strTemp(9) = StrToStr(strTemp(9), 7)
            '避免將全形字切成一半
            strTemp(10) = Left(strTemp(10), 9)
            'Add By Cheng 2001/12/28
            '若有條款內容/備註(CP49-->LW02)
            If Len(Trim("" & strTemp(10))) > 0 Then
               arrayTemp9 = Split(strTemp(10), ",")
               SQL = "Select LW02 FROM LAW WHERE LW01 = '" & Trim(arrayTemp9(0)) & "'"
               Rs.Open SQL, cnnConnection, adOpenStatic, adLockReadOnly
               '若有資料
               If Not Rs.EOF Then
                  '更換條款內容/備註
                  strTemp(10) = "" & Rs("LW02").Value
               End If
               If Rs.State <> adStateClosed Then Rs.Close
               Set Rs = Nothing
            End If
            
            strTemp(11) = ""
            strTemp(12) = StrToStr(strTemp(12), 4)
            PrintDatil
            If iPrint >= 10000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle
            End If
            .MoveNext
        Loop
    End If
End With
CheckOC
Printer.EndDoc
ShowPrintOk
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
txt1(0) = GetSystemKindByNick
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm020302 = Nothing
End Sub

Private Sub Option1_Click(Index As Integer)
Select Case Index
Case 0
     txt1(1).SetFocus
     txt1_GotFocus (1)
Case 1
     txt1(3).SetFocus
     txt1_GotFocus (3)
'Add By Sindy 2024/1/16
Case 2
     txt1(9).SetFocus
     txt1_GotFocus (9)
'2024/1/16 END
Case Else
End Select
End Sub

Private Sub txt1_GotFocus(Index As Integer)
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub Txt1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cmdOK(0).SetFocus
End If
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
            If strTemp2(i) = strTemp1(j) Then
                s = 1
                Exit For
            End If
        Next j
        If s = 0 Then
            s = MsgBox(strUserName & " 沒有 " & strTemp2(i) & " 的權限!! ", , "USER 權限問題")
            txt1(0).SetFocus
            txt1(0).SelStart = 0
            txt1(0).SelLength = Len(txt1(0))
            Exit Sub
        End If
     Next i
'Modify By Sindy 2024/1/16 +, 9, 10
Case 2, 4, 1, 3, 9, 10
   If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
      Me.txt1(Index).SetFocus
      txt1_GotFocus Index
      Exit Sub
   End If
     If Index = 2 Or Index = 4 Then
        If RunNick(txt1(Index - 1), txt1(Index)) Then
            txt1(Index - 1).SetFocus
            txt1_GotFocus (Index - 1)
            Exit Sub
         End If
     End If
Case 8
     If RunNick(txt1(Index - 1), txt1(Index)) Then
         txt1(Index - 1).SetFocus
         txt1_GotFocus (Index - 1)
         Exit Sub
      End If
Case 5
     Select Case Trim(txt1(5))
     Case "1", "2", "3", ""
     Case Else
          s = MsgBox("案件性質只能輸入 1 或 2 或 3 !!", , "USER 輸入錯誤")
          txt1(5).SetFocus
          txt1(5).SelStart = 0
          txt1(5).SelLength = Len(txt1(5))
          Exit Sub
     End Select
Case 6
     Lbl1 = GetPrjSalesNM(txt1(6))
     If Trim(txt1(Index)) <> "" Then
        If Trim(Lbl1.Caption) = "" Then
            s = MsgBox("承辦人輸入錯誤！", , "錯誤！")
            txt1(6).SetFocus
            txt1_GotFocus (6)
            Exit Sub
        End If
    End If
Case Else
End Select
End Sub
