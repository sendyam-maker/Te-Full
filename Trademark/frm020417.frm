VERSION 5.00
Begin VB.Form frm020417 
   BorderStyle     =   1  '單線固定
   Caption         =   "商標承辦人績效表"
   ClientHeight    =   2700
   ClientLeft      =   3672
   ClientTop       =   3300
   ClientWidth     =   3924
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   3924
   Begin VB.TextBox txt1 
      Height          =   285
      Index           =   2
      Left            =   2085
      MaxLength       =   5
      TabIndex        =   2
      Top             =   990
      Width           =   915
   End
   Begin VB.TextBox txt1 
      Height          =   285
      Index           =   1
      Left            =   1008
      MaxLength       =   5
      TabIndex        =   1
      Top             =   990
      Width           =   915
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   540
      Left            =   45
      TabIndex        =   12
      Top             =   2070
      Width           =   3825
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   765
         Style           =   2  '單純下拉式
         TabIndex        =   5
         Top             =   180
         Width           =   2880
      End
      Begin VB.Label Label4 
         Caption         =   "印表機"
         Height          =   180
         Left            =   105
         TabIndex        =   13
         Top             =   255
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   3108
      TabIndex        =   7
      Top             =   48
      Width           =   756
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2316
      TabIndex        =   6
      Top             =   48
      Width           =   756
   End
   Begin VB.TextBox txt1 
      Enabled         =   0   'False
      Height          =   264
      Index           =   4
      Left            =   1935
      MaxLength       =   4
      TabIndex        =   4
      Top             =   2385
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.TextBox txt1 
      Enabled         =   0   'False
      Height          =   264
      Index           =   3
      Left            =   1005
      MaxLength       =   4
      TabIndex        =   3
      Top             =   2370
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox txt1 
      Height          =   285
      Index           =   0
      Left            =   1008
      TabIndex        =   0
      Top             =   570
      Width           =   1740
   End
   Begin VB.Label Label2 
      Caption         =   $"frm020417.frx":0000
      ForeColor       =   &H000000C0&
      Height          =   465
      Left            =   210
      TabIndex        =   14
      Top             =   1500
      Width           =   3225
   End
   Begin VB.Line Line5 
      X1              =   1935
      X2              =   2205
      Y1              =   1140
      Y2              =   1140
   End
   Begin VB.Label Label1 
      Caption         =   "發文年月："
      Height          =   180
      Index           =   2
      Left            =   90
      TabIndex        =   9
      Top             =   1050
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "月"
      Height          =   180
      Left            =   2820
      TabIndex        =   11
      Top             =   2160
      Width           =   435
   End
   Begin VB.Line Line2 
      Visible         =   0   'False
      X1              =   1575
      X2              =   2325
      Y1              =   2490
      Y2              =   2490
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家："
      Enabled         =   0   'False
      Height          =   180
      Index           =   3
      Left            =   90
      TabIndex        =   10
      Top             =   2430
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   90
      TabIndex        =   8
      Top             =   630
      Width           =   915
   End
End
Attribute VB_Name = "frm020417"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/10 日期欄已修改
Option Explicit

Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, s As Integer, SavDay1 As String, SavDay2 As String, SavDay3 As String, SavDay4 As String
Dim strSQL2 As String, iPrint As Integer, Page As Integer, strTemp(0 To 15) As String, strTemp3 As String, TestOk As Boolean, StrTemp7(0 To 15) As String
Dim PLeft(0 To 15) As Integer, strTemp1 As Variant, strTemp2 As Variant, StrSQL6 As String, SeekPrint As Integer, SeekPrintL As Integer, ChangeNewPage As Boolean, BolChangePage As Boolean
Dim CALTMP1 As String, CALTMP2 As String, CALTMP3 As String
Dim TMPRSBYNICK As New ADODB.Recordset
Dim calToTal As String
Dim bolIsTMA As Boolean 'Added by Lydia 2024/11/15 資料改抓查名單(網中)

Private Sub cmdOK_Click(Index As Integer)
Select Case Index
Case 0 '確定
     'Modified by Morgan 2015/6/3
     'If Combo1.ListIndex >= SeekPrint Then
     '   j = Combo1.ListIndex + 1
     'Else
     '   j = Combo1.ListIndex
     'End If
     'Set Printer = Printers(j)
     PUB_RestorePrinter Combo1.Text
     'end 2015/6/3
     Printer.EndDoc 'Add By Sindy 2011/11/1
     'Modified by Moran 2015/6/1
     'Printer.PaperSize = 39
     Printer.PaperSize = PUB_GetPaperSize(15, 2)
     'end 2015/6/1
     DoEvents
     If Len(txt1(0)) = 0 Then
         s = MsgBox("系統類別不可空白!!", , "USER 輸入錯誤")
         txt1(0).SetFocus
         Exit Sub
     Else
         'Modify By Sindy 2019/12/6
'         If Len(txt1(2)) = 0 Or Len(txt1(1)) = 0 Then
'             s = MsgBox("發文年月不可空白!!", , "USER 輸入錯誤")
'             txt1(1).SetFocus
'             txt1_GotFocus (1)
'             Exit Sub
         If Len(txt1(1)) = 0 Then
             s = MsgBox("發文年月不可空白!!", , "USER 輸入錯誤")
             txt1(1).SetFocus
             txt1_GotFocus (1)
             Exit Sub
         ElseIf Len(txt1(2)) = 0 Then
             s = MsgBox("發文年月不可空白!!", , "USER 輸入錯誤")
             txt1(2).SetFocus
             txt1_GotFocus (2)
             Exit Sub
         '2019/12/6 END
         Else
             Screen.MousePointer = vbHourglass
             Me.Enabled = False
             ClearQueryLog (Me.Name) 'Add By Sindy 2010/10/21 清除查詢印表記錄檔欄位
             ProcessByNick
             Me.Enabled = True
             Screen.MousePointer = vbDefault
         End If
     End If
Case 1 '結束
     Unload Me
Case Else
End Select
End Sub

Sub ProcessByNick()
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

   'Added by Lydia 2024/11/15 資料改抓查名單(網中)
   If DBDATE(txt1(1) & "01") >= 查名單網中系統啟用日 Or DBDATE(txt1(2) & "01") >= 查名單網中系統啟用日 Then
      bolIsTMA = True
   Else
      bolIsTMA = False
   End If
   'end 2024//11/15
   
'***商申人員的資料表--R020417_1
'***商爭人員的資料表--R020417_2
Screen.MousePointer = vbHourglass
cnnConnection.Execute "DELETE FROM R020417_1 WHERE ID='" & strUserNum & "' "
cnnConnection.Execute "DELETE FROM R020417_11 WHERE ID='" & strUserNum & "' "
cnnConnection.Execute "DELETE FROM R020417_2 WHERE ID='" & strUserNum & "' "
cnnConnection.Execute "DELETE FROM R020417_21 WHERE ID='" & strUserNum & "' "
cnnConnection.Execute "DELETE FROM R020417_3 WHERE ID='" & strUserNum & "' "
cnnConnection.Execute "DELETE FROM R020417_4 WHERE ID='" & strUserNum & "' "
strSQL1 = ""
strSQL2 = ""
StrSQL6 = ""
If Len(txt1(0)) <> 0 Then
   strSQL1 = strSQL1 + " AND CP01 IN (" & SQLGrpStr(txt1(0), 2) & ") "
   strSQL2 = strSQL2 + " AND CP01 IN (" & SQLGrpStr(txt1(0), 5) & ") "
   pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(0) 'Add By Sindy 2010/10/21
End If
StrSQL6 = ""
If Len(txt1(3)) <> 0 Then
    strSQL1 = strSQL1 + " AND TM10>=" & txt1(3)
    strSQL2 = strSQL2 + " AND SP09>=" & txt1(3)
End If
If Len(txt1(4)) <> 0 Then
    strSQL1 = strSQL1 + " AND TM10<=" & txt1(4)
    strSQL2 = strSQL2 + " AND SP09<=" & txt1(4)
End If
If Len(txt1(1)) <> 0 Or Len(txt1(2)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(2) & txt1(1) & " ~ " & txt1(2) 'Add By Sindy 2010/10/21
End If
CheckOC
'案件進度檔
DoEvents
'Modify By Cheng 2002/04/09
'只抓員工等級為91~99的資料, 95列入商申人員, 其餘為商爭人員
'strSQL = "SELECT CP14,S1.ST05,cp12,CP14,TM10,CP01,CP10,cp18 FROM CASEPROGRESS,TRADEMARK,STAFF S1,STAFF S2 WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP13=S2.ST01(+) AND CP14=S1.ST01(+) AND CP27>=" & (Val(txt1(1)) + 1911) * 100 + Val(Format(txt1(2), "00")) & "01 AND CP27<=" & (Val(txt1(1)) + 1911) * 100 + Val(Format(txt1(2), "00")) & "31 AND CP26 IS NULL  " & strSQL1
'strSQL = strSQL + " union all select CP14,S1.ST05,cp12,CP14,SP09,CP01,CP10,cp18 FROM CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2 WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP13=S2.ST01(+) AND CP14=S1.ST01(+) AND CP27>=" & (Val(txt1(1)) + 1911) * 100 + Val(Format(txt1(2), "00")) & "01 AND CP27<=" & (Val(txt1(1)) + 1911) * 100 + Val(Format(txt1(2), "00")) & "31 AND CP26 IS NULL  " & strSQL2
'Modify By Sindy 2019/12/6 發文年月改為起迄年月
'                    strSql = "SELECT CP14,S1.ST05,cp12,CP14,TM10,CP01,CP10,cp18,CP09 FROM CASEPROGRESS,TRADEMARK,STAFF S1,STAFF S2 WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP13=S2.ST01(+) AND CP14=S1.ST01(+) AND S1.ST05 >='91' AND S1.ST05<='99' AND CP27>=" & (Val(txt1(1)) + 1911) * 100 + Val(Format(txt1(2), "00")) & "01 AND CP27<=" & (Val(txt1(1)) + 1911) * 100 + Val(Format(txt1(2), "00")) & "31 AND CP26 IS NULL  " & strSQL1
'strSql = strSql + " union all select CP14,S1.ST05,cp12,CP14,SP09,CP01,CP10,cp18,CP09 FROM CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2 WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP13=S2.ST01(+) AND CP14=S1.ST01(+) AND S1.ST05>='91' AND S1.ST05<='99' AND CP27>=" & (Val(txt1(1)) + 1911) * 100 + Val(Format(txt1(2), "00")) & "01 AND CP27<=" & (Val(txt1(1)) + 1911) * 100 + Val(Format(txt1(2), "00")) & "31 AND CP26 IS NULL  " & strSQL2
'Add By Sindy 2021/2/23 +,CP162
                    strSql = "SELECT CP14,S1.ST05,cp12,CP14,TM10,CP01,CP10,cp18,CP09,cp27,CP162 FROM CASEPROGRESS,TRADEMARK,STAFF S1,STAFF S2 WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP13=S2.ST01(+) AND CP14=S1.ST01(+) AND S1.ST05 >='91' AND S1.ST05<='99' AND CP27>=" & Val(txt1(1)) + 191100 & "01 AND CP27<=" & Val(txt1(2)) + 191100 & "31 AND CP26 IS NULL  " & strSQL1
strSql = strSql + " union all select CP14,S1.ST05,cp12,CP14,SP09,CP01,CP10,cp18,CP09,cp27,CP162 FROM CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2 WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP13=S2.ST01(+) AND CP14=S1.ST01(+) AND S1.ST05>='91' AND S1.ST05<='99' AND CP27>=" & Val(txt1(1)) + 191100 & "01 AND CP27<=" & Val(txt1(2)) + 191100 & "31 AND CP26 IS NULL  " & strSQL2
'2019/12/6 END
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        InsertQueryLog (.RecordCount) 'Add By Sindy 2010/10/21
        .MoveFirst
        Do While .EOF = False
            For i = 0 To 5
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            
            'Modify By Sindy 2021/2/23 點數
            strTemp(15) = Val(CheckStr(adoRecordset.Fields(7)))
            'Add By Sindy 2021/2/23 有案源單號無點數,並且總收文號為LOS01 P/T案總收文號者
            If "" & .Fields("cp162") <> "" And Val("" & .Fields("cp18")) = 0 Then
               StrSQLa = "SELECT los01,los10,los15,cp09,cp18 FROM lawofficesource,caseprogress" & _
                         " WHERE los15='" & .Fields("cp162").Value & "' AND los01='" & .Fields("cp09").Value & "'" & _
                         " AND los10=cp09(+) AND cp18>0"
               rsA.CursorLocation = adUseClient
               rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
               If rsA.RecordCount > 0 Then
                  strTemp(15) = Val("" & rsA("CP18").Value)
               End If
               If rsA.State <> adStateClosed Then rsA.Close
               Set rsA = Nothing
            'Add By Sindy 2021/2/23 有案源單號,要扣掉介紹規費cp17的點數
            ElseIf "" & .Fields("cp162") <> "" Then
               If Val("" & .Fields("cp18")) > 0 Then
                  StrSQLa = "SELECT los01,los10,los15,cp09,cp18,cp17 FROM lawofficesource,caseprogress" & _
                            " WHERE los15='" & .Fields("cp162").Value & "' AND los01='" & .Fields("cp09").Value & "'" & _
                            " AND los10=cp09(+) AND cp18>0"
                  rsA.CursorLocation = adUseClient
                  rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                  If rsA.RecordCount > 0 Then
                     strTemp(15) = strTemp(15) - Val(Format(Val("" & rsA("CP17").Value) / 1000, "0.0"))
                  End If
                  If rsA.State <> adStateClosed Then rsA.Close
                  Set rsA = Nothing
               End If
            End If
            '2021/2/23 END
            
            '個人目標
            'Modify By Sindy 2019/12/6
            'strTemp3 = GetPerformanceByNickPE06(Val(str(Val(txt1(1)) + 1911) & Format(txt1(2), "00")), Val(str(Val(txt1(1)) + 1911) & Format(txt1(2), "00")), "T", strTemp(3))
            'strTemp3 = GetPerformanceByNickPE06(Left(.Fields(9), 6), Left(.Fields(9), 6), "T", strTemp(3))
            strTemp3 = GetPerformanceByNickPE06(Val(txt1(1)) + 191100, Val(txt1(2)) + 191100, "T", strTemp(3))
            '2019/12/6 END
            '依系統類別與國家分類
            Select Case strTemp(5)
'T, FCT          *******************************************************************************************************
            'Modify By Cheng 2004/03/11
            '加FCT
'            Case "T"
            Case "T", "FCT"
            'End
                 ProcessT
'TS         **********************************************************************************************************************
            Case "TS"
                 ProcessTS
'TF         *********************************************************************************************************************************
            Case "TF"
                 ProcessTF
'TB         *************************************************************************************************************************************************
            Case "TB"
                 ProcessTB
'TM         ***********************************************************************************************************************************************************************************************
            Case "TM"
                 ProcessTM
'其他       ********************************************************************************************************************************************************************************
            Case Else
                 ProcessOther
            End Select
            .MoveNext
        Loop
    Else
        InsertQueryLog (0) 'Add By Sindy 2010/10/21
        ShowNoData
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
End With
CheckOC

'Add By Cheng 2004/03/11
'Performance 的其他發文點數加入其他項
'Modify By Sindy 2019/12/6
'StrSQLa = "Select PE01, PE08, ST05 From Performance, Staff Where PE01=ST01 And PE02='T' And PE03=" & (Val(Me.txt1(1).Text) + 1911) & Format(Me.txt1(2).Text, "00") & " And PE08 Is Not Null And PE08<>0 "
StrSQLa = "Select PE01, sum(PE08) PE08, ST05 From Performance, Staff Where PE01=ST01 And PE02='T' And PE03>=" & Val(Me.txt1(1).Text) + 191100 & " And PE03<=" & Val(Me.txt1(2).Text) + 191100 & " And PE08 Is Not Null And PE08<>0 group by PE01,ST05 "
'2019/12/6 END
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    While Not rsA.EOF
        '個人目標
        'Modify By Sindy 2019/12/6
        'strTemp3 = GetPerformanceByNickPE06(Val(str(Val(txt1(1)) + 1911) & Format(txt1(2), "00")), Val(str(Val(txt1(1)) + 1911) & Format(txt1(2), "00")), "T", "" & rsA.Fields(0).Value)
        strTemp3 = GetPerformanceByNickPE06(Val(txt1(1)) + 191100, Val(txt1(2)) + 191100, "T", "" & rsA.Fields(0).Value)
        '2019/12/6 END
        Select Case "" & rsA.Fields(2).Value
        Case 95 ', 15
            'Modified by Lydia 2023/12/01 ,8,=>,9,
            cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086012,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL("" & rsA.Fields(0).Value) & "'," & Val(strTemp3) & ",'其他'," & Val("" & rsA.Fields(1).Value) & "," & Val("" & rsA.Fields(1).Value) & ",'" & ChgSQL("" & rsA.Fields(0).Value) & "','*',9,'" & strUserNum & "') "
        Case Else
            'Modified by Lydia 2023/12/01 ,8,=>,9,
            cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088012,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL("" & rsA.Fields(0).Value) & "'," & Val(strTemp3) & ",'其他'," & Val("" & rsA.Fields(1).Value) & "," & Val("" & rsA.Fields(1).Value) & ",'" & ChgSQL("" & rsA.Fields(0).Value) & "','',9,'" & strUserNum & "') "
        End Select
        rsA.MoveNext
    Wend
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
'End

'900430 清資料庫，將要更新的員工目標清除
'strSQL = "SELECT TMQ10,S1.ST05,S2.ST03,TMQ10,TMQ07+TMQ08+TMQ09,'','',TMQ07,TMQ08,TMQ09 FROM TRADEMARKQUERY,STAFF S1,STAFF S2 WHERE TMQ02=S2.ST01(+) AND TMQ10=S1.ST01(+) AND TMQ11>=" & (Val(txt1(1)) + 1911) * 100 + Val(txt1(2)) & "01 AND TMQ11<=" & (Val(txt1(1)) + 1911) * 100 + Val(txt1(2)) & "31 "
'Modify By Sindy 2019/12/6
If Val(txt1(1)) = Val(txt1(2)) Then
   'cnnConnection.Execute "UPDATE PERFORMANCE SET PE12=0,PE13=0,PE14=0 WHERE PE01 IN (SELECT TMQ10 FROM TRADEMARKQUERY WHERE TMQ11>=" & (Val(txt1(1)) + 1911) * 100 + Val(txt1(2)) & "01 AND TMQ11<=" & (Val(txt1(1)) + 1911) * 100 + Val(txt1(2)) & "31 ) AND PE02='T' AND PE03=" & Val(str(Val(txt1(1)) + 1911) & Format(txt1(2), "00")) & " "
   'Added by Lydia 2024/11/15 查名單(網中)：排除1120904-1120928期間資料匯入>>TO_CHAR(TMA04,'YYYYMMDD')>='20240601'
   If bolIsTMA = True Then
      cnnConnection.Execute "UPDATE PERFORMANCE SET PE12=0,PE13=0,PE14=0 WHERE PE01 IN (SELECT TMA10 FROM TMQAPPFORM WHERE TMA14>=" & Val(txt1(1)) + 191100 & "01 AND TMA14<=" & Val(txt1(2)) + 191100 & "31 and TO_CHAR(TMA04,'YYYYMMDD')>='20240601') AND PE02='T' AND PE03=" & Val(txt1(1)) + 191100 & " "
   Else
   'end 2024/11/15
      cnnConnection.Execute "UPDATE PERFORMANCE SET PE12=0,PE13=0,PE14=0 WHERE PE01 IN (SELECT TMQ10 FROM TRADEMARKQUERY WHERE TMQ11>=" & Val(txt1(1)) + 191100 & "01 AND TMQ11<=" & Val(txt1(2)) + 191100 & "31 ) AND PE02='T' AND PE03=" & Val(txt1(1)) + 191100 & " "
   End If
   
End If
'2019/12/6 END
'商標委查檔--更新個人目標檔
DoEvents
'Modify By Sindy 2019/12/9
'strSql = "SELECT TMQ10,S1.ST05,S2.ST03,TMQ10,DECODE(TMQ07,NULL,0,TMQ07)+DECODE(TMQ08,NULL,0,TMQ08)+DECODE(TMQ09,NULL,0,TMQ09),'','',DECODE(TMQ07,NULL,0,TMQ07),DECODE(TMQ08,NULL,0,TMQ08),DECODE(TMQ09,NULL,0,TMQ09) FROM TRADEMARKQUERY,STAFF S1,STAFF S2 WHERE TMQ02=S2.ST01(+) AND TMQ10=S1.ST01(+) AND TMQ11>=" & (Val(txt1(1)) + 1911) * 100 + Val(txt1(2)) & "01 AND TMQ11<=" & (Val(txt1(1)) + 1911) * 100 + Val(txt1(2)) & "31 "
'Added by Lydia 2024/11/15 查名單(網中)：排除1120904-1120928期間資料匯入>>TO_CHAR(TMA04,'YYYYMMDD')>='20240601'
If bolIsTMA = True Then
   strSql = "SELECT TMA10,S1.ST05,S2.ST03,TMA10,DECODE(TMA36,NULL,0,TMA36)+DECODE(TMA37,NULL,0,TMA37)+DECODE(TMA38,NULL,0,TMA38),'','',DECODE(TMA36,NULL,0,TMA36),DECODE(TMA37,NULL,0,TMA37),DECODE(TMA38,NULL,0,TMA38),TMA14 FROM TMQAPPFORM,STAFF S1,STAFF S2 WHERE TMA08=S2.ST01(+) AND TMA10=S1.ST01(+) AND TMA14>=" & Val(txt1(1)) + 191100 & "01 AND TMA14<=" & Val(txt1(2)) + 191100 & "31 and TO_CHAR(TMA04,'YYYYMMDD')>='20240601' "
Else
'end 2024/11/15
   strSql = "SELECT TMQ10,S1.ST05,S2.ST03,TMQ10,DECODE(TMQ07,NULL,0,TMQ07)+DECODE(TMQ08,NULL,0,TMQ08)+DECODE(TMQ09,NULL,0,TMQ09),'','',DECODE(TMQ07,NULL,0,TMQ07),DECODE(TMQ08,NULL,0,TMQ08),DECODE(TMQ09,NULL,0,TMQ09),TMQ11 FROM TRADEMARKQUERY,STAFF S1,STAFF S2 WHERE TMQ02=S2.ST01(+) AND TMQ10=S1.ST01(+) AND TMQ11>=" & Val(txt1(1)) + 191100 & "01 AND TMQ11<=" & Val(txt1(2)) + 191100 & "31 "
End If
'2019/12/9 END
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        Do While .EOF = False
            For i = 0 To 9
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            'Modify By Sindy 2019/12/9
            'strSql = "SELECT NVL(PE12,0),NVL(PE13,0),NVL(PE14,0) FROM PERFORMANCE WHERE PE01='" & ChgSQL(strTemp(3)) & "' AND PE02='T' AND PE03=" & Val(str(Val(txt1(1)) + 1911) & Format(txt1(2), "00")) & " "
            strSql = "SELECT NVL(PE12,0),NVL(PE13,0),NVL(PE14,0) FROM PERFORMANCE WHERE PE01='" & ChgSQL(strTemp(3)) & "' AND PE02='T' AND PE03=" & Left(adoRecordset.Fields("TMQ11"), 6) & " "
            '2019/12/9 END
            Set TMPRSBYNICK = New ADODB.Recordset
            TMPRSBYNICK.CursorLocation = adUseClient
            TMPRSBYNICK.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If TMPRSBYNICK.RecordCount <> 0 Then
               CALTMP1 = CheckStr(TMPRSBYNICK.Fields(0))
               CALTMP2 = CheckStr(TMPRSBYNICK.Fields(1))
               CALTMP3 = CheckStr(TMPRSBYNICK.Fields(2))
            Else
               CALTMP1 = "0"
               CALTMP2 = "0"
               CALTMP3 = "0"
            End If
            'Modify By Sindy 2019/12/9
            If Val(txt1(1)) = Val(txt1(2)) Then
               'cnnConnection.Execute "UPDATE PERFORMANCE SET PE12=" & (Val(strTemp(7)) + Val(CALTMP1)) & ",PE13=" & (Val(strTemp(8)) + Val(CALTMP2)) & ",PE14=" & (Val(strTemp(9)) + Val(CALTMP3)) & " WHERE PE01='" & ChgSQL(strTemp(3)) & "' AND PE02='T' AND PE03=" & Val(str(Val(txt1(1)) + 1911) & Format(txt1(2), "00")) & " "
               cnnConnection.Execute "UPDATE PERFORMANCE SET PE12=" & (Val(strTemp(7)) + Val(CALTMP1)) & ",PE13=" & (Val(strTemp(8)) + Val(CALTMP2)) & ",PE14=" & (Val(strTemp(9)) + Val(CALTMP3)) & " WHERE PE01='" & ChgSQL(strTemp(3)) & "' AND PE02='T' AND PE03=" & Val(txt1(1)) + 191100 & " "
            End If
            '2019/12/9 END
            'Modify By Cheng 2002/05/02
            '***國內查名的點數計算方式改為(委查中文筆數TMQ07+委查英文筆數TMQ08)*0.1+(委查圖形筆數TMQ09+0.3)
            strTemp(4) = ((Val(strTemp(7)) + Val(strTemp(8))) * 0.1) + (Val(strTemp(9)) * 0.3)
            'Add By Cheng 2003/03/03
            '個人目標
            'Modify By Sindy 2019/12/9
            'strTemp3 = GetPerformanceByNickPE06(Val(str(Val(txt1(1)) + 1911) & Format(txt1(2), "00")), Val(str(Val(txt1(1)) + 1911) & Format(txt1(2), "00")), "T", strTemp(3))
            'strTemp3 = GetPerformanceByNickPE06(Left(adoRecordset.Fields("TMQ11"), 6), Left(adoRecordset.Fields("TMQ11"), 6), "T", strTemp(3))
            strTemp3 = GetPerformanceByNickPE06(Val(txt1(1)) + 191100, Val(txt1(2)) + 191100, "T", strTemp(3))
            '2019/12/9 END
            Select Case Mid(UCase(strTemp(2)), 1, 2)
            Case "S1"
                 Select Case UCase(strTemp(2))
                 Case "S11"
                      Select Case Val(strTemp(1))
                      '屬商申人員資料
                      Case 95 ', 15
                           cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086004,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'國內查名'," & Val(strTemp(4)) & "," & Val(strTemp(4)) & ",'" & ChgSQL(strTemp(3)) & "','*',3,'" & strUserNum & "') "
                      '屬商爭人員資料
                      Case Else
                           cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088004,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'國內查名'," & Val(strTemp(4)) & "," & Val(strTemp(4)) & ",'" & ChgSQL(strTemp(3)) & "','',3,'" & strUserNum & "') "
                      End Select
                 Case "S12"
                      Select Case Val(strTemp(1))
                      Case 95 ', 15
                           cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086005,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'國內查名'," & Val(strTemp(4)) & "," & Val(strTemp(4)) & ",'" & ChgSQL(strTemp(3)) & "','*',3,'" & strUserNum & "') "
                      Case Else
                           cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088005,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'國內查名'," & Val(strTemp(4)) & "," & Val(strTemp(4)) & ",'" & ChgSQL(strTemp(3)) & "','',3,'" & strUserNum & "') "
                      End Select
                 Case "S13"
                      Select Case Val(strTemp(1))
                      Case 95 ', 15
                           cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086006,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'國內查名'," & Val(strTemp(4)) & "," & Val(strTemp(4)) & ",'" & ChgSQL(strTemp(3)) & "','*',3,'" & strUserNum & "') "
                      Case Else
                           cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088006,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'國內查名'," & Val(strTemp(4)) & "," & Val(strTemp(4)) & ",'" & ChgSQL(strTemp(3)) & "','',3,'" & strUserNum & "') "
                      End Select
                 Case "S14"
                      Select Case Val(strTemp(1))
                      Case 95 ', 15
                           cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086007,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'國內查名'," & Val(strTemp(4)) & "," & Val(strTemp(4)) & ",'" & ChgSQL(strTemp(3)) & "','*',3,'" & strUserNum & "') "
                      Case Else
                           cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088007,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'國內查名'," & Val(strTemp(4)) & "," & Val(strTemp(4)) & ",'" & ChgSQL(strTemp(3)) & "','',3,'" & strUserNum & "') "
                      End Select
                 Case "S15"
                      Select Case Val(strTemp(1))
                      Case 95 ', 15
                           cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086008,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'國內查名'," & Val(strTemp(4)) & "," & Val(strTemp(4)) & ",'" & ChgSQL(strTemp(3)) & "','*',3,'" & strUserNum & "') "
                      Case Else
                           cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088008,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'國內查名'," & Val(strTemp(4)) & "," & Val(strTemp(4)) & ",'" & ChgSQL(strTemp(3)) & "','',3,'" & strUserNum & "') "
                      End Select
                 Case Else
                 End Select
            Case "S2"
                 Select Case Val(strTemp(1))
                 Case 95 ', 15
                      cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086009,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'國內查名'," & Val(strTemp(4)) & "," & Val(strTemp(4)) & ",'" & ChgSQL(strTemp(3)) & "','*',3,'" & strUserNum & "') "
                 Case Else
                      cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088009,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'國內查名'," & Val(strTemp(4)) & "," & Val(strTemp(4)) & ",'" & ChgSQL(strTemp(3)) & "','',3,'" & strUserNum & "') "
                 End Select
            Case "S3"
                 Select Case Val(strTemp(1))
                 Case 95 ', 15
                      cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086010,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'國內查名'," & Val(strTemp(4)) & "," & Val(strTemp(4)) & ",'" & ChgSQL(strTemp(3)) & "','*',3,'" & strUserNum & "') "
                 Case Else
                      cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088010,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'國內查名'," & Val(strTemp(4)) & "," & Val(strTemp(4)) & ",'" & ChgSQL(strTemp(3)) & "','',3,'" & strUserNum & "') "
                 End Select
            Case "S4"
                 Select Case Val(strTemp(1))
                 Case 95 ', 15
                      cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086011,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'國內查名'," & Val(strTemp(4)) & "," & Val(strTemp(4)) & ",'" & ChgSQL(strTemp(3)) & "','*',3,'" & strUserNum & "') "
                 Case Else
                      cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088011,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'國內查名'," & Val(strTemp(4)) & "," & Val(strTemp(4)) & ",'" & ChgSQL(strTemp(3)) & "','',3,'" & strUserNum & "') "
                 End Select
            Case Else
                 Select Case Val(strTemp(1))
                 Case 95 ', 15
                      cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086012,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'國內查名'," & Val(strTemp(4)) & "," & Val(strTemp(4)) & ",'" & ChgSQL(strTemp(3)) & "','*',3,'" & strUserNum & "') "
                 Case Else
                      cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088012,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'國內查名'," & Val(strTemp(4)) & "," & Val(strTemp(4)) & ",'" & ChgSQL(strTemp(3)) & "','',3,'" & strUserNum & "') "
                 End Select
            End Select
            .MoveNext
        Loop
    End If
End With

'Added by Lydia 2023/12/01 增加「查名」統計TS案件之發文點數-扣除已銷帳的點數
DoEvents
strSql = " select cp14,st05,cp12,sp09,cp01,cp10,cp18,sum(nvl(a1u07,0)/1000) a1u07,cp27" & _
         " From caseprogress, servicepractice, Staff, acc1u0" & _
         " WHERE cp14=st01(+) And ST05 >='91' AND ST05<='99' AND CP27>=" & Val(txt1(1)) + 191100 & "01 AND CP27<=" & Val(txt1(2)) + 191100 & "31" & _
         " AND CP26 IS NULL and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and sp01='TS' and sp09='000'" & _
         " and cp09=a1u03(+) and cp60=a1u02(+) group by cp14,st05,cp12,sp09,cp01,cp10,cp18,cp27"
CheckOC
With adoRecordset
   .CursorLocation = adUseClient
   .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If .RecordCount <> 0 And .RecordCount > 0 Then
      .MoveFirst
      Do While .EOF = False
         If Val(.Fields("cp18")) - Val(.Fields("a1u07")) > 0 Then
            For i = 0 To 8
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            strTemp(15) = Val(.Fields("cp18")) - Val(.Fields("a1u07"))
            '個人目標
            strTemp3 = GetPerformanceByNickPE06(Val(txt1(1)) + 191100, Val(txt1(2)) + 191100, "T", strTemp(0))
            Select Case Mid(UCase(strTemp(2)), 1, 2)
            Case "S1"
                 Select Case UCase(strTemp(2))
                 Case "S11"
                      Select Case Val(strTemp(1))
                      '屬商申人員資料
                      Case 95 ', 15
                           cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086004,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'查名'," & Val(strTemp(15)) & "," & Val(strTemp(15)) & ",'" & ChgSQL(strTemp(0)) & "','*',4,'" & strUserNum & "') "
                      '屬商爭人員資料
                      Case Else
                           cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088004,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'查名'," & Val(strTemp(15)) & "," & Val(strTemp(15)) & ",'" & ChgSQL(strTemp(0)) & "','',4,'" & strUserNum & "') "
                      End Select
                 Case "S12"
                      Select Case Val(strTemp(1))
                      Case 95 ', 15
                           cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086005,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'查名'," & Val(strTemp(15)) & "," & Val(strTemp(15)) & ",'" & ChgSQL(strTemp(0)) & "','*',4,'" & strUserNum & "') "
                      Case Else
                           cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088005,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'查名'," & Val(strTemp(15)) & "," & Val(strTemp(15)) & ",'" & ChgSQL(strTemp(0)) & "','',4,'" & strUserNum & "') "
                      End Select
                 Case "S13"
                      Select Case Val(strTemp(1))
                      Case 95 ', 15
                           cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086006,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'查名'," & Val(strTemp(15)) & "," & Val(strTemp(15)) & ",'" & ChgSQL(strTemp(0)) & "','*',4,'" & strUserNum & "') "
                      Case Else
                           cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088006,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'查名'," & Val(strTemp(15)) & "," & Val(strTemp(15)) & ",'" & ChgSQL(strTemp(0)) & "','',4,'" & strUserNum & "') "
                      End Select
                 Case "S14"
                      Select Case Val(strTemp(1))
                      Case 95 ', 15
                           cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086007,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'查名'," & Val(strTemp(15)) & "," & Val(strTemp(15)) & ",'" & ChgSQL(strTemp(0)) & "','*',4,'" & strUserNum & "') "
                      Case Else
                           cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088007,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'查名'," & Val(strTemp(15)) & "," & Val(strTemp(15)) & ",'" & ChgSQL(strTemp(0)) & "','',4,'" & strUserNum & "') "
                      End Select
                 Case "S15"
                      Select Case Val(strTemp(1))
                      Case 95 ', 15
                           cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086008,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'查名'," & Val(strTemp(15)) & "," & Val(strTemp(15)) & ",'" & ChgSQL(strTemp(0)) & "','*',4,'" & strUserNum & "') "
                      Case Else
                           cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088008,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'查名'," & Val(strTemp(15)) & "," & Val(strTemp(15)) & ",'" & ChgSQL(strTemp(0)) & "','',4,'" & strUserNum & "') "
                      End Select
                 Case Else
                 End Select
            Case "S2"
                 Select Case Val(strTemp(1))
                 Case 95 ', 15
                      cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086009,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'查名'," & Val(strTemp(15)) & "," & Val(strTemp(15)) & ",'" & ChgSQL(strTemp(0)) & "','*',4,'" & strUserNum & "') "
                 Case Else
                      cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088009,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'查名'," & Val(strTemp(15)) & "," & Val(strTemp(15)) & ",'" & ChgSQL(strTemp(0)) & "','',4,'" & strUserNum & "') "
                 End Select
            Case "S3"
                 Select Case Val(strTemp(1))
                 Case 95 ', 15
                      cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086010,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'查名'," & Val(strTemp(15)) & "," & Val(strTemp(15)) & ",'" & ChgSQL(strTemp(0)) & "','*',4,'" & strUserNum & "') "
                 Case Else
                      cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088010,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'查名'," & Val(strTemp(15)) & "," & Val(strTemp(15)) & ",'" & ChgSQL(strTemp(0)) & "','',4,'" & strUserNum & "') "
                 End Select
            Case "S4"
                 Select Case Val(strTemp(1))
                 Case 95 ', 15
                      cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086011,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'查名'," & Val(strTemp(15)) & "," & Val(strTemp(15)) & ",'" & ChgSQL(strTemp(0)) & "','*',4,'" & strUserNum & "') "
                 Case Else
                      cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088011,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'查名'," & Val(strTemp(15)) & "," & Val(strTemp(15)) & ",'" & ChgSQL(strTemp(0)) & "','',4,'" & strUserNum & "') "
                 End Select
            Case Else
                 Select Case Val(strTemp(1))
                 Case 95 ', 15
                      cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086012,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'查名'," & Val(strTemp(15)) & "," & Val(strTemp(15)) & ",'" & ChgSQL(strTemp(0)) & "','*',4,'" & strUserNum & "') "
                 Case Else
                      cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088012,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'查名'," & Val(strTemp(15)) & "," & Val(strTemp(15)) & ",'" & ChgSQL(strTemp(0)) & "','',4,'" & strUserNum & "') "
                 End Select
            End Select
         End If
         .MoveNext
      Loop
   End If
End With
'end 'Added by Lydia 2023/12/01 增加「查名」統計TS案件之發文點數-扣除已銷帳的點數

CheckOC
DoEvents
'第一張商申下半部
'過期
'Modify By Sindy 2019/12/9
'strSql = "select st01,count(*),st01 from trademarkquery,staff where tmq10=st01(+) and TMQ11>=" & (Val(txt1(1)) + 1911) * 100 + Val(txt1(2)) & "01 AND TMQ11<=" & (Val(txt1(1)) + 1911) * 100 + Val(txt1(2)) & "31 and tmq11>tmq06 group by st01"
'Added by Lydia 2024/11/15 查名單(網中)：排除1120904-1120928期間資料匯入>>TO_CHAR(TMA04,'YYYYMMDD')>='20240601'
If bolIsTMA = True Then
   strSql = "select st01,count(*),st01 from tmqappform,staff where tma10=st01(+) and TMA14>=" & Val(txt1(1)) + 191100 & "01 AND TMA14<=" & Val(txt1(2)) + 191100 & "31 and tma14>nvl(tma11,tma12) and TO_CHAR(TMA04,'YYYYMMDD')>='20240601' group by st01"
Else
'end 2024/11/15
   strSql = "select st01,count(*),st01 from trademarkquery,staff where tmq10=st01(+) and TMQ11>=" & Val(txt1(1)) + 191100 & "01 AND TMQ11<=" & Val(txt1(2)) + 191100 & "31 and tmq11>tmq06 group by st01"
End If
'2019/12/9 END
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        Do While .EOF = False
            For i = 0 To 2
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            strSql = "insert into r020417_11 (r087001,r087003,r087006,id) values('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp(1)) & ",'" & ChgSQL(strTemp(2)) & "','" & strUserNum & "') "
            cnnConnection.Execute strSql
            'Modify By Sindy 2019/12/9
            If Val(txt1(1)) = Val(txt1(2)) Then
               'cnnConnection.Execute "UPDATE PERFORMANCE SET PE15=" & Val(strTemp(1)) & " WHERE PE01='" & ChgSQL(strTemp(0)) & "' AND PE02='T' AND PE03=" & Val(str(Val(txt1(1)) + 1911) & Format(txt1(2), "00")) & " "
               cnnConnection.Execute "UPDATE PERFORMANCE SET PE15=" & Val(strTemp(1)) & " WHERE PE01='" & ChgSQL(strTemp(0)) & "' AND PE02='T' AND PE03=" & Val(txt1(1)) + 191100 & " "
            End If
            '2019/12/9 END
            .MoveNext
        Loop
    End If
End With
CheckOC
'未輸入
'Modify By Sindy 2019/12/9
'strSql = "select st01,count(*),st01 from trademarkquery,staff where tmq10=st01(+) and (tmq11 is null or tmq11=0) AND tmq06>=" & Val(str(Val(txt1(1)) + 1911) & Format(txt1(2), "00")) & "01 and tmq06<=" & Val(str(Val(txt1(1)) + 1911) & Format(txt1(2), "00")) & "31 group by st01 "
'Added by Lydia 2024/11/15 查名單(網中)：排除1120904-1120928期間資料匯入>>TO_CHAR(TMA04,'YYYYMMDD')>='20240601'
If bolIsTMA = True Then
   strSql = "select st01,count(*),st01 from tmqappform,staff where tma10=st01(+) and nvl(tma14,0)=0 AND nvl(tma11,tma12)>=" & Val(txt1(1)) + 191100 & "01 AND nvl(tma11,tma12)<=" & Val(txt1(2)) + 191100 & "31 and TO_CHAR(TMA04,'YYYYMMDD')>='20240601' group by st01 "
Else
'end 2024/11/15
   strSql = "select st01,count(*),st01 from trademarkquery,staff where tmq10=st01(+) and (tmq11 is null or tmq11=0) AND tmq06>=" & Val(txt1(1)) + 191100 & "01 AND tmq06<=" & Val(txt1(2)) + 191100 & "31 group by st01 "
End If
'2019/12/9 END
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        Do While .EOF = False
            For i = 0 To 2
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            strSql = "insert into r020417_11 (r087001,r087004,r087006,id) values('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp(1)) & ",'" & ChgSQL(strTemp(2)) & "','" & strUserNum & "') "
            cnnConnection.Execute strSql
            'Modify By Sindy 2019/12/9
            If Val(txt1(1)) = Val(txt1(2)) Then
               'cnnConnection.Execute "UPDATE PERFORMANCE SET PE16=" & Val(strTemp(1)) & " WHERE PE01='" & ChgSQL(strTemp(0)) & "' AND PE02='T' AND PE03=" & Val(str(Val(txt1(1)) + 1911) & Format(txt1(2), "00")) & " "
               cnnConnection.Execute "UPDATE PERFORMANCE SET PE16=" & Val(strTemp(1)) & " WHERE PE01='" & ChgSQL(strTemp(0)) & "' AND PE02='T' AND PE03=" & Val(txt1(1)) + 191100 & " "
            End If
            '2019/12/9 END
            .MoveNext
        Loop
    End If
End With
CheckOC
'查名失誤
'Modify By Sindy 2019/12/9
'strSql = "select st01,pe20||decode(pe21,null,'',','||pe21)||decode(pe22,null,'',','||pe22)||decode(pe23,null,'',','||pe23)||decode(pe24,null,'',','||pe24)||decode(pe25,null,'',','||pe25)||decode(pe26,null,'',','||pe26)||decode(pe27,null,'',','||pe27)||decode(pe28,null,'',','||pe28)||decode(pe29,null,'',','||pe29),st01 from performance,staff where pe01=st01(+) and pe03=" & (Val(txt1(1)) + 1911) * 100 + Val(txt1(2)) & " and pe02='T' "
strSql = "select st01,pe20||decode(pe21,null,'',','||pe21)||decode(pe22,null,'',','||pe22)||decode(pe23,null,'',','||pe23)||decode(pe24,null,'',','||pe24)||decode(pe25,null,'',','||pe25)||decode(pe26,null,'',','||pe26)||decode(pe27,null,'',','||pe27)||decode(pe28,null,'',','||pe28)||decode(pe29,null,'',','||pe29),st01 from performance,staff where pe01=st01(+) and pe03>=" & Val(txt1(1)) + 191100 & " and pe03<=" & Val(txt1(2)) + 191100 & " and pe02='T' "
'2019/12/9 END
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        Do While .EOF = False
            For i = 0 To 2
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            'Add By Sindy 2017/10/19 計算件數
            If strTemp(1) = "" Then
               strTemp(1) = 0
            Else
               strTemp(1) = UBound(Split(strTemp(1), ",")) + 1
            End If
            '2017/10/19 END
            strSql = "insert into r020417_11 (r087001,r087005,r087006,id) values('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & strUserNum & "') "
            cnnConnection.Execute strSql
            .MoveNext
        Loop
    End If
End With
CheckOC
'達成率
strSql = "select R086001,DECODE(MAX(R086002),0,0,SUM(R086013) / DECODE(MAX(R086002),0,1,MAX(R086002)/100)),R086014 from R020417_1 where ID='" & strUserNum & "' GROUP BY R086014,R086001 "
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        Do While .EOF = False
            For i = 0 To 2
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            strSql = "insert into r020417_11 (r087001,r087002,r087006,id) values('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp(1)) & ",'" & ChgSQL(strTemp(2)) & "','" & strUserNum & "') "
            cnnConnection.Execute strSql
            .MoveNext
        Loop
    End If
End With
CheckOC
DoEvents
'第二張商爭下半部
'達成率
strSql = "select R088001,decode(max(r088002),0,0,SUM(R088013) / DECODE(MAX(R088002),0,1,max(r088002)/100)),R088014 from R020417_2 where ID='" & strUserNum & "' GROUP BY R088014,R088001 "
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        Do While .EOF = False
            For i = 0 To 2
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            strSql = "insert into r020417_21 (r089001,r089002,r089006,id) values('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp(1)) & ",'" & ChgSQL(strTemp(2)) & "','" & strUserNum & "') "
            cnnConnection.Execute strSql
            .MoveNext
        Loop
    End If
End With
CheckOC
'預估準確率,勝訴率1,勝訴率2
'Modify By Sindy 2019/12/9
'strSql = "select st01,pe17,pe18,pe19,st01 from performance,staff where pe01=st01(+) and pe03=" & (Val(txt1(1)) + 1911) * 100 + Val(txt1(2)) & " and pe02='T' "
strSql = "select st01,pe17,pe18,pe19,st01 from performance,staff where pe01=st01(+) and pe03>=" & Val(txt1(1)) + 191100 & " and pe03<=" & Val(txt1(2)) + 191100 & " and pe02='T' "
'2019/12/9 END
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        Do While .EOF = False
            For i = 0 To 4
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            strSql = "insert into r020417_21 (r089001,r089003,r089004,r089005,r089006,id) values('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp(1)) & "," & Val(strTemp(2)) & "," & Val(strTemp(3)) & ",'" & ChgSQL(strTemp(4)) & "','" & strUserNum & "') "
            cnnConnection.Execute strSql
            .MoveNext
        Loop
    End If
End With
CheckOC
DoEvents
'重整第一張，第二張
strSql = "select r086001,MAX(r086002),r086003,sum(r086004),sum(r086005),sum(r086006),sum(r086007),sum(r086008),sum(r086009),sum(r086010),sum(r086011),sum(r086012),sum(r086013),r086014,r086015,max(r086016),id from r020417_1 where id='" & strUserNum & "' group by id,r086001,r086003,r086014,r086015 "
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        cnnConnection.Execute "delete from r020417_1 where id='" & strUserNum & "' "
        .MoveFirst
        Do While .EOF = False
            strSql = "insert into r020417_1 (r086001,r086002,r086003,r086004,r086005,r086006,r086007,r086008,r086009,r086010,r086011,r086012,r086013,r086014,r086015,r086016,id) values('" & ChgSQL(CheckStr(.Fields(0))) & "'," & Val(CheckStr(.Fields(1))) & ",'" & ChgSQL(CheckStr(.Fields(2))) & "'," & Val(CheckStr(.Fields(3))) & "," & Val(CheckStr(.Fields(4))) & "," & Val(CheckStr(.Fields(5))) & "," & Val(CheckStr(.Fields(6))) & "," & Val(CheckStr(.Fields(7))) & "," & Val(CheckStr(.Fields(8))) & "," & Val(CheckStr(.Fields(9))) & "," & Val(CheckStr(.Fields(10))) & "," & Val(CheckStr(.Fields(11))) & "," & Val(CheckStr(.Fields(12))) & ",'" & ChgSQL(CheckStr(.Fields(13))) & "','" & ChgSQL(CheckStr(.Fields(14))) & "'," & Val(CheckStr(.Fields(15))) & ",'" & ChgSQL(CheckStr(.Fields(16))) & "') "
            cnnConnection.Execute strSql
            .MoveNext
        Loop
    End If
End With
CheckOC
DoEvents
strSql = "select r088001,MAX(r088002),r088003,sum(r088004),sum(r088005),sum(r088006),sum(r088007),sum(r088008),sum(r088009),sum(r088010),sum(r088011),sum(r088012),sum(r088013),r088014,r088015,max(r088016),id from r020417_2 where id='" & strUserNum & "' group by id,r088001,r088003,r088014,r088015 "
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        cnnConnection.Execute "delete from r020417_2 where id='" & strUserNum & "' "
        .MoveFirst
        Do While .EOF = False
            strSql = "insert into r020417_2 (r088001,r088002,r088003,r088004,r088005,r088006,r088007,r088008,r088009,r088010,r088011,r088012,r088013,r088014,r088015,r088016,id) values('" & ChgSQL(CheckStr(.Fields(0))) & "'," & Val(CheckStr(.Fields(1))) & ",'" & ChgSQL(CheckStr(.Fields(2))) & "'," & Val(CheckStr(.Fields(3))) & "," & Val(CheckStr(.Fields(4))) & "," & Val(CheckStr(.Fields(5))) & "," & Val(CheckStr(.Fields(6))) & "," & Val(CheckStr(.Fields(7))) & "," & Val(CheckStr(.Fields(8))) & "," & Val(CheckStr(.Fields(9))) & "," & Val(CheckStr(.Fields(10))) & "," & Val(CheckStr(.Fields(11))) & "," & Val(CheckStr(.Fields(12))) & ",'" & ChgSQL(CheckStr(.Fields(13))) & "','" & ChgSQL(CheckStr(.Fields(14))) & "'," & Val(CheckStr(.Fields(15))) & ",'" & ChgSQL(CheckStr(.Fields(16))) & "') "
            cnnConnection.Execute strSql
            .MoveNext
        Loop
    End If
End With
CheckOC
DoEvents

'加入沒有案件性質的資料
strSql = "SELECT DISTINCT R086001,R086014 FROM R020417_1 WHERE ID='" & strUserNum & "' "
strSql = strSql + " union all select DISTINCT R088001,R088014 FROM R020417_2 WHERE ID='" & strUserNum & "' "
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        Do While .EOF = False
               cnnConnection.Execute "INSERT INTO R020417_1 VALUES ('" & ChgSQL(CheckStr(.Fields(0))) & "',0,'商申',0,0,0,0,0,0,0,0,0,0,'" & ChgSQL(CheckStr(.Fields(1))) & "','',1,'" & strUserNum & "') "
               cnnConnection.Execute "INSERT INTO R020417_1 VALUES ('" & ChgSQL(CheckStr(.Fields(0))) & "',0,'商爭',0,0,0,0,0,0,0,0,0,0,'" & ChgSQL(CheckStr(.Fields(1))) & "','',2,'" & strUserNum & "') "
               cnnConnection.Execute "INSERT INTO R020417_1 VALUES ('" & ChgSQL(CheckStr(.Fields(0))) & "',0,'國內查名',0,0,0,0,0,0,0,0,0,0,'" & ChgSQL(CheckStr(.Fields(1))) & "','',3,'" & strUserNum & "') "
               cnnConnection.Execute "INSERT INTO R020417_1 VALUES ('" & ChgSQL(CheckStr(.Fields(0))) & "',0,'查名',0,0,0,0,0,0,0,0,0,0,'" & ChgSQL(CheckStr(.Fields(1))) & "','',4,'" & strUserNum & "') " 'Added by Lydia 2023/12/01
               cnnConnection.Execute "INSERT INTO R020417_1 VALUES ('" & ChgSQL(CheckStr(.Fields(0))) & "',0,'大陸',0,0,0,0,0,0,0,0,0,0,'" & ChgSQL(CheckStr(.Fields(1))) & "','',5,'" & strUserNum & "') "  'Modified by Lydia 2023/12/01 4=>5;後面Index+1
               cnnConnection.Execute "INSERT INTO R020417_1 VALUES ('" & ChgSQL(CheckStr(.Fields(0))) & "',0,'馬德里',0,0,0,0,0,0,0,0,0,0,'" & ChgSQL(CheckStr(.Fields(1))) & "','',6,'" & strUserNum & "') "
               cnnConnection.Execute "INSERT INTO R020417_1 VALUES ('" & ChgSQL(CheckStr(.Fields(0))) & "',0,'條碼',0,0,0,0,0,0,0,0,0,0,'" & ChgSQL(CheckStr(.Fields(1))) & "','',7,'" & strUserNum & "') "
               cnnConnection.Execute "INSERT INTO R020417_1 VALUES ('" & ChgSQL(CheckStr(.Fields(0))) & "',0,'監視系統',0,0,0,0,0,0,0,0,0,0,'" & ChgSQL(CheckStr(.Fields(1))) & "','',8,'" & strUserNum & "') "
               cnnConnection.Execute "INSERT INTO R020417_1 VALUES ('" & ChgSQL(CheckStr(.Fields(0))) & "',0,'其他',0,0,0,0,0,0,0,0,0,0,'" & ChgSQL(CheckStr(.Fields(1))) & "','',9,'" & strUserNum & "') "
               
               cnnConnection.Execute "INSERT INTO R020417_2 VALUES ('" & ChgSQL(CheckStr(.Fields(0))) & "',0,'商申',0,0,0,0,0,0,0,0,0,0,'" & ChgSQL(CheckStr(.Fields(1))) & "','',1,'" & strUserNum & "') "
               cnnConnection.Execute "INSERT INTO R020417_2 VALUES ('" & ChgSQL(CheckStr(.Fields(0))) & "',0,'商爭',0,0,0,0,0,0,0,0,0,0,'" & ChgSQL(CheckStr(.Fields(1))) & "','',2,'" & strUserNum & "') "
               cnnConnection.Execute "INSERT INTO R020417_2 VALUES ('" & ChgSQL(CheckStr(.Fields(0))) & "',0,'國內查名',0,0,0,0,0,0,0,0,0,0,'" & ChgSQL(CheckStr(.Fields(1))) & "','',3,'" & strUserNum & "') "
               cnnConnection.Execute "INSERT INTO R020417_2 VALUES ('" & ChgSQL(CheckStr(.Fields(0))) & "',0,'查名',0,0,0,0,0,0,0,0,0,0,'" & ChgSQL(CheckStr(.Fields(1))) & "','',4,'" & strUserNum & "') " 'Added by Lydia 2023/12/01
               cnnConnection.Execute "INSERT INTO R020417_2 VALUES ('" & ChgSQL(CheckStr(.Fields(0))) & "',0,'大陸',0,0,0,0,0,0,0,0,0,0,'" & ChgSQL(CheckStr(.Fields(1))) & "','',5,'" & strUserNum & "') " 'Modified by Lydia 2023/12/01 4=>5;後面Index+1
               cnnConnection.Execute "INSERT INTO R020417_2 VALUES ('" & ChgSQL(CheckStr(.Fields(0))) & "',0,'馬德里',0,0,0,0,0,0,0,0,0,0,'" & ChgSQL(CheckStr(.Fields(1))) & "','',6,'" & strUserNum & "') "
               cnnConnection.Execute "INSERT INTO R020417_2 VALUES ('" & ChgSQL(CheckStr(.Fields(0))) & "',0,'條碼',0,0,0,0,0,0,0,0,0,0,'" & ChgSQL(CheckStr(.Fields(1))) & "','',7,'" & strUserNum & "') "
               cnnConnection.Execute "INSERT INTO R020417_2 VALUES ('" & ChgSQL(CheckStr(.Fields(0))) & "',0,'監視系統',0,0,0,0,0,0,0,0,0,0,'" & ChgSQL(CheckStr(.Fields(1))) & "','',8,'" & strUserNum & "') "
               cnnConnection.Execute "INSERT INTO R020417_2 VALUES ('" & ChgSQL(CheckStr(.Fields(0))) & "',0,'其他',0,0,0,0,0,0,0,0,0,0,'" & ChgSQL(CheckStr(.Fields(1))) & "','',9,'" & strUserNum & "') "
               
               cnnConnection.Execute "INSERT INTO R020417_11 VALUES ('" & ChgSQL(CheckStr(.Fields(0))) & "',0,0,0,'','" & ChgSQL(CheckStr(.Fields(1))) & "','','" & strUserNum & "') "
               
               cnnConnection.Execute "INSERT INTO R020417_21 VALUES ('" & ChgSQL(CheckStr(.Fields(0))) & "',0,0,0,0,'" & ChgSQL(CheckStr(.Fields(1))) & "','','" & strUserNum & "') "
            .MoveNext
        Loop
    End If
End With
CheckOC
'第三張
strSql = " select max(r086002),r086003,sum(r086004),sum(r086005),sum(r086006),sum(r086007),sum(r086008),sum(r086009),sum(r086010),sum(r086011),sum(r086012),sum(r086013),r086016 from r020417_1 where id='" & strUserNum & "' group by r086016,r086003 "
strSql = strSql + "union all select max(r088002),r088003,sum(r088004),sum(r088005),sum(r088006),sum(r088007),sum(r088008),sum(r088009),sum(r088010),sum(r088011),sum(r088012),sum(r088013),r088016 from r020417_2 where id='" & strUserNum & "' group by r088016,r088003 "
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        Do While .EOF = False
            For i = 0 To 12
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            strSql = "insert into r020417_3 values(" & Val(strTemp(0)) & ",'" & ChgSQL(strTemp(1)) & "'," & Val(strTemp(2)) & "," & Val(strTemp(3)) & "," & Val(strTemp(4)) & "," & Val(strTemp(5)) & "," & Val(strTemp(6)) & "," & Val(strTemp(7)) & "," & Val(strTemp(8)) & "," & Val(strTemp(9)) & "," & Val(strTemp(10)) & "," & Val(strTemp(11)) & "," & Val(strTemp(12)) & ",'" & strUserNum & "') "
            cnnConnection.Execute strSql
            .MoveNext
        Loop
    End If
End With
CheckOC
'第四張
'Modified by Lydia 2023/12/01 +R091013
'strSql = "select max(r086002),r086001,sum(r086013),0,0,0,0,0,0,0,sum(r086013),1,id from r020417_1 where id='" & strUserNum & "' and r086016=1 group by id,r086001 "
'strSql = strSql + " union all select max(r086002),r086001,0,sum(r086013),0,0,0,0,0,0,sum(r086013),2,id from r020417_1 where id='" & strUserNum & "' and r086016=2 group by id,r086001 "
'strSql = strSql + " union all select max(r086002),r086001,0,0,sum(r086013),0,0,0,0,0,sum(r086013),3,id from r020417_1 where id='" & strUserNum & "' and r086016=3 group by id,r086001 "
'strSql = strSql + " union all select max(r086002),r086001,0,0,0,sum(r086013),0,0,0,0,sum(r086013),4,id from r020417_1 where id='" & strUserNum & "' and r086016=4 group by id,r086001 "
'strSql = strSql + " union all select max(r086002),r086001,0,0,0,0,sum(r086013),0,0,0,sum(r086013),5,id from r020417_1 where id='" & strUserNum & "' and r086016=5 group by id,r086001 "
'strSql = strSql + " union all select max(r086002),r086001,0,0,0,0,0,sum(r086013),0,0,sum(r086013),6,id from r020417_1 where id='" & strUserNum & "' and r086016=6 group by id,r086001 "
'strSql = strSql + " union all select max(r086002),r086001,0,0,0,0,0,0,sum(r086013),0,sum(r086013),7,id from r020417_1 where id='" & strUserNum & "' and r086016=7 group by id,r086001 "
'strSql = strSql + " union all select max(r086002),r086001,0,0,0,0,0,0,0,sum(r086013),sum(r086013),8,id from r020417_1 where id='" & strUserNum & "' and r086016=8 group by id,r086001 "
'strSql = strSql + " union all select max(r088002),r088001,sum(r088013),0,0,0,0,0,0,0,sum(r088013),1,id from r020417_2 where id='" & strUserNum & "' and r088016=1 group by id,r088001 "
'strSql = strSql + " union all select max(r088002),r088001,0,sum(r088013),0,0,0,0,0,0,sum(r088013),2,id from r020417_2 where id='" & strUserNum & "' and r088016=2 group by id,r088001 "
'strSql = strSql + " union all select max(r088002),r088001,0,0,sum(r088013),0,0,0,0,0,sum(r088013),3,id from r020417_2 where id='" & strUserNum & "' and r088016=3 group by id,r088001 "
'strSql = strSql + " union all select max(r088002),r088001,0,0,0,sum(r088013),0,0,0,0,sum(r088013),4,id from r020417_2 where id='" & strUserNum & "' and r088016=4 group by id,r088001 "
'strSql = strSql + " union all select max(r088002),r088001,0,0,0,0,sum(r088013),0,0,0,sum(r088013),5,id from r020417_2 where id='" & strUserNum & "' and r088016=5 group by id,r088001 "
'strSql = strSql + " union all select max(r088002),r088001,0,0,0,0,0,sum(r088013),0,0,sum(r088013),6,id from r020417_2 where id='" & strUserNum & "' and r088016=6 group by id,r088001 "
'strSql = strSql + " union all select max(r088002),r088001,0,0,0,0,0,0,sum(r088013),0,sum(r088013),7,id from r020417_2 where id='" & strUserNum & "' and r088016=7 group by id,r088001 "
'strSql = strSql + " union all select max(r088002),r088001,0,0,0,0,0,0,0,sum(r088013),sum(r088013),8,id from r020417_2 where id='" & strUserNum & "' and r088016=8 group by id,r088001 "
''''strSql = "select max(r086002),r086001,sum(r086013),0,0,0,0,0,0,0,sum(r086013),1,id,0 from r020417_1 where id='" & strUserNum & "' and r086016=1 group by id,r086001 "
''''strSql = strSql + " union all select max(r086002),r086001,0,sum(r086013),0,0,0,0,0,0,sum(r086013),2,id,0 from r020417_1 where id='" & strUserNum & "' and r086016=2 group by id,r086001 "
''''strSql = strSql + " union all select max(r086002),r086001,0,0,sum(r086013),0,0,0,0,0,sum(r086013),3,id,0 from r020417_1 where id='" & strUserNum & "' and r086016=3 group by id,r086001 "
''''strSql = strSql + " union all select max(r086002),r086001,0,0,0,sum(r086013),0,0,0,0,0,4,id,sum(r086013) from r020417_1 where id='" & strUserNum & "' and r086016=4 group by id,r086001 "
''''strSql = strSql + " union all select max(r086002),r086001,0,0,0,0,sum(r086013),0,0,0,sum(r086013),5,id,0 from r020417_1 where id='" & strUserNum & "' and r086016=5 group by id,r086001 "
''''strSql = strSql + " union all select max(r086002),r086001,0,0,0,0,0,sum(r086013),0,0,sum(r086013),6,id,0 from r020417_1 where id='" & strUserNum & "' and r086016=6 group by id,r086001 "
''''strSql = strSql + " union all select max(r086002),r086001,0,0,0,0,0,0,sum(r086013),0,sum(r086013),7,id,0 from r020417_1 where id='" & strUserNum & "' and r086016=7 group by id,r086001 "
''''strSql = strSql + " union all select max(r086002),r086001,0,0,0,0,0,0,0,sum(r086013),sum(r086013),8,id,0 from r020417_1 where id='" & strUserNum & "' and r086016=8 group by id,r086001 "
''''strSql = strSql + " union all select max(r086002),r086001,0,0,0,0,0,0,0,sum(r086013),sum(r086013),9,id,0 from r020417_1 where id='" & strUserNum & "' and r086016=9 group by id,r086001 "
''''strSql = strSql + " union all select max(r088002),r088001,sum(r088013),0,0,0,0,0,0,0,sum(r088013),1,id,0 from r020417_2 where id='" & strUserNum & "' and r088016=1 group by id,r088001 "
''''strSql = strSql + " union all select max(r088002),r088001,0,sum(r088013),0,0,0,0,0,0,sum(r088013),2,id,0 from r020417_2 where id='" & strUserNum & "' and r088016=2 group by id,r088001 "
''''strSql = strSql + " union all select max(r088002),r088001,0,0,sum(r088013),0,0,0,0,0,sum(r088013),3,id,0 from r020417_2 where id='" & strUserNum & "' and r088016=3 group by id,r088001 "
''''strSql = strSql + " union all select max(r088002),r088001,0,0,0,0,0,0,0,0,sum(r088013),4,id,sum(r088013) from r020417_2 where id='" & strUserNum & "' and r088016=4 group by id,r088001 "
''''strSql = strSql + " union all select max(r088002),r088001,0,0,0,sum(r088013),0,0,0,0,sum(r088013),5,id,0 from r020417_2 where id='" & strUserNum & "' and r088016=5 group by id,r088001 "
''''strSql = strSql + " union all select max(r088002),r088001,0,0,0,0,sum(r088013),0,0,0,sum(r088013),6,id,0 from r020417_2 where id='" & strUserNum & "' and r088016=6 group by id,r088001 "
''''strSql = strSql + " union all select max(r088002),r088001,0,0,0,0,0,sum(r088013),0,0,sum(r088013),7,id,0 from r020417_2 where id='" & strUserNum & "' and r088016=7 group by id,r088001 "
''''strSql = strSql + " union all select max(r088002),r088001,0,0,0,0,0,0,sum(r088013),0,sum(r088013),8,id,0 from r020417_2 where id='" & strUserNum & "' and r088016=8 group by id,r088001 "
''''strSql = strSql + " union all select max(r088002),r088001,0,0,0,0,0,0,0,sum(r088013),sum(r088013),9,id,0 from r020417_2 where id='" & strUserNum & "' and r088016=9 group by id,r088001 "
'''''end 2023/12/01
strSql = "select max(r086002),r086001,sum(r086013),0,0,0,0,0,0,0,sum(r086013),1,id,0 from r020417_1 where id='" & strUserNum & "' and r086016=1 group by id,r086001 "
strSql = strSql + " union all select max(r086002),r086001,0,sum(r086013),0,0,0,0,0,0,sum(r086013),2,id,0 from r020417_1 where id='" & strUserNum & "' and r086016=2 group by id,r086001 "
strSql = strSql + " union all select max(r086002),r086001,0,0,sum(r086013),0,0,0,0,0,sum(r086013),3,id,0 from r020417_1 where id='" & strUserNum & "' and r086016=3 group by id,r086001 "
strSql = strSql + " union all select max(r086002),r086001,0,0,0,0,0,0,0,0,sum(r086013),4,id,sum(r086013) from r020417_1 where id='" & strUserNum & "' and r086016=4 group by id,r086001 "
strSql = strSql + " union all select max(r086002),r086001,0,0,0,sum(r086013),0,0,0,0,sum(r086013),5,id,0 from r020417_1 where id='" & strUserNum & "' and r086016=5 group by id,r086001 "
strSql = strSql + " union all select max(r086002),r086001,0,0,0,0,sum(r086013),0,0,0,sum(r086013),6,id,0 from r020417_1 where id='" & strUserNum & "' and r086016=6 group by id,r086001 "
strSql = strSql + " union all select max(r086002),r086001,0,0,0,0,0,sum(r086013),0,0,sum(r086013),7,id,0 from r020417_1 where id='" & strUserNum & "' and r086016=7 group by id,r086001 "
strSql = strSql + " union all select max(r086002),r086001,0,0,0,0,0,0,sum(r086013),0,sum(r086013),8,id,0 from r020417_1 where id='" & strUserNum & "' and r086016=8 group by id,r086001 "
strSql = strSql + " union all select max(r086002),r086001,0,0,0,0,0,0,0,sum(r086013),sum(r086013),9,id,0 from r020417_1 where id='" & strUserNum & "' and r086016=9 group by id,r086001 "
strSql = strSql + " union all select max(r088002),r088001,sum(r088013),0,0,0,0,0,0,0,sum(r088013),1,id,0 from r020417_2 where id='" & strUserNum & "' and r088016=1 group by id,r088001 "
strSql = strSql + " union all select max(r088002),r088001,0,sum(r088013),0,0,0,0,0,0,sum(r088013),2,id,0 from r020417_2 where id='" & strUserNum & "' and r088016=2 group by id,r088001 "
strSql = strSql + " union all select max(r088002),r088001,0,0,sum(r088013),0,0,0,0,0,sum(r088013),3,id,0 from r020417_2 where id='" & strUserNum & "' and r088016=3 group by id,r088001 "
strSql = strSql + " union all select max(r088002),r088001,0,0,0,0,0,0,0,0,sum(r088013),4,id,sum(r088013) from r020417_2 where id='" & strUserNum & "' and r088016=4 group by id,r088001 "
strSql = strSql + " union all select max(r088002),r088001,0,0,0,sum(r088013),0,0,0,0,sum(r088013),5,id,0 from r020417_2 where id='" & strUserNum & "' and r088016=5 group by id,r088001 "
strSql = strSql + " union all select max(r088002),r088001,0,0,0,0,sum(r088013),0,0,0,sum(r088013),6,id,0 from r020417_2 where id='" & strUserNum & "' and r088016=6 group by id,r088001 "
strSql = strSql + " union all select max(r088002),r088001,0,0,0,0,0,sum(r088013),0,0,sum(r088013),7,id,0 from r020417_2 where id='" & strUserNum & "' and r088016=7 group by id,r088001 "
strSql = strSql + " union all select max(r088002),r088001,0,0,0,0,0,0,sum(r088013),0,sum(r088013),8,id,0 from r020417_2 where id='" & strUserNum & "' and r088016=8 group by id,r088001 "
strSql = strSql + " union all select max(r088002),r088001,0,0,0,0,0,0,0,sum(r088013),sum(r088013),9,id,0 from r020417_2 where id='" & strUserNum & "' and r088016=9 group by id,r088001 "
strSql = "insert into r020417_4 " & strSql
cnnConnection.Execute strSql
DoEvents
'Add By Cheng 2002/04/10
'刪除總計為零的資料
strSql = "select r086001,sum(r086013) from r020417_1 where id='" & strUserNum & "' group by r086001 "
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        Do While .EOF = False
            If .Fields(1).Value = 0 Then
               cnnConnection.Execute "Delete From R020417_1 where R086001='" & .Fields(0).Value & "' and id='" & strUserNum & "'"
            End If
            .MoveNext
        Loop
    End If
End With
CheckOC
DoEvents
strSql = "select r088001,sum(r088013) from r020417_2 where id='" & strUserNum & "' group by r088001 "
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        Do While .EOF = False
            If .Fields(1).Value = 0 Then
               cnnConnection.Execute "Delete From R020417_2 where R088001='" & .Fields(0).Value & "' and id='" & strUserNum & "'"
            End If
            .MoveNext
        Loop
    End If
End With
CheckOC

DoEvents

'frm100.Tag = "5=5"
'frm100.StrMenu

'Removed by Morgan 2015/6/3
'DoEvents
'Printer.Orientation = 1
'DoEvents
'end 2015/6/3

PrintData1
'Removed by Morgan 2015/6/3
'Printer.Orientation = 1
'DoEvents
'If Combo1.ListIndex >= SeekPrint Then
''   j = Combo1.ListIndex + 1
'Else
'   j = Combo1.ListIndex
'End If
'Set Printer = Printers(j)
'Printer.PaperSize = 39
'end 2015/6/3

PrintData2
'Removed by Morgan 2015/6/3
'Printer.Orientation = 1
'DoEvents
'If Combo1.ListIndex >= SeekPrint Then
'   j = Combo1.ListIndex + 1
'Else
'   j = Combo1.ListIndex
'End If
'Set Printer = Printers(j)
'Printer.PaperSize = 39
'end 2015/6/3

PrintData3
'Removed by Morgan 2015/6/3
'Printer.Orientation = 1
'DoEvents
'If Combo1.ListIndex >= SeekPrint Then
'   j = Combo1.ListIndex + 1
'Else
'   j = Combo1.ListIndex
'End If
'Set Printer = Printers(j)
'Printer.PaperSize = 39
'end 2015/6/3

PrintData4
ShowPrintOk
Printer.EndDoc 'Add By Sindy 2011/11/1
Screen.MousePointer = vbDefault
End Sub

'Sub Process()
''因為當初一開始測試時，宋小姐有協議說改部分
''但忘記改哪裡了，而邱說把他改成與地雷條款相同，
''其他的不管，有問題她自己改     91/3/12 nick
'Screen.MousePointer = vbHourglass
'cnnConnection.Execute "DELETE FROM R020417_1 WHERE ID='" & strUserNum & "' "
'cnnConnection.Execute "DELETE FROM R020417_11 WHERE ID='" & strUserNum & "' "
'cnnConnection.Execute "DELETE FROM R020417_2 WHERE ID='" & strUserNum & "' "
'cnnConnection.Execute "DELETE FROM R020417_21 WHERE ID='" & strUserNum & "' "
'cnnConnection.Execute "DELETE FROM R020417_3 WHERE ID='" & strUserNum & "' "
'cnnConnection.Execute "DELETE FROM R020417_4 WHERE ID='" & strUserNum & "' "
'strSQL1 = ""
'strSQL2 = ""
'StrSQL6 = ""
'If Len(txt1(0)) <> 0 Then
'   strSQL1 = strSQL1 + " AND CP01 IN (" & SQLGrpStr(txt1(0), 2) & ") "
'   strSQL2 = strSQL2 + " AND CP01 IN (" & SQLGrpStr(txt1(0), 5) & ") "
'End If
'StrSQL6 = ""
''Modify By Cheng 2002/11/28
''取消申請國家範圍條件(另申請國家範圍欄位仍保留, 但隱藏在印表機選項後)
''If Len(txt1(3)) <> 0 Then
''    strSQL1 = strSQL1 + " AND TM10>=" & txt1(3)
''    strSQL2 = strSQL2 + " AND SP09>=" & txt1(3)
''End If
''If Len(txt1(4)) <> 0 Then
''    strSQL1 = strSQL1 + " AND TM10<=" & txt1(4)
''    strSQL2 = strSQL2 + " AND SP09<=" & txt1(4)
''End If
'CheckOC
''案件進度檔
'DoEvents
'strSql = "SELECT S1.ST02,S1.ST05,cp12,CP14,TM10,CP01,CP10,cp18 FROM CASEPROGRESS,TRADEMARK,STAFF S1,STAFF S2 WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP13=S2.ST01(+) AND CP14=S1.ST01(+) AND CP27>=" & (Val(txt1(1)) + 1911) * 100 + Val(Format(txt1(2), "00")) & "01 AND CP27<=" & (Val(txt1(1)) + 1911) * 100 + Val(Format(txt1(2), "00")) & "31 AND CP26 IS NULL  " & strSQL1
'strSql = strSql + " union all select S1.ST02,S1.ST05,cp12,CP14,SP09,CP01,CP10,cp18 FROM CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2 WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP13=S2.ST01(+) AND CP14=S1.ST01(+) AND CP27>=" & (Val(txt1(1)) + 1911) * 100 + Val(Format(txt1(2), "00")) & "01 AND CP27<=" & (Val(txt1(1)) + 1911) * 100 + Val(Format(txt1(2), "00")) & "31 AND CP26 IS NULL  " & strSQL2
'With adoRecordset
'    .CursorLocation = adUseClient
'    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'    If .RecordCount <> 0 And .RecordCount > 0 Then
'        .MoveFirst
'        Do While .EOF = False
'            For i = 0 To 5
'                strTemp(i) = CheckStr(.Fields(i))
'            Next i
'            '個人目標
'            strTemp3 = GetPerformanceByNickPE06(Val(str(Val(txt1(1)) + 1911) & Format(txt1(2), "00")), Val(str(Val(txt1(1)) + 1911) & Format(txt1(2), "00")), "T", strTemp(3))
'            '依系統類別與國家分類
'            Select Case strTemp(5)
''T              *******************************************************************************************************
'            Case "T"
'                 ProcessT
''TS         **********************************************************************************************************************
'            Case "TS"
'                 ProcessTS
''TF         *********************************************************************************************************************************
'            Case "TF"
'                 ProcessTF
''TB         *************************************************************************************************************************************************
'            Case "TB"
'                 ProcessTB
''TM         ***********************************************************************************************************************************************************************************************
'            Case "TM"
'                 ProcessTM
''其他       ********************************************************************************************************************************************************************************
'            Case Else
'                 ProcessOther
'            End Select
'            .MoveNext
'        Loop
'    Else
'        ShowNoData
'        Screen.MousePointer = vbDefault
'        Exit Sub
'    End If
'End With
'CheckOC
''商標委查檔--更新個人目標資料
'DoEvents
'strSql = "SELECT TMQ10,S1.ST05,S2.ST03,TMQ10,TMQ07+TMQ08+TMQ09,'','',TMQ07,TMQ08,TMQ09 FROM TRADEMARKQUERY,STAFF S1,STAFF S2 WHERE TMQ02=S2.ST01(+) AND TMQ10=S1.ST01(+) AND TMQ11>=" & (Val(txt1(1)) + 1911) * 100 + Val(txt1(2)) & "01 AND TMQ11<=" & (Val(txt1(1)) + 1911) * 100 + Val(txt1(2)) & "31 and TMQ11<=TMQ06 "
'CheckOC2
'With adoRecordset
'    .CursorLocation = adUseClient
'    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'    If .RecordCount <> 0 And .RecordCount > 0 Then
'        .MoveFirst
'        Do While .EOF = False
'            For i = 0 To 9
'                strTemp(i) = CheckStr(.Fields(i))
'            Next i
'            cnnConnection.Execute "UPDATE PERFORMANCE SET PE12=" & Val(strTemp(7)) & ",PE13=" & Val(strTemp(8)) & ",PE14=" & Val(strTemp(9)) & " WHERE PE01='" & ChgSQL(strTemp(3)) & "' AND PE02='T' AND PE03=" & Val(str(Val(txt1(1)) + 1911) & Format(txt1(2), "00")) & " "
'            .MoveNext
'        Loop
'    End If
'End With
'CheckOC
'
'DoEvents
''第一張商申下半部
''過期
'strSql = "select pe01,count(*),st01 from trademarkquery,staff where tmq10=st01(+) and  TMQ11>=" & (Val(txt1(1)) + 1911) * 100 + Val(txt1(2)) & "01 AND TMQ11<=" & (Val(txt1(1)) + 1911) * 100 + Val(txt1(2)) & "31 and tmq11>tmq06 group by st01,st02"
'CheckOC
'With adoRecordset
'    .CursorLocation = adUseClient
'    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'    If .RecordCount <> 0 And .RecordCount > 0 Then
'        .MoveFirst
'        Do While .EOF = False
'            For i = 0 To 2
'                strTemp(i) = CheckStr(.Fields(i))
'            Next i
'            strSql = "insert into r020417_11 (r087001,r087003,r087006,id) values('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp(1)) & ",'" & ChgSQL(strTemp(2)) & "','" & strUserNum & "') "
'            cnnConnection.Execute strSql
'            .MoveNext
'        Loop
'    End If
'End With
'CheckOC
''未輸入
'strSql = "select pe01,count(*),st01 from trademarkquery,staff where tmq10=st01(+) and (tmq11 is null or tmq11=0) group by st01,st02 "
'CheckOC
'With adoRecordset
'    .CursorLocation = adUseClient
'    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'    If .RecordCount <> 0 And .RecordCount > 0 Then
'        .MoveFirst
'        Do While .EOF = False
'            For i = 0 To 2
'                strTemp(i) = CheckStr(.Fields(i))
'            Next i
'            strSql = "insert into r020417_11 (r087001,r087004,r087006,id) values('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp(1)) & ",'" & ChgSQL(strTemp(2)) & "','" & strUserNum & "') "
'            cnnConnection.Execute strSql
'            .MoveNext
'        Loop
'    End If
'End With
'CheckOC
''查名失誤
'strSql = "select pe01,pe20||decode(pe21,null,'',','||pe21)||decode(pe22,null,'',','||pe22)||decode(pe23,null,'',','||pe23)||decode(pe24,null,'',','||pe24)||decode(pe25,null,'',','||pe25)||decode(pe26,null,'',','||pe26)||decode(pe27,null,'',','||pe27)||decode(pe28,null,'',','||pe28)||decode(pe29,null,'',','||pe29),st01 from performance,staff where pe01=st01(+) and pe03=" & (Val(txt1(1)) + 1911) * 100 + Val(txt1(2)) & " and pe02='T' "
'CheckOC
'With adoRecordset
'    .CursorLocation = adUseClient
'    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'    If .RecordCount <> 0 And .RecordCount > 0 Then
'        .MoveFirst
'        Do While .EOF = False
'            For i = 0 To 2
'                strTemp(i) = CheckStr(.Fields(i))
'            Next i
'            'Add By Sindy 2017/10/19 計算件數
'            If strTemp(1) = "" Then
'               strTemp(1) = 0
'            Else
'               strTemp(1) = UBound(Split(strTemp(1), ",")) + 1
'            End If
'            '2017/10/19 END
'            strSql = "insert into r020417_11 (r087001,r087005,r087006,id) values('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & strUserNum & "') "
'            cnnConnection.Execute strSql
'            .MoveNext
'        Loop
'    End If
'End With
'CheckOC
''達成率
'strSql = "select R086001,DECODE(MAX(R086002),0,0,SUM(R086013) / DECODE(MAX(R086002),0,1,MAX(R086002)/100)),R086014 from R020417_1 where ID='" & strUserNum & "' GROUP BY R086014,R086001 "
'CheckOC
'With adoRecordset
'    .CursorLocation = adUseClient
'    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'    If .RecordCount <> 0 And .RecordCount > 0 Then
'        .MoveFirst
'        Do While .EOF = False
'            For i = 0 To 2
'                strTemp(i) = CheckStr(.Fields(i))
'            Next i
'            strSql = "insert into r020417_11 (r087001,r087002,r087006,id) values('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp(1)) & ",'" & ChgSQL(strTemp(2)) & "','" & strUserNum & "') "
'            cnnConnection.Execute strSql
'            .MoveNext
'        Loop
'    End If
'End With
'CheckOC
''frm100.Tag = "5=3"
''frm100.StrMenu
'DoEvents
''第二張商爭下半部
''達成率
'strSql = "select R088001,decode(max(r088002),0,0,SUM(R088013) / DECODE(MAX(R088002),0,1,max(r088002)/100)),R088014 from R020417_2 where ID='" & strUserNum & "' GROUP BY R088014,R088001 "
'CheckOC
'With adoRecordset
'    .CursorLocation = adUseClient
'    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'    If .RecordCount <> 0 And .RecordCount > 0 Then
'        .MoveFirst
'        Do While .EOF = False
'            For i = 0 To 2
'                strTemp(i) = CheckStr(.Fields(i))
'            Next i
'            strSql = "insert into r020417_21 (r089001,r089002,r089006,id) values('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp(1)) & ",'" & ChgSQL(strTemp(2)) & "','" & strUserNum & "') "
'            cnnConnection.Execute strSql
'            .MoveNext
'        Loop
'    End If
'End With
'CheckOC
''預估準確率,勝訴率1,勝訴率2
'strSql = "select st01,pe17,pe18,pe19,st01 from performance,staff where pe01=st01(+) and pe03=" & (Val(txt1(1)) + 1911) * 100 + Val(txt1(2)) & " and pe02='T' "
'CheckOC
'With adoRecordset
'    .CursorLocation = adUseClient
'    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'    If .RecordCount <> 0 And .RecordCount > 0 Then
'        .MoveFirst
'        Do While .EOF = False
'            For i = 0 To 4
'                strTemp(i) = CheckStr(.Fields(i))
'            Next i
'            strSql = "insert into r020417_21 (r089001,r089003,r089004,r089005,r089006,id) values('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp(1)) & "," & Val(strTemp(2)) & "," & Val(strTemp(3)) & ",'" & ChgSQL(strTemp(4)) & "','" & strUserNum & "') "
'            cnnConnection.Execute strSql
'            .MoveNext
'        Loop
'    End If
'End With
'CheckOC
''frm100.Tag = "5=4"
''frm100.StrMenu
'DoEvents
''加入沒有案件性質的資料
'strSql = "SELECT DISTINCT R086001,R086014 FROM R020417_1 WHERE ID='" & strUserNum & "' "
'strSql = strSql + " union all select DISTINCT R088001,R088014 FROM R020417_2 WHERE ID='" & strUserNum & "' "
''StrSQL = StrSQL + " union all select DISTINCT R087001,R087006 FROM R020417_11 WHERE ID='" & strUserNum & "' "
''StrSQL = StrSQL + " union all select DISTINCT R089001,R089006 FROM R020417_21 WHERE ID='" & strUserNum & "' "
'CheckOC
'With adoRecordset
'    .CursorLocation = adUseClient
'    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'    If .RecordCount <> 0 And .RecordCount > 0 Then
'        .MoveFirst
'        Do While .EOF = False
'            cnnConnection.Execute "INSERT INTO R020417_1 VALUES ('" & ChgSQL(CheckStr(.Fields(0))) & "',0,'商申',0,0,0,0,0,0,0,0,0,0,'" & ChgSQL(CheckStr(.Fields(1))) & "','',1,'" & strUserNum & "') "
'            cnnConnection.Execute "INSERT INTO R020417_1 VALUES ('" & ChgSQL(CheckStr(.Fields(0))) & "',0,'商爭',0,0,0,0,0,0,0,0,0,0,'" & ChgSQL(CheckStr(.Fields(1))) & "','',2,'" & strUserNum & "') "
'            cnnConnection.Execute "INSERT INTO R020417_1 VALUES ('" & ChgSQL(CheckStr(.Fields(0))) & "',0,'國內查名',0,0,0,0,0,0,0,0,0,0,'" & ChgSQL(CheckStr(.Fields(1))) & "','',3,'" & strUserNum & "') "
'            cnnConnection.Execute "INSERT INTO R020417_1 VALUES ('" & ChgSQL(CheckStr(.Fields(0))) & "',0,'大陸',0,0,0,0,0,0,0,0,0,0,'" & ChgSQL(CheckStr(.Fields(1))) & "','',4,'" & strUserNum & "') "
'            cnnConnection.Execute "INSERT INTO R020417_1 VALUES ('" & ChgSQL(CheckStr(.Fields(0))) & "',0,'馬德里',0,0,0,0,0,0,0,0,0,0,'" & ChgSQL(CheckStr(.Fields(1))) & "','',5,'" & strUserNum & "') "
'            cnnConnection.Execute "INSERT INTO R020417_1 VALUES ('" & ChgSQL(CheckStr(.Fields(0))) & "',0,'條碼',0,0,0,0,0,0,0,0,0,0,'" & ChgSQL(CheckStr(.Fields(1))) & "','',6,'" & strUserNum & "') "
'            cnnConnection.Execute "INSERT INTO R020417_1 VALUES ('" & ChgSQL(CheckStr(.Fields(0))) & "',0,'監視系統',0,0,0,0,0,0,0,0,0,0,'" & ChgSQL(CheckStr(.Fields(1))) & "','',7,'" & strUserNum & "') "
'            cnnConnection.Execute "INSERT INTO R020417_1 VALUES ('" & ChgSQL(CheckStr(.Fields(0))) & "',0,'其他',0,0,0,0,0,0,0,0,0,0,'" & ChgSQL(CheckStr(.Fields(1))) & "','',8,'" & strUserNum & "') "
'
'            cnnConnection.Execute "INSERT INTO R020417_2 VALUES ('" & ChgSQL(CheckStr(.Fields(0))) & "',0,'商申',0,0,0,0,0,0,0,0,0,0,'" & ChgSQL(CheckStr(.Fields(1))) & "','',1,'" & strUserNum & "') "
'            cnnConnection.Execute "INSERT INTO R020417_2 VALUES ('" & ChgSQL(CheckStr(.Fields(0))) & "',0,'商爭',0,0,0,0,0,0,0,0,0,0,'" & ChgSQL(CheckStr(.Fields(1))) & "','',2,'" & strUserNum & "') "
'            cnnConnection.Execute "INSERT INTO R020417_2 VALUES ('" & ChgSQL(CheckStr(.Fields(0))) & "',0,'國內查名',0,0,0,0,0,0,0,0,0,0,'" & ChgSQL(CheckStr(.Fields(1))) & "','',3,'" & strUserNum & "') "
'            cnnConnection.Execute "INSERT INTO R020417_2 VALUES ('" & ChgSQL(CheckStr(.Fields(0))) & "',0,'大陸',0,0,0,0,0,0,0,0,0,0,'" & ChgSQL(CheckStr(.Fields(1))) & "','',4,'" & strUserNum & "') "
'            cnnConnection.Execute "INSERT INTO R020417_2 VALUES ('" & ChgSQL(CheckStr(.Fields(0))) & "',0,'馬德里',0,0,0,0,0,0,0,0,0,0,'" & ChgSQL(CheckStr(.Fields(1))) & "','',5,'" & strUserNum & "') "
'            cnnConnection.Execute "INSERT INTO R020417_2 VALUES ('" & ChgSQL(CheckStr(.Fields(0))) & "',0,'條碼',0,0,0,0,0,0,0,0,0,0,'" & ChgSQL(CheckStr(.Fields(1))) & "','',6,'" & strUserNum & "') "
'            cnnConnection.Execute "INSERT INTO R020417_2 VALUES ('" & ChgSQL(CheckStr(.Fields(0))) & "',0,'監視系統',0,0,0,0,0,0,0,0,0,0,'" & ChgSQL(CheckStr(.Fields(1))) & "','',7,'" & strUserNum & "') "
'            cnnConnection.Execute "INSERT INTO R020417_2 VALUES ('" & ChgSQL(CheckStr(.Fields(0))) & "',0,'其他',0,0,0,0,0,0,0,0,0,0,'" & ChgSQL(CheckStr(.Fields(1))) & "','',8,'" & strUserNum & "') "
'
'            cnnConnection.Execute "INSERT INTO R020417_11 VALUES ('" & ChgSQL(CheckStr(.Fields(0))) & "',0,0,0,'','" & ChgSQL(CheckStr(.Fields(1))) & "','','" & strUserNum & "') "
'
'            cnnConnection.Execute "INSERT INTO R020417_21 VALUES ('" & ChgSQL(CheckStr(.Fields(0))) & "',0,0,0,0,'" & ChgSQL(CheckStr(.Fields(1))) & "','','" & strUserNum & "') "
'
'            .MoveNext
'        Loop
'    End If
'End With
'CheckOC
''第三張
'strSql = " select max(r086002),r086003,sum(r086004),sum(r086005),sum(r086006),sum(r086007),sum(r086008),sum(r086009),sum(r086010),sum(r086011),sum(r086012),sum(r086013),r086016 from r020417_1 where id='" & strUserNum & "' group by r086016,r086003 "
'strSql = strSql + "union all select max(r088002),r088003,sum(r088004),sum(r088005),sum(r088006),sum(r088007),sum(r088008),sum(r088009),sum(r088010),sum(r088011),sum(r088012),sum(r088013),r088016 from r020417_2 where id='" & strUserNum & "' group by r088016,r088003 "
'CheckOC
'With adoRecordset
'    .CursorLocation = adUseClient
'    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'    If .RecordCount <> 0 And .RecordCount > 0 Then
'        .MoveFirst
'        Do While .EOF = False
'            For i = 0 To 12
'                strTemp(i) = CheckStr(.Fields(i))
'            Next i
'            strSql = "insert into r020417_3 values(" & Val(strTemp(0)) & ",'" & ChgSQL(strTemp(1)) & "'," & Val(strTemp(2)) & "," & Val(strTemp(3)) & "," & Val(strTemp(4)) & "," & Val(strTemp(5)) & "," & Val(strTemp(6)) & "," & Val(strTemp(7)) & "," & Val(strTemp(8)) & "," & Val(strTemp(9)) & "," & Val(strTemp(10)) & "," & Val(strTemp(11)) & "," & Val(strTemp(12)) & ",'" & strUserNum & "') "
'            cnnConnection.Execute strSql
'            .MoveNext
'        Loop
'    End If
'End With
'CheckOC
''第四張
'strSql = "select max(r086002),r086001,sum(r086013),0,0,0,0,0,0,0,sum(r086013),1,id from r020417_1 where id='" & strUserNum & "' and r086016=1 group by id,r086001 "
'strSql = strSql + " union all select max(r086002),r086001,0,sum(r086013),0,0,0,0,0,0,sum(r086013),2,id from r020417_1 where id='" & strUserNum & "' and r086016=2 group by id,r086001 "
'strSql = strSql + " union all select max(r086002),r086001,0,0,sum(r086013),0,0,0,0,0,sum(r086013),3,id from r020417_1 where id='" & strUserNum & "' and r086016=3 group by id,r086001 "
'strSql = strSql + " union all select max(r086002),r086001,0,0,0,sum(r086013),0,0,0,0,sum(r086013),4,id from r020417_1 where id='" & strUserNum & "' and r086016=4 group by id,r086001 "
'strSql = strSql + " union all select max(r086002),r086001,0,0,0,0,sum(r086013),0,0,0,sum(r086013),5,id from r020417_1 where id='" & strUserNum & "' and r086016=5 group by id,r086001 "
'strSql = strSql + " union all select max(r086002),r086001,0,0,0,0,0,sum(r086013),0,0,sum(r086013),6,id from r020417_1 where id='" & strUserNum & "' and r086016=6 group by id,r086001 "
'strSql = strSql + " union all select max(r086002),r086001,0,0,0,0,0,0,sum(r086013),0,sum(r086013),7,id from r020417_1 where id='" & strUserNum & "' and r086016=7 group by id,r086001 "
'strSql = strSql + " union all select max(r086002),r086001,0,0,0,0,0,0,0,sum(r086013),sum(r086013),8,id from r020417_1 where id='" & strUserNum & "' and r086016=8 group by id,r086001 "
'strSql = strSql + " union all select max(r088002),r088001,sum(r088013),0,0,0,0,0,0,0,sum(r088013),1,id from r020417_2 where id='" & strUserNum & "' and r088016=1 group by id,r088001 "
'strSql = strSql + " union all select max(r088002),r088001,0,sum(r088013),0,0,0,0,0,0,sum(r088013),2,id from r020417_2 where id='" & strUserNum & "' and r088016=2 group by id,r088001 "
'strSql = strSql + " union all select max(r088002),r088001,0,0,sum(r088013),0,0,0,0,0,sum(r088013),3,id from r020417_2 where id='" & strUserNum & "' and r088016=3 group by id,r088001 "
'strSql = strSql + " union all select max(r088002),r088001,0,0,0,sum(r088013),0,0,0,0,sum(r088013),4,id from r020417_2 where id='" & strUserNum & "' and r088016=4 group by id,r088001 "
'strSql = strSql + " union all select max(r088002),r088001,0,0,0,0,sum(r088013),0,0,0,sum(r088013),5,id from r020417_2 where id='" & strUserNum & "' and r088016=5 group by id,r088001 "
'strSql = strSql + " union all select max(r088002),r088001,0,0,0,0,0,sum(r088013),0,0,sum(r088013),6,id from r020417_2 where id='" & strUserNum & "' and r088016=6 group by id,r088001 "
'strSql = strSql + " union all select max(r088002),r088001,0,0,0,0,0,0,sum(r088013),0,sum(r088013),7,id from r020417_2 where id='" & strUserNum & "' and r088016=7 group by id,r088001 "
'strSql = strSql + " union all select max(r088002),r088001,0,0,0,0,0,0,0,sum(r088013),sum(r088013),8,id from r020417_2 where id='" & strUserNum & "' and r088016=8 group by id,r088001 "
'strSql = "insert into r020417_4 " & strSql
'cnnConnection.Execute strSql
''frm100.Tag = "5=5"
''frm100.StrMenu
'DoEvents
''Printer.Orientation = 1 'Removed by Morgan 2015/6/3
'DoEvents
'PrintData1
''Printer.Orientation = 1 'Removed by Morgan 2015/6/3
'DoEvents
'PrintData2
''Printer.Orientation = 1 'Removed by Morgan 2015/6/3
'DoEvents
'PrintData3
''Printer.Orientation = 1 'Removed by Morgan 2015/6/3
'DoEvents
'PrintData4
'ShowPrintOk
'Screen.MousePointer = vbDefault
'End Sub

Sub PrintData1()     '第一張
'Add By Cheng 2002/05/02
Dim Rs As New ADODB.Recordset
Dim strSQLc As String

'Modify By Cheng 2002/04/10
'strSQL = "SELECT st02,MAX(R086002),R086003,SUM(R086004),SUM(R086005),SUM(R086006),SUM(R086007),SUM(R086008),SUM(R086009),SUM(R086010),SUM(R086011),SUM(R086012),SUM(R086013),R086014,R086016,R086001,st05 FROM R020417_1,staff WHERE r086001=st01(+) and st05 in('95','15') and ID='" & strUserNum & "' GROUP BY R086014,R086001,R086016,R086003,st02,st05 ORDER BY st05,R086001,R086016 "
strSql = "SELECT st02,MAX(R086002),R086003,SUM(R086004),SUM(R086005),SUM(R086006),SUM(R086007),SUM(R086008),SUM(R086009),SUM(R086010),SUM(R086011),SUM(R086012),SUM(R086013),R086014,R086016,R086001,st05 FROM R020417_1,staff WHERE r086001=st01(+) and st05='95' and ID='" & strUserNum & "' GROUP BY R086014,R086001,R086016,R086003,st02,st05 ORDER BY st05,R086001,R086016 "
CheckOC
calToTal = ""
Page = 1
ChangeNewPage = True
BolChangePage = True
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        SavDay1 = CheckStr(.Fields(15))
        SavDay2 = CheckStr(.Fields(1))
      'Add By Cheng 2002/05/02
        '取得承辦人目標數
        strSQLc = "SELECT MAX(R086002) FROM R020417_1 WHERE r086001='" & SavDay1 & "' And ID='" & strUserNum & "'"
        If Rs.State <> adStateClosed Then Rs.Close
        Set Rs = Nothing
        Rs.CursorLocation = adUseClient
        Rs.Open strSQLc, cnnConnection, adOpenStatic, adLockReadOnly
        If Rs.RecordCount > 0 Then
           SavDay2 = CheckStr(Rs.Fields(0))
        End If
        If Rs.State <> adStateClosed Then Rs.Close
        Set Rs = Nothing
        
        SavDay3 = CheckStr(.Fields(13))
        calToTal = str(Val(calToTal) + Val(SavDay2))
        PrintTitle1
        Do While .EOF = False
            For i = 0 To 12
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            '若承辦人不同
            If SavDay3 <> CheckStr(.Fields(13)) Then
                ShowLine
                PrintEnd1
                ShowLine
                Page = Page + 1
                If ChangeNewPage = False Then
                    Printer.NewPage
                End If
                SavDay1 = CheckStr(.Fields(15))
                SavDay2 = strTemp(1)
               'Add By Cheng 2002/05/02
                 strSQLc = "SELECT MAX(R086002) FROM R020417_1 WHERE r086001='" & SavDay1 & "' And ID='" & strUserNum & "'"
                 If Rs.State <> adStateClosed Then Rs.Close
                 Set Rs = Nothing
                 Rs.CursorLocation = adUseClient
                 Rs.Open strSQLc, cnnConnection, adOpenStatic, adLockReadOnly
                 If Rs.RecordCount > 0 Then
                    SavDay2 = CheckStr(Rs.Fields(0))
                 End If
                 If Rs.State <> adStateClosed Then Rs.Close
                 Set Rs = Nothing
                
                SavDay3 = CheckStr(.Fields(13))
               If ChangeNewPage = True Then
                  ChangeNewPage = False
               Else
                  ChangeNewPage = True
               End If
                PrintTitle1
            End If
            PrintDatil1
            .MoveNext
        Loop
    End If
End With
ShowLine
PrintEnd1
ShowLine
'Modified by Morgan 2015/6/3
'Printer.EndDoc
Printer.NewPage
'end 2015/6/3
End Sub

Sub PrintTitle1()
GetPleft1
If ChangeNewPage = True Then
iPrint = 0
'Printer.Orientation = 1 'Removed by Morgan 2015/6/3
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 6000
Printer.CurrentY = iPrint
Printer.Print GetTitleNick & "商標承辦人績效表(1)"
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 500
Printer.CurrentX = 7000
Printer.CurrentY = iPrint
'Modify By Sindy 2019/12/9
'Printer.Print "發文年月：" & txt1(1) & "/" & txt1(2)
Printer.Print "發文年月：" & IIf(Len(txt1(1)) = 4, Left(txt1(1), 2), Left(txt1(1), 3)) & _
                           "/" & Right(txt1(1), 2) & " ~ " & _
                           IIf(Len(txt1(2)) = 4, Left(txt1(2), 2), Left(txt1(2), 3)) & _
                           "/" & Right(txt1(2), 2)
'2019/12/9 END
'Printer.CurrentX = 9200
'Printer.CurrentY = iPrint
'Printer.Print "－" & txt1(1) & "/" & txt1(2)
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 16500
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
iPrint = iPrint + 300
End If
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "承辦人：" & GetPrjSalesNM(SavDay1) & "    (商申人員)"
Printer.CurrentX = 5000
Printer.CurrentY = iPrint
Printer.Print "目標：" & SavDay2
Printer.CurrentX = 16500
Printer.CurrentY = iPrint
Printer.Print "頁    次：" & str(Page)
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(19000, iPrint + 150)
iPrint = iPrint + 300
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "案件性質"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "北一"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "北二"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "北三"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "北四"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "北五"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print "中所"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
Printer.Print "南所"
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iPrint
Printer.Print "高所"
Printer.CurrentX = PLeft(11)
Printer.CurrentY = iPrint
Printer.Print "其他"
Printer.CurrentX = PLeft(12)
Printer.CurrentY = iPrint
Printer.Print "全所總計"
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(19000, iPrint + 150)
iPrint = iPrint + 300
End Sub

Sub PrintDatil1()
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print strTemp(2)
For i = 3 To 12
    Printer.CurrentX = PLeft(i) + 500 - Printer.TextWidth(Format(strTemp(i), "###,###,##0.00"))
    Printer.CurrentY = iPrint
    Printer.Print Format(strTemp(i), "###,###,##0.00")
Next i
iPrint = iPrint + 300
End Sub

Sub GetPleft1()
Erase PLeft
PLeft(0) = 0
PLeft(1) = 0
PLeft(2) = 0
For i = 3 To 12
    PLeft(i) = 2500 + ((i - 3) * 1650)
Next i
End Sub

Sub PrintEnd1()
strSql = "SELECT '總計',SUM(R086004),SUM(R086005),SUM(R086006),SUM(R086007),SUM(R086008),SUM(R086009),SUM(R086010),SUM(R086011),SUM(R086012),SUM(R086013) FROM R020417_1 WHERE ID='" & strUserNum & "' " & IIf(Len(SavDay3) = 0, " AND R086014 IS NULL ", " AND R086014='" & SavDay3 & "' ")
CheckOC2
With adoRecordset1
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        Do While .EOF = False
            For i = 0 To 10
                StrTemp7(i) = CheckStr(.Fields(i))
            Next i
            Printer.CurrentX = PLeft(2)
            Printer.CurrentY = iPrint
            Printer.Print StrTemp7(0)
            For i = 1 To 10
                Printer.CurrentX = PLeft(i + 2) + 500 - Printer.TextWidth(Format(StrTemp7(i), "###,###,##0.00"))
                Printer.CurrentY = iPrint
                Printer.Print Format(StrTemp7(i), "###,###,##0.00")
            Next i
            iPrint = iPrint + 300
            .MoveNext
        Loop
    End If
End With
ShowLine
CheckOC2
strSql = "SELECT SUM(R087002),SUM(R087003),SUm(R087004),MAX(nvl(R087005,0)) FROM R020417_11 WHERE ID='" & strUserNum & "' " & IIf(Len(SavDay3) = 0, " AND R087006 IS NULL ", " AND R087006='" & SavDay3 & "' ")
CheckOC2
With adoRecordset1
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        Do While .EOF = False
            Printer.CurrentX = 0
            Printer.CurrentY = iPrint
            Printer.Print "達成率     " & Format(CheckStr(.Fields(0)), "##0.00") & "%"
            Printer.CurrentX = 5000
            Printer.CurrentY = iPrint
            Printer.Print "過期       " & CheckStr(.Fields(1))
            Printer.CurrentX = 10000
            Printer.CurrentY = iPrint
            Printer.Print "未輸入     " & CheckStr(.Fields(2))
            iPrint = iPrint + 300
            ShowLine
            Printer.CurrentX = 0
            Printer.CurrentY = iPrint
            Printer.Print "查名失誤   " & CheckStr(.Fields(3))
            iPrint = iPrint + 300
            .MoveNext
        Loop
    End If
End With
CheckOC2
iPrint = iPrint + 900
End Sub

Sub PrintData2()     '第二張
'Add By Cheng 2002/05/02
Dim Rs As New ADODB.Recordset
Dim strSQLc As String

'Modify By Cheng 2002/04/10
'strSQL = "SELECT r088001,MAX(r088002),r088003,SUM(r088004),SUM(r088005),SUM(r088006),SUM(r088007),SUM(r088008),SUM(r088009),SUM(r088010),SUM(r088011),SUM(r088012),SUM(r088013),r088014,r088016,st05 FROM R020417_2,staff WHERE r088001 = st01(+) and st05 not in ('95','15') and ID='" & strUserNum & "' GROUP BY st05,r088014,r088001,r088016,r088003 ORDER BY st05,r088001,r088016 "
strSql = "SELECT r088001,MAX(r088002),r088003,SUM(r088004),SUM(r088005),SUM(r088006),SUM(r088007),SUM(r088008),SUM(r088009),SUM(r088010),SUM(r088011),SUM(r088012),SUM(r088013),r088014,r088016,st05 FROM R020417_2,staff WHERE r088001 = st01(+) and st05<>'95' and ID='" & strUserNum & "' GROUP BY st05,r088014,r088001,r088016,r088003 ORDER BY st05,r088001,r088016 "
CheckOC
Page = 1
ChangeNewPage = True
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        SavDay1 = CheckStr(.Fields(0))
        SavDay2 = CheckStr(.Fields(1))
      
      'Add By Cheng 2002/05/02
        strSQLc = "SELECT MAX(r088002) FROM R020417_2 WHERE r088001='" & SavDay1 & "' And ID='" & strUserNum & "'"
        If Rs.State <> adStateClosed Then Rs.Close
        Set Rs = Nothing
        Rs.CursorLocation = adUseClient
        Rs.Open strSQLc, cnnConnection, adOpenStatic, adLockReadOnly
        If Rs.RecordCount > 0 Then
           SavDay2 = CheckStr(Rs.Fields(0))
        End If
        If Rs.State <> adStateClosed Then Rs.Close
        Set Rs = Nothing
        
        SavDay3 = CheckStr(.Fields(13))
        calToTal = str(Val(calToTal) + Val(SavDay2))
        PrintTitle2
        Do While .EOF = False
            For i = 0 To 12
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            If SavDay3 <> CheckStr(.Fields(13)) Then
                ShowLine
                PrintEnd2
                ShowLine
                Page = Page + 1
                If ChangeNewPage = False Then
                  Printer.NewPage
                End If
                SavDay1 = strTemp(0)
                SavDay2 = strTemp(1)
               'Add By Cheng 2002/05/02
                 strSQLc = "SELECT MAX(r088002) FROM R020417_2 WHERE r088001='" & SavDay1 & "' And ID='" & strUserNum & "'"
                 If Rs.State <> adStateClosed Then Rs.Close
                 Set Rs = Nothing
                 Rs.CursorLocation = adUseClient
                 Rs.Open strSQLc, cnnConnection, adOpenStatic, adLockReadOnly
                 If Rs.RecordCount > 0 Then
                    SavDay2 = CheckStr(Rs.Fields(0))
                 End If
                 If Rs.State <> adStateClosed Then Rs.Close
                 Set Rs = Nothing
                
                SavDay3 = CheckStr(.Fields(13))
               If ChangeNewPage = True Then
                  ChangeNewPage = False
               Else
                  ChangeNewPage = True
               End If
                PrintTitle2
            End If
            PrintDatil2
            .MoveNext
        Loop
    End If
End With
ShowLine
PrintEnd2
ShowLine
'Modified by Morgan 2015/6/3
'Printer.EndDoc
Printer.NewPage
'end 2015/6/3
End Sub

Sub PrintTitle2()
GetPleft2
If ChangeNewPage = True Then
iPrint = 0
'Printer.Orientation = 1 'Removed by Morgan 2015/6/3
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 6000
Printer.CurrentY = iPrint
Printer.Print GetTitleNick & "商標承辦人績效表(2)"
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 500
Printer.CurrentX = 7000
Printer.CurrentY = iPrint
'Modify By Sindy 2019/12/9
'Printer.Print "發文年月：" & txt1(1) & "/" & txt1(2) '& "－" & txt1(1) & "/" & txt1(2)
Printer.Print "發文年月：" & IIf(Len(txt1(1)) = 4, Left(txt1(1), 2), Left(txt1(1), 3)) & _
                           "/" & Right(txt1(1), 2) & " ~ " & _
                           IIf(Len(txt1(2)) = 4, Left(txt1(2), 2), Left(txt1(2), 3)) & _
                           "/" & Right(txt1(2), 2)
'2019/12/9 END
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 16500
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
iPrint = iPrint + 300
End If
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "承辦人：" & GetPrjSalesNM(SavDay1) & "    (商爭人員)"
Printer.CurrentX = 5000
Printer.CurrentY = iPrint
Printer.Print "目標：" & SavDay2
Printer.CurrentX = 16500
Printer.CurrentY = iPrint
Printer.Print "頁    次：" & str(Page)
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(19000, iPrint + 150)
iPrint = iPrint + 300
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "案件性質"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "北一"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "北二"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "北三"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "北四"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "北五"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print "中所"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
Printer.Print "南所"
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iPrint
Printer.Print "高所"
Printer.CurrentX = PLeft(11)
Printer.CurrentY = iPrint
Printer.Print "其他"
Printer.CurrentX = PLeft(12)
Printer.CurrentY = iPrint
Printer.Print "全所總計"
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(19000, iPrint + 150)
iPrint = iPrint + 300
End Sub

Sub PrintDatil2()
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print strTemp(2)
For i = 3 To 12
    Printer.CurrentX = PLeft(i) + 500 - Printer.TextWidth(Format(strTemp(i), "###,###,##0.00"))
    Printer.CurrentY = iPrint
    Printer.Print Format(strTemp(i), "###,###,##0.00")
Next i
iPrint = iPrint + 300
End Sub

Sub GetPleft2()
Erase PLeft
PLeft(0) = 0
PLeft(1) = 0
PLeft(2) = 0
For i = 3 To 12
    PLeft(i) = 2500 + ((i - 3) * 1650)
Next i
End Sub

Sub PrintEnd2()
strSql = "SELECT '總計',SUM(r088004),SUM(r088005),SUM(r088006),SUM(r088007),SUM(r088008),SUM(r088009),SUM(r088010),SUM(r088011),SUM(r088012),SUM(r088013) FROM R020417_2 WHERE ID='" & strUserNum & "' " & IIf(Len(SavDay3) = 0, " AND R088014 IS NULL ", " AND r088014='" & SavDay3 & "' ")
CheckOC2
With adoRecordset1
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        Do While .EOF = False
            For i = 0 To 10
                StrTemp7(i) = CheckStr(.Fields(i))
            Next i
            Printer.CurrentX = PLeft(2)
            Printer.CurrentY = iPrint
            Printer.Print StrTemp7(0)
            For i = 1 To 10
                Printer.CurrentX = PLeft(i + 2) + 500 - Printer.TextWidth(Format(StrTemp7(i), "###,###,##0.00"))
                Printer.CurrentY = iPrint
                Printer.Print Format(StrTemp7(i), "###,###,##0.00")
            Next i
            iPrint = iPrint + 300
            .MoveNext
        Loop
    End If
End With
ShowLine
CheckOC2
strSql = "SELECT SUM(R089002),SUM(R089003),SUm(R089004),SUM(R089005) FROM R020417_21 WHERE ID='" & strUserNum & "' " & IIf(Len(SavDay3) = 0, " AND R089006 IS NULL ", " AND R089006='" & SavDay3 & "' ")
CheckOC2
With adoRecordset1
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        Do While .EOF = False
            Printer.CurrentX = 0
            Printer.CurrentY = iPrint
            Printer.Print "達成率     " & Format(CheckStr(.Fields(0)), "##0.00") & "%"
            Printer.CurrentX = 5000
            Printer.CurrentY = iPrint
            '2012/4/6 MODIFY BY SONIA 葉經理說預估準確率及二個勝訴率只做台灣案故加註,並修改frm020408,frm020409
            Printer.Print "台灣案預估準確率 " & ChgSQL(CheckStr(.Fields(1))) & "%"
            Printer.CurrentX = 10000
            Printer.CurrentY = iPrint
            'Modify By Cheng 2003/04/01
'            Printer.Print "勝訴率(不含未答)" & ChgSQL(CheckStr(.Fields(2))) & "%"
            Printer.Print "台灣案勝訴率(含未答)" & ChgSQL(CheckStr(.Fields(2))) & "%"
            Printer.CurrentX = 15000
            Printer.CurrentY = iPrint
            'Modify By Cheng 2003/03/
'            Printer.Print "勝訴率(含未答)" & CheckStr(.Fields(3)) & "%"
            Printer.Print "台灣案勝訴率(不含未答)" & CheckStr(.Fields(3)) & "%"
            iPrint = iPrint + 300
            .MoveNext
        Loop
    End If
End With
CheckOC2
iPrint = iPrint + 900
End Sub

Sub PrintData3()     '第三張
strSql = "SELECT SUM(R090001),R090002,SUM(R090003),SUM(R090004),SUM(R090005),SUM(R090006),SUM(R090007),SUM(R090008),SUM(R090009),SUM(R090010),SUM(R090011),SUM(R090012),R090013 FROM R020417_3 WHERE ID='" & strUserNum & "'  GROUP BY R090013,R090002 ORDER BY R090013,R090002 "
CheckOC
Page = 1
SavDay4 = ""
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        SavDay1 = ""
        
         'Modify By Cheng 2002/04/10
'        Do While .EOF = False
'            SavDay1 = str(Val(CheckStr(.Fields(0))) + Val(SavDay1))
'            .MoveNext
'        Loop
        'Modify By Sindy 2019/12/9
        'SavDay1 = GetPerformanceByNickPE06_1(Val(str(Val(txt1(1)) + 1911) & Format(txt1(2), "00")), Val(str(Val(txt1(1)) + 1911) & Format(txt1(2), "00")))
        SavDay1 = GetPerformanceByNickPE06_1(Val(txt1(1)) + 191100, Val(txt1(2)) + 191100)
        '2019/12/9 END
        .MoveFirst
        PrintTitle3
        Do While .EOF = False
            For i = 0 To 11
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            SavDay4 = strTemp(0)
            PrintDatil3
            .MoveNext
        Loop
    End If
End With
ShowLine
PrintEnd3
ShowLine
CheckOC
'Modified by Morgan 2015/6/3
'Printer.EndDoc
Printer.NewPage
'end 2015/6/3
End Sub

Sub PrintTitle3()
GetPleft3
iPrint = 0
'Printer.Orientation = 1 'Removed by Morgan 2015/6/3
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 6500
Printer.CurrentY = iPrint
Printer.Print GetTitleNick & "商標績效表(1)"
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 500
Printer.CurrentX = 7000
Printer.CurrentY = iPrint
'Modify By Sindy 2019/12/9
'Printer.Print "發文年月：" & txt1(1) & "/" & txt1(2) '& "－" & txt1(1) & "/" & txt1(2)
Printer.Print "發文年月：" & IIf(Len(txt1(1)) = 4, Left(txt1(1), 2), Left(txt1(1), 3)) & _
                           "/" & Right(txt1(1), 2) & " ~ " & _
                           IIf(Len(txt1(2)) = 4, Left(txt1(2), 2), Left(txt1(2), 3)) & _
                           "/" & Right(txt1(2), 2)
'2019/12/9 END
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 16500
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "總目標：" & SavDay1
'Printer.Print "總目標：" & calToTal
Printer.CurrentX = 16500
Printer.CurrentY = iPrint
Printer.Print "頁    次：" & str(Page)
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(19000, iPrint + 150)
iPrint = iPrint + 300
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "案件性質"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "北一"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "北二"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "北三"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "北四"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "北五"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "中所"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print "南所"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
Printer.Print "高所"
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iPrint
Printer.Print "其他"
Printer.CurrentX = PLeft(11)
Printer.CurrentY = iPrint
Printer.Print "全所總計"
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(19000, iPrint + 150)
iPrint = iPrint + 300
End Sub

Sub PrintDatil3()
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print strTemp(1)
For i = 2 To 11
    Printer.CurrentX = PLeft(i) + 500 - Printer.TextWidth(Format(strTemp(i), "###,###,##0.00"))
    Printer.CurrentY = iPrint
    Printer.Print Format(strTemp(i), "###,###,##0.00")
Next i
iPrint = iPrint + 300
End Sub

Sub GetPleft3()
Erase PLeft
PLeft(0) = 0
PLeft(1) = 0
For i = 2 To 11
    PLeft(i) = 2500 + ((i - 2) * 1650)
Next i
End Sub

Sub PrintEnd3()
strSql = "SELECT '總計',SUM(R090003),SUM(R090004),SUM(R090005),SUM(R090006),SUM(R090007),SUM(R090008),SUM(R090009),SUM(R090010),SUM(R090011),SUM(R090012) FROM R020417_3 WHERE ID='" & strUserNum & "' "
CheckOC2
SavDay2 = ""
With adoRecordset1
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        Do While .EOF = False
            For i = 0 To 10
                StrTemp7(i) = CheckStr(.Fields(i))
            Next i
            Printer.CurrentX = PLeft(1)
            Printer.CurrentY = iPrint
            Printer.Print StrTemp7(0)
            For i = 1 To 10
                Printer.CurrentX = PLeft(i + 1) + 500 - Printer.TextWidth(Format(StrTemp7(i), "###,###,##0.00"))
                Printer.CurrentY = iPrint
                Printer.Print Format(StrTemp7(i), "###,###,##0.00")
            Next i
            iPrint = iPrint + 300
            ShowLine
            StrTemp7(0) = SavDay1
            Printer.CurrentX = 0
            Printer.CurrentY = iPrint
            If Val(StrTemp7(10)) = 0 Then
                Printer.Print "達成率     " & "0 %"
            Else
                Printer.Print "達成率     " & Format(Trim(str(Val(StrTemp7(10)) / Val(StrTemp7(0))) * 100), "##0.00") & " %"
            End If
            iPrint = iPrint + 300
            .MoveNext
        Loop
    End If
End With
CheckOC2
End Sub

Sub PrintData4()     '第四張
'Modified by Lydia 2023/12/01
'strSql = "SELECT '',nvl(st02,r091002),SUM(R091003),SUM(R091004),SUM(R091005),SUM(R091006),SUM(R091007),SUM(R091008),SUM(R091009),SUM(R091010),SUM(R091011),R091002 FROM staff ,R020417_4 where R091002=st01(+) and id='" & strUserNum & "' GROUP BY R091002,nvl(st02,r091002) ORDER BY R091002,nvl(st02,r091002)"
strSql = "SELECT '',nvl(st02,r091002),SUM(R091003),SUM(R091004),SUM(R091005),SUM(R091013),SUM(R091006),SUM(R091007),SUM(R091008),SUM(R091009),SUM(R091010),SUM(R091011),R091002 FROM staff ,R020417_4 where R091002=st01(+) and id='" & strUserNum & "' GROUP BY R091002,nvl(st02,r091002) ORDER BY R091002,nvl(st02,r091002)"
CheckOC
Page = 1
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        PrintTitle4
        Do While .EOF = False
            'Modified by Lydia 2023/12/01 10=>11
            For i = 0 To 11
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            PrintDatil4
            If iPrint >= 14000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle4
            End If
            .MoveNext
        Loop
    End If
End With
CheckOC
ShowLine
PrintEnd4
ShowLine
Printer.EndDoc
End Sub

Sub PrintTitle4()
GetPleft4
iPrint = 0
'Printer.Orientation = 1 'Removed by Morgan 2015/6/3
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 6500
Printer.CurrentY = iPrint
Printer.Print GetTitleNick & "商標績效表(2)"
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 500
Printer.CurrentX = 7000
Printer.CurrentY = iPrint
'Modify By Sindy 2019/12/9
'Printer.Print "發文年月：" & txt1(1) & "/" & txt1(2) '& "－" & txt1(1) & "/" & txt1(2)
Printer.Print "發文年月：" & IIf(Len(txt1(1)) = 4, Left(txt1(1), 2), Left(txt1(1), 3)) & _
                           "/" & Right(txt1(1), 2) & " ~ " & _
                           IIf(Len(txt1(2)) = 4, Left(txt1(2), 2), Left(txt1(2), 3)) & _
                           "/" & Right(txt1(2), 2)
'2019/12/9 END
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 16500
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "總目標：" & SavDay1
'Printer.Print "總目標：" & calToTal
Printer.CurrentX = 16500
Printer.CurrentY = iPrint
Printer.Print "頁    次：" & str(Page)
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(19000, iPrint + 150)
iPrint = iPrint + 300
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "承辦人"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "商申"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "商爭"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "國內查名"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
'Added by Lydia 2023/12/01
Printer.Print "查名"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
'Modified by Lydia 2023/12/01 後面Index+1
Printer.Print "大陸"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "馬德里"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print "條碼"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
Printer.Print "監視系統"
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iPrint
Printer.Print "其他"
Printer.CurrentX = PLeft(11)
'end 2023/12/01
Printer.CurrentY = iPrint
Printer.Print "全所總計"
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(19000, iPrint + 150)
iPrint = iPrint + 300
End Sub

Sub PrintDatil4()
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print strTemp(1)
'Modified by Lydia 2023/12/01 10=>11
For i = 2 To 11
    Printer.CurrentX = PLeft(i) + 500 - Printer.TextWidth(Format(strTemp(i), "###,###,##0.00"))
    Printer.CurrentY = iPrint
    Printer.Print Format(strTemp(i), "###,###,##0.00")
Next i
iPrint = iPrint + 300
End Sub

Sub GetPleft4()
Erase PLeft
PLeft(0) = 0
PLeft(1) = 0
'Modified by Lydia 2023/12/01 10=>11
For i = 2 To 11
    PLeft(i) = 2500 + ((i - 2) * 1650)
Next i
End Sub

Sub PrintEnd4()
'Modified by Lydia 2023/12/01
'strSql = "SELECT '總計',SUM(R091003),SUM(R091004),SUM(R091005),SUM(R091006),SUM(R091007),SUM(R091008),SUM(R091009),SUM(R091010),SUM(R091011) FROM R020417_4 WHERE ID='" & strUserNum & "' "
strSql = "SELECT '總計',SUM(R091003),SUM(R091004),SUM(R091005),SUM(R091013),SUM(R091006),SUM(R091007),SUM(R091008),SUM(R091009),SUM(R091010),SUM(R091011) FROM R020417_4 WHERE ID='" & strUserNum & "' "
CheckOC2
With adoRecordset1
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        Do While .EOF = False
            'Modified by Lydia 2023/12/01 9=>10
            For i = 0 To 10
                StrTemp7(i) = CheckStr(.Fields(i))
            Next i
            Printer.CurrentX = PLeft(1)
            Printer.CurrentY = iPrint
            Printer.Print StrTemp7(0)
            'Modified by Lydia 2023/12/01 9=>10
            For i = 1 To 10
                Printer.CurrentX = PLeft(i + 1) + 500 - Printer.TextWidth(Format(StrTemp7(i), "###,###,##0.00"))
                Printer.CurrentY = iPrint
                Printer.Print Format(StrTemp7(i), "###,###,##0.00")
            Next i
            iPrint = iPrint + 300
            .MoveNext
        Loop
    End If
End With
CheckOC2
End Sub

Sub ProcessOther()
Select Case Mid(UCase(strTemp(2)), 1, 2)
    Case "S1"
        Select Case UCase(strTemp(2))
        Case "S11"
            Select Case Val(strTemp(1))
            Case 95 ', 15
                'Modified by Lydia 2023/12/01 ,8,=>,9, (後面Index+1)
                cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086004,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'其他'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',9,'" & strUserNum & "') "
            Case Else
                cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088004,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'其他'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',9,'" & strUserNum & "') "
            End Select
        Case "S12"
            Select Case Val(strTemp(1))
            Case 95 ', 15
                cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086005,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'其他'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',9,'" & strUserNum & "') "
            Case Else
                cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088005,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'其他'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',9,'" & strUserNum & "') "
            End Select
        Case "S13"
            Select Case Val(strTemp(1))
            Case 95 ', 15
                cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086006,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'其他'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',9,'" & strUserNum & "') "
            Case Else
                cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088006,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'其他'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',9,'" & strUserNum & "') "
            End Select
        Case "S14"
            Select Case Val(strTemp(1))
            Case 95 ', 15
                cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086007,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'其他'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',9,'" & strUserNum & "') "
            Case Else
                cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088007,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'其他'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',9,'" & strUserNum & "') "
            End Select
        Case "S15"
            Select Case Val(strTemp(1))
            Case 95 ', 15
                cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086008,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'其他'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',9,'" & strUserNum & "') "
            Case Else
                cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088008,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'其他'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',9,'" & strUserNum & "') "
            End Select
        Case Else
        End Select
    Case "S2"
        Select Case Val(strTemp(1))
        Case 95 ', 15
            cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086009,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'其他'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',9,'" & strUserNum & "') "
        Case Else
            cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088009,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'其他'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',9,'" & strUserNum & "') "
        End Select
    Case "S3"
        Select Case Val(strTemp(1))
        Case 95 ', 15
            cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086010,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'其他'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',9,'" & strUserNum & "') "
        Case Else
            cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088010,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'其他'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',9,'" & strUserNum & "') "
        End Select
    Case "S4"
        Select Case Val(strTemp(1))
        Case 95 ', 15
            cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086011,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'其他'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',9,'" & strUserNum & "') "
        Case Else
            cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088011,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'其他'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',9,'" & strUserNum & "') "
        End Select
    Case Else
        Select Case Val(strTemp(1))
        Case 95 ', 15
            cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086012,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'其他'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',9,'" & strUserNum & "') "
        Case Else
            cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088012,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'其他'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',9,'" & strUserNum & "') "
        End Select
End Select
End Sub

Sub ProcessTM()
Select Case Mid(UCase(strTemp(2)), 1, 2)
Case "S1"
    Select Case UCase(strTemp(2))
    Case "S11"
        Select Case Val(strTemp(1))
        Case 95 ', 15
            'Modified by Lydia 2023/12/01 ,7,=>,8, (後面Index+1)
            cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086004,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'監視系統'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',8,'" & strUserNum & "') "
        Case Else
            cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088004,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'監視系統'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',8,'" & strUserNum & "') "
        End Select
    Case "S12"
        Select Case Val(strTemp(1))
        Case 95 ', 15
            cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086005,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'監視系統'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',8,'" & strUserNum & "') "
        Case Else
            cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088005,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'監視系統'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',8,'" & strUserNum & "') "
        End Select
    Case "S13"
        Select Case Val(strTemp(1))
        Case 95 ', 15
            cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086006,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'監視系統'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',8,'" & strUserNum & "') "
        Case Else
            cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088006,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'監視系統'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',8,'" & strUserNum & "') "
        End Select
    Case "S14"
        Select Case Val(strTemp(1))
        Case 95 ', 15
            cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086007,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'監視系統'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',8,'" & strUserNum & "') "
        Case Else
            cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088007,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'監視系統'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',8,'" & strUserNum & "') "
        End Select
    Case "S15"
        Select Case Val(strTemp(1))
        Case 95 ', 15
            cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086008,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'監視系統'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',8,'" & strUserNum & "') "
        Case Else
            cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088008,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'監視系統'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',8,'" & strUserNum & "') "
        End Select
    Case Else
    End Select
Case "S2"
    Select Case Val(strTemp(1))
    Case 95 ', 15
        cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086009,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'監視系統'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',8,'" & strUserNum & "') "
    Case Else
        cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088009,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'監視系統'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',8,'" & strUserNum & "') "
    End Select
Case "S3"
    Select Case Val(strTemp(1))
    Case 95 ', 15
        cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086010,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'監視系統'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',8,'" & strUserNum & "') "
    Case Else
        cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088010,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'監視系統'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',8,'" & strUserNum & "') "
    End Select
Case "S4"
    Select Case Val(strTemp(1))
    Case 95 ', 15
        cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086011,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'監視系統'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',8,'" & strUserNum & "') "
    Case Else
        cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088011,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'監視系統'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',8,'" & strUserNum & "') "
    End Select
Case Else
    Select Case Val(strTemp(1))
    Case 95 ', 15
        cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086012,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'監視系統'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',8,'" & strUserNum & "') "
    Case Else
        cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088012,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'監視系統'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',8,'" & strUserNum & "') "
    End Select
End Select
End Sub

Sub ProcessTB()
Select Case Mid(UCase(strTemp(2)), 1, 2)
Case "S1"
    Select Case UCase(strTemp(2))
    Case "S11"
        Select Case Val(strTemp(1))
        Case 95 ', 15
            'Modified by Lydia 2023/12/01 ,6,=>,7, (後面Index+1)
            cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086004,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'條碼'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',7,'" & strUserNum & "') "
        Case Else
            cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088004,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'條碼'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',7,'" & strUserNum & "') "
        End Select
    Case "S12"
        Select Case Val(strTemp(1))
        Case 95 ', 15
            cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086005,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'條碼'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',7,'" & strUserNum & "') "
        Case Else
            cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088005,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'條碼'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',7,'" & strUserNum & "') "
        End Select
    Case "S13"
        Select Case Val(strTemp(1))
        Case 95 ', 15
            cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086006,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'條碼'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',7,'" & strUserNum & "') "
        Case Else
            cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088006,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'條碼'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',7,'" & strUserNum & "') "
        End Select
    Case "S14"
        Select Case Val(strTemp(1))
        Case 95 ', 15
            cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086007,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'條碼'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',7,'" & strUserNum & "') "
        Case Else
            cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088007,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'條碼'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',7,'" & strUserNum & "') "
        End Select
    Case "S15"
        Select Case Val(strTemp(1))
        Case 95 ', 15
            cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086008,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'條碼'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',7,'" & strUserNum & "') "
        Case Else
            cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088008,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'條碼'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',7,'" & strUserNum & "') "
        End Select
    Case Else
    End Select
Case "S2"
    Select Case Val(strTemp(1))
    Case 95 ', 15
        cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086009,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'條碼'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',7,'" & strUserNum & "') "
    Case Else
        cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088009,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'條碼'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',7,'" & strUserNum & "') "
    End Select
Case "S3"
    Select Case Val(strTemp(1))
    Case 95 ', 15
        cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086010,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'條碼'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',7,'" & strUserNum & "') "
    Case Else
        cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088010,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'條碼'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',7,'" & strUserNum & "') "
    End Select
Case "S4"
    Select Case Val(strTemp(1))
    Case 95 ', 15
        cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086011,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'條碼'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',7,'" & strUserNum & "') "
    Case Else
        cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088011,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'條碼'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',7,'" & strUserNum & "') "
    End Select
Case Else
    Select Case Val(strTemp(1))
    Case 95 ', 15
        cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086012,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'條碼'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',7,'" & strUserNum & "') "
    Case Else
        cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088012,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'條碼'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',7,'" & strUserNum & "') "
    End Select
End Select

End Sub

Sub ProcessTF()
Select Case Mid(UCase(strTemp(2)), 1, 2)
Case "S1"
    Select Case UCase(strTemp(2))
    Case "S11"
        Select Case Val(strTemp(1))
        Case 95 ', 15
            'Modified by Lydia 2023/12/01 ,5,=>,6, (後面Index+1)
            cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086004,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'馬德里'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',6,'" & strUserNum & "') "
        Case Else
            cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088004,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'馬德里'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',6,'" & strUserNum & "') "
        End Select
    Case "S12"
        Select Case Val(strTemp(1))
        Case 95 ', 15
            cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086005,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'馬德里'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',6,'" & strUserNum & "') "
        Case Else
            cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088005,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'馬德里'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',6,'" & strUserNum & "') "
        End Select
    Case "S13"
        Select Case Val(strTemp(1))
        Case 95 ', 15
            cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086006,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'馬德里'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',6,'" & strUserNum & "') "
        Case Else
            cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088006,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'馬德里'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',6,'" & strUserNum & "') "
        End Select
    Case "S14"
        Select Case Val(strTemp(1))
        Case 95 ', 15
            cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086007,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'馬德里'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',6,'" & strUserNum & "') "
        Case Else
            cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088007,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'馬德里'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',6,'" & strUserNum & "') "
        End Select
    Case "S15"
        Select Case Val(strTemp(1))
        Case 95 ', 15
            cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086008,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'馬德里'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',6,'" & strUserNum & "') "
        Case Else
            cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088008,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'馬德里'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',6,'" & strUserNum & "') "
        End Select
    Case Else
    End Select
Case "S2"
    Select Case Val(strTemp(1))
    Case 95 ', 15
        cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086009,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'馬德里'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',6,'" & strUserNum & "') "
    Case Else
        cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088009,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'馬德里'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',6,'" & strUserNum & "') "
    End Select
Case "S3"
    Select Case Val(strTemp(1))
    Case 95 ', 15
        cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086010,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'馬德里'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',6,'" & strUserNum & "') "
    Case Else
        cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088010,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'馬德里'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',6,'" & strUserNum & "') "
    End Select
Case "S4"
    Select Case Val(strTemp(1))
    Case 95 ', 15
        cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086011,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'馬德里'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',6,'" & strUserNum & "') "
    Case Else
        cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086011,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'馬德里',0,0,'" & ChgSQL(strTemp(3)) & "','*',6,'" & strUserNum & "') "
    End Select
Case Else
    Select Case Val(strTemp(1))
    Case 95 ', 15
        cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086012,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'馬德里'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',6,'" & strUserNum & "') "
    Case Else
        cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088012,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'馬德里'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',6,'" & strUserNum & "') "
    End Select
End Select

End Sub

Sub ProcessTS()
Select Case Val(strTemp(4))
Case 0     '國內查名    改用商標委查計算
Case Is > 0
    Select Case Mid(UCase(strTemp(2)), 1, 2)
    '北所
    Case "S1"
        Select Case UCase(strTemp(2))
        '北一
        Case "S11"
            Select Case Val(strTemp(1))
            '申
            Case 95 ', 15
                'Modified by Lydia 2023/12/01 ,4,=>,5, (後面Index+1)
                cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086004,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'大陸'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',5,'" & strUserNum & "') "
            '爭
            Case Else
                cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088004,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'大陸'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',5,'" & strUserNum & "') "
            End Select
        '北二
        Case "S12"
            Select Case Val(strTemp(1))
            '申
            Case 95 ', 15
                cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086005,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'大陸'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',5,'" & strUserNum & "') "
            '爭
            Case Else
                cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088005,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'大陸'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',5,'" & strUserNum & "') "
            End Select
        '北三
        Case "S13"
            Select Case Val(strTemp(1))
            '申
            Case 95 ', 15
                cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086006,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'大陸'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',5,'" & strUserNum & "') "
            '爭
            Case Else
                cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088006,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'大陸'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',5,'" & strUserNum & "') "
            End Select
        '北四
        Case "S14"
            Select Case Val(strTemp(1))
            '申
            Case 95 ', 15
                cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086007,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'大陸'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',5,'" & strUserNum & "') "
            '爭
            Case Else
                cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088007,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'大陸'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',5,'" & strUserNum & "') "
            End Select
        '北五
        Case "S15"
            Select Case Val(strTemp(1))
            '申
            Case 95 ', 15
                cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086008,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'大陸'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',5,'" & strUserNum & "') "
            '爭
            Case Else
                cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088008,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'大陸'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',5,'" & strUserNum & "') "
            End Select
        Case Else
        End Select
    '中所
    Case "S2"
        Select Case Val(strTemp(1))
        '申
        Case 95 ', 15
            cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086009,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'大陸'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',5,'" & strUserNum & "') "
        '爭
        Case Else
            cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088009,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'大陸'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',5,'" & strUserNum & "') "
        End Select
    '南所
    Case "S3"
        Select Case Val(strTemp(1))
        '申
        Case 95 ', 15
            cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086010,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'大陸'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',5,'" & strUserNum & "') "
        '爭
        Case Else
            cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088010,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'大陸'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',5,'" & strUserNum & "') "
        End Select
    '高所
    Case "S4"
        Select Case Val(strTemp(1))
        '申
        Case 95 ', 15
            cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086011,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'大陸'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',5,'" & strUserNum & "') "
        '爭
        Case Else
            cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088011,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'大陸'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',5,'" & strUserNum & "') "
        End Select
    '其他
    Case Else
        Select Case Val(strTemp(1))
        '申
        Case 95 ', 15
            cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086012,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'大陸'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',5,'" & strUserNum & "') "
        '爭
        Case Else
            cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088012,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'大陸'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',5,'" & strUserNum & "') "
        End Select
    End Select

End Select

End Sub

Sub ProcessT()
Select Case Val(strTemp(4))
    '國內
Case 0
    Select Case Mid(UCase(strTemp(2)), 1, 2)
         '北所
    Case "S1"
        Select Case UCase(strTemp(2))
        '北一
        Case "S11"
            Select Case Val(strTemp(1))
            '屬於商申人員的資料
            Case 95 ', 15
               '2011/12/7 modify by sonia 加210陳述意見書的控制同202申請意見書
               '商申案
               If Left(adoRecordset("CP10"), 1) <> "4" And Left(adoRecordset("CP10"), 1) <> "6" And adoRecordset("CP10") <> "202" And adoRecordset("CP10") <> "210" And adoRecordset("CP10") <> "204" And adoRecordset("CP10") <> "205" And adoRecordset("CP10") < "7" And adoRecordset("CP09") < "C" Then
                  cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086004,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'商申'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',1,'" & strUserNum & "') "
               '2009/7/1 modify by sonia 將其他及商爭順序顛倒,否則CP10>="7"者不會列入其他反而列入商爭,例:第一期註冊費
               '其他
               ElseIf Left(adoRecordset("CP10"), 1) >= "7" Then
                  ProcessOther
               '商爭案
               ElseIf (Left(adoRecordset("CP10"), 1) = "4" Or Left(adoRecordset("CP10"), 1) = "6" Or adoRecordset("CP10") = "202" Or adoRecordset("CP10") = "210" Or adoRecordset("CP10") = "204" Or adoRecordset("CP10") = "205" Or adoRecordset("CP10") >= "7") And adoRecordset("CP09") < "C" Then
                  cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086004,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'商爭'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',2,'" & strUserNum & "') "
               End If
            '屬於商爭人員的資料
            Case Else
               '商申案
               If Left(adoRecordset("CP10"), 1) <> "4" And Left(adoRecordset("CP10"), 1) <> "6" And adoRecordset("CP10") <> "202" And adoRecordset("CP10") <> "210" And adoRecordset("CP10") <> "204" And adoRecordset("CP10") <> "205" And adoRecordset("CP10") < "7" And adoRecordset("CP09") < "C" Then
                  cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088004,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'商申'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',1,'" & strUserNum & "') "
               '2009/7/1 modify by sonia 將其他及商爭順序顛倒,否則CP10>="7"者不會列入其他反而列入商爭,例:第一期註冊費
               '其他
               ElseIf Left(adoRecordset("CP10"), 1) >= "7" Then
                  ProcessOther
               '商爭案
               ElseIf (Left(adoRecordset("CP10"), 1) = "4" Or Left(adoRecordset("CP10"), 1) = "6" Or adoRecordset("CP10") = "202" Or adoRecordset("CP10") = "210" Or adoRecordset("CP10") = "204" Or adoRecordset("CP10") = "205" Or adoRecordset("CP10") >= "7") And adoRecordset("CP09") < "C" Then
                  cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088004,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'商爭'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',2,'" & strUserNum & "') "
               End If
            End Select
        '北二
        Case "S12"
            Select Case Val(strTemp(1))
            '申
            Case 95 ', 15
               '商申案
               If Left(adoRecordset("CP10"), 1) <> "4" And Left(adoRecordset("CP10"), 1) <> "6" And adoRecordset("CP10") <> "202" And adoRecordset("CP10") <> "210" And adoRecordset("CP10") <> "204" And adoRecordset("CP10") <> "205" And adoRecordset("CP10") < "7" And adoRecordset("CP09") < "C" Then
                  cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086005,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'商申'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',1,'" & strUserNum & "') "
               '2009/7/1 modify by sonia 將其他及商爭順序顛倒,否則CP10>="7"者不會列入其他反而列入商爭,例:第一期註冊費
               '其他
               ElseIf Left(adoRecordset("CP10"), 1) >= "7" Then
                  ProcessOther
               '商爭案
               ElseIf (Left(adoRecordset("CP10"), 1) = "4" Or Left(adoRecordset("CP10"), 1) = "6" Or adoRecordset("CP10") = "202" Or adoRecordset("CP10") = "210" Or adoRecordset("CP10") = "204" Or adoRecordset("CP10") = "205" Or adoRecordset("CP10") >= "7") And adoRecordset("CP09") < "C" Then
                  cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086005,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'商爭'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',2,'" & strUserNum & "') "
               End If
            '爭
            Case Else
               '商申案
               If Left(adoRecordset("CP10"), 1) <> "4" And Left(adoRecordset("CP10"), 1) <> "6" And adoRecordset("CP10") <> "202" And adoRecordset("CP10") <> "210" And adoRecordset("CP10") <> "204" And adoRecordset("CP10") <> "205" And adoRecordset("CP10") < "7" And adoRecordset("CP09") < "C" Then
                  cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088005,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'商申'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',1,'" & strUserNum & "') "
               '2009/7/1 modify by sonia 將其他及商爭順序顛倒,否則CP10>="7"者不會列入其他反而列入商爭,例:第一期註冊費
               '其他
               ElseIf Left(adoRecordset("CP10"), 1) >= "7" Then
                  ProcessOther
               '商爭案
               ElseIf (Left(adoRecordset("CP10"), 1) = "4" Or Left(adoRecordset("CP10"), 1) = "6" Or adoRecordset("CP10") = "202" Or adoRecordset("CP10") = "210" Or adoRecordset("CP10") = "204" Or adoRecordset("CP10") = "205" Or adoRecordset("CP10") >= "7") And adoRecordset("CP09") < "C" Then
                  cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088005,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'商爭'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',2,'" & strUserNum & "') "
               End If
            End Select
        '北三
        Case "S13"
            Select Case Val(strTemp(1))
            '申
            Case 95 ', 15
               '商申案
               If Left(adoRecordset("CP10"), 1) <> "4" And Left(adoRecordset("CP10"), 1) <> "6" And adoRecordset("CP10") <> "202" And adoRecordset("CP10") <> "210" And adoRecordset("CP10") <> "204" And adoRecordset("CP10") <> "205" And adoRecordset("CP10") < "7" And adoRecordset("CP09") < "C" Then
                  cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086006,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'商申'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',1,'" & strUserNum & "') "
               '2009/7/1 modify by sonia 將其他及商爭順序顛倒,否則CP10>="7"者不會列入其他反而列入商爭,例:第一期註冊費
               '其他
               ElseIf Left(adoRecordset("CP10"), 1) >= "7" Then
                  ProcessOther
               '商爭案
               ElseIf (Left(adoRecordset("CP10"), 1) = "4" Or Left(adoRecordset("CP10"), 1) = "6" Or adoRecordset("CP10") = "202" Or adoRecordset("CP10") = "210" Or adoRecordset("CP10") = "204" Or adoRecordset("CP10") = "205" Or adoRecordset("CP10") >= "7") And adoRecordset("CP09") < "C" Then
                  cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086006,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'商爭'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',2,'" & strUserNum & "') "
               End If
            '爭
            Case Else
               '商申案
               If Left(adoRecordset("CP10"), 1) <> "4" And Left(adoRecordset("CP10"), 1) <> "6" And adoRecordset("CP10") <> "202" And adoRecordset("CP10") <> "210" And adoRecordset("CP10") <> "204" And adoRecordset("CP10") <> "205" And adoRecordset("CP10") < "7" And adoRecordset("CP09") < "C" Then
                  cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088006,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'商申'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',1,'" & strUserNum & "') "
               '2009/7/1 modify by sonia 將其他及商爭順序顛倒,否則CP10>="7"者不會列入其他反而列入商爭,例:第一期註冊費
               '其他
               ElseIf Left(adoRecordset("CP10"), 1) >= "7" Then
                  ProcessOther
               '商爭案
               ElseIf (Left(adoRecordset("CP10"), 1) = "4" Or Left(adoRecordset("CP10"), 1) = "6" Or adoRecordset("CP10") = "202" Or adoRecordset("CP10") = "210" Or adoRecordset("CP10") = "204" Or adoRecordset("CP10") = "205" Or adoRecordset("CP10") >= "7") And adoRecordset("CP09") < "C" Then
                  cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088006,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'商爭'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',2,'" & strUserNum & "') "
               End If
            End Select
        '北四
        Case "S14"
            Select Case Val(strTemp(1))
            '申
            Case 95 ', 15
               '商申案
               If Left(adoRecordset("CP10"), 1) <> "4" And Left(adoRecordset("CP10"), 1) <> "6" And adoRecordset("CP10") <> "202" And adoRecordset("CP10") <> "210" And adoRecordset("CP10") <> "204" And adoRecordset("CP10") <> "205" And adoRecordset("CP10") < "7" And adoRecordset("CP09") < "C" Then
                  cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086007,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'商申'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',1,'" & strUserNum & "') "
               '2009/7/1 modify by sonia 將其他及商爭順序顛倒,否則CP10>="7"者不會列入其他反而列入商爭,例:第一期註冊費
               '其他
               ElseIf Left(adoRecordset("CP10"), 1) >= "7" Then
                  ProcessOther
               '商爭案
               ElseIf (Left(adoRecordset("CP10"), 1) = "4" Or Left(adoRecordset("CP10"), 1) = "6" Or adoRecordset("CP10") = "202" Or adoRecordset("CP10") = "210" Or adoRecordset("CP10") = "204" Or adoRecordset("CP10") = "205" Or adoRecordset("CP10") >= "7") And adoRecordset("CP09") < "C" Then
                  cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086007,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'商爭'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',2,'" & strUserNum & "') "
               End If
            '爭
            Case Else
               '商申案
               If Left(adoRecordset("CP10"), 1) <> "4" And Left(adoRecordset("CP10"), 1) <> "6" And adoRecordset("CP10") <> "202" And adoRecordset("CP10") <> "210" And adoRecordset("CP10") <> "204" And adoRecordset("CP10") <> "205" And adoRecordset("CP10") < "7" And adoRecordset("CP09") < "C" Then
                  cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088007,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'商申'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',1,'" & strUserNum & "') "
               '2009/7/1 modify by sonia 將其他及商爭順序顛倒,否則CP10>="7"者不會列入其他反而列入商爭,例:第一期註冊費
               '其他
               ElseIf Left(adoRecordset("CP10"), 1) >= "7" Then
                  ProcessOther
               '商爭案
               ElseIf (Left(adoRecordset("CP10"), 1) = "4" Or Left(adoRecordset("CP10"), 1) = "6" Or adoRecordset("CP10") = "202" Or adoRecordset("CP10") = "210" Or adoRecordset("CP10") = "204" Or adoRecordset("CP10") = "205" Or adoRecordset("CP10") >= "7") And adoRecordset("CP09") < "C" Then
                  cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088007,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'商爭'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',2,'" & strUserNum & "') "
               End If
            End Select
        '北五
        Case "S15"
            Select Case Val(strTemp(1))
            '申
            Case 95 ', 15
               'Modify By Cheng 2002/04/10
'                cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086007,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'商申'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',1,'" & strUserNum & "') "
               '商申案
               If Left(adoRecordset("CP10"), 1) <> "4" And Left(adoRecordset("CP10"), 1) <> "6" And adoRecordset("CP10") <> "202" And adoRecordset("CP10") <> "210" And adoRecordset("CP10") <> "204" And adoRecordset("CP10") <> "205" And adoRecordset("CP10") < "7" And adoRecordset("CP09") < "C" Then
                  cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086008,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'商申'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',1,'" & strUserNum & "') "
               '2009/7/1 modify by sonia 將其他及商爭順序顛倒,否則CP10>="7"者不會列入其他反而列入商爭,例:第一期註冊費
               '其他
               ElseIf Left(adoRecordset("CP10"), 1) >= "7" Then
                  ProcessOther
               '商爭案
               ElseIf (Left(adoRecordset("CP10"), 1) = "4" Or Left(adoRecordset("CP10"), 1) = "6" Or adoRecordset("CP10") = "202" Or adoRecordset("CP10") = "210" Or adoRecordset("CP10") = "204" Or adoRecordset("CP10") = "205" Or adoRecordset("CP10") >= "7") And adoRecordset("CP09") < "C" Then
                  cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086008,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'商爭'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',2,'" & strUserNum & "') "
               End If
            '爭
            Case Else
               '商申案
               If Left(adoRecordset("CP10"), 1) <> "4" And Left(adoRecordset("CP10"), 1) <> "6" And adoRecordset("CP10") <> "202" And adoRecordset("CP10") <> "210" And adoRecordset("CP10") <> "204" And adoRecordset("CP10") <> "205" And adoRecordset("CP10") < "7" And adoRecordset("CP09") < "C" Then
                  cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088008,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'商申'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',1,'" & strUserNum & "') "
               '2009/7/1 modify by sonia 將其他及商爭順序顛倒,否則CP10>="7"者不會列入其他反而列入商爭,例:第一期註冊費
               '其他
               ElseIf Left(adoRecordset("CP10"), 1) >= "7" Then
                  ProcessOther
               '商爭案
               ElseIf (Left(adoRecordset("CP10"), 1) = "4" Or Left(adoRecordset("CP10"), 1) = "6" Or adoRecordset("CP10") = "202" Or adoRecordset("CP10") = "210" Or adoRecordset("CP10") = "204" Or adoRecordset("CP10") = "205" Or adoRecordset("CP10") >= "7") And adoRecordset("CP09") < "C" Then
                  cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088008,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'商爭'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',2,'" & strUserNum & "') "
               End If
            End Select
        Case Else
        End Select
         '中所
    Case "S2"
        Select Case Val(strTemp(1))
        '申
        Case 95 ', 15
            '商申案
            If Left(adoRecordset("CP10"), 1) <> "4" And Left(adoRecordset("CP10"), 1) <> "6" And adoRecordset("CP10") <> "202" And adoRecordset("CP10") <> "210" And adoRecordset("CP10") <> "204" And adoRecordset("CP10") <> "205" And adoRecordset("CP10") < "7" And adoRecordset("CP09") < "C" Then
               cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086009,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'商申'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',1,'" & strUserNum & "') "
            '2009/7/1 modify by sonia 將其他及商爭順序顛倒,否則CP10>="7"者不會列入其他反而列入商爭,例:第一期註冊費
            '其他
            ElseIf Left(adoRecordset("CP10"), 1) >= "7" Then
               ProcessOther
            '商爭案
            ElseIf (Left(adoRecordset("CP10"), 1) = "4" Or Left(adoRecordset("CP10"), 1) = "6" Or adoRecordset("CP10") = "202" Or adoRecordset("CP10") = "210" Or adoRecordset("CP10") = "204" Or adoRecordset("CP10") = "205" Or adoRecordset("CP10") >= "7") And adoRecordset("CP09") < "C" Then
               cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086009,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'商爭'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',2,'" & strUserNum & "') "
            End If
        '爭
        Case Else
            '商申案
            If Left(adoRecordset("CP10"), 1) <> "4" And Left(adoRecordset("CP10"), 1) <> "6" And adoRecordset("CP10") <> "202" And adoRecordset("CP10") <> "210" And adoRecordset("CP10") <> "204" And adoRecordset("CP10") <> "205" And adoRecordset("CP10") < "7" And adoRecordset("CP09") < "C" Then
               cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088009,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'商申'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',1,'" & strUserNum & "') "
            '2009/7/1 modify by sonia 將其他及商爭順序顛倒,否則CP10>="7"者不會列入其他反而列入商爭,例:第一期註冊費
            '其他
            ElseIf Left(adoRecordset("CP10"), 1) >= "7" Then
               ProcessOther
            '商爭案
            ElseIf (Left(adoRecordset("CP10"), 1) = "4" Or Left(adoRecordset("CP10"), 1) = "6" Or adoRecordset("CP10") = "202" Or adoRecordset("CP10") = "210" Or adoRecordset("CP10") = "204" Or adoRecordset("CP10") = "205" Or adoRecordset("CP10") >= "7") And adoRecordset("CP09") < "C" Then
               cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088009,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'商爭'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',2,'" & strUserNum & "') "
            End If
        End Select
    '南所
    Case "S3"
        Select Case Val(strTemp(1))
        '申
        Case 95 ', 15
            '商申案
            If Left(adoRecordset("CP10"), 1) <> "4" And Left(adoRecordset("CP10"), 1) <> "6" And adoRecordset("CP10") <> "202" And adoRecordset("CP10") <> "210" And adoRecordset("CP10") <> "204" And adoRecordset("CP10") <> "205" And adoRecordset("CP10") < "7" And adoRecordset("CP09") < "C" Then
               cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086010,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'商申'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',1,'" & strUserNum & "') "
            '2009/7/1 modify by sonia 將其他及商爭順序顛倒,否則CP10>="7"者不會列入其他反而列入商爭,例:第一期註冊費
            '其他
            ElseIf Left(adoRecordset("CP10"), 1) >= "7" Then
               ProcessOther
            '商爭案
            ElseIf (Left(adoRecordset("CP10"), 1) = "4" Or Left(adoRecordset("CP10"), 1) = "6" Or adoRecordset("CP10") = "202" Or adoRecordset("CP10") = "210" Or adoRecordset("CP10") = "204" Or adoRecordset("CP10") = "205" Or adoRecordset("CP10") >= "7") And adoRecordset("CP09") < "C" Then
               cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086010,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'商爭'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',2,'" & strUserNum & "') "
            End If
        '爭
        Case Else
            '商申案
            If Left(adoRecordset("CP10"), 1) <> "4" And Left(adoRecordset("CP10"), 1) <> "6" And adoRecordset("CP10") <> "202" And adoRecordset("CP10") <> "210" And adoRecordset("CP10") <> "204" And adoRecordset("CP10") <> "205" And adoRecordset("CP10") < "7" And adoRecordset("CP09") < "C" Then
               cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088010,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'商申'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',1,'" & strUserNum & "') "
            '2009/7/1 modify by sonia 將其他及商爭順序顛倒,否則CP10>="7"者不會列入其他反而列入商爭,例:第一期註冊費
            '其他
            ElseIf Left(adoRecordset("CP10"), 1) >= "7" Then
               ProcessOther
            '商爭案
            ElseIf (Left(adoRecordset("CP10"), 1) = "4" Or Left(adoRecordset("CP10"), 1) = "6" Or adoRecordset("CP10") = "202" Or adoRecordset("CP10") = "210" Or adoRecordset("CP10") = "204" Or adoRecordset("CP10") = "205" Or adoRecordset("CP10") >= "7") And adoRecordset("CP09") < "C" Then
               cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088010,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'商爭'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',2,'" & strUserNum & "') "
            End If
        End Select
    '高所
    Case "S4"
        Select Case Val(strTemp(1))
        '申
        Case 95 ', 15
            '商申案
            If Left(adoRecordset("CP10"), 1) <> "4" And Left(adoRecordset("CP10"), 1) <> "6" And adoRecordset("CP10") <> "202" And adoRecordset("CP10") <> "210" And adoRecordset("CP10") <> "204" And adoRecordset("CP10") <> "205" And adoRecordset("CP10") < "7" And adoRecordset("CP09") < "C" Then
               cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086011,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'商申'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',1,'" & strUserNum & "') "
            '2009/7/1 modify by sonia 將其他及商爭順序顛倒,否則CP10>="7"者不會列入其他反而列入商爭,例:第一期註冊費
            '其他
            ElseIf Left(adoRecordset("CP10"), 1) >= "7" Then
               ProcessOther
            '商爭案
            ElseIf (Left(adoRecordset("CP10"), 1) = "4" Or Left(adoRecordset("CP10"), 1) = "6" Or adoRecordset("CP10") = "202" Or adoRecordset("CP10") = "210" Or adoRecordset("CP10") = "204" Or adoRecordset("CP10") = "205" Or adoRecordset("CP10") >= "7") And adoRecordset("CP09") < "C" Then
               cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086011,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'商爭'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',2,'" & strUserNum & "') "
            End If
        '爭
        Case Else
            '商申案
            If Left(adoRecordset("CP10"), 1) <> "4" And Left(adoRecordset("CP10"), 1) <> "6" And adoRecordset("CP10") <> "202" And adoRecordset("CP10") <> "210" And adoRecordset("CP10") <> "204" And adoRecordset("CP10") <> "205" And adoRecordset("CP10") < "7" And adoRecordset("CP09") < "C" Then
               cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088011,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'商申'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',1,'" & strUserNum & "') "
            '2009/7/1 modify by sonia 將其他及商爭順序顛倒,否則CP10>="7"者不會列入其他反而列入商爭,例:第一期註冊費
            '其他
            ElseIf Left(adoRecordset("CP10"), 1) >= "7" Then
               ProcessOther
            '商爭案
            ElseIf (Left(adoRecordset("CP10"), 1) = "4" Or Left(adoRecordset("CP10"), 1) = "6" Or adoRecordset("CP10") = "202" Or adoRecordset("CP10") = "210" Or adoRecordset("CP10") = "204" Or adoRecordset("CP10") = "205" Or adoRecordset("CP10") >= "7") And adoRecordset("CP09") < "C" Then
               cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088011,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'商爭'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',2,'" & strUserNum & "') "
            End If
        End Select
    '其他
    Case Else
        Select Case Val(strTemp(1))
        '申
        Case 95 ', 15
            '商申案
            If Left(adoRecordset("CP10"), 1) <> "4" And Left(adoRecordset("CP10"), 1) <> "6" And adoRecordset("CP10") <> "202" And adoRecordset("CP10") <> "210" And adoRecordset("CP10") <> "204" And adoRecordset("CP10") <> "205" And adoRecordset("CP10") < "7" And adoRecordset("CP09") < "C" Then
               cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086012,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'商申'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',1,'" & strUserNum & "') "
            '2009/7/1 modify by sonia 將其他及商爭順序顛倒,否則CP10>="7"者不會列入其他反而列入商爭,例:第一期註冊費
            '其他
            ElseIf Left(adoRecordset("CP10"), 1) >= "7" Then
               ProcessOther
            '商爭案
            ElseIf (Left(adoRecordset("CP10"), 1) = "4" Or Left(adoRecordset("CP10"), 1) = "6" Or adoRecordset("CP10") = "202" Or adoRecordset("CP10") = "210" Or adoRecordset("CP10") = "204" Or adoRecordset("CP10") = "205" Or adoRecordset("CP10") >= "7") And adoRecordset("CP09") < "C" Then
               cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086012,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'商爭'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',2,'" & strUserNum & "') "
            End If
        '爭
        Case Else
            '商申案
            If Left(adoRecordset("CP10"), 1) <> "4" And Left(adoRecordset("CP10"), 1) <> "6" And adoRecordset("CP10") <> "202" And adoRecordset("CP10") <> "210" And adoRecordset("CP10") <> "204" And adoRecordset("CP10") <> "205" And adoRecordset("CP10") < "7" And adoRecordset("CP09") < "C" Then
               cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088012,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'商申'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',1,'" & strUserNum & "') "
            '2009/7/1 modify by sonia 將其他及商爭順序顛倒,否則CP10>="7"者不會列入其他反而列入商爭,例:第一期註冊費
            '其他
            ElseIf Left(adoRecordset("CP10"), 1) >= "7" Then
               ProcessOther
            '商爭案
            ElseIf (Left(adoRecordset("CP10"), 1) = "4" Or Left(adoRecordset("CP10"), 1) = "6" Or adoRecordset("CP10") = "202" Or adoRecordset("CP10") = "210" Or adoRecordset("CP10") = "204" Or adoRecordset("CP10") = "205" Or adoRecordset("CP10") >= "7") And adoRecordset("CP09") < "C" Then
               cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088012,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'商爭'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',2,'" & strUserNum & "') "
            End If
        End Select
    End Select
'國外
Case Is > 0
    Select Case Mid(UCase(strTemp(2)), 1, 2)
    '北所
    Case "S1"
        Select Case UCase(strTemp(2))
        '北一
        Case "S11"
            Select Case Val(strTemp(1))
            '申
            Case 95 ', 15
                'Modified by Lydia 2023/12/01 ,4,=>,5,
                cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086004,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'大陸'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',5,'" & strUserNum & "') "
            '爭
            Case Else
                'Modified by Lydia 2023/12/01 ,4,=>,5,
                cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088004,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'大陸'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',5,'" & strUserNum & "') "
            End Select
        '北二
        Case "S12"
            Select Case Val(strTemp(1))
            '申
            Case 95 ', 15
                'Modified by Lydia 2023/12/01 ,4,=>,5,
                cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086005,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'大陸'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',5,'" & strUserNum & "') "
            '爭
            Case Else
                'Modified by Lydia 2023/12/01 ,4,=>,5,
                cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088005,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'大陸'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',5,'" & strUserNum & "') "
            End Select
        '北三
        Case "S13"
            Select Case Val(strTemp(1))
            '申
            Case 95 ', 15
                'Modified by Lydia 2023/12/01 ,4,=>,5,
                cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086006,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'大陸'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',5,'" & strUserNum & "') "
            '爭
            Case Else
                'Modified by Lydia 2023/12/01 ,4,=>,5,
                cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088006,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'大陸'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',5,'" & strUserNum & "') "
            End Select
        '北四
        Case "S14"
            Select Case Val(strTemp(1))
            '申
            Case 95 ', 15
                'Modified by Lydia 2023/12/01 ,4,=>,5,
                cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086007,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'大陸'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',5,'" & strUserNum & "') "
            '爭
            Case Else
                'Modified by Lydia 2023/12/01 ,4,=>,5,
                cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088007,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'大陸'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',5,'" & strUserNum & "') "
            End Select
        '北五
        Case "S15"
            Select Case Val(strTemp(1))
            '申
            Case 95 ', 15
                'Modified by Lydia 2023/12/01 ,4,=>,5,
                cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086008,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'大陸'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',5,'" & strUserNum & "') "
            '爭
            Case Else
                'Modified by Lydia 2023/12/01 ,4,=>,5,
                cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088008,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'大陸'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',5,'" & strUserNum & "') "
            End Select
        Case Else
        End Select
    '中所
    Case "S2"
        Select Case Val(strTemp(1))
        '申
        Case 95 ', 15
            'Modified by Lydia 2023/12/01 ,4,=>,5,
            cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086009,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'大陸'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',5,'" & strUserNum & "') "
        '爭
        Case Else
            'Modified by Lydia 2023/12/01 ,4,=>,5,
            cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088009,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'大陸'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',5,'" & strUserNum & "') "
        End Select
    '南所
    Case "S3"
        Select Case Val(strTemp(1))
        '申
        Case 95 ', 15
            'Modified by Lydia 2023/12/01 ,4,=>,5,
            cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086010,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'大陸'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',5,'" & strUserNum & "') "
        '爭
        Case Else
            'Modified by Lydia 2023/12/01 ,4,=>,5,
            cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088010,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'大陸'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',5,'" & strUserNum & "') "
        End Select
    '高所
    Case "S4"
        Select Case Val(strTemp(1))
        '申
        Case 95 ', 15
            'Modified by Lydia 2023/12/01 ,4,=>,5,
            cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086011,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'大陸'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',5,'" & strUserNum & "') "
        '爭
        Case Else
            'Modified by Lydia 2023/12/01 ,4,=>,5,
            cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088011,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'大陸'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',5,'" & strUserNum & "') "
        End Select
    '其他
    Case Else
        Select Case Val(strTemp(1))
        '申
        Case 95 ', 15
            'Modified by Lydia 2023/12/01 ,4,=>,5,
            cnnConnection.Execute "INSERT INTO R020417_1 (R086001,R086002,R086003,R086012,R086013,R086014,R086015,R086016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'大陸'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','*',5,'" & strUserNum & "') "
        '爭
        Case Else
            'Modified by Lydia 2023/12/01 ,4,=>,5,
            cnnConnection.Execute "INSERT INTO R020417_2 (R088001,R088002,R088003,R088012,R088013,R088014,R088015,R088016,ID) VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp3) & ",'大陸'," & strTemp(15) & "," & strTemp(15) & ",'" & ChgSQL(strTemp(3)) & "','',5,'" & strUserNum & "') "
        End Select
    End Select
End Select
End Sub

Private Sub Combo1_Click()
If Combo1.ListIndex >= SeekPrint Then
   j = Combo1.ListIndex + 1
Else
   j = Combo1.ListIndex
End If
Set Printer = Printers(j)

End Sub

Private Sub Form_Load()
MoveFormToCenter Me
txt1(0) = GetSystemKindByNickT

SeekPrintL = Printer.Orientation
'Modified by Morgan 2015/6/3
'j = 0
'strSql = Printer.DeviceName
'For i = 0 To Printers.Count - 1
'    Set Printer = Printers(i)
'    If Printer.DeviceName <> strSql Then
'        Combo1.AddItem Printer.DeviceName, j
'        j = j + 1
'    End If
'    If Printer.DeviceName = strSql Then
'        SeekPrint = i
'    End If
'Next i
'Combo1.Text = Combo1.List(0)
PUB_SetPrinter Me.Name, Combo1, , , SeekPrint
'end 2015/6/3
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Added by Morgan 2015/6/3
'若印表機變動, 則更新列印設定
If Me.Combo1.Text <> Me.Combo1.Tag Then
    PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
End If
'end 2015/6/3
Set Printer = Printers(SeekPrint)
Printer.Orientation = SeekPrintL
Set frm020417 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub Txt1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cmdok(0).SetFocus
End If
End Sub
Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txt1_LostFocus(Index As Integer)

Select Case Index
Case 0
     strTemp1 = Split(Replace(UCase(GetSystemKindByNickT), ",,", ""), ",")
     strTemp2 = Split(Replace(UCase(txt1(0)), ",,", ""), ",")
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
'Modify By Sindy 2019/12/6
Case 1, 2
   If txt1(Index) = MsgText(601) Then Exit Sub
   If ChkDate(txt1(Index) & "01") = False Then
      txt1(Index).SetFocus
      txt1_GotFocus Index
      Exit Sub
   End If
   If Index = 2 Then
      If RunNick2(txt1(1), txt1(2)) = True Then
         txt1(Index).SetFocus
         txt1_GotFocus Index
         Exit Sub
      End If
   End If
'Case 1
'    'Modify By Cheng 2002/11/28
'    '若有輸入資料才檢查
'    If Me.txt1(1).Text <> "" Then
'        If Val(txt1(1)) <= 0 Or Val(txt1(1)) > (Val(Format(Date, "YYYY")) - 1911) Then
'            s = MsgBox("年輸入錯誤!!", , "USER 輸入錯誤")
'            txt1(1).SetFocus
'            txt1(1).SelStart = 0
'            txt1(1).SelLength = Len(txt1(1))
'            Exit Sub
'        End If
'    End If
'Case 2
'    'Modify By Cheng 2002/11/28
'    '若有輸入資料才檢查
'    If Me.txt1(2).Text <> "" Then
'        If Val(txt1(2)) <= 0 Or Val(txt1(2)) >= 13 Or IIf(Val(txt1(1)) = Val(Format(Date, "YYYY")) - 1911, IIf(Val(txt1(2)) > Val(Format(Date, "mm")), True, False), False) Then
'            s = MsgBox("月份輸入錯誤!!", , "USER 輸入錯誤")
'            txt1(2).SetFocus
'            txt1(2).SelStart = 0
'            txt1(2).SelLength = Len(txt1(2))
'            Exit Sub
'        End If
'    End If
Case 4
     If RunNick(txt1(Index - 1), txt1(Index)) Then
         txt1(Index - 1).SetFocus
         txt1_GotFocus (Index - 1)
         Exit Sub
      End If

Case Else
End Select
End Sub

Sub ShowLine()
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(19000, iPrint + 150)
iPrint = iPrint + 300
End Sub

