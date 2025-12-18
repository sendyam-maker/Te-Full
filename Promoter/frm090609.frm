VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090609 
   BorderStyle     =   1  '單線固定
   Caption         =   "承辦人工作量查詢"
   ClientHeight    =   2364
   ClientLeft      =   2412
   ClientTop       =   1728
   ClientWidth     =   4404
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2364
   ScaleWidth      =   4404
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Index           =   1
      Left            =   3156
      TabIndex        =   14
      Top             =   20
      Width           =   1200
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2376
      TabIndex        =   13
      Top             =   20
      Width           =   756
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   7
      Left            =   1560
      MaxLength       =   1
      TabIndex        =   12
      Top             =   1860
      Width           =   300
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   6
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   11
      Top             =   1860
      Width           =   330
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   5
      Left            =   1872
      MaxLength       =   3
      TabIndex        =   10
      Top             =   1524
      Width           =   735
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   4
      Left            =   1080
      MaxLength       =   3
      TabIndex        =   9
      Top             =   1536
      Width           =   705
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   3
      Left            =   1080
      MaxLength       =   6
      TabIndex        =   8
      Top             =   1188
      Width           =   855
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   2
      Left            =   1872
      MaxLength       =   4
      TabIndex        =   7
      Top             =   852
      Width           =   705
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   1
      Left            =   1080
      MaxLength       =   4
      TabIndex        =   6
      Top             =   852
      Width           =   705
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   0
      Left            =   1080
      TabIndex        =   5
      Top             =   516
      Width           =   1995
   End
   Begin MSForms.Label lbl1 
      Height          =   300
      Left            =   1980
      TabIndex        =   16
      Top             =   1185
      Width           =   1470
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2593;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "(1.北 2.中 3.南 4.高 5.其他)"
      Height          =   180
      Index           =   7
      Left            =   1872
      TabIndex        =   15
      Top             =   1908
      Width           =   2412
   End
   Begin VB.Line Line3 
      X1              =   1230
      X2              =   1695
      Y1              =   1980
      Y2              =   1980
   End
   Begin VB.Line Line2 
      X1              =   1512
      X2              =   2187
      Y1              =   1668
      Y2              =   1668
   End
   Begin VB.Line Line1 
      X1              =   1488
      X2              =   2118
      Y1              =   1008
      Y2              =   1008
   End
   Begin VB.Label Label1 
      Caption         =   "所別："
      Height          =   180
      Index           =   4
      Left            =   84
      TabIndex        =   4
      Top             =   1920
      Width           =   660
   End
   Begin VB.Label Label1 
      Caption         =   "部門代號："
      Height          =   180
      Index           =   3
      Left            =   84
      TabIndex        =   3
      Top             =   1584
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "承辦人："
      Height          =   180
      Index           =   2
      Left            =   84
      TabIndex        =   2
      Top             =   1248
      Width           =   828
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家："
      Height          =   180
      Index           =   1
      Left            =   84
      TabIndex        =   1
      Top             =   912
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   84
      TabIndex        =   0
      Top             =   576
      Width           =   1092
   End
End
Attribute VB_Name = "frm090609"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/12 改成Form2.0 ; lbl1
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
'Modified by Morgan 2013/4/19 原105條件要再加125
'modify by sonia 2016/9/6 cp27欄的判斷改用cp158,cp57改用cp159
Option Explicit
Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, s As Integer, SavDay1 As String, SavDay2 As String, SavDay3 As String, StrSQL7 As String, StrSQL4 As String, strSQL5 As String
Dim strSQL2 As String, iPrint As Integer, Page As Integer, strTemp(0 To 21) As String, strTemp3 As String, TestOk As Boolean, StrTemp7(0 To 13) As String, StrTemp99(0 To 7) As String, StrSQL3 As String, k As Integer
Dim PLeft(0 To 21) As Integer, strTemp1 As Variant, strTemp2 As Variant, StrSQL6 As String, Str020401SysKind As String, PLeft1(1 To 9) As Integer, SeekPrint As Integer, SeekPrintL As Integer
Public ObjForm As Integer
Dim bol911001checkRange As Boolean
Public m_bolRedo As Boolean 'Added by Morgan 2024/4/17

Private Sub cmdok_Click(Index As Integer)
Select Case Index
Case 0
     If Len(txt1(0)) = 0 Then
         s = MsgBox("系統類別不可空白!!", , "USER 輸入錯誤")
         txt1(0).SetFocus
         Exit Sub
     Else
         Screen.MousePointer = vbHourglass
         Me.Enabled = False
         ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/17 清除查詢印表記錄檔欄位
         Process
         Me.Enabled = True
         Screen.MousePointer = vbDefault
     End If
Case 1
     Unload Me
Case Else
End Select
End Sub

Sub Process()
Dim stPA As String, stTM As String, stLA As String, stHC As String, stSP As String 'Add by Morgan 2011/5/2
Dim StrSQL7 As String, strSQL8 As String 'Add by Morgan 2011/5/2
Dim intQ As Integer, tmpArr As Variant 'Added by Lydia 2017/12/25
Dim stColSQL As String 'Added by Morgan 2024/3/7

cnnConnection.Execute "DELETE FROM R090609_1 WHERE ID='" & strUserNum & "' "
cnnConnection.Execute "DELETE FROM R090609_2 WHERE ID='" & strUserNum & "' "
strSQL1 = ""
strSQL2 = ""
StrSQL3 = ""
StrSQL4 = ""
strSQL5 = ""
StrSQL6 = ""
StrSQL7 = "" '支援
strSQL8 = "" '收文點數
If Len(txt1(1)) <> 0 Then
    strSQL1 = strSQL1 + " AND PA09>='" & txt1(1) & "' "
    strSQL2 = strSQL2 + " AND TM10>='" & txt1(1) & "' "
    StrSQL3 = StrSQL3 + " AND LC15>='" & txt1(1) & "' "
    StrSQL4 = StrSQL4
    strSQL5 = strSQL5 + " AND SP09>='" & txt1(1) & "' "
End If
If Len(txt1(2)) <> 0 Then
    strSQL1 = strSQL1 + " AND PA09<='" & txt1(2) & "' "
    strSQL2 = strSQL2 + " AND TM10<='" & txt1(2) & "' "
    StrSQL3 = StrSQL3 + " AND LC15<='" & txt1(2) & "' "
    StrSQL4 = StrSQL4
    strSQL5 = strSQL5 + " AND SP09<='" & txt1(2) & "' "
End If
If Len(txt1(1)) <> 0 Or Len(txt1(2)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(1) & txt1(1) & "-" & txt1(2) 'Add By Sindy 2010/12/17
End If

If Len(txt1(4)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND S1.ST03>='" & txt1(4) & "' "
End If
If Len(txt1(5)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND S1.ST03<='" & txt1(5) & "' "
End If
If Len(txt1(4)) <> 0 Or Len(txt1(5)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(3) & txt1(4) & "-" & txt1(5) 'Add By Sindy 2010/12/17
End If
If Len(txt1(6)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND S1.ST06>='" & txt1(6) & "' "
End If
If Len(txt1(7)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND S1.ST06<='" & txt1(7) & "' "
End If
If Len(txt1(6)) <> 0 Or Len(txt1(7)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(4) & txt1(6) & "-" & txt1(7) & Label1(7) 'Add By Sindy 2010/12/17
End If

'add by nickc 2005/04/11 加離職不出現
StrSQL6 = StrSQL6 & " and S1.st04='1' "

StrSQL7 = StrSQL6 'Add by Morgan 2011/5/2 員工檔的條件共用(進度檔的要分開,支援及收文點數不同)

If Len(txt1(3)) <> 0 Then
    
    StrSQL6 = StrSQL6 + " AND CP14='" & txt1(3) & "' "
    'Add by Morgan 2011/5/2
    StrSQL6 = StrSQL6 + " AND S1.ST01='" & txt1(3) & "' "
    StrSQL7 = StrSQL7 + " AND S1.ST01='" & txt1(3) & "' "
    'end 2011/5/2
    
    pub_QL05 = pub_QL05 & ";" & Label1(2) & txt1(3) & LBL1 'Add By Sindy 2010/12/17
End If


StrSQL6 = StrSQL6 + " and CP159=0 "
pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(0) 'Add By Sindy 2010/12/17


'Modify by Morgan 2011/5/2 語法調整
'Memo by Morgan 2024/3/7 舊程式已刪除
stPA = SQLGrpStr(txt1(0), 1)
stTM = SQLGrpStr(txt1(0), 2)
stLA = SQLGrpStr(txt1(0), 3)
stHC = SQLGrpStr(txt1(0), 4)
stSP = SQLGrpStr(txt1(0), 5)

'超過法定
strSql = "select cp14,Sum(Decode(CP26, Null, 1, 0)),0,0,0,0,0,0,0,0,0,0,'" & strUserNum & "',0,0,0,0,0,0,0,0 from caseprogress,PATENT,staff S1 where CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp14=S1.ST01(+) and cp06>=" & ChangeWDateStringToWString(DateAdd("m", -6, ChangeWStringToWDateString(strSrvDate(1)))) & " and cp06<=" & strSrvDate(1) & "  " & strSQL1 & StrSQL6 & " AND CP01 IN (" & stPA & ") and cp158=0  GROUP BY cp14 "
If stTM <> "' '" Then strSql = strSql + " UNION all  SELECT cp14,Sum(Decode(CP26, Null, 1, 0)),0,0,0,0,0,0,0,0,0,0,'" & strUserNum & "',0,0,0,0,0,0,0,0 from caseprogress,trademark,staff S1 where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND cp14=S1.ST01(+)  and cp06>=" & ChangeWDateStringToWString(DateAdd("m", -6, ChangeWStringToWDateString(strSrvDate(1)))) & " and cp06<=" & strSrvDate(1) & "  " & strSQL2 + StrSQL6 & " AND CP01 IN (" & stTM & ") and cp158=0   GROUP BY cp14 "
If stLA <> "' '" Then strSql = strSql + " UNION all  SELECT cp14,Sum(Decode(CP26, Null, 1, 0)),0,0,0,0,0,0,0,0,0,0,'" & strUserNum & "',0,0,0,0,0,0,0,0 from caseprogress,LAWCASE,staff S1 where CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND cp14=S1.ST01(+) and cp06>=" & ChangeWDateStringToWString(DateAdd("m", -6, ChangeWStringToWDateString(strSrvDate(1)))) & " and cp06<=" & strSrvDate(1) & "  " & StrSQL3 + StrSQL6 & " AND CP01 IN (" & stLA & ") and cp158=0  GROUP BY cp14 "
If stHC <> "' '" Then strSql = strSql + " UNION all  SELECT cp14,Sum(Decode(CP26, Null, 1, 0)),0,0,0,0,0,0,0,0,0,0,'" & strUserNum & "',0,0,0,0,0,0,0,0 from caseprogress,HIRECASE,staff S1 where CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND cp14=S1.ST01(+) and cp06>=" & ChangeWDateStringToWString(DateAdd("m", -6, ChangeWStringToWDateString(strSrvDate(1)))) & " and cp06<=" & strSrvDate(1) & "  " & StrSQL4 + StrSQL6 & " AND CP01 IN (" & stHC & ") and cp158=0   GROUP BY cp14 "
If stSP <> "' '" Then strSql = strSql + " UNION all  SELECT cp14,Sum(Decode(CP26, Null, 1, 0)),0,0,0,0,0,0,0,0,0,0,'" & strUserNum & "',0,0,0,0,0,0,0,0 from caseprogress,SERVICEPRACTICE,staff S1 where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND cp14=S1.ST01(+) and cp06>=" & ChangeWDateStringToWString(DateAdd("m", -6, ChangeWStringToWDateString(strSrvDate(1)))) & " and cp06<=" & strSrvDate(1) & "  " & strSQL5 + StrSQL6 & " AND CP01 IN (" & stSP & ") and cp158=0 GROUP BY cp14 "

'承辦量 1 設計只有專利有
strSql = strSql + " UNION all  SELECT cp14,0,Sum(Decode(CP26, Null, 1, 0)),0,0,0,0,0,0,0,0,0,'" & strUserNum & "',sum(decode(cp112,'Y',round(nvl(cp97,0) * nvl(cp98,0) * nvl(cp111,1),2),round(cp97 * cp98,2))),0,0,0,0,0,0,0 from caseprogress,engineerprogress,PATENT,staff S1 where ep02=cp09(+) and CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp14=S1.ST01(+) " & strSQL1 + StrSQL6 & " AND EP09>=" & Mid(GetTodayDate, 1, 6) & "01 and ep09<=" & Mid(GetTodayDate, 1, 6) & "31 AND CP10 IN ('103','105','125') " & " AND CP01 IN (" & stPA & ") and ((cp158>=" & Mid(GetTodayDate, 1, 6) & "01 and cp158<=" & Mid(GetTodayDate, 1, 6) & "31 ) or cp158=0 ) GROUP BY cp14 "

'承辦量 2
'Modified by Morgan 2024/3/7 專利非設計案的案件數改為>=特定考核值的件數(機構組:1,電子組:0.83,生化組:0.67)
'strSql = strSql + " UNION all  SELECT cp14,0,0,Sum(Decode(CP26, Null, 1, 0)),0,0,0,0,0,0,0,0,'" & strUserNum & "',0,sum(decode(cp112,'Y',round(nvl(cp97,0) * nvl(cp98,0) * nvl(cp111,1),2),round(cp97 * cp98,2))),0,0,0,0,0,0 from caseprogress,engineerprogress,PATENT,staff S1 where ep02=cp09(+) and CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp14=S1.ST01(+) " & strSQL1 + StrSQL6 & " AND EP09>=" & Mid(strSrvDate(1), 1, 6) & "01 and ep09<=" & Mid(strSrvDate(1), 1, 6) & "31 AND CP10 NOT IN ('103','105','125') " & " AND CP01 IN (" & stPA & ") and ((cp158>=" & Mid(strSrvDate(1), 1, 6) & "01 and cp158<=" & Mid(strSrvDate(1), 1, 6) & "31 ) or cp158=0 ) GROUP BY cp14 "
'Modified by Morgan 2025/1/8 應該要用工程師的組別判斷而不是案件屬性
'stColSQL = "Decode(sign(cp97 * cp98 * decode(cp112,'Y',nvl(cp111,1),1)-decode(pa158,'1',1,'2',0.83,'3',0.67)),0,1,1,1,0)"
stColSQL = "Decode(sign(cp97 * cp98 * decode(cp112,'Y',nvl(cp111,1),1)-decode(s1.st70,'1',1,'2',0.83,'3',0.67)),0,1,1,1,0)"
'end 2025/1/8
strSql = strSql + " UNION all  SELECT cp14,0,0,Sum(" & stColSQL & "),0,0,0,0,0,0,0,0,'" & strUserNum & "',0,sum(decode(cp112,'Y',round(nvl(cp97,0) * nvl(cp98,0) * nvl(cp111,1),2),round(cp97 * cp98,2))),0,0,0,0,0,0 from caseprogress,engineerprogress,PATENT,staff S1 where ep02=cp09(+) and CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp14=S1.ST01(+) " & strSQL1 + StrSQL6 & " AND EP09>=" & Mid(strSrvDate(1), 1, 6) & "01 and ep09<=" & Mid(strSrvDate(1), 1, 6) & "31 AND CP10 NOT IN ('103','105','125') " & " AND CP01 IN (" & stPA & ") and ((cp158>=" & Mid(strSrvDate(1), 1, 6) & "01 and cp158<=" & Mid(strSrvDate(1), 1, 6) & "31 ) or cp158=0 ) GROUP BY cp14 "
If stTM <> "' '" Then strSql = strSql + " UNION all  SELECT cp14,0,0,Sum(Decode(CP26, Null, 1, 0)),0,0,0,0,0,0,0,0,'" & strUserNum & "',0,sum(decode(cp112,'Y',round(nvl(cp97,0) * nvl(cp98,0) * nvl(cp111,1),2),round(cp97 * cp98,2))),0,0,0,0,0,0 from caseprogress,engineerprogress,trademark,staff S1 where ep02=cp09(+) and CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND cp14=S1.ST01(+) " & strSQL2 + StrSQL6 & " AND EP09>=" & Mid(strSrvDate(1), 1, 6) & "01 and ep09<=" & Mid(strSrvDate(1), 1, 6) & "31 AND CP01 IN (" & stTM & ") and ((cp158>=" & Mid(strSrvDate(1), 1, 6) & "01 and cp158<=" & Mid(strSrvDate(1), 1, 6) & "31 ) or cp158=0 ) GROUP BY cp14 "
If stLA <> "' '" Then strSql = strSql + " UNION all  SELECT cp14,0,0,Sum(Decode(CP26, Null, 1, 0)),0,0,0,0,0,0,0,0,'" & strUserNum & "',0,sum(decode(cp112,'Y',round(nvl(cp97,0) * nvl(cp98,0) * nvl(cp111,1),2),round(cp97 * cp98,2))),0,0,0,0,0,0 from caseprogress,engineerprogress,LAWCASE,staff S1 where ep02=cp09(+) and CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND cp14=S1.ST01(+) " & StrSQL3 + StrSQL6 & " AND EP09>=" & Mid(strSrvDate(1), 1, 6) & "01 and ep09<=" & Mid(strSrvDate(1), 1, 6) & "31  AND CP01 IN (" & stLA & ") and ((cp158>=" & Mid(strSrvDate(1), 1, 6) & "01 and cp158<=" & Mid(strSrvDate(1), 1, 6) & "31 ) or cp158=0 ) GROUP BY cp14 "
If stHC <> "' '" Then strSql = strSql + " UNION all  SELECT cp14,0,0,Sum(Decode(CP26, Null, 1, 0)),0,0,0,0,0,0,0,0,'" & strUserNum & "',0,sum(decode(cp112,'Y',round(nvl(cp97,0) * nvl(cp98,0) * nvl(cp111,1),2),round(cp97 * cp98,2))),0,0,0,0,0,0 from caseprogress,engineerprogress,HIRECASE,staff S1 where ep02=cp09(+) and CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND cp14=S1.ST01(+) " & StrSQL4 + StrSQL6 & " AND EP09>=" & Mid(strSrvDate(1), 1, 6) & "01 and ep09<=" & Mid(strSrvDate(1), 1, 6) & "31  AND CP01 IN (" & stHC & ") and ((cp158>=" & Mid(strSrvDate(1), 1, 6) & "01 and cp158<=" & Mid(strSrvDate(1), 1, 6) & "31 ) or cp158=0 ) GROUP BY cp14 "
If stSP <> "' '" Then strSql = strSql + " UNION all  SELECT cp14,0,0,Sum(Decode(CP26, Null, 1, 0)),0,0,0,0,0,0,0,0,'" & strUserNum & "',0,sum(decode(cp112,'Y',round(nvl(cp97,0) * nvl(cp98,0) * nvl(cp111,1),2),round(cp97 * cp98,2))),0,0,0,0,0,0 from caseprogress,engineerprogress,SERVICEPRACTICE,staff S1 where ep02=cp09(+) and CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND cp14=S1.ST01(+) " & strSQL5 + StrSQL6 & " AND EP09>=" & Mid(strSrvDate(1), 1, 6) & "01 and ep09<=" & Mid(strSrvDate(1), 1, 6) & "31 AND CP01 IN (" & stSP & ") and ((cp158>=" & Mid(strSrvDate(1), 1, 6) & "01 and cp158<=" & Mid(strSrvDate(1), 1, 6) & "31 ) or cp158=0 ) GROUP BY cp14 "

'Add by Morgan 2011/5/2
'支援(列非設計且不考慮系統及國家條件)
'Modified by Morgan 2014/3/20 2014/4/1 起支援改每小時折計0.2基數
'strSql = strSql + " UNION all  SELECT SH02,0,0,0,0,0,0,0,0,0,0,0,'" & strUserNum & "',0,Sum(Round(Decode(SH06, 'CFP', Nvl(SH05,0)/3, Nvl(SH05,0)/4) ,2)),0,0,0,0,0,0 from staff S1,SupportHour where SH02=S1.ST01(+) " & StrSQL7 & " AND SH01>=" & Mid(GetTodayDate, 1, 6) & "01 and SH01<=" & Mid(GetTodayDate, 1, 6) & "31  And SH11='V' GROUP BY SH02 "
'Modified by Morgan 2019/4/9 108考核支援時數轉換要除組別參數
'strSql = strSql + " UNION all  SELECT SH02,0,0,0,0,0,0,0,0,0,0,0,'" & strUserNum & "',0,Sum(Round(" & Sh2EPtCode & " ,2)),0,0,0,0,0,0 from staff S1,SupportHour where SH02=S1.ST01(+) " & StrSQL7 & " AND SH01>=" & Mid(GetTodayDate, 1, 6) & "01 and SH01<=" & Mid(GetTodayDate, 1, 6) & "31  And SH11='V' GROUP BY SH02 "
strSql = strSql + " UNION all  SELECT SH02,0,0,0,0,0,0,0,0,0,0,0,'" & strUserNum & "',0,Sum(Round(" & Sh2EPtCode & " / GetDivNum(st70,sh01) ,2)),0,0,0,0,0,0 from staff S1,SupportHour where SH02=S1.ST01(+) " & StrSQL7 & " AND SH01>=" & Mid(GetTodayDate, 1, 6) & "01 and SH01<=" & Mid(GetTodayDate, 1, 6) & "31  And SH11='V' GROUP BY SH02 "
'end 2014/3/20
'收文點數
'Modify by Morgan 2011/6/1 若有建點數分配資料時點數改分配點數(目前有225提供書狀意見及226配合開庭)
'Modified by Morgan 2014/3/20 --2014/4/1起非智權收文改每點折算0.04基數
'Memo by Morgan 2024/3/7 舊程式已刪除
If PUB_108RuleDate > strSrvDate(1) Then 'Added by Morgan 2019/4/23 108考核(取消收文點數轉換)
   strSql = strSql + " UNION all  SELECT cp13,0,0,0,0,0,0,0,0,0,0,0,'" & strUserNum & "',Sum(" & Pt2EPtCode & "),0,0,0,0,0,0,0 from staff S1,caseprogress,PATENT ,acc0n0 where a0n02(+)=cp09 and CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp13(+)=S1.ST01 " & strSQL1 + StrSQL7 & " AND cP05>=" & Mid(GetTodayDate, 1, 6) & "01 and cp05<=" & Mid(GetTodayDate, 1, 6) & "31 AND CP10 IN ('103','105','125') " & " AND CP01 IN (" & stPA & ") And nvl(a0n03/1000,cp18)>0 and cp20 is null and cp159=0 and substr(cp12,1,1)<>'S' GROUP BY cp13 "
   strSql = strSql + " UNION all  SELECT cp13,0,0,0,0,0,0,0,0,0,0,0,'" & strUserNum & "',0,Sum(" & Pt2EPtCode & "),0,0,0,0,0,0 from staff S1,caseprogress,PATENT ,acc0n0 where a0n02(+)=cp09 and CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp13(+)=S1.ST01 " & strSQL1 + StrSQL7 & " AND cP05>=" & Mid(GetTodayDate, 1, 6) & "01 and cp05<=" & Mid(GetTodayDate, 1, 6) & "31 AND CP10 NOT IN ('103','105','125') " & " AND CP01 IN (" & stPA & ") And nvl(a0n03/1000,cp18)>0 and cp20 is null and cp159=0 and substr(cp12,1,1)<>'S' GROUP BY cp13 "
   If stTM <> "' '" Then strSql = strSql + " UNION all  SELECT cp13,0,0,0,0,0,0,0,0,0,0,0,'" & strUserNum & "',0,Sum(" & Pt2EPtCode & "),0,0,0,0,0,0 from staff S1,caseprogress,trademark,acc0n0 where a0n02(+)=cp09 and CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND cp14=S1.ST01(+) " & strSQL2 + StrSQL7 & " AND cP05>=" & Mid(strSrvDate(1), 1, 6) & "01 and cP05<=" & Mid(strSrvDate(1), 1, 6) & "31 AND CP01 IN (" & stTM & ") And nvl(a0n03/1000,cp18)>0 and cp20 is null and cp159=0 and substr(cp12,1,1)<>'S' GROUP BY cp13 "
   If stLA <> "' '" Then strSql = strSql + " UNION all  SELECT cp13,0,0,0,0,0,0,0,0,0,0,0,'" & strUserNum & "',0,Sum(" & Pt2EPtCode & "),0,0,0,0,0,0 from staff S1,caseprogress,LAWCASE ,acc0n0 where a0n02(+)=cp09 and CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND cp14=S1.ST01(+) " & StrSQL3 + StrSQL7 & " AND cP05>=" & Mid(strSrvDate(1), 1, 6) & "01 and cP05<=" & Mid(strSrvDate(1), 1, 6) & "31  AND CP01 IN (" & stLA & ") And nvl(a0n03/1000,cp18)>0 and cp20 is null and cp159=0 and substr(cp12,1,1)<>'S' GROUP BY cp13 "
   If stHC <> "' '" Then strSql = strSql + " UNION all  SELECT cp13,0,0,0,0,0,0,0,0,0,0,0,'" & strUserNum & "',0,Sum(" & Pt2EPtCode & "),0,0,0,0,0,0 from staff S1,caseprogress,HIRECASE ,acc0n0 where a0n02(+)=cp09 and CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND cp14=S1.ST01(+) " & StrSQL4 + StrSQL7 & " AND cP05>=" & Mid(strSrvDate(1), 1, 6) & "01 and cP05<=" & Mid(strSrvDate(1), 1, 6) & "31  AND CP01 IN (" & stHC & ") And nvl(a0n03/1000,cp18)>0 and cp20 is null and cp159=0 and substr(cp12,1,1)<>'S' GROUP BY cp13 "
   If stSP <> "' '" Then strSql = strSql + " UNION all  SELECT cp13,0,0,0,0,0,0,0,0,0,0,0,'" & strUserNum & "',0,Sum(" & Pt2EPtCode & "),0,0,0,0,0,0 from staff S1,caseprogress,SERVICEPRACTICE ,acc0n0 where a0n02(+)=cp09 and CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND cp14=S1.ST01(+) " & strSQL5 + StrSQL7 & " AND cP05>=" & Mid(strSrvDate(1), 1, 6) & "01 and cP05<=" & Mid(strSrvDate(1), 1, 6) & "31 AND CP01 IN (" & stSP & ") And nvl(a0n03/1000,cp18)>0 and cp20 is null and cp159=0 and substr(cp12,1,1)<>'S' GROUP BY cp13 "
End If 'Added by Morgan 2019/4/23

'end 2014/3/20
'END 2011/6/1

'可辦量 1 設計只有專利有
strSql = strSql + " UNION all  SELECT cp14,0,0,0,Sum(Decode(CP26, Null, 1, 0)),0,0,0,0,0,0,0,'" & strUserNum & "',0,0,sum(decode(cp112,'Y',round(nvl(cp97,0) * nvl(cp98,0) * nvl(cp111,1),2),round(cp97 * cp98,2))),0,0,0,0,0 from staff S1,engineerprogress,caseprogress,PATENT where ep02=cp09(+) and CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND ep05(+)=S1.ST01 " & strSQL1 + StrSQL6 & " AND EP06>0 AND CP10 IN ('103','105','125') AND EP09 IS NULL " & " AND CP01 IN (" & stPA & ") and cp158=0 GROUP BY cp14 "

'可辦量 2
'Modified by Morgan 2024/3/7 專利非設計案的案件數改為>=特定考核值的件數
'strSql = strSql + " UNION all  SELECT cp14,0,0,0,0,Sum(Decode(CP26, Null, 1, 0)),0,0,0,0,0,0,'" & strUserNum & "',0,0,0,sum(decode(cp112,'Y',round(nvl(cp97,0) * nvl(cp98,0) * nvl(cp111,1),2),round(cp97 * cp98,2))),0,0,0,0 from staff S1,engineerprogress,caseprogress,PATENT where ep02=cp09(+) and CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND ep05(+)=S1.ST01 " & strSQL1 + StrSQL6 & " AND EP06>0 AND CP10 NOT IN ('103','105','125') AND EP09 IS NULL AND cp159=0 " & " AND CP01 IN (" & stPA & ") and cp158=0  GROUP BY cp14 "
strSql = strSql + " UNION all  SELECT cp14,0,0,0,0,Sum(" & stColSQL & "),0,0,0,0,0,0,'" & strUserNum & "',0,0,0,sum(decode(cp112,'Y',round(nvl(cp97,0) * nvl(cp98,0) * nvl(cp111,1),2),round(cp97 * cp98,2))),0,0,0,0 from staff S1,engineerprogress,caseprogress,PATENT where ep02=cp09(+) and CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND ep05(+)=S1.ST01 " & strSQL1 + StrSQL6 & " AND EP06>0 AND CP10 NOT IN ('103','105','125') AND EP09 IS NULL AND cp159=0 " & " AND CP01 IN (" & stPA & ") and cp158=0  GROUP BY cp14 "
If stTM <> "' '" Then strSql = strSql + " UNION all  SELECT cp14,0,0,0,0,Sum(Decode(CP26, Null, 1, 0)),0,0,0,0,0,0,'" & strUserNum & "',0,0,0,sum(decode(cp112,'Y',round(nvl(cp97,0) * nvl(cp98,0) * nvl(cp111,1),2),round(cp97 * cp98,2))),0,0,0,0 from staff S1,engineerprogress,caseprogress,trademark where ep02=cp09(+) and CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND ep05(+)=S1.ST01 " & strSQL2 + StrSQL6 & " AND EP06>0 AND EP09 IS NULL AND cp159=0 " & " AND CP01 IN (" & stTM & ") and cp158=0  GROUP BY cp14 "
If stLA <> "' '" Then strSql = strSql + " UNION all  SELECT cp14,0,0,0,0,Sum(Decode(CP26, Null, 1, 0)),0,0,0,0,0,0,'" & strUserNum & "',0,0,0,sum(decode(cp112,'Y',round(nvl(cp97,0) * nvl(cp98,0) * nvl(cp111,1),2),round(cp97 * cp98,2))),0,0,0,0 from staff S1,engineerprogress,caseprogress,LAWCASE where ep02=cp09(+) and CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND ep05(+)=S1.ST01 " & StrSQL3 + StrSQL6 & " AND EP06>0 AND EP09 IS NULL AND cp159=0 " & " AND CP01 IN (" & stLA & ") and cp158=0  GROUP BY cp14 "
If stHC <> "' '" Then strSql = strSql + " UNION all  SELECT cp14,0,0,0,0,Sum(Decode(CP26, Null, 1, 0)),0,0,0,0,0,0,'" & strUserNum & "',0,0,0,sum(decode(cp112,'Y',round(nvl(cp97,0) * nvl(cp98,0) * nvl(cp111,1),2),round(cp97 * cp98,2))),0,0,0,0 from staff S1,engineerprogress,caseprogress,HIRECASE where ep02=cp09(+) and CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND ep05(+)=S1.ST01 " & StrSQL4 + StrSQL6 & " AND EP06>0  AND EP09 IS NULL AND cp159=0 " & " AND CP01 IN (" & stHC & ") and cp158=0  GROUP BY cp14 "
If stSP <> "' '" Then strSql = strSql + " UNION all  SELECT cp14,0,0,0,0,Sum(Decode(CP26, Null, 1, 0)),0,0,0,0,0,0,'" & strUserNum & "',0,0,0,sum(decode(cp112,'Y',round(nvl(cp97,0) * nvl(cp98,0) * nvl(cp111,1),2),round(cp97 * cp98,2))),0,0,0,0 from staff S1,engineerprogress,caseprogress,SERVICEPRACTICE where ep02=cp09(+) and CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND ep05(+)=S1.ST01 " & strSQL5 + StrSQL6 & " AND EP06>0 AND EP09 IS NULL AND cp159=0 " & " AND CP01 IN (" & stSP & ") and cp158=0  GROUP BY cp14 "

'分案量 1 設計只有專利有
strSql = strSql + " UNION all  SELECT cp14,0,0,0,0,0,Sum(Decode(CP26, Null, 1, 0)),0,0,0,0,0,'" & strUserNum & "',0,0,0,0,sum(decode(cp112,'Y',round(nvl(cp97,0) * nvl(cp98,0) * nvl(cp111,1),2),round(cp97 * cp98,2))),0,0,0 from caseprogress,engineerprogress,PATENT,staff S1 where cp09=ep02(+) and CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp14=S1.ST01(+) " & strSQL1 + StrSQL6 & " AND CP05>=" & Mid(strSrvDate(1), 1, 6) & "01 and cp05<=" & Mid(strSrvDate(1), 1, 6) & "31 AND CP10 IN ('103','105','125') " & " AND CP01 IN (" & stPA & ") and ((cp158>=" & Mid(strSrvDate(1), 1, 6) & "01 and cp158<=" & Mid(strSrvDate(1), 1, 6) & "31 ) or cp158=0 ) GROUP BY cp14 "

'分案量 2
'Modified by Morgan 2024/3/7 專利非設計案的案件數改為>=特定考核值的件數
'strSql = strSql + " UNION all  SELECT cp14,0,0,0,0,0,0,Sum(Decode(CP26, Null, 1, 0)),0,0,0,0,'" & strUserNum & "',0,0,0,0,0,sum(decode(cp112,'Y',round(nvl(cp97,0) * nvl(cp98,0) * nvl(cp111,1),2),round(cp97 * cp98,2))),0,0 from caseprogress,engineerprogress,PATENT,staff S1 where cp09=ep02(+) and CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp14=S1.ST01(+) " & strSQL1 + StrSQL6 & " AND CP05>=" & Mid(strSrvDate(1), 1, 6) & "01 and cp05<=" & Mid(strSrvDate(1), 1, 6) & "31 AND CP10 NOT IN ('103','105','125') " & " AND CP01 IN (" & stPA & ") and ((cp158>=" & Mid(strSrvDate(1), 1, 6) & "01 and cp158<=" & Mid(strSrvDate(1), 1, 6) & "31 ) or cp158=0 ) GROUP BY cp14 "
strSql = strSql + " UNION all  SELECT cp14,0,0,0,0,0,0,Sum(" & stColSQL & "),0,0,0,0,'" & strUserNum & "',0,0,0,0,0,sum(decode(cp112,'Y',round(nvl(cp97,0) * nvl(cp98,0) * nvl(cp111,1),2),round(cp97 * cp98,2))),0,0 from caseprogress,engineerprogress,PATENT,staff S1 where cp09=ep02(+) and CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp14=S1.ST01(+) " & strSQL1 + StrSQL6 & " AND CP05>=" & Mid(strSrvDate(1), 1, 6) & "01 and cp05<=" & Mid(strSrvDate(1), 1, 6) & "31 AND CP10 NOT IN ('103','105','125') " & " AND CP01 IN (" & stPA & ") and ((cp158>=" & Mid(strSrvDate(1), 1, 6) & "01 and cp158<=" & Mid(strSrvDate(1), 1, 6) & "31 ) or cp158=0 ) GROUP BY cp14 "
If stTM <> "' '" Then strSql = strSql + " UNION all  SELECT cp14,0,0,0,0,0,0,Sum(Decode(CP26, Null, 1, 0)),0,0,0,0,'" & strUserNum & "',0,0,0,0,0,sum(decode(cp112,'Y',round(nvl(cp97,0) * nvl(cp98,0) * nvl(cp111,1),2),round(cp97 * cp98,2))),0,0 from caseprogress,engineerprogress,trademark,staff S1 where cp09=ep02(+) and CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND cp14=S1.ST01(+) " & strSQL2 + StrSQL6 & " AND CP05>=" & Mid(strSrvDate(1), 1, 6) & "01 and cp05<=" & Mid(strSrvDate(1), 1, 6) & "31  AND CP01 IN (" & stTM & ") and ((cp158>=" & Mid(strSrvDate(1), 1, 6) & "01 and cp158<=" & Mid(strSrvDate(1), 1, 6) & "31 ) or cp158=0 ) GROUP BY cp14 "
If stLA <> "' '" Then strSql = strSql + " UNION all  SELECT cp14,0,0,0,0,0,0,Sum(Decode(CP26, Null, 1, 0)),0,0,0,0,'" & strUserNum & "',0,0,0,0,0,sum(decode(cp112,'Y',round(nvl(cp97,0) * nvl(cp98,0) * nvl(cp111,1),2),round(cp97 * cp98,2))),0,0 from caseprogress,engineerprogress,LAWCASE,staff S1 where cp09=ep02(+) and CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND cp14=S1.ST01(+) " & StrSQL3 + StrSQL6 & " AND CP05>=" & Mid(strSrvDate(1), 1, 6) & "01 and cp05<=" & Mid(strSrvDate(1), 1, 6) & "31  AND CP01 IN (" & stLA & ") and ((cp158>=" & Mid(strSrvDate(1), 1, 6) & "01 and cp158<=" & Mid(strSrvDate(1), 1, 6) & "31 ) or cp158=0 ) GROUP BY cp14 "
If stHC <> "' '" Then strSql = strSql + " UNION all  SELECT cp14,0,0,0,0,0,0,Sum(Decode(CP26, Null, 1, 0)),0,0,0,0,'" & strUserNum & "',0,0,0,0,0,sum(decode(cp112,'Y',round(nvl(cp97,0) * nvl(cp98,0) * nvl(cp111,1),2),round(cp97 * cp98,2))),0,0 from caseprogress,engineerprogress,HIRECASE,staff S1 where cp09=ep02(+) and CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND cp14=S1.ST01(+) " & StrSQL4 + StrSQL6 & " AND CP05>=" & Mid(strSrvDate(1), 1, 6) & "01 and cp05<=" & Mid(strSrvDate(1), 1, 6) & "31 AND CP01 IN (" & stHC & ") and ((cp158>=" & Mid(strSrvDate(1), 1, 6) & "01 and cp158<=" & Mid(strSrvDate(1), 1, 6) & "31 ) or cp158=0 ) GROUP BY cp14 "
If stSP <> "' '" Then strSql = strSql + " UNION all  SELECT cp14,0,0,0,0,0,0,Sum(Decode(CP26, Null, 1, 0)),0,0,0,0,'" & strUserNum & "',0,0,0,0,0,sum(decode(cp112,'Y',round(nvl(cp97,0) * nvl(cp98,0) * nvl(cp111,1),2),round(cp97 * cp98,2))),0,0 from caseprogress,engineerprogress,SERVICEPRACTICE,staff S1 where cp09=ep02(+) and CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND cp14=S1.ST01(+) " & strSQL5 + StrSQL6 & " AND CP05>=" & Mid(strSrvDate(1), 1, 6) & "01 and cp05<=" & Mid(strSrvDate(1), 1, 6) & "31  AND CP01 IN (" & stSP & ") and ((cp158>=" & Mid(strSrvDate(1), 1, 6) & "01 and cp158<=" & Mid(strSrvDate(1), 1, 6) & "31 ) or cp158=0)  GROUP BY cp14 "

'發文件數 1 設計只有專利有
strSql = strSql + " UNION all  SELECT cp14,0,0,0,0,0,0,0,Sum(Decode(CP26, Null, 1, 0)),0,0,0,'" & strUserNum & "',0,0,0,0,0,0,sum(decode(cp112,'Y',round(nvl(cp97,0) * nvl(cp98,0) * nvl(cp111,1),2),round(cp97 * cp98,2))),0 from caseprogress,engineerprogress,PATENT,staff S1 where cp09=ep02(+) and CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp14=S1.ST01(+) " & strSQL1 + StrSQL6 & " AND CP10 IN ('103','105','125') " & " AND CP01 IN (" & stPA & ") and ((cp158>=" & Mid(strSrvDate(1), 1, 6) & "01 and cp158<=" & Mid(strSrvDate(1), 1, 6) & "31 )  ) GROUP BY cp14 "

'發文件數 2
strSql = strSql + " UNION all  SELECT cp14,0,0,0,0,0,0,0,0,Sum(Decode(CP26, Null, 1, 0)),0,0,'" & strUserNum & "',0,0,0,0,0,0,0,sum(decode(cp112,'Y',round(nvl(cp97,0) * nvl(cp98,0) * nvl(cp111,1),2),round(cp97 * cp98,2))) from caseprogress,engineerprogress,PATENT,staff S1 where cp09=ep02(+) and CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp14=S1.ST01(+) " & strSQL1 + StrSQL6 & "  AND CP10 NOT IN ('103','105','125') " & " AND CP01 IN (" & stPA & ") and ((cp158>=" & Mid(strSrvDate(1), 1, 6) & "01 and cp158<=" & Mid(strSrvDate(1), 1, 6) & "31 )  ) GROUP BY cp14 "
If stTM <> "' '" Then strSql = strSql + " UNION all  SELECT cp14,0,0,0,0,0,0,0,0,Sum(Decode(CP26, Null, 1, 0)),0,0,'" & strUserNum & "',0,0,0,0,0,0,0,sum(decode(cp112,'Y',round(nvl(cp97,0) * nvl(cp98,0) * nvl(cp111,1),2),round(cp97 * cp98,2))) from caseprogress,engineerprogress,trademark,staff S1 where cp09=ep02(+) and CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND cp14=S1.ST01(+) " & strSQL2 + StrSQL6 & "  AND CP01 IN (" & stTM & ") and ((cp158>=" & Mid(strSrvDate(1), 1, 6) & "01 and cp158<=" & Mid(strSrvDate(1), 1, 6) & "31 )  ) GROUP BY cp14 "
If stLA <> "' '" Then strSql = strSql + " UNION all  SELECT cp14,0,0,0,0,0,0,0,0,Sum(Decode(CP26, Null, 1, 0)),0,0,'" & strUserNum & "',0,0,0,0,0,0,0,sum(decode(cp112,'Y',round(nvl(cp97,0) * nvl(cp98,0) * nvl(cp111,1),2),round(cp97 * cp98,2))) from caseprogress,engineerprogress,LAWCASE,staff S1 where cp09=ep02(+) and CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND cp14=S1.ST01(+) " & StrSQL3 + StrSQL6 & "   AND CP01 IN (" & stLA & ") and ((cp158>=" & Mid(strSrvDate(1), 1, 6) & "01 and cp158<=" & Mid(strSrvDate(1), 1, 6) & "31 )  ) GROUP BY cp14 "
If stHC <> "' '" Then strSql = strSql + " UNION all  SELECT cp14,0,0,0,0,0,0,0,0,Sum(Decode(CP26, Null, 1, 0)),0,0,'" & strUserNum & "',0,0,0,0,0,0,0,sum(decode(cp112,'Y',round(nvl(cp97,0) * nvl(cp98,0) * nvl(cp111,1),2),round(cp97 * cp98,2))) from caseprogress,engineerprogress,HIRECASE,staff S1 where cp09=ep02(+) and CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND cp14=S1.ST01(+) " & StrSQL4 + StrSQL6 & "   AND CP01 IN (" & stHC & ") and ((cp158>=" & Mid(strSrvDate(1), 1, 6) & "01 and cp158<=" & Mid(strSrvDate(1), 1, 6) & "31 )  ) GROUP BY cp14 "
If stSP <> "' '" Then strSql = strSql + " UNION all  SELECT cp14,0,0,0,0,0,0,0,0,Sum(Decode(CP26, Null, 1, 0)),0,0,'" & strUserNum & "',0,0,0,0,0,0,0,sum(decode(cp112,'Y',round(nvl(cp97,0) * nvl(cp98,0) * nvl(cp111,1),2),round(cp97 * cp98,2))) from caseprogress,engineerprogress,SERVICEPRACTICE,staff S1 where cp09=ep02(+) and CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND cp14=S1.ST01(+) " & strSQL5 + StrSQL6 & "   AND CP01 IN (" & stSP & ") and ((cp158>=" & Mid(strSrvDate(1), 1, 6) & "01 and cp158<=" & Mid(strSrvDate(1), 1, 6) & "31 )  ) GROUP BY cp14 "

'發文點數 1 設計只有專利有/
'Add by Morgan 2011/5/2 不知為何畫面有但原程式沒有
'Modify by Morgan 2011/5/30 若有建點數分配資料時點數改分配點數(目前有225提供書狀意見及226配合開庭)
'strSQL = strSQL + " UNION all  SELECT cp14,0,0,0,0,0,0,0,0,0,SUM(CP18),0,'" & strUserNum & "',0,0,0,0,0,0,0,0 from caseprogress,engineerprogress,PATENT,staff S1 where cp09=ep02(+) and CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp14=S1.ST01(+) " & strSQL1 + StrSQL6 & " AND CP10 IN ('103','105') " & " AND CP01 IN (" & stPA & ") and ((cp27>=" & Mid(strSrvDate(1), 1, 6) & "01 and cp27<=" & Mid(strSrvDate(1), 1, 6) & "31 )  ) GROUP BY cp14 "
strSql = strSql + " UNION all  SELECT cp14,0,0,0,0,0,0,0,0,0,SUM(nvl(a0n03/1000,cp18)),0,'" & strUserNum & "',0,0,0,0,0,0,0,0 from caseprogress,engineerprogress,PATENT,staff S1 ,acc0n0 where a0n02(+)=cp09 and cp09=ep02(+) and CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp14=S1.ST01(+) " & strSQL1 + StrSQL6 & " AND CP10 IN ('103','105','125') " & " AND CP01 IN (" & stPA & ") and ((cp158>=" & Mid(strSrvDate(1), 1, 6) & "01 and cp158<=" & Mid(strSrvDate(1), 1, 6) & "31 )  ) GROUP BY cp14 "

'發文點數 2
'Modify by Morgan 2011/6/1 若有建點數分配資料時點數改分配點數(目前有225提供書狀意見及226配合開庭)
'strSQL = strSQL + " UNION all  SELECT cp14,0,0,0,0,0,0,0,0,0,0,SUM(CP18),'" & strUserNum & "',0,0,0,0,0,0,0,0 from caseprogress,engineerprogress,PATENT,staff S1 where cp09=ep02(+) and CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp14=S1.ST01(+) " & strSQL1 + StrSQL6 & " AND CP10 NOT IN ('103','105') " & " AND CP01 IN (" & stPA & ") and ((cp27>=" & Mid(strSrvDate(1), 1, 6) & "01 and cp27<=" & Mid(strSrvDate(1), 1, 6) & "31 )  ) GROUP BY cp14 "
'If stTM <> "' '" Then strSQL = strSQL + " UNION all  SELECT cp14,0,0,0,0,0,0,0,0,0,0,SUM(CP18),'" & strUserNum & "',0,0,0,0,0,0,0,0 from caseprogress,engineerprogress,trademark,staff S1 where cp09=ep02(+) and CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND cp14=S1.ST01(+) " & strSQL2 + StrSQL6 & "  AND CP01 IN (" & stTM & ") and ((cp27>=" & Mid(strSrvDate(1), 1, 6) & "01 and cp27<=" & Mid(strSrvDate(1), 1, 6) & "31 )  ) GROUP BY cp14 "
'If stLA <> "' '" Then strSQL = strSQL + " UNION all  SELECT cp14,0,0,0,0,0,0,0,0,0,0,SUM(CP18),'" & strUserNum & "',0,0,0,0,0,0,0,0 from caseprogress,engineerprogress,LAWCASE,staff S1 where cp09=ep02(+) and CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND cp14=S1.ST01(+) " & StrSQL3 + StrSQL6 & "  AND CP01 IN (" & stLA & ") and ((cp27>=" & Mid(strSrvDate(1), 1, 6) & "01 and cp27<=" & Mid(strSrvDate(1), 1, 6) & "31 )  ) GROUP BY cp14 "
'If stHC <> "' '" Then strSQL = strSQL + " UNION all  SELECT cp14,0,0,0,0,0,0,0,0,0,0,SUM(CP18),'" & strUserNum & "',0,0,0,0,0,0,0,0 from caseprogress,engineerprogress,HIRECASE,staff S1 where cp09=ep02(+) and CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND cp14=S1.ST01(+) " & StrSQL4 + StrSQL6 & "  AND CP01 IN (" & stHC & ") and ((cp27>=" & Mid(strSrvDate(1), 1, 6) & "01 and cp27<=" & Mid(strSrvDate(1), 1, 6) & "31 )  ) GROUP BY cp14 "
'If stSP <> "' '" Then strSQL = strSQL + " UNION all  SELECT cp14,0,0,0,0,0,0,0,0,0,0,SUM(CP18),'" & strUserNum & "',0,0,0,0,0,0,0,0 from caseprogress,engineerprogress,SERVICEPRACTICE,staff S1 where cp09=ep02(+) and CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND cp14=S1.ST01(+) " & strSQL5 + StrSQL6 & "  AND CP01 IN (" & stSP & ") and ((cp27>=" & Mid(strSrvDate(1), 1, 6) & "01 and cp27<=" & Mid(strSrvDate(1), 1, 6) & "31 ) ) GROUP BY cp14 "
strSql = strSql + " UNION all  SELECT cp14,0,0,0,0,0,0,0,0,0,0,SUM(nvl(a0n03/1000,cp18)),'" & strUserNum & "',0,0,0,0,0,0,0,0 from caseprogress,engineerprogress,PATENT,staff S1 ,acc0n0 where a0n02(+)=cp09 and cp09=ep02(+) and CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp14=S1.ST01(+) " & strSQL1 + StrSQL6 & " AND CP10 NOT IN ('103','105','125') " & " AND CP01 IN (" & stPA & ") and ((cp158>=" & Mid(strSrvDate(1), 1, 6) & "01 and cp158<=" & Mid(strSrvDate(1), 1, 6) & "31 )  ) GROUP BY cp14 "
If stTM <> "' '" Then strSql = strSql + " UNION all  SELECT cp14,0,0,0,0,0,0,0,0,0,0,SUM(nvl(a0n03/1000,cp18)),'" & strUserNum & "',0,0,0,0,0,0,0,0 from caseprogress,engineerprogress,trademark,staff S1 ,acc0n0 where a0n02(+)=cp09 and cp09=ep02(+) and CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND cp14=S1.ST01(+) " & strSQL2 + StrSQL6 & "  AND CP01 IN (" & stTM & ") and ((cp158>=" & Mid(strSrvDate(1), 1, 6) & "01 and cp158<=" & Mid(strSrvDate(1), 1, 6) & "31 )  ) GROUP BY cp14 "
If stLA <> "' '" Then strSql = strSql + " UNION all  SELECT cp14,0,0,0,0,0,0,0,0,0,0,SUM(nvl(a0n03/1000,cp18)),'" & strUserNum & "',0,0,0,0,0,0,0,0 from caseprogress,engineerprogress,LAWCASE,staff S1 ,acc0n0 where a0n02(+)=cp09 and cp09=ep02(+) and CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND cp14=S1.ST01(+) " & StrSQL3 + StrSQL6 & "  AND CP01 IN (" & stLA & ") and ((cp158>=" & Mid(strSrvDate(1), 1, 6) & "01 and cp158<=" & Mid(strSrvDate(1), 1, 6) & "31 )  ) GROUP BY cp14 "
If stHC <> "' '" Then strSql = strSql + " UNION all  SELECT cp14,0,0,0,0,0,0,0,0,0,0,SUM(nvl(a0n03/1000,cp18)),'" & strUserNum & "',0,0,0,0,0,0,0,0 from caseprogress,engineerprogress,HIRECASE,staff S1 ,acc0n0 where a0n02(+)=cp09 and cp09=ep02(+) and CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND cp14=S1.ST01(+) " & StrSQL4 + StrSQL6 & "  AND CP01 IN (" & stHC & ") and ((cp158>=" & Mid(strSrvDate(1), 1, 6) & "01 and cp158<=" & Mid(strSrvDate(1), 1, 6) & "31 )  ) GROUP BY cp14 "
If stSP <> "' '" Then strSql = strSql + " UNION all  SELECT cp14,0,0,0,0,0,0,0,0,0,0,SUM(nvl(a0n03/1000,cp18)),'" & strUserNum & "',0,0,0,0,0,0,0,0 from caseprogress,engineerprogress,SERVICEPRACTICE,staff S1 ,acc0n0 where a0n02(+)=cp09 and cp09=ep02(+) and CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND cp14=S1.ST01(+) " & strSQL5 + StrSQL6 & "  AND CP01 IN (" & stSP & ") and ((cp158>=" & Mid(strSrvDate(1), 1, 6) & "01 and cp158<=" & Mid(strSrvDate(1), 1, 6) & "31 ) ) GROUP BY cp14 "
'end 2011/6/1
'end 2011/5/2

'Modify by Morgan 2010/6/23 新增新制計算欄位
'cnnConnection.Execute " INSERT INTO R090609_1 " & strSql
'Modified by Lydia 2017/12/25 O8速度變慢,拆成多筆執行(106/10/25 承辦人工作量查詢變慢，王副總說最近要跑半小時)
'strSql = " INSERT INTO R090609_1 (R103001,R103002,R103003,R103004,R103005,R103006,R103007" & _
'   ",R103008,R103009,R103010,R103011,R103012,ID,R103013,R103014,R103015,R103016,R103017" & _
'   ",R103018,R103019,R103020)" & strSql
'cnnConnection.Execute strSql
strExc(1) = UCase(strSql)
tmpArr = Split(strExc(1), "UNION ALL")
For intQ = 0 To UBound(tmpArr)
    If Trim(tmpArr(intQ)) <> "" Then
        strExc(2) = " INSERT INTO R090609_1 (R103001,R103002,R103003,R103004,R103005,R103006,R103007" & _
           ",R103008,R103009,R103010,R103011,R103012,ID,R103013,R103014,R103015,R103016,R103017" & _
           ",R103018,R103019,R103020)" & Trim(tmpArr(intQ))
        cnnConnection.Execute strExc(2), intI
    End If
Next intQ
'end 2017/12/25

'Modified by Lydia 2015/01/21 所有在職的人員都要列出, 不管當時是否有符合條件的資料
'Modified by Morgan 2025/1/6 排除已有資料的語法要加id否則會抓到其他人的暫存資料而漏抓編號
'strExc(0) = "SELECT S1.ST01,0,0,0,0,0,0,0,0,0,0,0,'" & strUserNum & "',0,0,0,0,0,0,0,0 from staff S1 where substr(S1.st01,1,1) between '6' and 'F' " & _
          "and S1.st01 not in (select r103001 from r090609_1) " & StrSQL7
strExc(0) = "SELECT S1.ST01,0,0,0,0,0,0,0,0,0,0,0,'" & strUserNum & "',0,0,0,0,0,0,0,0 from staff S1 where substr(S1.st01,1,1) between '6' and 'F' " & _
          "and S1.st01 not in (select r103001 from r090609_1 where ID='" & strUserNum & "') " & StrSQL7
'end 2025/1/6
strSql = " INSERT INTO R090609_1 (R103001,R103002,R103003,R103004,R103005,R103006,R103007" & _
   ",R103008,R103009,R103010,R103011,R103012,ID,R103013,R103014,R103015,R103016,R103017" & _
   ",R103018,R103019,R103020)" & strExc(0)
cnnConnection.Execute strSql

'逾法定期限
'若己閉卷, 則在本所案號後加"*"號
strSql = "SELECT cp14,EP01,SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)),CP26,DECODE(PA09,'000',PTM03,PTM04),decode(pa09,'000',cpm03,cpm04)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",NVL(S3.ST02,EP04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",0,EP12,NVL(S2.ST02,CP13),CP09,'" & strUserNum & "' FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP WHERE cp09=ep02(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) and ((cp158>=" & Mid(strSrvDate(1), 1, 6) & "01 and cp158<=" & Mid(strSrvDate(1), 1, 6) & "31 ) or cp158=0 ) " & strSQL1 & StrSQL6 & _
         " AND CP01 IN (" & stPA & ") " & _
         " and cp06>=" & ChangeWDateStringToWString(DateAdd("m", -6, ChangeWStringToWDateString(strSrvDate(1)))) & " and cp06<=" & strSrvDate(1) & " "
If stTM <> "' '" Then strSql = strSql + " UNION all  SELECT cp14,EP01,SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊',''),NVL(TM05,NVL(TM06,TM07)),CP26,decode(tm10,'000',ptm03,ptm04),decode(tm10,'000',cpm03,cpm04)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",NVL(S3.ST02,EP04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",0,EP12,NVL(S2.ST02,CP13),CP09,'" & strUserNum & "' FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,trADEMARK,CASEPROPERTYMAP,PATENTTRADEMARKMAP WHERE cp09=ep02(+) AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND cp14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '2'=PTM01(+) AND tm08=PTM02(+) and ((cp158>=" & Mid(strSrvDate(1), 1, 6) & "01 and cp158<=" & Mid(strSrvDate(1), 1, 6) & "31 ) or cp158=0 ) " & strSQL2 & StrSQL6 & _
         " AND CP01 IN (" & stTM & ") " & _
         " and cp06>=" & ChangeWDateStringToWString(DateAdd("m", -6, ChangeWStringToWDateString(strSrvDate(1)))) & " and cp06<=" & strSrvDate(1) & " "
If stLA <> "' '" Then strSql = strSql + " UNION all  SELECT cp14,EP01,SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊',''),NVL(LC05,NVL(LC06,LC07)),CP26,'',decode(lc15,'000',cpm03,cpm04)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",NVL(S3.ST02,EP04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",0,EP12,NVL(S2.ST02,CP13),CP09,'" & strUserNum & "' FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,LAWCASE,CASEPROPERTYMAP WHERE cp09=ep02(+) AND CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND cp14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) and ((cp158>=" & Mid(strSrvDate(1), 1, 6) & "01 and cp158<=" & Mid(strSrvDate(1), 1, 6) & "31 ) or cp158=0 ) " & StrSQL3 & StrSQL6 & " AND CP01 IN (" & stLA & ") " & _
         " and cp06>=" & ChangeWDateStringToWString(DateAdd("m", -6, ChangeWStringToWDateString(strSrvDate(1)))) & " and cp06<=" & strSrvDate(1) & " "
If stHC <> "' '" Then strSql = strSql + " UNION all  SELECT cp14,EP01,SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊',''),HC06,CP26,'',CPM03," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",NVL(S3.ST02,EP04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",0,EP12,NVL(S2.ST02,CP13),CP09,'" & strUserNum & "' FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,HIRECASE,CASEPROPERTYMAP WHERE cp09=ep02(+) AND CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND cp14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) and ((cp158>=" & Mid(strSrvDate(1), 1, 6) & "01 and cp158<=" & Mid(strSrvDate(1), 1, 6) & "31 ) or cp158=0 ) " & StrSQL4 & StrSQL6 & "  AND CP01 IN (" & stHC & ") " & _
         " and cp06>=" & ChangeWDateStringToWString(DateAdd("m", -6, ChangeWStringToWDateString(strSrvDate(1)))) & " and cp06<=" & strSrvDate(1) & " "
If stSP <> "' '" Then strSql = strSql + " UNION all  SELECT cp14,EP01,SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊',''),NVL(SP05,NVL(SP06,SP07)),CP26,'',decode(sp09,'000',cpm03,cpm04)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",NVL(S3.ST02,EP04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",0,EP12,NVL(S2.ST02,CP13),CP09,'" & strUserNum & "' FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,SERVICEPRACTICE,CASEPROPERTYMAP WHERE cp09=ep02(+) AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND cp14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) and ((cp158>=" & Mid(strSrvDate(1), 1, 6) & "01 and cp158<=" & Mid(strSrvDate(1), 1, 6) & "31 ) or cp158=0 ) " & strSQL5 & StrSQL6 & " AND CP01 IN (" & stSP & ") " & _
         " and cp06>=" & ChangeWDateStringToWString(DateAdd("m", -6, ChangeWStringToWDateString(strSrvDate(1)))) & " and cp06<=" & strSrvDate(1) & " "
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        DoEvents
        Do While .EOF = False
            For i = 0 To 21
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            '計算承辦天數
            If Len(strTemp(14)) <> 0 And Len(strTemp(12)) <> 0 And Val(strTemp(14)) <> 0 And Val(strTemp(12)) <> 0 Then
                 strTemp(18) = Trim(str(GetWorkDay(ChangeTStringToWString(ChangeTDateStringToTString(strTemp(14))), ChangeTStringToWString(ChangeTDateStringToTString(strTemp(12))))))
            Else
                If Len(strTemp(13)) <> 0 And Len(strTemp(12)) <> 0 And Val(strTemp(13)) <> 0 And Val(strTemp(12)) <> 0 Then
                    strTemp(18) = Trim(str(GetWorkDay(ChangeTStringToWString(ChangeTDateStringToTString(strTemp(13))), ChangeTStringToWString(ChangeTDateStringToTString(strTemp(12))))))
                End If
            End If
            '92.04.16 nick 20 20 單引號問題
            'strSQL = "INSERT INTO R090609_2 VALUES ('" & strTemp(0) & "'," & Val(strTemp(1)) & ",'" & strTemp(2) & "','" & strTemp(3) & "','" & strTemp(4) & "','" & strTemp(5) & "','" & strTemp(6) & "','" & strTemp(7) & "','" & strTemp(8) & "','" & strTemp(9) & "','" & strTemp(10) & "','" & strTemp(11) & "','" & strTemp(12) & "','" & strTemp(13) & "','" & strTemp(14) & "','" & strTemp(15) & "','" & strTemp(16) & "','" & strTemp(17) & "'," & Val(strTemp(18)) & ",'" & strTemp(19) & "','" & strTemp(20) & "','" & strTemp(21) & "','" & strUserNum & "') "
            strSql = "INSERT INTO R090609_2 VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp(1)) & ",'" & ChgSQL(strTemp(2)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & ChgSQL(strTemp(8)) & "','" & ChgSQL(strTemp(9)) & "','" & ChgSQL(strTemp(10)) & "','" & ChgSQL(strTemp(11)) & "','" & ChgSQL(strTemp(12)) & "','" & ChgSQL(strTemp(13)) & "','" & ChgSQL(strTemp(14)) & "','" & ChgSQL(strTemp(15)) & "','" & ChgSQL(strTemp(16)) & "','" & ChgSQL(strTemp(17)) & "'," & Val(strTemp(18)) & ",'" & ChgSQL(strTemp(19)) & "','" & ChgSQL(strTemp(20)) & "','" & ChgSQL(strTemp(21)) & "','" & strUserNum & "') "
            cnnConnection.Execute strSql
            .MoveNext
            DoEvents
        Loop
    End If
End With
CheckOC
strSql = "select sum(r103002) from r090609_1 where id='" & strUserNum & "' "
CheckOC
TestOk = False
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        If Val(CheckStr(.Fields(0))) <> 0 Then
            TestOk = True
        End If
    End If
End With
CheckOC

If TestOk Then
    ObjForm = 2
    Me.Hide
    frm090609_2.Show
    
    'Added by Morgan 2024/4/19
    If m_bolRedo = True Then
      frm090609_2.cmdOK(1).Value = True
    End If
    'end 2024/4/19
Else
    ObjForm = 1
    Me.Hide
    frm090609_1.cmdOK(0).Enabled = False
    frm090609_1.Show
End If
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
txt1(0) = Systemkind_g
'2007/5/15 ADD BY SONIA
Select Case Mid(GetStaffDepartment(strUserNum), 1, 2)
   Case "P1"
      txt1(4) = "P10"
      txt1(5) = "P11"
   Case "P2"
      txt1(4) = "P20"
      txt1(5) = "P21"
   Case Else
      txt1(4) = ""
      txt1(5) = ""
End Select
'2007/5/15 END
bol911001checkRange = True

LBL1.Caption = "" 'Added by Lydia 2022/01/12
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090609 = Nothing
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
Case 0 '系統類別
      'Add By Cheng 2002/01/07
      Me.txt1(Index).Text = GetAllSysKind(Me.txt1(Index))
     strTemp1 = Split(UCase(Systemkind_g), ",")
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
Case 3
     LBL1 = GetPrjSalesNM(txt1(3))
     If Trim(txt1(Index)) <> "" Then
        If Trim(LBL1.Caption) = "" Then
            s = MsgBox("承辦人輸入錯誤！", , "錯誤！")
            txt1(Index).SetFocus
            txt1_GotFocus (Index)
            Exit Sub
        End If
     End If
Case 6
     bol911001checkRange = True
     Select Case Trim(txt1(6))
     Case "1", "2", "3", "4", "5", ""
     Case Else
          s = MsgBox("所別只能輸入 1 到 5 !!", , "USER 輸入錯誤")
          txt1(6).SetFocus
          txt1(6).SelStart = 0
          txt1(6).SelLength = Len(txt1(6))
          bol911001checkRange = False
          Exit Sub
     End Select
Case 7
     If bol911001checkRange = True Then
          Select Case Trim(txt1(7))
          Case "1", "2", "3", "4", "5", ""
          Case Else
               s = MsgBox("所別只能輸入 1 到 5 !!", , "USER 輸入錯誤")
               txt1(7).SetFocus
               txt1(7).SelStart = 0
               txt1(7).SelLength = Len(txt1(7))
               Exit Sub
          End Select
        If RunNick(txt1(Index - 1), txt1(Index)) Then
            txt1(Index - 1).SetFocus
            txt1_GotFocus (Index - 1)
            Exit Sub
        End If
     End If
     bol911001checkRange = True
Case 2, 5
   If RunNick(txt1(Index - 1), txt1(Index)) Then
       txt1(Index - 1).SetFocus
       txt1_GotFocus (Index - 1)
       Exit Sub
   End If
Case Else
End Select
End Sub
