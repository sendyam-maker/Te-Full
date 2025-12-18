VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090702 
   BorderStyle     =   1  '單線固定
   Caption         =   "繪圖人員工作量查詢"
   ClientHeight    =   2310
   ClientLeft      =   1380
   ClientTop       =   1590
   ClientWidth     =   4200
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   4200
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   0
      Left            =   1104
      TabIndex        =   0
      Top             =   570
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   1
      Left            =   1104
      MaxLength       =   4
      TabIndex        =   1
      Top             =   955
      Width           =   705
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   2
      Left            =   1908
      MaxLength       =   4
      TabIndex        =   2
      Top             =   955
      Width           =   705
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   3
      Left            =   1104
      MaxLength       =   6
      TabIndex        =   3
      Top             =   1340
      Width           =   855
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   4
      Left            =   1104
      MaxLength       =   1
      TabIndex        =   4
      Top             =   1725
      Width           =   330
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   5
      Left            =   1608
      MaxLength       =   1
      TabIndex        =   5
      Top             =   1725
      Width           =   300
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2088
      TabIndex        =   6
      Top             =   20
      Width           =   756
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Index           =   1
      Left            =   2868
      TabIndex        =   7
      Top             =   20
      Width           =   1200
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Left            =   2010
      TabIndex        =   13
      Top             =   1363
      Width           =   1920
      VariousPropertyBits=   27
      Size            =   "3387;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   12
      Top             =   630
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家："
      Height          =   180
      Index           =   1
      Left            =   180
      TabIndex        =   11
      Top             =   1015
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "繪圖人員："
      Height          =   180
      Index           =   2
      Left            =   180
      TabIndex        =   10
      Top             =   1400
      Width           =   1050
   End
   Begin VB.Label Label1 
      Caption         =   "所別："
      Height          =   180
      Index           =   4
      Left            =   180
      TabIndex        =   9
      Top             =   1740
      Width           =   660
   End
   Begin VB.Line Line1 
      X1              =   1515
      X2              =   2145
      Y1              =   1098
      Y2              =   1098
   End
   Begin VB.Line Line3 
      X1              =   1275
      X2              =   1740
      Y1              =   1860
      Y2              =   1860
   End
   Begin VB.Label Label1 
      Caption         =   "(1.北 2.中 3.南 4.高 5.其他)"
      Height          =   180
      Index           =   7
      Left            =   1950
      TabIndex        =   8
      Top             =   1740
      Width           =   2130
   End
End
Attribute VB_Name = "frm090702"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/07 改成Form2.0 ; lbl1
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/17 日期欄已修改
Option Explicit
Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, s As Integer, SavDay1 As String, SavDay2 As String, SavDay3 As String, StrSQL7 As String, StrSQL4 As String, strSQL5 As String
Dim strSQL2 As String, iPrint As Integer, Page As Integer, strTemp(0 To 21) As String, strTemp3 As String, TestOk As Boolean, StrTemp7(0 To 13) As String, StrTemp99(0 To 7) As String, k As Integer
Dim PLeft(0 To 21) As Integer, strTemp1 As Variant, strTemp2 As Variant, StrSQL6 As String, Str020401SysKind As String, PLeft1(1 To 9) As Integer, SeekPrint As Integer, SeekPrintL As Integer
Public ObjForm As Integer

Private Sub cmdOK_Click(Index As Integer)
Select Case Index
Case 0
     'If Len(txt1(0)) = 0 Then
     '    S = MsgBox("系統類別不可空白!!", , "USER 輸入錯誤")
     '    txt1(0).SetFocus
     '    Exit Sub
     'Else
         Screen.MousePointer = vbHourglass
         Me.Enabled = False
         ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/20 清除查詢印表記錄檔欄位
         Process
         Me.Enabled = True
         Screen.MousePointer = vbDefault
     'End If
Case 1
     Unload Me
Case Else
End Select
End Sub

Sub Process()
Dim NickRS As New ADODB.Recordset
Dim strBeginDate As String
Dim strEndDate As String
Dim strInsert As String
Dim strFromDate As String
Dim str1stDate As String
Dim strLog As String

strInsert = " INSERT INTO R090702_1(R103001,R103002,R103003,R103004,R103005,R103006,R103007" & _
   ",R103008,R103009,R103010,ID,R103011,R103012,R103013,R103014,R103015,R103016,R103017,R103018,R103019,R103020) "
'本月1號
str1stDate = Left(strSrvDate(1), 6) & "01"
'所限起始日,可辦量齊備日起始日
strFromDate = CompDate(1, -3, str1stDate)

cnnConnection.Execute "DELETE FROM R090702_1 WHERE ID='" & strUserNum & "' "
cnnConnection.Execute "DELETE FROM R090702_2 WHERE ID='" & strUserNum & "' "

strSQL1 = ""
StrSQL6 = ""
'If Len(txt1(0)) <> 0 Then
'   StrSQL1 = StrSQL1 + " and CP01 in (" & SQLGrpStr(txt1(0), 1) & ") "
'End If

StrSQL6 = "": StrSQL7 = ""
If Len(txt1(1)) <> 0 Then
    strSQL1 = strSQL1 + " AND PA09>='" & txt1(1) & "' "
End If
If Len(txt1(2)) <> 0 Then
    strSQL1 = strSQL1 + " AND PA09<='" & txt1(2) & "' "
End If
If Len(txt1(1)) <> 0 Or Len(txt1(2)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(1) & txt1(1) & "-" & txt1(2) 'Add By Sindy 2010/12/20
End If
If Len(txt1(3)) <> 0 Then
    'Modified by Morgan 2016/4/26 +CP29
    StrSQL6 = StrSQL6 + " AND EP13='" & txt1(3) & "' AND CP29='" & txt1(3) & "' "
    StrSQL7 = StrSQL7 + " AND SH02='" & txt1(3) & "' "
    pub_QL05 = pub_QL05 & ";" & Label1(2) & txt1(3) & lbl1 'Add By Sindy 2010/12/20
End If
If Len(txt1(4)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND S1.ST06>='" & txt1(4) & "' "
    StrSQL7 = StrSQL7 + " AND S1.ST06>='" & txt1(4) & "' "
End If
If Len(txt1(5)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND S1.ST06<='" & txt1(5) & "' "
    StrSQL7 = StrSQL7 + " AND S1.ST06<='" & txt1(5) & "' "
End If
If Len(txt1(4)) <> 0 Or Len(txt1(5)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(4) & txt1(4) & "-" & txt1(5) & Label1(7) 'Add By Sindy 2010/12/20
End If
strBeginDate = ChangeWDateStringToWString(DateAdd("m", -1, ChangeWStringToWDateString(Left(strSrvDate(1), 6) & "01")))
strEndDate = strSrvDate(1)

StrSQL6 = StrSQL6 & " and ((cp21='Y' and (ep20 is null or ep29 is null)) or cp21 is null)  "

'逾時件數
'Modified by Morgan 2016/4/26 1.加新制欄位,2.加未完稿或當月完稿條件減少回傳筆數可縮短執行時間
'Modified by Morgan 2016/4/29 改只抓當月完稿的案件來統計--翔龍,游經理
'strSql = "SELECT S1.ST01, CP10, EP14, '1', EP15, CP07, CP27, CP09, CP57,CP100,CP101 FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT, Staff S1 WHERE CP01 IN ('P','CFP','FCP') AND EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP13=S1.ST01(+) And EP20 Is Null And EP13 Is Not Null And (EP14>=" & strBeginDate & " And EP14<=" & strEndDate & ") and nvl(ep15," & strSrvDate(1) & ")>=" & str1stDate & strSQL1 & StrSQL6
'strSql = strSql & " Union SELECT S1.ST01, CP10, Nvl(EP17, EP08), '2', EP18, CP07, CP27, CP09, CP57,CP103,CP104 FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT, Staff S1 WHERE CP01 IN ('P','CFP','FCP') AND EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP13=S1.ST01(+) And EP29 Is Null And EP13 Is Not Null And (EP17>=" & strBeginDate & " And EP17<=" & strEndDate & ") and nvl(ep18," & strSrvDate(1) & ")>=" & str1stDate & strSQL1 & StrSQL6
strSql = "SELECT S1.ST01, CP10, EP14, '1', EP15, CP07, CP27, CP09, CP57,CP100,CP101 FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT, Staff S1 WHERE CP01 IN ('P','CFP','FCP') AND EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP13=S1.ST01(+) And EP20 Is Null And EP13 Is Not Null And EP15>=" & str1stDate & strSQL1 & StrSQL6
strSql = strSql & " Union SELECT S1.ST01, CP10, Nvl(EP17, EP08), '2', EP18, CP07, CP27, CP09, CP57,CP103,CP104 FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT, Staff S1 WHERE CP01 IN ('P','CFP','FCP') AND EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP13=S1.ST01(+) And EP29 Is Null And EP13 Is Not Null And EP18>=" & str1stDate & str1stDate & strSQL1 & StrSQL6
'End
If NickRS.State = 1 Then NickRS.Close
NickRS.CursorLocation = adUseClient
NickRS.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly

If NickRS.RecordCount <> 0 Then
    NickRS.MoveFirst
    Do While NickRS.EOF = False
            Select Case CheckStr(NickRS.Fields(1))
            Case "103", "105" '設計申請
                If CheckStr(NickRS.Fields(3)) = "1" Then  '草圖
                    '若有草齊日及草完日
                    If "" & NickRS.Fields(2).Value <> "" And "" & NickRS.Fields(4).Value <> "" Then
                        '若草完日與系統日同月份
                        If Left(Val(DBDATE("" & NickRS.Fields(4).Value)), 6) = Left(strSrvDate(1), 6) Then
                            If GetWorkDay(CheckStr(NickRS.Fields(4)), "" & NickRS.Fields(2).Value) > 5 Then
                                strSql = " INSERT INTO R090702_1(R103001,R103002,R103003,R103004,R103005,R103006,R103007" & _
                                    ",R103008,R103009,R103010,ID) Values('" & NickRS.Fields(0).Value & "', 1, 0, 0, 0, 0, 0, 0, 0, 0,'" & strUserNum & "' ) "
                                cnnConnection.Execute strSql, intI
                            End If
                        End If
                    '若有草齊日無草完日
                    ElseIf "" & NickRS.Fields(2).Value <> "" And "" & NickRS.Fields(4).Value = "" Then
                        If GetWorkDay(strSrvDate(1), CheckStr(NickRS.Fields(2))) > 5 Then
                           strSql = " INSERT INTO R090702_1(R103001,R103002,R103003,R103004,R103005,R103006,R103007" & _
                                    ",R103008,R103009,R103010,ID) Values('" & NickRS.Fields(0).Value & "', 1, 0, 0, 0, 0, 0, 0, 0, 0,'" & strUserNum & "' ) "
                           cnnConnection.Execute strSql, intI
                        End If
                    End If
                Else '墨圖
                    '若有墨齊日及墨完日
                    If "" & NickRS.Fields(2).Value <> "" And "" & NickRS.Fields(4).Value <> "" Then
                        '若墨完日與系統日同月份
                        If Left(Val(DBDATE("" & NickRS.Fields(4).Value)), 6) = Left(strSrvDate(1), 6) Then
                            If GetWorkDay(CheckStr(NickRS.Fields(4)), "" & NickRS.Fields(2).Value) > 3 Then
                              strSql = " INSERT INTO R090702_1(R103001,R103002,R103003,R103004,R103005,R103006,R103007" & _
                                    ",R103008,R103009,R103010,ID) Values('" & NickRS.Fields(0).Value & "', 0, 1, 0, 0, 0, 0, 0, 0, 0,'" & strUserNum & "' ) "
                              cnnConnection.Execute strSql, intI
                            End If
                        End If
                    '若有墨齊日無墨完日
                    ElseIf "" & NickRS.Fields(2).Value <> "" And "" & NickRS.Fields(4).Value = "" Then
                        If GetWorkDay(strSrvDate(1), CheckStr(NickRS.Fields(2))) > 3 Then
                           strSql = " INSERT INTO R090702_1(R103001,R103002,R103003,R103004,R103005,R103006,R103007" & _
                                    ",R103008,R103009,R103010,ID) Values('" & NickRS.Fields(0).Value & "', 0, 1, 0, 0, 0, 0, 0, 0, 0,'" & strUserNum & "' ) "
                           cnnConnection.Execute strSql, intI
                        End If
                    End If
                End If
            Case Else
                If CheckStr(NickRS.Fields(3)) = "1" Then  '草圖
                    '若有草齊日及草完日
                    If "" & NickRS.Fields(2).Value <> "" And "" & NickRS.Fields(4).Value <> "" Then
                        '若草完日與系統日同月份
                        If Left(Val(DBDATE("" & NickRS.Fields(4).Value)), 6) = Left(strSrvDate(1), 6) Then
                            If GetWorkDay(CheckStr(NickRS.Fields(4)), "" & NickRS.Fields(2).Value) > 4 Then
                              strSql = " INSERT INTO R090702_1(R103001,R103002,R103003,R103004,R103005,R103006,R103007" & _
                                    ",R103008,R103009,R103010,ID) Values('" & NickRS.Fields(0).Value & "', 1, 0, 0, 0, 0, 0, 0, 0, 0,'" & strUserNum & "' ) "
                              cnnConnection.Execute strSql, intI
                            End If
                        End If
                    '若有草齊日無草完日
                    ElseIf "" & NickRS.Fields(2).Value <> "" And "" & NickRS.Fields(4).Value = "" Then
                        If GetWorkDay(strSrvDate(1), CheckStr(NickRS.Fields(2))) > 4 Then
                           strSql = " INSERT INTO R090702_1(R103001,R103002,R103003,R103004,R103005,R103006,R103007" & _
                                    ",R103008,R103009,R103010,ID) Values('" & NickRS.Fields(0).Value & "', 1, 0, 0, 0, 0, 0, 0, 0, 0,'" & strUserNum & "' ) "
                           cnnConnection.Execute strSql, intI
                        End If
                    End If
                Else '墨圖
                    '若有墨齊日及墨完日
                    If "" & NickRS.Fields(2).Value <> "" And "" & NickRS.Fields(4).Value <> "" Then
                        '若墨完日與系統日同月份
                        If Left(Val(DBDATE("" & NickRS.Fields(4).Value)), 6) = Left(strSrvDate(1), 6) Then
                            If GetWorkDay((CheckStr(NickRS.Fields(4))), "" & NickRS.Fields(2).Value) > 3 Then
                              strSql = " INSERT INTO R090702_1(R103001,R103002,R103003,R103004,R103005,R103006,R103007" & _
                                    ",R103008,R103009,R103010,ID) Values('" & NickRS.Fields(0).Value & "', 0, 1, 0, 0, 0, 0, 0, 0, 0,'" & strUserNum & "' ) "
                              cnnConnection.Execute strSql, intI
                            End If
                        End If
                    '若有墨齊日無墨完日
                    ElseIf "" & NickRS.Fields(2).Value <> "" And "" & NickRS.Fields(4).Value = "" Then
                        If GetWorkDay(strSrvDate(1), CheckStr(NickRS.Fields(2))) > 3 Then
                           strSql = " INSERT INTO R090702_1(R103001,R103002,R103003,R103004,R103005,R103006,R103007" & _
                                    ",R103008,R103009,R103010,ID) Values('" & NickRS.Fields(0).Value & "', 0, 1, 0, 0, 0, 0, 0, 0, 0,'" & strUserNum & "' ) "
                           cnnConnection.Execute strSql, intI
                        End If
                    End If
                End If
            End Select
        NickRS.MoveNext
    Loop
End If
If NickRS.State = 1 Then NickRS.Close

'承辦量
strSql = strInsert & " SELECT S1.st01,0,0,Sum(1),0,0,0,0,0,0,'" & strUserNum & "',0,sum(cp100*cp101),0,0,0,0,0,0,0,0 from caseprogress,engineerprogress,PATENT,staff S1 where ep02=cp09(+) and cp01=pa01(+) and CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP13=ST01(+) AND ST05 IN ('79','81','82','AC') " & strSQL1 + StrSQL6 & " AND EP15>=" & Left(strSrvDate(1), 6) & "01 and ep20 is null GROUP BY S1.ST01 "
strSql = strSql + " UNION all  SELECT S1.st01,0,0,0,Sum(1),0,0,0,0,0,'" & strUserNum & "',0,0,sum(cp103*cp104),0,0,0,0,0,0,0 from caseprogress,engineerprogress,PATENT,staff S1 where ep02=cp09(+) and cp01=pa01(+) and CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP13=ST01(+) AND ST05 IN ('79','81','82','AC') " & strSQL1 + StrSQL6 & " AND EP18>=" & Left(strSrvDate(1), 6) & "01 and EP29 is null GROUP BY S1.ST01 "
cnnConnection.Execute strSql, intI

'承辦量(支援記錄)
strSql = strInsert & " SELECT S1.st01,0,0,Sum(Nvl(SH05,0)/4),0,0,0,0,0,0,'" & strUserNum & "',0,Sum(Nvl(SH05,0)*0.2),0,0,0,0,0,0,0,0 from SupportHour,PATENT,staff S1 where SH06=pa01(+) and SH07=PA02(+) AND SH08=PA03(+) AND SH09=PA04(+) AND SH02=ST01(+) AND ST05 IN ('79','81','82','AC') " & strSQL1 + StrSQL7 & " AND SH01>=" & Left(strSrvDate(1), 6) & "01 And SH11='V' GROUP BY S1.ST01 "
strSql = strSql + " UNION all  SELECT S1.st01,0,0,0,Sum(Nvl(SH05,0)/4),0,0,0,0,0,'" & strUserNum & "',0,0,Sum(Nvl(SH05,0)*0.2),0,0,0,0,0,0,0 from SupportHour,PATENT,staff S1 where SH06=pa01(+) and SH07=PA02(+) AND SH08=PA03(+) AND SH09=PA04(+) AND SH02=ST01(+) AND ST05 IN ('79','81','82','AC') " & strSQL1 + StrSQL7 & " AND SH01>=" & Left(strSrvDate(1), 6) & "01 And SH11='V' GROUP BY S1.ST01 "
cnnConnection.Execute strSql, intI

'可辦量
'草圖
strSql = strInsert & "SELECT S1.st01,0,0,0,0,Sum(1),0,0,0,0,'" & strUserNum & "',0,0,0,sum(cp100*cp101),0,0,0,0,0,0 from caseprogress,engineerprogress,PATENT,staff S1 where ep02=cp09(+) and cp01=pa01(+) and CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP13=ST01(+) AND ST05 IN ('79','81','82','AC') " & strSQL1 + StrSQL6 & " And CP27||CP57 Is Null And EP14>=" & strFromDate & " And EP15 Is Null And EP20 Is Null GROUP BY S1.ST01 "
cnnConnection.Execute strSql, intI

'墨圖
strSql = strInsert & "SELECT S1.st01,0,0,0,0,0,Sum(1),0,0,0,'" & strUserNum & "',0,0,0,0,sum(cp103*cp104),0,0,0,0,0 from caseprogress,engineerprogress,PATENT,staff S1 where ep02=cp09(+) and cp01=pa01(+) and CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP13=ST01(+) AND ST05 IN ('79','81','82','AC') " & strSQL1 + StrSQL6 & " And CP27||CP57 Is Null And EP17>=" & strFromDate & " And EP18 Is Null And EP29 Is Null GROUP BY S1.ST01 "
cnnConnection.Execute strSql, intI

'修改圖式
'Modified by Morgan 2016/5/4 +墨完
strSql = strInsert & "SELECT S1.st01,0,0,0,0,0,0,0,0,0,'" & strUserNum & "',sum(1),0,0,0,0,0,0,0,0,0 from caseprogress,engineerprogress,empelectronprocess e1,PATENT,staff S1 where ep02=cp09(+) and cp01=pa01(+) and CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP13=ST01(+) AND ST05 IN ('79','81','82','AC') " & strSQL1 + StrSQL6 & " And CP27||CP57 Is Null And EP14>=" & strFromDate & " and eep01(+)=ep02 and eep04='" & EMP_修改圖式 & "' And EP20 Is Null and not exists(select * from empelectronprocess e2 where e2.eep01=e1.eep01 and e2.eep02>e1.eep02 and e2.eep04 in ('" & EMP_草完 & "','" & EMP_標號 & "','" & EMP_墨完 & "')) GROUP BY S1.ST01"
cnnConnection.Execute strSql, intI

'分案量(以草圖計算)
strSql = strInsert & "SELECT S1.st01,0,0,0,0,0,0,Sum(1),0,0,'" & strUserNum & "',0,0,0,0,0,sum(cp100*cp101),0,0,0,0 from caseprogress,engineerprogress,PATENT,staff S1 where ep02=cp09(+) and cp01=pa01(+) and CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP13=ST01(+) AND ST05 IN ('79','81','82','AC') " & strSQL1 + StrSQL6 & " AND EP06>=" & Left(strSrvDate(1), 6) & "01 AND EP20 IS NULL and ep13 is not null " & " GROUP BY S1.ST01 "
cnnConnection.Execute strSql, intI

'發文件數
strSql = strInsert & "SELECT S1.st01,0,0,0,0,0,0,0,Sum(1),0,'" & strUserNum & "',0,0,0,0,0,0,sum(cp103*cp104),0,0,0 from caseprogress,engineerprogress,PATENT,staff S1 where ep02(+)=cp09 and cp01=pa01(+) and CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP13=ST01(+) AND ST05 IN ('79','81','82','AC') " & strSQL1 + StrSQL6 & " AND CP27>=" & Left(strSrvDate(1), 6) & "01 and cp29 is not null And CP57 Is Null GROUP BY S1.ST01 "
cnnConnection.Execute strSql, intI

'發文點數
strSql = strInsert & "SELECT S1.st01,0,0,0,0,0,0,0,0,SUM(CP18),'" & strUserNum & "',0,0,0,0,0,0,0,0,0,0 from caseprogress,engineerprogress,PATENT,staff S1 where ep02(+)=cp09 and cp01=pa01(+) and CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP13=ST01(+) AND ST05 IN ('79','81','82','AC') " & strSQL1 + StrSQL6 & " AND CP27>=" & Left(strSrvDate(1), 6) & "01 and cp29 is not null And CP57 Is Null And EP29 Is Null GROUP BY S1.ST01 "
cnnConnection.Execute strSql, intI

'Modified by Morgan 2016/4/27 新增新制計算欄位
'R103011:修改圖式,R103012:新制草圖承辦量,R103013:新制墨圖承辦量,R103014:新制草圖可辦量,R103015:新制墨圖可辦量,R103016:新制分案量
'cnnConnection.Execute " INSERT INTO R090702_1 " & strSql
'cnnConnection.Execute " INSERT INTO R090702_1(R103001,R103002,R103003,R103004,R103005,R103006,R103007" & _
   ",R103008,R103009,R103010,ID,R103011,R103012,R103013,R103014,R103015,R103016,R103017,R103018,R103019,R103020) " & strSql


'逾本所期限案件明細
'若已閉卷, 則在本所案號後加"*"號
'Modified by Morgan 2016/4/27 +未取消收文
strSql = "SELECT S1.ST01,SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)),CP26,decode(pa09,'000',cpm03,cpm04),s2.st02," & SQLDate("CP48") & ",cp18," & SQLDate("ep14") & "," & SQLDate("eP15") & ",0," & SQLDate("EP17") & "," & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",eP26,s3.st02,CP09 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP WHERE EP02(+)=CP09 AND cp01=pa01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP29=S1.ST01(+) AND S1.ST05 IN ('79','81','82','AC') AND eP05=s2.ST01(+) AND cP13=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND CP57||CP27 IS NULL AND CP06>=" & strFromDate & " AND CP06<=" & Val(strSrvDate(1)) & " and cp29>'6' And (EP15>0 Or EP18 > 0) " & strSQL1 & StrSQL6
CheckOC
TestOk = False
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        k = 0
        'FRM100.Show
        'FRM100.Tag = Trim(str(.RecordCount)) & "=0"
        DoEvents
        Do While .EOF = False
            TestOk = True
            For i = 0 To 20
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            '計算草圖作業天數
            If Len(strTemp(10)) <> 0 And Len(strTemp(11)) <> 0 And Val(strTemp(10)) <> 0 And Val(strTemp(11)) <> 0 Then
                 strTemp(12) = Trim(str(GetWorkDay(ChangeTStringToWString(ChangeTDateStringToTString(strTemp(11))), ChangeTStringToWString(ChangeTDateStringToTString(strTemp(10))))))
            End If
            '計算墨圖作業天數
            If Len(strTemp(14)) <> 0 And Len(strTemp(13)) <> 0 And Val(strTemp(14)) <> 0 And Val(strTemp(13)) <> 0 Then
                 strTemp(15) = Trim(str(GetWorkDay(ChangeTStringToWString(ChangeTDateStringToTString(strTemp(14))), ChangeTStringToWString(ChangeTDateStringToTString(strTemp(13))))))
            End If
            strSql = "INSERT INTO R090702_2 VALUES ('" & strTemp(0) & "','" & strTemp(1) & "','" & strTemp(2) & "','" & strTemp(3) & "','" & strTemp(4) & "','" & strTemp(5) & "','" & strTemp(6) & "','" & strTemp(7) & "','" & strTemp(8) & "','" & strTemp(9) & "','" & strTemp(10) & "','" & strTemp(11) & "','" & strTemp(12) & "','" & strTemp(13) & "','" & strTemp(14) & "','" & strTemp(15) & "','" & strTemp(16) & "','" & strTemp(17) & "'," & Val(strTemp(18)) & ",'" & strTemp(19) & "','" & strTemp(20) & "','" & strUserNum & "') "
            cnnConnection.Execute strSql
            .MoveNext
            k = k + 1
            'FRM100.Tag = Trim(str(.RecordCount)) & "=" & Trim(str(K))
            'FRM100.StrMenu
            DoEvents
        Loop
    End If
    
End With

CheckOC
''UNLOAD FRM100
'Marked By Cheng 2004/03/05
''判斷有逾本所期限資料
''strSQL = "select sum(r103002)+sum(r103003) from r090702_1 where id='" & strUserNum & "' "
'strSQL = "select S1.st01,Sum(Decode(EP29, Null , 1, 0)),0,0,0,0,0,0,0,0,'" & strUserNum & "' from caseprogress,engineerprogress,PATENT,staff S1 where ep02=cp09(+) and cp01=pa01(+) and CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP13=ST01(+) AND ST05 IN ('79','81','82','AC') AND CP27 IS NULL  and (ep15 is not null or ep15<> 0) AND CP06<=" & Val(strSrvDate(1)) & strSQL1 + StrSQL6 & " GROUP BY S1.ST01 "
'strSQL = strSQL + " Union all select S1.st01,0,Sum(Decode(EP29, Null, 1, 0)),0,0,0,0,0,0,0,'" & strUserNum & "' from caseprogress,engineerprogress,PATENT,staff S1 where ep02=cp09(+) and cp01=pa01(+) and CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP13=ST01(+) AND ST05 IN ('79','81','82','AC') AND CP27 IS NULL  AND (ep18 is not null or ep18<> 0) and CP06<=" & Val(strSrvDate(1)) & strSQL1 + StrSQL6 & " GROUP BY S1.ST01 "
'CheckOC
'TestOk = False
'With adoRecordset
'    .CursorLocation = adUseClient
'    .Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'    If .RecordCount <> 0 And .RecordCount > 0 Then
'        If Val(CheckStr(.Fields(1).Value)) + Val("" & .Fields(2).Value) <> 0 Then
'            TestOk = True
'        End If
'    End If
'End With
'CheckOC
'End
If TestOk Then
    ObjForm = 2
    Me.Hide
    frm090702_2.Show
Else
    ObjForm = 1
    Me.Hide
    frm090702_1.cmdok(0).Enabled = False
    frm090702_1.Show
End If
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
txt1(0) = Systemkind_g_P
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090702 = Nothing
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
Case 0 '系統類別
      'Add By Cheng 2002/01/07
      Me.txt1(Index).Text = GetAllSysKind(Me.txt1(Index))
     strTemp1 = Split(UCase(Systemkind_g_P), ",")
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
     lbl1 = GetPrjSales(txt1(3))
Case 4
     Select Case Trim(txt1(4))
     Case "1", "2", "3", "4", "5", ""
     Case Else
          s = MsgBox("所別只能輸入 1 到 5 !!", , "USER 輸入錯誤")
          txt1(4).SetFocus
          txt1(4).SelStart = 0
          txt1(4).SelLength = Len(txt1(4))
          Exit Sub
     End Select
Case 5
     Select Case Trim(txt1(5))
     Case "1", "2", "3", "4", "5", ""
     Case Else
          s = MsgBox("所別只能輸入 1 到 5 !!", , "USER 輸入錯誤")
          txt1(5).SetFocus
          txt1(5).SelStart = 0
          txt1(5).SelLength = Len(txt1(5))
          Exit Sub
     End Select
Case Else
End Select
End Sub

