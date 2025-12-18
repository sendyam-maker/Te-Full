VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmHTAauto 
   Caption         =   "考勤機刷卡紀錄自動接收作業"
   ClientHeight    =   4080
   ClientLeft      =   60
   ClientTop       =   350
   ClientWidth     =   7900
   Icon            =   "frmHTAauto.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4080
   ScaleWidth      =   7900
   StartUpPosition =   3  '系統預設值
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   3660
      Top             =   1770
      _ExtentX        =   988
      _ExtentY        =   988
      _Version        =   393216
   End
   Begin VB.CommandButton cmdEP08 
      Caption         =   "系統自動上會稿完成日"
      Height          =   405
      Left            =   5580
      TabIndex        =   7
      Top             =   3690
      Width           =   2175
   End
   Begin VB.CommandButton cmdSetTime 
      Caption         =   "校時"
      Height          =   405
      Left            =   2385
      TabIndex        =   6
      Top             =   90
      Width           =   915
   End
   Begin VB.CommandButton cmdChk 
      Caption         =   "手動打卡異常檢查"
      Height          =   405
      Left            =   3795
      TabIndex        =   5
      Top             =   90
      Width           =   1770
   End
   Begin VB.PictureBox Picture1 
      Height          =   345
      Left            =   1035
      ScaleHeight     =   310
      ScaleWidth      =   910
      TabIndex        =   4
      Top             =   60
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Timer tmrClock 
      Left            =   585
      Top             =   30
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "立即接收"
      Height          =   405
      Left            =   5625
      TabIndex        =   3
      Top             =   90
      Width           =   1050
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   405
      Left            =   6705
      TabIndex        =   2
      Top             =   90
      Width           =   1050
   End
   Begin VB.ListBox lstHistory 
      Height          =   2560
      Left            =   135
      TabIndex        =   1
      Top             =   540
      Width           =   7635
   End
   Begin VB.Timer tmrPolling 
      Left            =   135
      Top             =   30
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  '對齊表單下方
      Height          =   320
      Left            =   0
      TabIndex        =   0
      Top             =   3760
      Width           =   7900
      _ExtentX        =   13935
      _ExtentY        =   564
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5080
            MinWidth        =   5080
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuShow 
      Caption         =   "menu"
      Visible         =   0   'False
      Begin VB.Menu mnuDisplay 
         Caption         =   "顯示"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "結束"
      End
   End
End
Attribute VB_Name = "frmHTAauto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bolActived As Boolean
Dim mlngID As Long
Dim intChkErrRow As Integer
Dim bolRunUpdEP08 As Boolean, intUpdEP08okCnt As Integer 'Add By Sindy 2013/10/18
Dim intUpdEP06okCnt As Integer 'Add By Sindy 2016/3/23
Dim bolRunChkAutoEmail As Boolean 'Add By Sindy 2017/6/14
Dim bolRun_T_UpdEP08 As Boolean, intUpd_T_EP08okCnt As Integer 'Add By Sindy 2019/3/21
'Add By Sindy 2025/4/30
Dim m_strFileName As String, m_strDate As String, m_strKind
Dim xlsSalesPoint As New Excel.Application
Dim wksaccrpt114 As New Worksheet
Dim lngCounter As Long
'2025/4/30 END


Private Sub cmdDown_Click()
   PollingData True
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdSetTime_Click()
   Dim arrIpList
   Dim ii As Integer
   
   HTAips = GetHtaIP()
   If HTAips <> "" Then
      arrIpList = Split(HTAips, ";")
      For ii = LBound(arrIpList) To UBound(arrIpList)
         HTAip = arrIpList(ii)
         If HTAip <> "" Then
            If HTAWriteTime(True) = True Then
               WLog "(" & HTAip & ") 指紋機時間已同步！"
            Else
               WLog "(" & HTAip & ") 指紋機時間同步失敗！"
            End If
         End If
      Next
   Else
      WLog "考勤機IP未設定！"
   End If
End Sub

'系統自動上會稿完成日
Private Sub cmdEP08_Click()
   WLog "（手動）系統自動上會稿完成日開始！"
   intUpdEP08okCnt = 0
   intUpdEP06okCnt = 0 'Add By Sindy 2016/3/23
   intUpd_T_EP08okCnt = 0 'Add By Sindy 2019/3/21
   Call RunUpdateEP08(Format(Time, "HHMMSS"))
   Call Run_T_UpdateEP08(Format(Time, "HHMMSS")) 'Add By Sindy 2019/3/21
   WLog "（手動）系統自動上會稿完成日結束！更新筆數：" & intUpdEP08okCnt & " 筆"
   WLog "（手動）系統自動上齊備日結束！更新筆數：" & intUpdEP06okCnt & " 筆" 'Add By Sindy 2016/3/23
   WLog "（手動）系統自動上會完結束！筆數：" & intUpd_T_EP08okCnt & " 筆" 'Add By Sindy 2019/3/21
End Sub

'Add By Sindy 2013/10/15 系統自動上會稿完成日-專利處
'Modify By Sindy 2016/3/22 系統自動上齊備日
Private Function RunUpdateEP08(strTime As String) As Boolean
Dim rsTmp As New ADODB.Recordset
Dim bolGetCnn As Boolean
Dim strChkDate As String
Dim strChkTime As String
Dim strEP02 As String
Dim strEP38 As String
   
   'Add By Sindy 2013/10/28 非工作日不可更新會稿完成日
   If ChkWorkDay(Format(Now, "YYYYMMDD")) = False Then
      WLog "非工作日不可更新會稿完成日，程式結束！"
      Exit Function
   End If
   
   If strTime >= "100000" And strTime <= "105959" Then
      strChkDate = Format(Format(DateAdd("d", -1, Now)), "YYYYMMDD")
      strChkTime = "150000"
   ElseIf strTime >= "110000" And strTime <= "115959" Then
      strChkDate = Format(Format(DateAdd("d", -1, Now)), "YYYYMMDD")
      strChkTime = "160000"
   ElseIf strTime >= "120000" And strTime <= "125959" Then
      strChkDate = Format(Format(DateAdd("d", -1, Now)), "YYYYMMDD")
      strChkTime = "170000"
   ElseIf strTime >= "150000" And strTime <= "155959" Then
      strChkDate = Format(Now, "YYYYMMDD")
      strChkTime = "090000"
   ElseIf strTime >= "160000" And strTime <= "165959" Then
      strChkDate = Format(Now, "YYYYMMDD")
      strChkTime = "100000"
   ElseIf strTime >= "170000" And strTime <= "175959" Then
      strChkDate = Format(Now, "YYYYMMDD")
      strChkTime = "110000"
   ElseIf strTime >= "180000" And strTime <= "185959" Then
      strChkDate = Format(Now, "YYYYMMDD")
      strChkTime = "140000"
   Else
      MsgBox "尚未到確認會稿完成日的時段。"
      WLog "尚未到確認會稿完成日的時段。"
      Exit Function
   End If
   
   RunUpdateEP08 = False
   Screen.MousePointer = vbHourglass
   
   '智權同仁已上會稿完成, 但工程師仍無動作之案件
   'SELECT SUBSTR(S1.ST02,1,3) 工程師, CP01||'-'||CP02||'-'||CP03||'-'||CP04 本所案號,SUBSTR(S2.ST02,1,3) 智權人員,SUBSTR(DECODE(PA09,'000',CPM03,CPM04),1,6) 案件性質,SUBSTR(SQLDATET(EP38),1,10) 會完日,SUBSTR(sqltime(E1.EEP07),1,10) 時間
'   strSql = "SELECT ep02,ep38,cp14,E1.EEP06,E1.EEP07" & _
'            " FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,CASEPROPERTYMAP,PATENT,EmpElectronProcess E1,EmpElectronProcess E2" & _
'            " WHERE EP38<TO_CHAR(sysdate,'YYYYMMDD')" & _
'            " AND NVL(EP08,0)=0 AND NVL(EP07,0)>0" & _
'            " AND EP02=CP09(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+)" & _
'            " AND EP02=E1.EEP01(+) AND '" & EMP_會完 & "'=E1.EEP04(+) AND E1.EEP01 IS NOT NULL" & _
'            " AND E1.EEP01=E2.EEP01(+) AND '" & EMP_不自動更新會完日 & "'=E2.EEP04(+) AND (E2.EEP02 IS NULL OR (E2.EEP02<E1.EEP02 AND E2.EEP01 IS NOT NULL))" & _
'            " AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+)" & _
'            " AND (E1.EEP06<" & strChkDate & " or (E1.EEP06=" & strChkDate & " AND E1.EEP07<=" & strChkTime & "))" & _
'            " ORDER BY CP14,EP38"
   'Modify By Sindy 2014/6/25 上列Sql有誤,修改逐筆判斷
   'Modified by Morgan 2015/10/8
   'strSql = "SELECT ep02,ep38,cp14,cp01,cp02,cp03,cp04" & _
            " FROM ENGINEERPROGRESS,CASEPROGRESS" & _
            " WHERE EP38<=" & strSrvDate(1) & _
            " AND NVL(EP38,0)>0 AND NVL(EP08,0)=0 AND NVL(EP07,0)>0" & _
            " AND EP02=CP09(+) and cp57 is null" & _
            " ORDER BY CP14,EP38"
   strSql = "SELECT ep02,ep38,cp14,cp01,cp02,cp03,cp04" & _
            " FROM ENGINEERPROGRESS,CASEPROGRESS,patent" & _
            " WHERE EP38>0 AND NVL(EP08,0)=0 AND NVL(EP07,0)>0" & _
            " AND EP02=CP09(+) and cp57 is null" & _
            " and pa01=cp01 and pa02=cp02 and pa03=cp03 and pa04=cp04" & _
            " ORDER BY CP14,EP38"
   'end 2015/10/8
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      Do While Not rsTmp.EOF
         If Val("" & rsTmp.Fields("ep38")) > 0 Then '有智權人員會稿完成日
            strEP02 = rsTmp.Fields("ep02")
            strEP38 = rsTmp.Fields("ep38")
            '檢查最近的送會相關流程是否為會完
            strSql = "select * from EmpElectronProcess where eep01='" & strEP02 & "' and eep04 in('" & EMP_送會 & "','" & EMP_會完 & "','" & EMP_會修 & "','" & EMP_不自動更新會完日 & "') order by EEP02 desc"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               RsTemp.MoveFirst
'               If RsTemp.Fields("eep04") = EMP_會完 Then
'                  If CheckIsPersonRest("97038", strSrvDate(1), Left(Right("000000" & ServerTime, 6), 2) & ":" & Mid(Right("000000" & ServerTime, 6), 3, 2)) = False Then
'                     PUB_SendMail "administrator", "97038", "", "系統自動更新會稿完成日 : 檢查的日期及時間=" & strChkDate & " " & strChkTime, _
'                                                                "本所案號：" & rsTmp.Fields("cp01") & rsTmp.Fields("cp02") & rsTmp.Fields("cp03") & rsTmp.Fields("cp04") & vbCrLf & _
'                                                                "總收文號：" & strEP02 & vbCrLf & _
'                                                                "智會完日：" & strEP38 & vbCrLf & _
'                                                                RsTemp.Fields("eep06") & " " & RsTemp.Fields("eep07") & " / " & strChkDate & " " & strChkTime, , , , , , , "administrator", "系統管理員", , True
'                  End If
'               End If
               If RsTemp.Fields("eep04") = EMP_會完 And _
                  (Val(RsTemp.Fields("eep06")) < Val(strChkDate) Or (Val(RsTemp.Fields("eep06")) = Val(strChkDate) And Val(RsTemp.Fields("eep07")) <= Val(strChkTime))) Then
                  cnnConnection.BeginTrans
                  bolGetCnn = True
                  UpdateEp08 strEP02, strEP38 '更新相關會稿完成日資料
                  cnnConnection.CommitTrans
                  bolGetCnn = False
                  WLog "更新文號" & strEP02 & "-->" & strEP38
                  PUB_SendMailCache '發郵件 (相關會稿完成日的郵件)
                  intUpdEP08okCnt = intUpdEP08okCnt + 1
               End If
            End If
         End If
         rsTmp.MoveNext
      Loop
   End If
   rsTmp.Close
   
   'Modify By Sindy 2016/3/22 系統自動上齊備日
   strSql = "select ep02,ep06,cp14,cp01,cp02,cp03,cp04,eep02,eep06,eep07" & _
            " From empelectronprocess,engineerprogress,caseprogress" & _
            " where eep01||eep02 in(select eep01||max(eep02) from empelectronprocess where eep04='" & EMP_圖完 & "'" & _
            " and instr(eep11,'會圖完成')=0 and instr(eep11,'會圖文完成')=0 and instr(eep11,'會(圖/文)完成')=0 group by eep01)" & _
            " and eep01=ep02(+) and ep02=cp09(+)" & _
            " and ep06 is null" & _
            " and cp57 is null" & _
            " ORDER BY EP02"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      Do While Not rsTmp.EOF
         '超過4個工作小時未確認，則系統自動以「圖完」的日期建為新的「齊備日」
         If (Val(rsTmp.Fields("eep06")) < Val(strChkDate) Or (Val(rsTmp.Fields("eep06")) = Val(strChkDate) And Val(rsTmp.Fields("eep07")) <= Val(strChkTime))) Then
            cnnConnection.BeginTrans
            bolGetCnn = True
            
            '更新齊備日
            strSql = "update engineerprogress set ep06=" & rsTmp.Fields("eep06") & " where ep02='" & rsTmp.Fields("ep02") & "'"
            cnnConnection.Execute strSql
            
            '記錄已處理過
            strSql = "update empelectronprocess set eep11='會(圖/文)完成;'||eep11 where eep01='" & rsTmp.Fields("ep02") & "' and eep02=" & rsTmp.Fields("eep02")
            cnnConnection.Execute strSql
            
            cnnConnection.CommitTrans
            WLog "更新文號" & rsTmp.Fields("ep02") & "(" & rsTmp.Fields("eep02") & ")"
            bolGetCnn = False
            
            intUpdEP06okCnt = intUpdEP06okCnt + 1
         End If
         rsTmp.MoveNext
      Loop
   End If
   rsTmp.Close
   '2016/3/22 END
   
   RunUpdateEP08 = True
   Screen.MousePointer = vbDefault
   Set rsTmp = Nothing
   Exit Function

ErrHand:
   Screen.MousePointer = vbDefault
   If bolGetCnn = True Then
      cnnConnection.RollbackTrans
   End If
   WLog "更新會稿完成日失敗！" & Err.Description
End Function

'Add By Sindy 2019/3/21 系統自動上會完-商標處
'OA分析函電子化之系統規則 : (內商承辦人辦理的C類,自動會完,要加發Mail通知智權人員和承辦人)
'所有內商承辦案件 (所有T字頭的台灣大陸案及FCT爭議案)
'的主管機關來函, 由承辦人分析且要送會的案件:
'交智權人員會稿後，若隔日下午五點前未會稿完成者,系統即"自動會回"承辦人
Private Function Run_T_UpdateEP08(strTime As String) As Boolean
Dim rsTmp As New ADODB.Recordset
Dim bolGetCnn As Boolean
Dim strChkDate As String
'Dim strChkTime As String
Dim strEEP01 As String
Dim strCP13 As String, strCP14 As String
Dim intMaxEEP02 As Integer
Dim strCaseNo As String, strCaseName As String, strEP01 As String, strCP10Nm As String
Dim strSubject As String, strContent As String
   
   '非工作日不可更新會稿完成日
   If ChkWorkDay(Format(Now, "YYYYMMDD")) = False Then
      WLog "非工作日不可更新會稿完成日，程式結束！"
      Exit Function
   End If
   
   If strTime >= "170000" And strTime <= "175959" Then
      strChkDate = Format(Format(DateAdd("d", -1, Now)), "YYYYMMDD")
      'strChkTime = "150000"
   Else
      MsgBox "尚未到確認會稿完成的時段。"
      WLog "尚未到確認會稿完成的時段。"
      Exit Function
   End If
   
   Run_T_UpdateEP08 = False
   Screen.MousePointer = vbHourglass
   
   '商標處人員送會中C類案件
   strSql = "select * from empelectronprocess,caseprogress,staff" & _
            " where eep04='08' and eep09='Y' and eep01=cp09(+)" & _
            " and eep03=st01(+)" & _
            " and substr(st03,1,2)='P2'" & _
            " and substr(eep01,1,1)='C'" & _
            " and eep06<=" & strChkDate & _
            " and cp158=0 and cp159=0"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      Do While Not rsTmp.EOF
         strEEP01 = rsTmp.Fields("EEP01")
         strCP13 = rsTmp.Fields("CP13")
         strCP14 = rsTmp.Fields("CP14")
         
         cnnConnection.BeginTrans
         bolGetCnn = True
         
         '送會中的待回覆取消
         strSql = "update empelectronprocess set eep09=null" & _
                  " where EEP01='" & strEEP01 & "' and eep02=" & rsTmp.Fields("EEP02")
         cnnConnection.Execute strSql
         '取得最大序號
         intMaxEEP02 = 0
         strSql = "select eep02 From empelectronprocess where eep01='" & strEEP01 & "' order by eep02 desc"
         intI = 1
         CheckOC3
         Set AdoRecordSet3 = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            AdoRecordSet3.MoveFirst
            If AdoRecordSet3.RecordCount > 0 Then
               intMaxEEP02 = AdoRecordSet3.Fields(0)
            End If
         End If
         '新增會完歷程
         strSql = "insert into empelectronprocess(eep01,eep02,eep03,eep04,eep05,eep06,eep07,eep08,eep10) values(" & _
                  CNULL(strEEP01) & "," & intMaxEEP02 + 1 & ",'QPGMR'," & _
                  CNULL(EMP_會完) & "," & _
                  CNULL(strCP14) & "," & _
                  strSrvDate(1) & "," & Right("000000" & ServerTime, 6) & ",'已逾時系統自動會完','" & strCP13 & "')"
         cnnConnection.Execute strSql
         
         '更新EP38.智權人員會稿完成日,會稿完成日,預設通知客戶
         'Modify By Sindy 2021/1/13 若"多案單筆歷程"也要一併更新 ex:T-228845
         If "" & rsTmp.Fields("CP163") <> "" Then
            strSql = "update engineerprogress set" & _
                     " EP38=" & strSrvDate(1) & _
                     ",EP08=" & strSrvDate(1) & _
                     ",EP11='Y'" & _
                     " where ep02 in(select cp09 from caseprogress where cp163='" & strEEP01 & "')"
         Else
         '2021/1/13 END
            strSql = "update engineerprogress set" & _
                     " EP38=" & strSrvDate(1) & _
                     ",EP08=" & strSrvDate(1) & _
                     ",EP11='Y'" & _
                     " where ep02='" & strEEP01 & "'"
         End If
         cnnConnection.Execute strSql
         
         cnnConnection.CommitTrans
         bolGetCnn = False
         
         '******************************
         '      基本資料檔
         '******************************
         strSql = "select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 本所案號,tm05 as 案件名稱" & _
                  ",nvl(DECODE(tm10,'000',cpm03,cpm04),cp10) as 案件性質,ep01 as 目次,cp13,cp14,tm10,TM44,TM45" & _
                  " From caseprogress,trademark,engineerprogress,casepropertymap" & _
                  " where cp09='" & strEEP01 & "'" & _
                  " and cp01=tm01 and cp02=tm02 and cp03=tm03 and cp04=tm04" & _
                  " and cp09=ep02(+)" & _
                  " and cp01=cpm01(+) and cp10=cpm02(+)" & _
                  " union select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 本所案號,NVL(SP05,NVL(SP06,SP07)) as 案件名稱" & _
                  ",nvl(DECODE(SP09,'000',cpm03,cpm04),cp10) as 案件性質,ep01 as 目次,cp13,cp14,SP09,SP26,SP27" & _
                  " From caseprogress,servicepractice,engineerprogress,casepropertymap" & _
                  " where cp09='" & strEEP01 & "'" & _
                  " and cp01=SP01 and cp02=SP02 and cp03=SP03 and cp04=SP04" & _
                  " and cp09=ep02(+)" & _
                  " and cp01=cpm01(+) and cp10=cpm02(+)"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            RsTemp.MoveFirst
            If RsTemp.RecordCount > 0 Then
               strCaseNo = RsTemp.Fields("本所案號")
               strCaseName = RsTemp.Fields("案件名稱")
               strEP01 = RsTemp.Fields("目次")
               strCP10Nm = RsTemp.Fields("案件性質")
            End If
         End If
         
         strSubject = Replace(strCaseNo, "-0-00", "") & "(核會流程)-->已逾時系統自動會完"
         strContent = "當月目次：" & strEP01 & vbCrLf
         If "" & RsTemp.Fields("TM44") <> "" And "" & RsTemp.Fields("TM10") = "000" Then
            strContent = strContent & "貴方卷號：" & "" & RsTemp.Fields("TM45") & vbCrLf
         End If
         strContent = strContent & "本所案號：" & strCaseNo & vbCrLf
         strContent = strContent & _
                      "案件名稱：" & strCaseName & vbCrLf & _
                      "案件性質：" & strCP10Nm & vbCrLf & _
                      "流程狀態：會完" & vbCrLf
         strContent = strContent & "內　　容：已逾時系統自動會完" & vbCrLf
         strContent = strContent & vbCrLf & vbCrLf & vbCrLf & _
                      "請至系統的下列位置進行：" & vbCrLf & vbCrLf & _
                      " 承　辦　人　員 ：承辦人->工作進度資料維護->待辦歷程" & vbCrLf & _
                      " 核　判　人　員 ：承辦人->待核判區" & vbCrLf & _
                      " 智　權　人　員 ：智權部->日常作業->待會稿區"
         '發E-Mail通知承辦人和智權人員
         PUB_SendMail strUserNum, strCP14, "", strSubject, strContent, , , , , , strCP13
         
         WLog "新增會完" & strEEP01 & "-->" & strCP14 & "-" & strCP13 & "-" & strSrvDate(1)
         intUpd_T_EP08okCnt = intUpd_T_EP08okCnt + 1

         rsTmp.MoveNext
      Loop
   End If
   rsTmp.Close
   
   Run_T_UpdateEP08 = True
   Screen.MousePointer = vbDefault
   Set rsTmp = Nothing
   Exit Function

ErrHand:
   Screen.MousePointer = vbDefault
   If bolGetCnn = True Then
      cnnConnection.RollbackTrans
   End If
   WLog "更新會完歷程失敗！" & Err.Description
End Function

Private Sub Form_Activate()
'   Dim lngRt As Long, strUser As String * 100
   
   Screen.MousePointer = vbHourglass
   If bolActived = False Then
      Me.Top = (Screen.Height - Me.Height) / 2
      Me.Left = (Screen.Width - Me.Width) / 2
      
      'Modify By Sindy 2013/8/1
      Call ConnectDB
      
      'Add By Sindy 2013/8/2
      '檢查是否有設定接收刷卡資料的排程
      strSql = "select count(*) from PollSchedule order by ps01 asc"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         If RsTemp.Fields(0) <= 0 Then
            MsgBox "尚未設定接收刷卡資料的排程！", , "注意！"
            Call cmdExit_Click
            Exit Sub
         End If
      End If
      
      g_strWriteSysLogFilePath = App.path & "\HTAautolog.log" '欲記錄Log的完整路徑及檔名 Add By Sindy 2018/5/28
      strSrvDate(1) = ServerDate
      strSrvDate(2) = strSrvDate(1) - 19110000
      bolMailFailNoAlert = True 'Add by Sindy 2014/3/5 寄信都不要彈錯誤訊息
      '關閉鈕 鎖 x 變灰色
      DisableControl frmHTAauto
      bolActived = True
   End If
   Screen.MousePointer = vbDefault
End Sub

'Modify By Sindy 2013/8/1
Private Sub ConnectDB()
   Me.Caption = Me.Tag 'Added by Morgan 2013/9//5
   If InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") = 0 Then 'Run執行檔
      If fConnect() = False Then
         Unload Me
         End
      End If
   Else
      If PUB_Connect2DB() = False Then
         Unload Me
         End
      End If
      Forms(0).Caption = Forms(0).Caption & PUB_GetDbTerminal
   End If
   
   pub_HostName = PUB_ReadHostName 'Added by Morgan 2013/9/5 要記錄電腦名稱否則寄信會失敗
         
   'Added by Morgan 2015/10/8
   strSrvDate(1) = ServerDate
   strSrvDate(2) = strSrvDate(1) - 19110000
   'end 2015/10/8
   
   Debug.Print PUB_GetDbTerminal
   PUB_SetSystemVar 'Add By Sindy 2017/9/6 設定系統變數
   
   If ClsPDSetUserData(strUserNum, strUserName, strGroup) = False Then
       End
   End If
End Sub

Private Sub Form_Load()
   tmrClock.Interval = 1000
   tmrPolling.Interval = 60000
   lstHistory.Clear
   If mlngID = 0 Then mlngID = AddToSystemTray(Picture1.hWnd, WM_MOUSEMOVE, Me.Icon, Me.Caption)
   Me.Tag = Me.Caption 'Added by Morgan 2013/9//5
End Sub

Private Sub Form_Resize()
If Me.WindowState = "1" Then Me.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call SaveExcelFile("", "") 'Add By Sindy 2025/4/30
   If mlngID <> 0 Then
      DeleteFromSystemTray mlngID
      mlngID = 0
   End If
   Set frmHTAauto = Nothing
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim MSG As Long

If Me.ScaleMode = 1 Then
   MSG = x / Screen.TwipsPerPixelX
End If
Select Case MSG
      Case WM_MOUSEMOVE '移動滑鼠
          'Label1.Caption = "正在移動滑鼠"
      Case WM_LBUTTONDBLCLK '連點滑鼠左鍵
          'Label1.Caption = "連點滑鼠左鍵"
          Me.WindowState = "0"
          Me.Visible = True
      Case WM_LBUTTONDOWN '按下滑鼠左鍵
          'Label1.Caption = "按下滑鼠左鍵"
      Case WM_LBUTTONUP '放開滑鼠左鍵
          'Label1.Caption = "放開滑鼠左鍵"
      Case WM_RBUTTONDBLCLK '連點滑鼠右鍵
          'Label1.Caption = "連點滑鼠右鍵"
      Case WM_RBUTTONDOWN '按下滑鼠右鍵
          'Label1.Caption = "按下滑鼠右鍵"
          Me.PopupMenu mnuShow, vbPopupMenuLeftAlign + vbPopupMenuRightButton
      Case WM_RBUTTONUP '放開滑鼠右鍵
          ''Label1.Caption = "放開滑鼠右鍵"
End Select
End Sub

Private Sub tmrClock_Timer()
   StatusBar1.Panels.Item(2).Text = Time
End Sub

Private Sub tmrPolling_Timer()
'Static iCount As Integer
Dim sTime As String
'Add By Sindy 2013/7/4
Dim strDate As String, strKind As String
Static bolChkRunEndA_1 As Boolean, bolChkRunEndP_1 As Boolean
Static bolChkRunEndA_2 As Boolean, bolChkRunEndP_2 As Boolean
Static bolChkRunEndA_3 As Boolean, bolChkRunEndP_3 As Boolean
Static bolChkRunEndA_4 As Boolean, bolChkRunEndP_4 As Boolean
Dim strChkAStarTime As String, strChkAEndTime As String
Dim strChkPStarTime As String, strChkPEndTime As String
Dim intChkTimeKind As Integer
'2013/7/4 END
Dim dblPS01 As Double, dblPS02 As Double
Dim rsTmp As New ADODB.Recordset
Dim intReadCnt As Integer
Dim ii As Integer
Dim strST06 As String
Dim strST06Nm As String
Dim i As Integer
'Added by Morgan 2019/7/10
Static bolImportRate As Boolean
Dim strErrMsg As String
'end 2019/7/10
Static bolUpdateHoliday As Boolean 'Added by Morgan 2020/9/8
   
   strDate = Format(Now, "YYYYMMDD")
   sTime = Format(Time, "HHMMSS")
   
'   'Add By Sindy 2013/8/1 摩根發現Table資料庫結構變,必須斷線再重新連線,不然會有錯誤
'   If cnnConnection.State = adStateOpen Then
'      Forms(0).StatusBar1.Panels(1).Text = "強迫斷線..."
'      cnnConnection.Close
'      WLog "tmrPolling_Timer : 強迫斷線..." 'Add By Sindy 2015/8/14
'      '再連線
'      Call ConnectDB
'   End If
'   '2013/8/1 END
   'Modify By Sindy 2015/8/18 同上原因改為早上7:00再連線,晚上10:00斷線則不再連資料庫
   If (sTime >= "070000" And sTime < "220000") And cnnConnection.State = adStateClosed Then
      Forms(0).StatusBar1.Panels(1).Text = "連線資料庫..."
      WLog "tmrPolling_Timer : 連線資料庫..."
      '再連線
      Call ConnectDB
      WLog "tmrPolling_Timer : 已連線..."
   ElseIf (sTime >= "220000" Or sTime < "070000") Then
      If cnnConnection.State = adStateOpen Then
         Forms(0).StatusBar1.Panels(1).Text = "強迫斷線..."
         cnnConnection.Close
         WLog "tmrPolling_Timer : 強迫斷線..."
      End If
      Exit Sub '這段時間休息,不須執行程式
   End If
   '2015/8/18 END
     
   'Added by Morgan 2020/9/8
   If Right(strDate, 4) = "1231" Then
      If sTime >= "210000" And sTime <= "213000" Then
         If bolUpdateHoliday = False Then
            bolUpdateHoliday = True
            WLog "tmrPolling_Timer : 更新門禁機假日表..."
            If PUB_WriteHoliday(True, strErrMsg) = False Then
               WLog "tmrPolling_Timer : 更新門禁機假日表失敗(" & strErrMsg & ")"
               PUB_SendMail strUserNum, "74001;92012", "", strErrMsg, "如旨"
            Else
               WLog "tmrPolling_Timer : 更新門禁機假日表完成"
            End If
         End If
      ElseIf sTime > "213000" Then
         bolUpdateHoliday = False
      End If
   End If
   'end 2020/9/8
   
   'Add By Sindy 2013/11/5 非工作天,則不需執行Timer,不然Log假日還是繼續寫但程式中其實都判斷不執行
   If ChkWorkDay(DBDATE(strDate)) = False Then
      Exit Sub
   End If
   '2013/11/5 END
   
   'Add By Sindy 2013/10/18 系統自動確認會稿完成日
   ' 9:00-12:00
   '14:00-18:00
   If (sTime >= "100000" And sTime <= "105000") Or _
      (sTime >= "110000" And sTime <= "115000") Or _
      (sTime >= "120000" And sTime <= "125000") Or _
      (sTime >= "150000" And sTime <= "155000") Or _
      (sTime >= "160000" And sTime <= "165000") Or _
      (sTime >= "170000" And sTime <= "175000") Or _
      (sTime >= "180000" And sTime <= "185000") Then
      If bolRunUpdEP08 = False Then
         'MsgBox "Run RunUpdateEP08 !!!"
         WLog "系統自動上會稿完成日開始！"
         intUpdEP08okCnt = 0
         intUpdEP06okCnt = 0 'Add By Sindy 2016/3/23
         Call RunUpdateEP08(sTime)
         WLog "系統自動上會稿完成日結束！更新筆數：" & intUpdEP08okCnt & " 筆"
         WLog "系統自動上齊備日結束！更新筆數：" & intUpdEP06okCnt & " 筆" 'Add By Sindy 2016/3/23
         bolRunUpdEP08 = True
      End If
   End If
   '清空已執行過EP08的註記變數值,等待下一次更新
   If Not (sTime >= "100000" And sTime <= "105000") And _
      Not (sTime >= "110000" And sTime <= "115000") And _
      Not (sTime >= "120000" And sTime <= "125000") And _
      Not (sTime >= "150000" And sTime <= "155000") And _
      Not (sTime >= "160000" And sTime <= "165000") And _
      Not (sTime >= "170000" And sTime <= "175000") And _
      Not (sTime >= "180000" And sTime <= "185000") Then
      bolRunUpdEP08 = False
   End If
   '2013/10/18 END
   
   'Add By Sindy 2019/3/21 系統自動上會完
   ' 9:00-12:00
   '14:00-18:00
   If (sTime >= "170000" And sTime <= "175000") Then
      If bolRun_T_UpdEP08 = False Then
         WLog "系統自動上會完開始！"
         intUpd_T_EP08okCnt = 0
         Call Run_T_UpdateEP08(sTime)
         WLog "系統自動上會完結束！筆數：" & intUpd_T_EP08okCnt & " 筆"
         bolRun_T_UpdEP08 = True
      End If
   End If
   '清空已執行過EP08的註記變數值,等待下一次更新
   If Not (sTime >= "170000" And sTime <= "175000") Then
      bolRun_T_UpdEP08 = False
   End If
   '2019/3/21 END
   
   'Add By Sindy 2021/11/23 判斷 彈性上下班時段 設定 接收刷卡資料排程(A.全所上班異常 P.全所下班異常)
   Call PUB_ChkByPassWork("1", strSrvDate(1))
   If intByPassArea = 6 Then '分6班制,10:15分啟動異常檢查
      strSql = "update PollSchedule set PS03='A',PS04='P' where PS01=101500"
      cnnConnection.Execute strSql, intI
      strSql = "update PollSchedule set PS03=null,PS04=null where PS01=91500"
      cnnConnection.Execute strSql, intI
   Else '分3班制,9:15分啟動異常檢查
      strSql = "update PollSchedule set PS03=null,PS04=null where PS01=101500"
      cnnConnection.Execute strSql, intI
      strSql = "update PollSchedule set PS03='A',PS04='P' where PS01=91500"
      cnnConnection.Execute strSql
   End If
   '2021/11/23 END
   
   '接收刷卡資料
   strSql = "select * from PollSchedule order by ps01 asc"
   intI = 1
   Set rsTmp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      rsTmp.MoveFirst
      Do While Not rsTmp.EOF
         dblPS01 = rsTmp.Fields("PS01")
         dblPS02 = Val(rsTmp.Fields("PS01")) + 100
         If sTime >= dblPS01 And sTime < dblPS02 Then
            PollingData , True
         End If
         rsTmp.MoveNext
      Loop
   End If
   
'   iCount = iCount + 1
'   If sTime > "072500" And sTime < "093000" Then
'      'PollingData
'      iCount = 0
   'Else
   '   ElseIf iCount > 59 Then
'      PollingData , True
'      iCount = 0
'   End If
   
   Call PUB_KillTempFile("*加班關心提醒.xls") 'Add By Sindy 2025/5/12 清除舊檔
   
   '檢查打卡是否異常
   strSql = "select * from PollSchedule where ps03 is not null or ps04 is not null order by ps01 asc"
   intI = 1
   Set rsTmp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      rsTmp.MoveFirst
      Do While Not rsTmp.EOF
         intChkTimeKind = 0
         dblPS01 = rsTmp.Fields("PS01")
         If IsNull(rsTmp.Fields("PS02")) Then
            dblPS02 = Val(rsTmp.Fields("PS01")) + 3000
         Else
            dblPS02 = Val(rsTmp.Fields("PS02"))
         End If
         If sTime >= dblPS01 And sTime <= dblPS02 Then
            '下班異常
            If Not IsNull(rsTmp.Fields("PS04")) Then
               If rsTmp.Fields("PS04") <> "" Then
                  'Add By Sindy 2013/7/4 檢查下班打卡是否有異常
                  strChkPStarTime = Format(dblPS01, "000000")
                  strChkPEndTime = Format(dblPS02, "000000")
                  If bolChkRunEndP_1 = False Or bolChkRunEndP_2 = False Or bolChkRunEndP_3 = False Or bolChkRunEndP_4 = False Then
                     strKind = "P"
                     'Modify By Sindy 2013/11/4 下班異常也要在工作天時才Run,若星期六Run星期五的下班異常,中所資料可能尚未讀入完整
                     'strDate = Format(Format(DateAdd("d", -1, Now)), "YYYYMMDD")
                     strDate = Format(Now, "YYYYMMDD")
                     If ChkWorkDay(DBDATE(strDate)) = True Then
                        For i = 1 To 22
                           strDate = Format(Format(DateAdd("d", i * -1, Now)), "YYYYMMDD")
                           If ChkWorkDay(DBDATE(strDate)) = True Then Exit For
                        Next i
                     '2013/11/4 END
                        
                        'Add By Sindy 2014/2/7 再執行一次接收,以免尚有資料未接收下來時,造成比對資料有誤
                        '2014/2/7 上午 08:45:53  ==>  (192.168.4.1) 開始接收刷卡紀錄
                        '2014/2/7 上午 08:45:54  ==>  (192.168.4.1) 接收完成共5筆！
                        '2014/2/7 上午 09:16:00  ==>  20140206 下班打卡異常檢查開始！
                        '2014/2/7 上午 09:16:03  ==>  20140206 下班打卡異常檢查結束！（北所）異常筆數：9 筆
                        PollingData , True
                        '2014/2/7 END
                        
                        'Modify By Sindy 2021/12/13 劉柏翰經理:12/17因辦理尾牙活動,取消下班時段打卡異常
                        If CheckDataValidate(strDate, strKind, 0) = True And strDate <> "20211217" Then
                           'Add By Sindy 2025/6/5 檢查是否今天已有Run過,防止重覆執行
                           strSql = "select * from executelog where el01='" & Me.Name & "-" & UCase(rsTmp.Fields("PS04")) & "' and el02='" & strUserNum & "' and el03=" & strSrvDate(1)
                           intI = 1
                           Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                           If intI = 0 Then
                           '2025/6/5 END
                              WLog strDate & " 下班打卡異常檢查開始！"
                              '檢查是否有接收到刷卡資料,若無,通知經理和秀玲
                              'If (strKind = "A" And intChkTimeKind = 1) Or strKind = "P" Then '上班第一階段或下班才須檢查
                              For ii = 1 To 4
                                 If ii = 1 Then strST06 = "1": strST06Nm = "（北所）"
                                 'Add By Sindy 2013/10/31
                                 If ii = 2 Then strST06 = "2": strST06Nm = "（中所）"
                                 If ii = 2 And Val(strDate) < 20131101 Then Exit For
                                 '2013/10/31 END
                                 'Add By Sindy 2013/11/20
                                 If ii = 3 Then strST06 = "4": strST06Nm = "（高所）"
                                 If ii = 3 And Val(strDate) < 20131201 Then Exit For
                                 '2013/11/20 END
                                 'Add By Sindy 2013/12/10
                                 If ii = 4 Then strST06 = "3": strST06Nm = "（南所）"
                                 If ii = 4 And Val(strDate) < 20131216 Then Exit For
                                 '2013/12/10 END
                                 strSql = "select count(*) from pollrecord,staffcarddata,staff where pr03=scd02(+) and scd01=st01(+) and pr01=" & DBDATE(strDate) & _
                                          " and pr02>=170000 and st06='" & strST06 & "'"
                                 intI = 1: intReadCnt = 0
                                 Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                                 If intI = 1 Then
                                    intReadCnt = RsTemp.Fields(0)
                                 End If
                                 If intI = 0 Or intReadCnt = 0 Then
                                    'Modify By Sindy 2018/8/27
                                    If ChkWorkDay(DBDATE(strDate), , True, strST06) = True Then '檢查颱風假問題
                                    '2018/8/27 END
                                       PUB_SendMail "administrator", "74001;83002", "", ChangeWStringToTDateString(strDate) & strST06Nm & "下班刷卡資料無法下載，請確認！", "同主旨", , , , , , , "administrator", "系統管理員", , True
                                       WLog ChangeWStringToTDateString(strDate) & strST06Nm & "下班刷卡資料無法下載，請確認！"
                                    End If
                                 Else
                                    If ChkWorkTime(strDate, strKind, 0, strST06) = True Then
                                       WLog strDate & " 下班打卡異常檢查結束！" & strST06Nm & "異常筆數：" & intChkErrRow & " 筆"
                                    End If
                                 End If
                                 If ii = 1 Then bolChkRunEndP_1 = True
                                 If ii = 2 Then bolChkRunEndP_2 = True
                                 If ii = 3 Then bolChkRunEndP_3 = True
                                 If ii = 4 Then bolChkRunEndP_4 = True
                              Next ii
                              PUB_AddExcuteLog Me.Name & "-" & UCase(rsTmp.Fields("PS04")) 'Add By Sindy 2025/6/5
                           End If '2025/6/5 +
                        Else
                           bolChkRunEndP_1 = True
                           bolChkRunEndP_2 = True
                           bolChkRunEndP_3 = True
                           bolChkRunEndP_4 = True
                        End If
                     End If
                  End If
               End If
            End If
            '上班異常
            If Not IsNull(rsTmp.Fields("PS03")) Then
               If rsTmp.Fields("PS03") <> "" Then
                  'Add By Sindy 2013/7/4 檢查上班打卡是否有異常
                  strChkAStarTime = Format(dblPS01, "000000")
                  strChkAEndTime = Format(dblPS02, "000000")
                  If UCase(rsTmp.Fields("PS03")) = "A" Then
                     intChkTimeKind = 1
                  ElseIf UCase(rsTmp.Fields("PS03")) = "A2" Then
                     intChkTimeKind = 2
                  ElseIf UCase(rsTmp.Fields("PS03")) = "A3" Then
                     intChkTimeKind = 3
                  End If
                  If bolChkRunEndA_1 = False Or bolChkRunEndA_2 = False Or bolChkRunEndA_3 = False Or bolChkRunEndA_4 = False Then
                     strKind = "A"
                     strDate = Format(Now, "YYYYMMDD")
                     
                     'Add By Sindy 2014/2/7 再執行一次接收,以免尚有資料未接收下來時,造成比對資料有誤
                     '2014/2/7 上午 08:45:53  ==>  (192.168.4.1) 開始接收刷卡紀錄
                     '2014/2/7 上午 08:45:54  ==>  (192.168.4.1) 接收完成共5筆！
                     '2014/2/7 上午 09:16:00  ==>  20140206 下班打卡異常檢查開始！
                     '2014/2/7 上午 09:16:03  ==>  20140206 下班打卡異常檢查結束！（北所）異常筆數：9 筆
                     PollingData , True
                     '2014/2/7 END
                     
                     If CheckDataValidate(strDate, strKind, intChkTimeKind) = True Then
                        'Add By Sindy 2025/6/5 檢查是否今天已有Run過,防止重覆執行
                        strSql = "select * from executelog where el01='" & Me.Name & "-" & UCase(rsTmp.Fields("PS03")) & "' and el02='" & strUserNum & "' and el03=" & strSrvDate(1)
                        intI = 1
                        Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                        If intI = 0 Then
                        '2025/6/5 END
                           WLog strDate & " 上班（" & intChkTimeKind & "）打卡異常檢查開始！"
                           '檢查是否有接收到刷卡資料,若無,通知經理和秀玲
                           'If (strKind = "A" And intChkTimeKind = 1) Or strKind = "P" Then '上班第一階段或下班才須檢查
                           For ii = 1 To 4
                              If ii = 1 Then strST06 = "1": strST06Nm = "（北所）"
                              If ii = 2 Then strST06 = "2": strST06Nm = "（中所）" 'Add By Sindy 2013/10/31
                              'Add By Sindy 2013/11/20
                              If ii = 3 Then strST06 = "4": strST06Nm = "（高所）"
                              If ii = 3 And Val(strDate) < 20131201 Then Exit For
                              '2013/11/20 END
                              'Add By Sindy 2013/12/10
                              If ii = 4 Then strST06 = "3": strST06Nm = "（南所）"
                              If ii = 4 And Val(strDate) < 20131216 Then Exit For
                              '2013/12/10 END
                              If (intChkTimeKind = 2 Or intChkTimeKind = 3) And strST06 <> "1" Then Exit For '北所才有出異常的特殊時段
                              strSql = "select count(*) from pollrecord,staffcarddata,staff where pr03=scd02(+) and scd01=st01(+) and pr01=" & DBDATE(strDate) & _
                                       " and pr02<=90000 and st06='" & strST06 & "'"
                              intI = 1: intReadCnt = 0
                              Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                              If intI = 1 Then
                                 intReadCnt = RsTemp.Fields(0)
                              End If
                              If intI = 0 Or intReadCnt = 0 Then
                                 If intChkTimeKind = 1 Then
                                    'Modify By Sindy 2018/8/27
                                    If ChkWorkDay(DBDATE(strDate), , True, strST06) = True Then '檢查颱風假問題
                                    '2018/8/27 END
                                       PUB_SendMail "administrator", "74001;83002", "", ChangeWStringToTDateString(strDate) & strST06Nm & "上班刷卡資料無法下載，請確認！", "同主旨", , , , , , , "administrator", "系統管理員", , True
                                       WLog ChangeWStringToTDateString(strDate) & strST06Nm & "上班刷卡資料無法下載，請確認！"
                                    End If
                                 End If
                              Else
                                 If ChkWorkTime(strDate, strKind, intChkTimeKind, strST06) = True Then
                                    WLog strDate & " 上班（" & intChkTimeKind & "）打卡異常檢查結束！" & strST06Nm & "異常筆數：" & intChkErrRow & " 筆"
                                 End If
                              End If
                              If ii = 1 Then bolChkRunEndA_1 = True
                              If ii = 2 Or intChkTimeKind <> 1 Then bolChkRunEndA_2 = True
                              If ii = 3 Or intChkTimeKind <> 1 Then bolChkRunEndA_3 = True
                              If ii = 4 Or intChkTimeKind <> 1 Then bolChkRunEndA_4 = True
                           Next ii
                           PUB_AddExcuteLog Me.Name & "-" & UCase(rsTmp.Fields("PS03")) 'Add By Sindy 2025/6/5
                        End If '2025/6/5 +
                     Else
                        bolChkRunEndA_1 = True
                        bolChkRunEndA_2 = True
                        bolChkRunEndA_3 = True
                        bolChkRunEndA_4 = True
                     End If
                  End If
               End If
            End If
            Exit Do
         End If
         rsTmp.MoveNext
      Loop
   End If
   
'   intChkTimeKind = 0
'   If sTime >= 94500 And sTime <= 101500 Then
'      '檢查下班打卡
'      strChkPStarTime = "094500"
'      strChkPEndTime = "101500"
'      '檢查上班打卡第1階段
'      strChkAStarTime = "094500"
'      strChkAEndTime = "101500"
'      intChkTimeKind = 1
'   ElseIf sTime >= 120000 And sTime <= 123000 Then
'      '檢查上班打卡第2階段
'      strChkAStarTime = "120000"
'      strChkAEndTime = "123000"
'      intChkTimeKind = 2
'   ElseIf sTime >= 140000 And sTime <= 143000 Then
'      '檢查上班打卡第3階段
'      strChkAStarTime = "140000"
'      strChkAEndTime = "143000"
'      intChkTimeKind = 3
''   Else
''      strChkPStarTime = "000000"
''      strChkPEndTime = "010000"
'   End If
   
   '清空註記執行過打卡異常程式的變數值
   If Not (sTime >= strChkAStarTime And sTime <= strChkAEndTime) Then
      bolChkRunEndA_1 = False
      bolChkRunEndA_2 = False
      bolChkRunEndA_3 = False
      bolChkRunEndA_4 = False
   End If
   If Not (sTime >= strChkPStarTime And sTime <= strChkPEndTime) Then
      bolChkRunEndP_1 = False
      bolChkRunEndP_2 = False
      bolChkRunEndP_3 = False
      bolChkRunEndP_4 = False
   End If
   
   'Add By Sindy 2017/6/13
   '檢查系統自動接收郵件,是否有正常Run...
   If strSrvDate(1) >= 20170706 Then
      sTime = Format(Time, "HHMMSS")
         'Modify By Sindy 2022/5/27 取消這3個時段的檢查
'        (sTime >= "121500" And sTime <= "122500") Or _
'        (sTime >= "124500" And sTime <= "125500") Or _
'        (sTime >= "131500" And sTime <= "132500") Or
      If (sTime >= "071500" And sTime <= "072500") Or _
         (sTime >= "074500" And sTime <= "075500") Or _
         (sTime >= "081500" And sTime <= "082500") Or _
         (sTime >= "084500" And sTime <= "085500") Or _
         (sTime >= "091500" And sTime <= "092500") Or _
         (sTime >= "094500" And sTime <= "095500") Or _
         (sTime >= "101500" And sTime <= "102500") Or _
         (sTime >= "104500" And sTime <= "105500") Or _
         (sTime >= "111500" And sTime <= "112500") Or _
         (sTime >= "114500" And sTime <= "115500") Or _
         (sTime >= "134500" And sTime <= "135500") Or _
         (sTime >= "141500" And sTime <= "142500") Or _
         (sTime >= "144500" And sTime <= "145500") Or _
         (sTime >= "151500" And sTime <= "152500") Or _
         (sTime >= "154500" And sTime <= "155500") Or _
         (sTime >= "161500" And sTime <= "162500") Or _
         (sTime >= "164500" And sTime <= "165500") Or _
         (sTime >= "171500" And sTime <= "172500") Or _
         (sTime >= "174500" And sTime <= "175500") Or _
         (sTime >= "181500" And sTime <= "182500") Then
         If bolRunChkAutoEmail = False Then
            WLog "檢查系統自動接收郵件,是否有正常Run...開始！"
            bolRunChkAutoEmail = RunChkAutoEmail(strDate, sTime)
            If bolRunChkAutoEmail = True Then
               WLog "檢查系統自動接收郵件,是否有正常Run...結束！"
            End If
         End If
      End If
      '清空已執行過檢查郵件的註記變數值,等待下一次檢查
         'Modify By Sindy 2022/5/27 取消這3個時段的檢查
'         Not (sTime >= "121500" And sTime <= "122500") And _
'         Not (sTime >= "124500" And sTime <= "125500") And _
'         Not (sTime >= "131500" And sTime <= "132500") And
      If Not (sTime >= "071500" And sTime <= "072500") And _
         Not (sTime >= "074500" And sTime <= "075500") And _
         Not (sTime >= "081500" And sTime <= "082500") And _
         Not (sTime >= "084500" And sTime <= "085500") And _
         Not (sTime >= "091500" And sTime <= "092500") And _
         Not (sTime >= "094500" And sTime <= "095500") And _
         Not (sTime >= "101500" And sTime <= "102500") And _
         Not (sTime >= "104500" And sTime <= "105500") And _
         Not (sTime >= "111500" And sTime <= "112500") And _
         Not (sTime >= "114500" And sTime <= "115500") And _
         Not (sTime >= "134500" And sTime <= "135500") And _
         Not (sTime >= "141500" And sTime <= "142500") And _
         Not (sTime >= "144500" And sTime <= "145500") And _
         Not (sTime >= "151500" And sTime <= "152500") And _
         Not (sTime >= "154500" And sTime <= "155500") And _
         Not (sTime >= "161500" And sTime <= "162500") And _
         Not (sTime >= "164500" And sTime <= "165500") And _
         Not (sTime >= "171500" And sTime <= "172500") And _
         Not (sTime >= "174500" And sTime <= "175500") And _
         Not (sTime >= "181500" And sTime <= "182500") Then
         bolRunChkAutoEmail = False
      End If
   End If
      
   If rsTmp.State <> 0 Then rsTmp.Close
   Set rsTmp = Nothing
   
   'Modify By Sindy 2024/8/1 Mark: @TEST此寫法,因經理有調整DB設定無法使用
'   '＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊
'   '測試用-暫時 不能加execDelSql.sql是因更新正式資料庫又會再倒回測試DB,資料就蓋掉了
'   '＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊
'   strSql = "UPDATE STAFF_PWD@TEST SET SP03='B6366' WHERE SP01='A4021'"
'   cnnConnection.Execute strSql, intI
'   strSql = "UPDATE STAFF_PWD@TEST SET SP03='9636<' WHERE SP01='84027'"
'   cnnConnection.Execute strSql, intI
'   strSql = "UPDATE STAFF_PWD@TEST SET SP03='9838=' WHERE SP01='86048'"
'   cnnConnection.Execute strSql, intI
'   strSql = "UPDATE STAFF_PWD@TEST SET SP03='B736;' WHERE SP01='A5026'"
'   cnnConnection.Execute strSql, intI
'   strSql = "UPDATE STAFF_PWD@TEST SET SP03='B8379' WHERE SP01='A6034'"
'   cnnConnection.Execute strSql, intI
'   strSql = "UPDATE STAFF_PWD@TEST SET SP03='B435<' WHERE SP01='A2017'"
'   cnnConnection.Execute strSql, intI
'   '＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊
   
   'Added by Morgan 2019/7/10
   '下載台銀當日結匯匯率CSV檔並匯入
   If (sTime >= "163000" And sTime <= "170000") Then
      If bolImportRate = False Then
         bolImportRate = True
         WLog ""
         If ImportTWBankRate(strErrMsg) = False Then
            PUB_SendMail strUserNum, "account", "", "每日自動匯入台銀匯率失敗!!!(" & strSrvDate(2) & ")", strErrMsg, , , , , , "92012", , , , , False, , , False
         End If
         WLog ""
      End If
   End If
   If Not (sTime >= "163000" And sTime <= "170000") Then
      bolImportRate = False
   End If
   Exit Sub
   'end 2019/7/10
End Sub

'Add By Sindy 2017/6/13
'檢查系統自動接收郵件,是否有正常Run...
'RunChkAutoEmail = False.要繼續檢查 True.OK等下一時段再檢查
Private Function RunChkAutoEmail(ByVal strDate As String, ByVal strTime As String) As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strChkTime As String
Dim strChkTimeS As String, strChkTimeE As String
Dim strSubject As String
Dim strContext As String
Dim ii As Integer
   
   RunChkAutoEmail = True
   '非工作日不須檢查,因為人員休假
   If ChkWorkDay(Format(Now, "YYYYMMDD")) = False Then
      WLog "非工作日不須檢查(系統自動接收郵件系統)，程式結束！"
      Exit Function
   End If
   
   If Mid(Format(strTime, "0#####"), 3, 2) < 30 Then
      strChkTimeS = Left(Format(strTime, "0#####"), 2) & "0000"
      strChkTimeE = Left(Format(strTime, "0#####"), 2) & "3000"
   Else
      strChkTimeS = Left(Format(strTime, "0#####"), 2) & "3000"
      strChkTimeE = Left(Format(strTime, "0#####"), 2) + 1 & "0000"
   End If
   
   '檢查是否已有接收過信件資料:
   '信箱均無任何接收資訊,立即回報電腦中心
   strSql = "select mrl03,mrl04 from MailReceiveLog" & _
            " where mrl02=" & strDate & _
            " and mrl03 between " & strChkTimeS & " and " & strChkTimeE
   intI = 1
   Set rsTmp = ClsLawReadRstMsg(intI, strSql)
   If intI = 0 Then
      strSubject = "系統自動接收郵件，信箱均無任何接收資訊，請至" & UCase(Pub_GetSpecMan("分信主機名稱")) & "查看目前郵件系統是否有正常！"
      strContext = "同主旨" & vbCrLf & _
                   "strChkTimeS=" & strChkTimeS & vbCrLf & _
                   "strChkTimeE=" & strChkTimeE
      PUB_SendMail "administrator", Pub_GetSpecMan("電腦中心郵件檢核人員"), "", strSubject, strContext, , , , , , , "administrator", "系統管理員", , , False, , , False, , , False
      RunChkAutoEmail = True
      Exit Function
   End If
   '信箱已接收完成
   For ii = 1 To 4 '2
      strSql = "select mrl03,mrl04 from MailReceiveLog" & _
               " where mrl01='" & Format(ii, "0#") & "'" & _
               " and mrl02=" & strDate & _
               " and mrl03 between " & strChkTimeS & " and " & strChkTimeE & _
               " and mrl09='E'"
      intI = 1
      Set rsTmp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 And ii = 4 Then
         'Modify By Sindy 2021/10/1 信件有遺失,轉寄資訊正常,但確實寄信備份網頁系統找不到信件
         'select ii08,ii09,ii20,ii21,ii22,ii17 from ipdeptinput where ii01='20181025' and ii03 in('F0292','F0304','F0293','F0262');
         '/*
         '      II08       II09 II20                       II21       II22 II17
         '---------- ---------- -------------------- ---------- ---------- --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
         '  20181025     141308 Y                      20181025     141310 未傳遞的主旨: Mail Delivery Failure
         '  20181026     143250 Y                      20181026     143256 Mail Delivery Failure
         '  20181026     143249 Y                      20181026     143255 IMPORTANT NOTICE
         '  20181026     143249 Y                      20181026     143254 Out of Office Notice
         '*/
         strExc(0) = "select count(*) from ipdeptinput where ii20<>'Y' and ii20 is not null" & _
                     " and ii01>=20181001" & _
                     " order by ii01,ii02"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp.Fields(0) > 0 Then
               PUB_SendMail strUserNum, Pub_GetSpecMan("電腦中心郵件檢核人員"), "", "檢查信件是否有遺失(" & RsTemp.Fields(0) & "筆)", strExc(0), , , , , , , "administrator", "系統管理員", , True, False, , , False, , , False
            End If
         End If
         '2021/10/1 END
         
         RunChkAutoEmail = True
         Exit Function
      End If
   Next ii
   '檢查正在執行中的Timer
   strSql = "select mrl01,mrl02,mrl03,mrl04,mrl05 from MailReceiveLog" & _
            " where mrl02=" & strDate & _
            " and mrl09='Y'"
   intI = 1
   Set rsTmp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      RunChkAutoEmail = False
      rsTmp.MoveFirst
      Do While Not rsTmp.EOF
         '如果已20分鐘尚未結束,則通知電腦中心人員
         strChkTime = Format(rsTmp.Fields("mrl03"), "0#####")
         If Mid(strChkTime, 3, 2) + 20 = 59 Then
            strChkTime = CStr(Left(strChkTime, 2) + 1) & "00" & CStr(Right(strChkTime, 2))
         ElseIf Mid(strChkTime, 3, 2) + 20 > 59 Then
            strChkTime = CStr(Left(strChkTime, 2) + 1) & Format(CStr(Mid(strChkTime, 3, 2) + 20 - 60), "0#") & CStr(Right(strChkTime, 2))
         Else
            strChkTime = CStr(Left(strChkTime, 2)) & Format(CStr(Mid(strChkTime, 3, 2) + 20), "0#") & CStr(Right(strChkTime, 2))
         End If
         If Val(strChkTime) <= Val(Format(Time, "HHMMSS")) Then
            strSubject = "系統自動接收郵件，有接收信箱(" & rsTmp.Fields("mrl01") & ")正在執行中,已20分鐘尚未結束,是否有異常，請至" & UCase(Pub_GetSpecMan("分信主機名稱")) & "查看目前郵件系統是否有正常！"
            strContext = "*信箱代碼：01.IPDept_inbound 02.IPDept_backup 03.Patent 04.TM" & vbCrLf & _
                         "接收日期：" & rsTmp.Fields("mrl02") & vbCrLf & _
                         "接收起始時間：" & rsTmp.Fields("mrl03") & vbCrLf & _
                         "接收截止時間：" & rsTmp.Fields("mrl04") & vbCrLf & _
                         "操作人員：" & rsTmp.Fields("mrl05") & " " & GetPrjSalesNM(rsTmp.Fields("mrl05")) & vbCrLf & _
                         "strChkTimeS=" & strChkTimeS & vbCrLf & _
                         "strChkTimeE=" & strChkTimeE & vbCrLf & _
                         "strChkTime =" & strChkTime
            PUB_SendMail "administrator", Pub_GetSpecMan("電腦中心郵件檢核人員"), "", strSubject, strContext, , , , , , , "administrator", "系統管理員", , , False, , , False, , , False
            RunChkAutoEmail = True
            Exit Function
         End If
         rsTmp.MoveNext
      Loop
   End If
   
   Set rsTmp = Nothing
   Exit Function
   
ErrHand:
   WLog "檢查系統自動接收郵件,是否有正常Run...失敗！" & Err.Description
End Function

Private Sub PollingData(Optional bNotAuto As Boolean, Optional bUpdateTime As Boolean)
   Dim iRecs As Integer
   Dim arrIpList
   Dim ii As Integer
   
   'Modify By Sindy 2018/1/2
   'If UCase(PUB_GetDbTerminal) = UCase("(M51-1)") Then
   If UCase(pub_DbTerminalName) <> UCase(正式資料庫電腦名稱) Then
'      If MsgBox("現在是(" & pub_DbTerminalName & ")測試資料庫，確定還要下載刷卡資料嗎？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
   '2018/1/2 END
         MsgBox "現在是(" & pub_DbTerminalName & ")測試資料庫，不可下載刷卡資料", vbExclamation
         Exit Sub
'      End If
   End If
   
   HTAips = GetHtaIP()
   If HTAips <> "" Then
      arrIpList = Split(HTAips, ";")
      For ii = LBound(arrIpList) To UBound(arrIpList)
         HTAip = arrIpList(ii)
         If HTAip <> "" Then
         
            If bUpdateTime = True Then
               If HTAWriteTime(True) = True Then
                  WLog "(" & HTAip & ") 指紋機時間已同步！"
               Else
                  WLog "(" & HTAip & ") 指紋機時間同步失敗！"
               End If
            End If
            
            WLog "(" & HTAip & ") 開始接收刷卡紀錄" & IIf(bNotAuto, "(手動)", "")
            
            If ghComm <> 0 Then HTAclose True 'Added by Morgan 2017/6/28
            
            If HTAPolling(iRecs, True) = True Then
               If iRecs = 0 Then
                  WLog "(" & HTAip & ") 沒有新刷卡紀錄可接收！"
               Else
                  WLog "(" & HTAip & ") 接收完成共" & iRecs & "筆！"
               End If
            Else
               WLog "(" & HTAip & ") 刷卡紀錄接收失敗！"
            End If
         End If
      Next
   Else
      WLog "考勤機IP未設定！"
   End If
End Sub

Private Sub mnuDisplay_Click()
Me.WindowState = "0"
Me.Visible = True
End Sub

Private Sub mnuQuit_Click()
Unload Me
End Sub

'Add By Sindy 2013/7/4
Private Sub cmdChk_Click()
   Dim strDate As String, strKind As String, intChkTimeKind As Integer
   Dim strPS03_1 As String, strPS03_2 As String, strPS03_3 As String
   Dim strST06 As String
   Dim strST06Nm As String
   
   'Add By Sindy 2013/10/24
   strST06 = Trim(InputBox("請輸入所別：1.北所 2.中所 3.南所 4.高所？"))
   If strST06 <> "1" And strST06 <> "2" And strST06 <> "3" And strST06 <> "4" Then
      MsgBox "請輸入所別（1.北所 2.中所 3.南所 4.高所）！"
      Exit Sub
   End If
   'Add By Sindy 2013/11/4
   If strST06 = "1" Then strST06Nm = "（北所）"
   If strST06 = "2" Then strST06Nm = "（中所）"
   If strST06 = "3" Then strST06Nm = "（南所）"
   If strST06 = "4" Then strST06Nm = "（高所）"
   '2013/11/4 END
   strDate = InputBox("請輸入欲檢查打卡異常的日期？")
   If ChkDate(DBDATE(strDate)) = False Then
      Exit Sub
   End If
   If ChkWorkDay(DBDATE(strDate)) = False Then
      MsgBox "非工作日不需要檢查打卡是否異常！"
      Exit Sub
   End If
   strDate = DBDATE(strDate)
   
   strKind = UCase(InputBox("請輸入欲檢查那一時段的打卡異常資料：A.上班 P.下班？"))
   If strKind <> "A" And strKind <> "P" Then
      MsgBox "請輸入之時段為 A 或 P！"
      Exit Sub
   End If
   '上班異常檢查不可大於系統日
   If strKind = "A" Then
      '讀取檢查上班打卡異常的起始時間
      strSql = "select * from PollSchedule where ps03 is not null"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         RsTemp.MoveFirst
         Do While Not RsTemp.EOF
            If UCase(RsTemp.Fields("ps03")) = "A" Then
               strPS03_1 = Format(RsTemp.Fields("ps01"), "000000")
            ElseIf UCase(RsTemp.Fields("ps03")) = "A2" Then
               strPS03_2 = Format(RsTemp.Fields("ps01"), "000000")
            ElseIf UCase(RsTemp.Fields("ps03")) = "A3" Then
               strPS03_3 = Format(RsTemp.Fields("ps01"), "000000")
            End If
            RsTemp.MoveNext
         Loop
      End If
      
      If Val(DBDATE(strDate)) > Val(Format(Now, "YYYYMMDD")) Then
         MsgBox "上班異常檢查不可大於系統日！"
         Exit Sub
      ElseIf Val(DBDATE(strDate)) = Val(Format(Now, "YYYYMMDD")) And Val(Format(Time, "HHMMSS")) < Val(strPS03_1) Then
         MsgBox "當日上班異常檢查必須" & Format(Left(strPS03_1, 4), "00:00") & "後才可進行！"
         Exit Sub
      End If
   End If
   '下班異常檢查不可大於等於系統日
   If strKind = "P" And Val(DBDATE(strDate)) >= Val(Format(Now, "YYYYMMDD")) Then
      MsgBox "下班異常檢查不可大於等於系統日！"
      Exit Sub
   End If
   
   intChkTimeKind = 0
   If strKind = "A" Then
      If strST06 = "1" Then
         intChkTimeKind = UCase(InputBox("請輸入欲檢查 [上班] 時段的那一區間上班資料：1.9點前上班 2.11:30上班 3.13:30上班？"))
         If intChkTimeKind <> 1 And intChkTimeKind <> 2 And intChkTimeKind <> 3 Then
            MsgBox "請輸入那一區間上班資料為 1 或 2 或 3！"
            Exit Sub
         End If
         If intChkTimeKind = 2 Then
            If Val(DBDATE(strDate)) = Val(Format(Now, "YYYYMMDD")) And Val(Format(Time, "HHMMSS")) < Val(strPS03_2) Then
               MsgBox "當日上班異常檢查必須" & Format(Left(strPS03_2, 4), "00:00") & "後才可進行！"
               Exit Sub
            End If
         End If
         If intChkTimeKind = 3 Then
            If Val(DBDATE(strDate)) = Val(Format(Now, "YYYYMMDD")) And Val(Format(Time, "HHMMSS")) < Val(strPS03_3) Then
               MsgBox "當日上班異常檢查必須" & Format(Left(strPS03_3, 4), "00:00") & "後才可進行！"
               Exit Sub
            End If
         End If
      Else
         intChkTimeKind = 1
      End If
   End If
   
   Call cmdDown_Click '先執行接收,以免尚有資料未接收下來時,造成比對資料有誤
   
   WLog "（手動）" & strDate & IIf(strKind = "A", " 上班（" & intChkTimeKind & "）", " 下班") & "打卡異常檢查開始！"
   If CheckDataValidate(strDate, strKind, intChkTimeKind) = True Then
      If ChkWorkTime(strDate, strKind, intChkTimeKind, strST06) = True Then
         WLog "（手動）" & strDate & IIf(strKind = "A", " 上班（" & intChkTimeKind & "）", " 下班") & "打卡異常檢查結束！" & strST06Nm & "異常筆數：" & intChkErrRow & " 筆"
      End If
      'Add By Sindy 2025/4/30
      If m_strFileName <> "" Then
         If MsgBox("有加班關心提醒清單，要寄出嗎？", vbYesNo + vbInformation + vbDefaultButton2) = vbYes Then
            Call SaveExcelFile("", "")
         End If
      End If
      '2025/4/30 END
   Else
      MsgBox "輸入資料有誤，請查看Log資訊！"
      Exit Sub
   End If
End Sub

Private Function CheckDataValidate(strDate As String, strKind As String, intChkTimeKind As Integer) As Boolean
   
   CheckDataValidate = False
   
   '必須為工作日才需要檢查是否有打卡異常
   If ChkWorkDay(DBDATE(strDate)) = False Then
      WLog "非工作日不需要檢查打卡是否異常，程式結束！"
      'tmrPolling.Interval = 0
      Exit Function
   End If
   '上班異常檢查不可大於系統日
   If strKind = "A" Then
      If Val(DBDATE(strDate)) > Val(Format(Now, "YYYYMMDD")) Then
         WLog "上班異常檢查不可大於系統日，程式結束！"
         'tmrPolling.Interval = 0
         Exit Function
'      ElseIf Val(DBDATE(strDate)) = Val(Format(Now, "YYYYMMDD")) And Val(Format(Time, "HHMMSS")) < Val("094500") Then
'         WLog "當日上班異常檢查必須9:45後才可進行，程式結束！"
'         'tmrPolling.Interval = 0
'         Exit Function
      End If
   End If
   '下班異常檢查不可大於等於系統日
   If strKind = "P" And Val(DBDATE(strDate)) >= Val(Format(Now, "YYYYMMDD")) Then
      WLog "下班異常檢查不可大於等於系統日，程式結束！"
      'tmrPolling.Interval = 0
      Exit Function
   End If
   
   CheckDataValidate = True
End Function

'Add By Sindy 2013/7/4
'打卡異常檢查
'strKind : A.上班 P.下班
'intChkTimeKind : 因有人員特殊上班時間 99029.iain上班時間為11:30-17:30 96006.朱苡甄上班時間為13:30-20:30
'                 1. 8:00(9:00) - 17:00(18:00)
'                 2. 11:30 - 17:30
'                 3. 13:30 - 20:30
'Modify By Sindy 2013/10/24 +strST06
Private Function ChkWorkTime(strDate As String, strKind As String, intChkTimeKind As Integer, _
                             strST06 As String) As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim strST01 As String, strB1404 As String
Dim min_pr02 As String
Dim max_pr02 As String
Dim chkMinTime As Double
Dim chkMaxTime As Double
Dim bolSendMail As Boolean
Dim bolRest1Day As Boolean, strRestKind As String
Dim bolInsert As Boolean
Dim bolErr As Boolean
Dim strUpdDate As String, strUpdTime As String
Dim bolGetCnn As Boolean
Dim bolExists As Boolean
Dim bolIsAExists As Boolean
Dim strWorkStar As String, strWorkEnd As String
Dim strStarTime As String, strEndTime As String
Dim strStarWorkTime  As String, strEndWorkTime As String
Dim bolUpdateB14Data As Boolean 'Add By Sindy 2016/8/9
Dim strST20 As String, strSubject As String, strContent As String, strUseText As String 'Add By Sindy 2025/4/25
Dim strData As String, varTemp As Variant, i As Integer, strCC As String 'Add By Sindy 2025/5/28
Dim bolOvertime As Boolean 'Add By Sindy 2025/8/19 是否有填寫加班單
   
   ChkWorkTime = False
   
   Screen.MousePointer = vbHourglass
   
   Call SaveExcelFile(strKind, strDate) 'Add By Sindy 2025/4/30
   
   strUpdDate = Format(Now, "YYYYMMDD")
   strUpdTime = Format(Time, "HHMMSS")
   
   'Modify By Sindy 2013/11/11
   'If strST06 = "1" And (strKind = "P" Or intChkTimeKind = 2 Or intChkTimeKind = 3) Then
   If strST06 = "1" And strKind = "A" Then
   '2013/11/11 END
      'Add By Sindy 2013/8/9
      '逐筆檢查是否有系統需要直接確認的
      strSql = "select * from ABS014 where b1411 is null"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         Do While Not rsTmp.EOF
            Call PUB_UpdateB14Data(rsTmp.Fields("b1401"), rsTmp.Fields("b1402"))
            rsTmp.MoveNext
         Loop
      End If
      rsTmp.Close
   End If
   
   intChkErrRow = 0
   '北所在職人員 st06 in('1','2','3','4') / and st03='M51'
   '68007.林信昌上班以實際時間計.
   '63001.董事長及67004.副董事長及法務處律師不用打卡
   '73029.廖宗岳上班時段特殊排除檢查打卡異常
   'Modify By Sindy 2013/8/20 增加員工在職日的判斷
   'Modify By Sindy 2015/9/8 新人A4032.李建鋒在台南暫調高雄支援,上班時段:8:00~16:30,8:30~17:00,9:00~17:30
   'Modify By Sindy 2016/1/7 A4032排除檢查,已取消,開始要檢查其打卡異常
   'Modify By Sindy 2016/10/11 L01.法務處律師只有3個人不打卡,其他都要(and st03 not in ('R04','L01')取消L01)
   'Modify By Sindy 2021/1/29 + "and substr(ST01,4,1)<>'9' " 排除員編第4碼為9的
   'Modify By Sindy 2024/8/6 + 增加排除 'F51','F52' 二個單位 ex:F5724=張舒翔 'Modify By Sindy 2024/9/9 發現
   '                           在新增內翻外翻單位時,SD02就會存入P或F了; 所以不用寫死這二個單位
   If strKind = "P" Then '下班異常檢查
'      strSql = "select st01,st02,st06,a0902,st22,st03,st20" & _
'               " from staff,acc090,SalaryData" & _
'               " where ST01=SD01" & _
'               " and (SD02 not in('P','F') or SD02 is null)" & _
'               " and st01 not in('73029','68007','63001','67004','A2099','" & Replace(Pub_GetSpecMan("不用打卡的律師"), ";", "','") & "')" & _
'               " and st03 not in ('R04') and st06='" & strST06 & "'" & _
'               " and st03=a0901(+) and st04='1'" & _
'               " and st13 is not null and st13<=" & strDate & _
'               " and substr(ST01,4,1)<>'9'"
'      'Modify By Sindy 2025/9/1 增加抓有上班打卡,隔日欲檢查下班時,人員已上離職了
'      strSql = strSql & " union select st01,st02,st06,a0902,st22,st03,st20" & _
'               " from staff,acc090,abs014" & _
'               " where b1403='A' and b1402=" & strDate & _
'               " and b1401=st01(+)" & _
'               " and st03 not in ('R04') and st06='" & strST06 & "'" & _
'               " and st03=a0901(+) and st04='2'" & _
'               " order by st03,st01"
      'Modify By Sindy 2025/9/11 改抓該日期的"刷卡資料"和"異動資料"為資料的來源
      '   1.才能解決上班打卡,隔日離職的狀況
      '   2.和當日到職的狀況
      'Modify By Sindy 2025/9/22 排除來賓 +and substr(scd01,1,4)<>'9999'
      strSql = "select st01,st02,st06,a0902,st22,st03,st20 from (" & _
               "select b1401,b1402 From abs014 where b1402=" & strDate & _
               " Union select distinct scd01 as b1401,pr01 as b1402 From pollrecord, staffcarddata Where pr03=scd02 And pr01=" & strDate & " and substr(scd01,1,4)<>'9999'" & _
               "),staff,acc090" & _
               " where b1401=st01(+) and st06='" & strST06 & "'" & _
               " and st03=a0901(+)" & _
               " and st01 not in('73029','68007','63001','67004','A2099','" & Replace(Pub_GetSpecMan("不用打卡的律師"), ";", "','") & "')" & _
               " order by st03,st01"
      '2025/9/11 END
   Else '上班異常檢查
      If intChkTimeKind = 1 Then
         'Modify By Sindy 2016/10/11 L01.法務處律師只有3個人不打卡,其他都要(and st03 not in ('R04','L01')取消L01)
         'Modify By Sindy 2018/4/27 + A7007.iain上班時間為13:30-17:30
         'Modify By Sindy 2020/6/4 + 鄭皓云(A9004)13:30-17:30 (同Iain)
         'Modify By Sindy 2023/7/26 + B2024劉美英 接朱小姐的工作
         'Modify By Sindy 2024/8/6 + 增加排除 'F51','F52' 二個單位 ex:F5724=張舒翔 'Modify By Sindy 2024/9/9 發現
   '                           在新增內翻外翻單位時,SD02就會存入P或F了; 所以不用寫死這二個單位
         strSql = "select st01,st02,st06,a0902,st22,st03,st20 " & _
                  "from staff,acc090,SalaryData " & _
                  "where ST01=SD01 " & _
                  "and (SD02 not in('P','F') or SD02 is null) " & _
                  "and st01 not in('B2024','A9004','A7007','99029','96006','73029','68007','63001','67004','A2099','" & Replace(Pub_GetSpecMan("不用打卡的律師"), ";", "','") & "') " & _
                  "and st03 not in ('R04') and st06='" & strST06 & "' " & _
                  "and st03=a0901(+) and st04='1' " & _
                  "and st13 is not null and st13<=" & strDate & " " & _
                  "and substr(ST01,4,1)<>'9' " & _
                  "order by st03,st01 "
      '99029.iain上班時間為11:30-17:30
      ElseIf intChkTimeKind = 2 Then
         strSql = "select st01,st02,st06,a0902,st22,st03,st20 " & _
                  "from staff,acc090 " & _
                  "where st01 in('99029') " & _
                  "and st03=a0901(+) and st04='1' " & _
                  "and st13 is not null and st13<=" & strDate & " " & _
                  "and substr(ST01,4,1)<>'9' " & _
                  "order by st03,st01 "
      '96006.朱苡甄上班時間為13:30-20:30
      'Modify By Sindy 2018/4/27 + A7007.iain上班時間為13:30-17:30
      'Modify By Sindy 2020/6/4 + 鄭皓云(A9004)13:30-17:30 (同Iain)
      'Modify By Sindy 2023/7/26 + B2024劉美英 接朱小姐的工作
      ElseIf intChkTimeKind = 3 Then
         strSql = "select st01,st02,st06,a0902,st22,st03,st20 " & _
                  "from staff,acc090 " & _
                  "where st01 in('A9004','A7007','96006','B2024') " & _
                  "and st03=a0901(+) and st04='1' " & _
                  "and st13 is not null and st13<=" & strDate & " " & _
                  "and substr(ST01,4,1)<>'9' " & _
                  "order by st03,st01 "
      End If
   End If
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      Do While Not rsTmp.EOF
         strST01 = rsTmp.Fields("st01")
         strST20 = "" & rsTmp.Fields("st20") 'Add By Sindy 2025/4/25 職稱
'         If strST01 = "A4016" Or strST01 = "99029" Then
'            MsgBox "strST01=" & strST01
'         Else
'            GoTo ReadNext
'         End If
         
         'Modify By Sindy 2013/10/30
'         '工作時段
'         If strST01 = "99029" Then
'            strWorkStar = "11:30"
'            strWorkEnd = "17:30"
'         ElseIf strST01 = "96006" Then
'            strWorkStar = "13:30"
'            strWorkEnd = "20:30"
'         Else
'            strWorkStar = "09:00"
'            strWorkEnd = "17:00"
'         End If
         Call Pub_GetSpecWorkHour(strST01, strDate, strWorkStar, strWorkEnd)
         strWorkStar = Format(strWorkStar, "00:00")
         strWorkEnd = Format(strWorkEnd, "00:00")
         If strWorkStar = "00:00" Or strWorkStar = "" Then
            'Modify By Sindy 2021/5/17
            m_bolByPassWork = PUB_ChkByPassWork(PUB_GetST06(strST01), strDate, "", "", "", "")
            '2021/5/17 END
            strWorkStar = strByPassStarTime(intByPassArea) '最晚的上班時段
            strWorkEnd = strByPassEndTime(1) '最早的下班時段
'            strWorkStar = "09:00"
'            strWorkEnd = "17:00"
         End If
         '2013/10/30 END
         
         strSql = "select nvl(min(pr02),0) as min_pr02,nvl(max(pr02),0) as max_pr02 from pollrecord,staffcarddata where pr03=scd02(+)" & _
                  " and scd01='" & strST01 & "'" & _
                  " and pr01=" & strDate & _
                  " order by pr02 asc"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         min_pr02 = "0"
         max_pr02 = "0"
         If intI = 1 Then
            If RsTemp.RecordCount > 0 Then
               min_pr02 = "" & RsTemp.Fields("min_pr02")
               max_pr02 = "" & RsTemp.Fields("max_pr02")
            End If
         End If
         chkMinTime = Val(Left(Format(min_pr02, "000000"), 4))
         chkMaxTime = Val(Left(Format(max_pr02, "000000"), 4))
         
         '預設值
         bolInsert = False: bolRest1Day = False: strB1404 = "": bolErr = False
         
         '上班異常
         If strKind = "A" Then
            '遲到或未打卡
            If chkMinTime >= Val(Format(strWorkStar, "hhmm")) Or chkMinTime = 0 Then
               bolInsert = True
               If Val(min_pr02) <> 0 Then strB1404 = min_pr02
'               '檢查有無請假資料
'               If CheckIsPersonRestSector(strST01, strDate, "00:00", strDate, "24:00", "") = True Then
'                  bolSendMail = False 'Add By Sindy 2013/9/12 因王副總,有假單不發E-Mail
'                  '取得是否整日休假
'                  '整日休或打卡時間在請假區間內的都示為正常打卡
'                  If Val(strB1404) > 0 Then
'                     If CheckIsPersonRest(strST01, strDate, Left(Format(strB1404, "00:00:00"), 5), strRestKind, bolRest1Day, True, strStarTime, strEndTime, strStarWorkTime, strEndWorkTime) = True Then
'                        bolRest1Day = True
'                     End If
'                  Else
'                     'Modify By Sindy 2013/8/27 未打卡,若整日休假才系統確認
'                     If CheckIsPersonRest(strST01, strDate, strWorkStar, strRestKind, bolRest1Day, True, strStarTime, strEndTime, strStarWorkTime, strEndWorkTime) = True Then
'                        'Modify By Sindy 2013/8/27
'                        'bolRest1Day = True
'                        If bolRest1Day = True Then
'                           bolRest1Day = True
'                        Else
'                           bolRest1Day = False
'                        End If
'                        '2013/8/27 END
'                     End If
'                  End If
''                  '正常區間內請假不發Mail及系統自動確認
''                  If bolRest1Day = True Then
''                     bolSendMail = False
''                  Else
''                     bolSendMail = True
''                  End If
'               Else
'                  bolSendMail = True
'               End If
            End If
         '下班異常
         Else
            '檢查上午是否有異常資料
            bolIsAExists = False
            strSql = "select * from abs014 where b1401='" & strST01 & "'" & _
                     " and b1402=" & strDate & _
                     " and b1403='A'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               If RsTemp.RecordCount > 0 Then
                  bolIsAExists = True
                  min_pr02 = "" & RsTemp.Fields("b1404")
'                  '是否已有補入上班時段
'                  If Val(Format("" & RsTemp.Fields("b1406"), "0000")) > 0 Then
'                     chkMinTime = Val(Format("" & RsTemp.Fields("b1406"), "0000"))
'                  Else
'                     chkMinTime = Val(Left(Format("" & RsTemp.Fields("b1404"), "000000"), 4))
'                  End If
               End If
            End If
            '檢查上下班時段是否有不符規定的
            '若有上班異常就不需檢查下班異常
            If chkMaxTime >= Val(Format(strWorkEnd, "hhmm")) And bolIsAExists = False Then
               bolErr = True '預設值
               If (chkMinTime < Val(Format(strWorkStar, "hhmm")) And chkMinTime <> 0) Then
                  If ChkTaieWorkingHour(chkMinTime, chkMaxTime, strWorkStar, strWorkEnd) = True Then bolErr = False
'                     Else
'                        '檢查上午異常資料是否已有補入上班時段
'                        strSql = "select * from abs014 where b1401='" & strST01 & "'" & _
'                                 " and b1402=" & strDate & _
'                                 " and b1403='A'"
'                        intI = 1
'                        Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'                        If intI = 1 Then
'                           If RsTemp.RecordCount > 0 Then
'                              chkMinTime = Val(Format("" & RsTemp.Fields("b1406"), "0000"))
'                              If ChkTaieWorkingHour(chkMinTime, chkMaxTime) = True Then bolErr = False
'                           End If
'                        End If
               End If
            Else
               'Add By Sindy 2018/1/4 檢查最小的刷卡時間為9點後,未於下午6點後下班者
               'Modify By Sindy 2020/6/10
               If PUB_bWkSpec = True Then
                  If Val(min_pr02) > Val(Format(strWorkStar, "hhmmss")) And Val(max_pr02) < Val(Format(strWorkEnd, "hhmmss")) Then
                     bolInsert = True
                  End If
               Else
               '2020/6/10 END
                  'Modify By Sindy 2021/5/17
                  'If Val(min_pr02) > 90000 And Val(max_pr02) < 180000 Then
                  If Val(min_pr02) > Val(Format(strByPassStarTime(intByPassArea), "HHMM") & "00") And _
                     Val(max_pr02) < Val(Format(strByPassEndTime(intByPassArea), "HHMM") & "00") Then
                  '2021/5/17 END
                     bolInsert = True
                  End If
               End If
               '2018/1/4 END
            End If
            '檢查是否整日未打卡
'            If strST01 = "A1001" Then
'               MsgBox strST01
'            End If
            If Val(min_pr02) = 0 And Val(max_pr02) = 0 Then
               strB1404 = "" '整日未打卡,下班也要出異常
            ElseIf Val(min_pr02) = Val(max_pr02) Then
               strB1404 = "" '下班未打卡
            Else
               strB1404 = max_pr02
            End If
            '早退或未打卡或上下班時段不符規定者
            If chkMaxTime < Val(Format(strWorkEnd, "hhmm")) Or bolErr = True Or Val(strB1404) = 0 Then
               bolInsert = True
'               '檢查有無請假資料
'               If CheckIsPersonRestSector(strST01, strDate, "00:00", strDate, "24:00", "") = True Then
'                  bolSendMail = False 'Add By Sindy 2013/9/12 因王副總,有假單不發E-Mail
'                  '取得是否整日休假
'                  '整日休或打卡時間在請假區間內的都示為正常打卡
'                  If Val(strB1404) > 0 Then
'                     If strB1404 >= 121001 And strB1404 <= 132959 Then '中午休息時間
'                        '中午休息時間下班者
'                        If CheckIsPersonRest(strST01, strDate, "13:30", strRestKind, bolRest1Day, True, strStarTime, strEndTime, strStarWorkTime, strEndWorkTime) = True Then
'                           bolRest1Day = True
'                        End If
'                     Else
'                        If CheckIsPersonRest(strST01, strDate, Left(Format(strB1404, "00:00:00"), 5), strRestKind, bolRest1Day, True, strStarTime, strEndTime, strStarWorkTime, strEndWorkTime) = True Then
'                           bolRest1Day = True
'                        End If
'                     End If
'                  Else
'                     'Modify By Sindy 2013/8/27 未打卡,若整日休假才系統確認
'                     If CheckIsPersonRest(strST01, strDate, strWorkEnd, strRestKind, bolRest1Day, True, strStarTime, strEndTime, strStarWorkTime, strEndWorkTime) = True Then
'                        'Modify By Sindy 2013/8/27
'                        'bolRest1Day = True
'                        If bolRest1Day = True Then
'                           bolRest1Day = True
'                        Else
'                           bolRest1Day = False
'                        End If
'                        '2013/8/27 END
'                     End If
'                  End If
''                  '正常區間內請假不發Mail及系統自動確認
''                  If bolRest1Day = True Then
''                     bolSendMail = False
''                  Else
''                     bolSendMail = True
''                  End If
'               Else
'                  bolSendMail = True
'               End If
            End If
         End If
         '檢查是否已有該筆異常資料存在
         bolExists = False
         strSql = "select * from ABS014 where b1401='" & strST01 & "'" & _
                  " and b1402=" & strDate & _
                  " and b1403='" & strKind & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            If RsTemp.RecordCount > 0 Then
               bolExists = True
            End If
         End If
         '有異常且資料尚未存在時才新增資料
         If bolInsert = True And bolExists = False Then
            If ChkWorkDay(DBDATE(strDate), strST01, True) = True Then '針對個人檢查工作日(颱風假問題)
               cnnConnection.BeginTrans
               bolGetCnn = True
               'Modify By Sindy 2016/8/9 改成先寫異常記錄,再檢核是否需要系統自動核銷
'               strSql = "insert into ABS014(b1401,b1402,b1403,b1404,b1405,b1411,b1412,b1413)" & _
'                        " values(" & CNULL(strST01) & "," & CNULL(strDate) & "," & CNULL(strKind) & "," & _
'                        CNULL(strB1404) & "," & CNULL(IIf(bolRest1Day = True, IIf(strRestKind = "3", "6", "1"), "")) & "," & CNULL(IIf(bolRest1Day = True, "A", "")) & "," & _
'                        CNULL(IIf(bolRest1Day = True, strUpdDate, "")) & "," & CNULL(IIf(bolRest1Day = True, strUpdTime, "")) & ")"
               strSql = "insert into ABS014(b1401,b1402,b1403,b1404)" & _
                        " values(" & CNULL(strST01) & "," & CNULL(strDate) & "," & CNULL(strKind) & "," & CNULL(strB1404) & ")"
               cnnConnection.Execute strSql
               '檢查系統是否自動核銷
               bolUpdateB14Data = PUB_UpdateB14Data(strST01, strDate)
               '2016/8/9 END
               cnnConnection.CommitTrans
               intChkErrRow = intChkErrRow + 1
               bolGetCnn = False
               'Modify By Sindy 2016/7/28 原有請假就不發E-Mail(71011/1020911/上班異常),改為有系統自動核銷才不發E-Mail ex:92012/1050725/下班異常
               'Modify By Sindy 2016/8/22 當時無假單且未核銷者才發E-Mail
               '檢查有無請假資料
               bolSendMail = True
               If CheckIsPersonRestSector(strST01, strDate, "00:00", strDate, "24:00", "") = True Then
                  bolSendMail = False
               End If
               If (bolSendMail = True And bolUpdateB14Data = False) Then '當時無假單且未核銷者才發E-Mail
               '2016/8/22 END
               'If bolRest1Day = False Then
               'If bolUpdateB14Data = False Then
               '2016/7/28 END
                  '81040副所長不發異常E-Mail
                  'Modify By Sindy 2014/9/24 閻副所長未打卡,原不以MAIL通知,修改為要通知
                  'If strST01 <> "81040" Then
                  '2014/9/24 END
                     If Val(DBDATE(strDate)) >= 20130801 Then
                        Call StaffCardErrSendMail(strST01, strDate, strKind, strB1404, chkMinTime, chkMaxTime, min_pr02, max_pr02)
                     End If
                  'End If
   '                     If strKind = "A" Then
   '                        strSubject = ChangeWStringToTDateString(strDate) & " 上班打卡異常通知(整日未打卡)"
   '                     Else
   '                        strSubject = ChangeWStringToTDateString(strDate) & " 下班打卡異常通知"
   '                     End If
   '                     strContent = "打卡時間：" & IIf(strB1404 = "", "未打卡", Format(strB1404, "00:00:00")) & vbCrLf
   '                     If strKind = "P" And bolErr = True Then
   '                        strContent = strContent & "（上下班時段不符規定：上班" & Format(chkMinTime, "00:00") & "下班" & Format(chkMaxTime, "00:00") & "）" & vbCrLf
   '                     End If
   '                     strContent = strContent & "處　　理：請至出缺勤系統執行打卡異常個人處理且(或)請假。" & vbCrLf
   '                     'strContent = strContent & "員工代號：" & strST01 & vbCrLf
   '                     PUB_SendMail "administrator", strST01, "", strSubject, strContent, , , , , , , , , , True
               End If
            End If
         End If
         
         'Add By Sindy 2025/4/25 下班無異常時,才需檢查是否要出【下班提醒通知】
         '（寄發對象排除協理級以上主管（含）與智權人員）
         '排除上班時間為特殊時段的同仁
         'Modify By Sindy 2025/8/9 原排除不寄發關心提醒信件的人員,改為列入給人事處的清單裡。一樣不寄提醒信
         '  如:林美宏律師、蘇英偉律師、蔣瑞安律師、智權部全體同仁，以及林柄佑協理、簡偉倫協理、杜燕文協理。
'         If bolInsert = False And strKind = "P" And _
'            PUB_bWkSpec = False _
'            And Left(Trim(PUB_GetST93(strST01)), 1) <> "S" And _
'            (Val(strST20) > 34 Or strST20 = "") Then
         'Modify By Sindy 2025/10/2 排除11=所長、12=副所長 + And Val(strST20) > 12
         If bolInsert = False And strKind = "P" And _
            PUB_bWkSpec = False And (Val(strST20) > 12 Or strST20 = "") Then
            
            strSubject = "": strContent = "": strCC = ""
            '查詢前一個工作日的出缺勤資料，
            '依據同仁下班時間寄發【下班時間逾30分鐘】、【超過20:00下班】之通知信
            '關心同仁下班情況，提醒若因公務晚下班，記得填寫加班單
            '【同仁加班關心提醒】兩者通知不重複
            If Val(chkMaxTime) >= 2000 Then
               strSubject = "同仁加班關心提醒"
               strContent = "您好：" & vbCrLf & vbCrLf & _
                        "人事處查閱出勤紀錄，看到您" & Mid(strDate, 5, 2) & "/" & Mid(strDate, 7, 2) & "下班時間較晚(超過20:00)，故以本函提醒您，辛苦您了！" & vbCrLf & _
                        "事務所關心同仁在工作與生活上的平衡，維護同仁身心健康，工作上有需要加班時，提醒避免超過20:00" & vbCrLf & _
                        "若是工作負荷較重也請務必與主管提出，以調整工作分配。" & vbCrLf & _
                        "若有其他需協助之處，也請不吝通知人事處。謝謝！" & vbCrLf & vbCrLf & vbCrLf & "　　　　　　　　　　　　人事處啟" & vbCrLf
            '【下班逾30分鐘關懷通知】
            Else
               If Val(chkMinTime) > 0 And Val(chkMaxTime) > 0 Then
                  '用上班時間檢查那一個下班時段
                  bolExists = False
                  For intI = 1 To intByPassArea
                     If Val(chkMinTime & "00") < Val(Format(strByPassStarTime(intI), "HHMM") & "00") Then
                        bolExists = True
                        Exit For
                     End If
                  Next intI
                  strUseText = "您好：" & vbCrLf & vbCrLf & _
                               "人事處查閱出勤紀錄，看到您" & Mid(strDate, 5, 2) & "/" & Mid(strDate, 7, 2) & "下班時間逾正常下班時段30分鐘以上，故以本函提醒您，" & vbCrLf & _
                               "若是因工作需要而加班，請記得填寫加班申請；若是因處理個人事務也請記得在結束工作時先行打卡，以完整記錄工作時間。" & vbCrLf & _
                               "若有其他需協助之處，也請不吝通知人事處。謝謝！" & vbCrLf & vbCrLf & vbCrLf & "　　　　　　　　　　　　人事處啟" & vbCrLf
                  If bolExists = False Then
                     '檢查是否有請假單
                     '假單簽核檔
'                     strSql = "select * from abs010 where b1003='" & strST01 & "' and b1002='01'" & _
'                              " and " & strDate & " between b1004 and b1006" & _
'                              " and b1006 not in('" & 註銷 & "','" & 退回 & "')"
                     'Modify By Sindy 2025/8/8 發現瑞婷114/8/8請單有修改上下班時段,所以要改抓假單主檔
                     strSql = "select sa16,sa17 from Staff_Absence where sa01='" & strST01 & "' and " & strDate & " between sa02 and sa04" & _
                              " Union " & _
                              "select sb17,sb18 from Staff_Busi_Trip where sb01='" & strST01 & "' and " & strDate & " between sb02 and sb04"
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                     If intI = 1 Then
                     '2025/8/8 END
                        '使用假單上的下班時段
                        'If Val("" & RsTemp.Fields("b1029")) > 0 Then
                        If Val("" & RsTemp.Fields("sa17")) > 0 Then
                           'If DateDiff("n", Format(Val(Format(RsTemp.Fields("b1029"), "0000") & "00"), "00:00:00"), Format(Val(chkMaxTime & "00"), "00:00:00")) >= 35 Then
                           'Modify By Sindy 2025/8/19 原先寄發提醒信的時間為逾正常下班時間35分鐘（彈性5分鐘），調整為30分鐘。
                           If DateDiff("n", Format(Val(Format(RsTemp.Fields("sa17"), "0000") & "00"), "00:00:00"), Format(Val(chkMaxTime & "00"), "00:00:00")) >= 30 Then
                              strSubject = "下班逾30分鐘關懷通知"
                              strContent = strUseText
                           End If
                        End If
                     Else
                        '當做應該18:00下班
                        'Modify By Sindy 2025/8/19 原先寄發提醒信的時間為逾正常下班時間35分鐘（彈性5分鐘），調整為30分鐘。
                        If DateDiff("n", Format(Val(Format(strByPassEndTime(3), "HHMM") & "00"), "00:00:00"), Format(Val(chkMaxTime & "00"), "00:00:00")) >= 30 Then
                           strSubject = "下班逾30分鐘關懷通知"
                           strContent = strUseText
                        End If
                     End If
                  Else
                  'If bolExists = True Then
                     'Modify By Sindy 2025/8/19 原先寄發提醒信的時間為逾正常下班時間35分鐘（彈性5分鐘），調整為30分鐘。
                     If DateDiff("n", Format(Val(Format(strByPassEndTime(intI), "HHMM") & "00"), "00:00:00"), Format(Val(chkMaxTime & "00"), "00:00:00")) >= 30 Then
                        strSubject = "下班逾30分鐘關懷通知"
                        strContent = strUseText
                     End If
                  End If
               End If
            End If
            If strContent <> "" Then
               'Modify By Sindy 2025/10/15
               'strContent = strContent & "請至案件管理系統的一般作業\出缺勤作業\表單\下班逾30分鐘原因確認，進行操作。" & vbCrLf
               '2025/10/15 END
               If m_strFileName = "" Then
                  m_strFileName = App.path & "\" & TransDate(strDate, 1) & "加班關心提醒.xls"
                  m_strDate = strDate
                  m_strKind = strKind
                  If Dir(m_strFileName) <> MsgText(601) Then
                     Kill m_strFileName
                     DoEvents
                  End If
                  xlsSalesPoint.SheetsInNewWorkbook = 1
                  xlsSalesPoint.Workbooks.add
                  Set wksaccrpt114 = xlsSalesPoint.Worksheets(1)
                  wksaccrpt114.Columns("a:a").ColumnWidth = 12
                  wksaccrpt114.Columns("b:b").ColumnWidth = 12
                  wksaccrpt114.Columns("c:c").ColumnWidth = 12
                  wksaccrpt114.Columns("d:d").ColumnWidth = 14
                  wksaccrpt114.Columns("e:e").ColumnWidth = 14
                  wksaccrpt114.Columns("f:f").ColumnWidth = 14
                  wksaccrpt114.Range("a1").Value = "員工編號"
                  wksaccrpt114.Range("b1").Value = "姓名"
                  wksaccrpt114.Range("c1").Value = "日期"
                  wksaccrpt114.Range("d1").Value = "打卡起迄時間"
                  wksaccrpt114.Range("e1").Value = "通知種類"
                  wksaccrpt114.Range("f1").Value = "已填加班單"
                  lngCounter = 1
               End If
               lngCounter = lngCounter + 1
               wksaccrpt114.Range("a" & lngCounter).Value = strST01
               wksaccrpt114.Range("a" & lngCounter).HorizontalAlignment = xlLeft '靠左
               wksaccrpt114.Range("b" & lngCounter).Value = GetPrjSalesNM(strST01)
               wksaccrpt114.Range("c" & lngCounter).Value = ChangeWStringToTDateString(strDate)
               wksaccrpt114.Range("d" & lngCounter).Value = Format(Right("0000" & chkMinTime, 4), "00:00") & "~" & Format(Right("0000" & chkMaxTime, 4), "00:00")
               wksaccrpt114.Range("e" & lngCounter).Value = IIf(Val(chkMaxTime) >= 2000, "超過20:00", "下班逾30分鐘")
               
               '檢查是否已填寫加班單
               bolOvertime = False
               strSql = "select so01 from Staff_Overtime where so01='" & strST01 & "'" & _
                        " and so02=" & strDate & _
                        " union " & _
                        "select B1003 from ABS010 where B1002='02' and B1003='" & strST01 & "'" & _
                        " and B1004=" & strDate & " and B1018 not in('" & 註銷 & "')"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strSql)
               If intI = 1 Then
                  bolOvertime = True
               End If
               
               'Modify By Sindy 2025/8/9 原排除不寄發關心提醒信件的人員,改為列入給人事處的清單裡。一樣不寄提醒信
               '  如:林美宏律師、蘇英偉律師、蔣瑞安律師、智權部全體同仁，以及林柄佑協理、簡偉倫協理、杜燕文協理。
               If Left(Trim(PUB_GetST93(strST01)), 1) <> "S" And _
                  (Val(strST20) > 34 Or strST20 = "") Then
               '2025/8/9 END
                  'Add By Sindy 2025/5/28 下班提醒通知信增加「同時發副本給該主管」
                  strData = GetABS001_2(strST01) '取得審核主管(,區隔)
                  If strData <> "" Then
                     '排除L單位, 逐一判斷"簽核主管", 規則為下:
                     If Left(Trim(PUB_GetST93(strST01)), 1) <> "L" Then
                        '1.同仁的下班提醒通知，副本發給簽核主管1和簽核主管2
                        '2.如簽核主管1或2為總經理及閻所長的同仁，下班提醒通知的副本請發給人事處嘉渝副理和佳叡
                        varTemp = Split(strData, ",")
                        For i = 0 To UBound(varTemp)
                           If i > 1 Then Exit For '只抓簽核主管1和2
                           If varTemp(i) = Pub_GetSpecMan("總經理員工編號") Or varTemp(i) = Pub_GetSpecMan("所長員工編號") Then
                              If strCC = "" Then strCC = Pub_GetSpecMan("人事室出缺勤電子簽核")
                              Exit For
                           Else
                              strCC = strCC & ";" & varTemp(i)
                           End If
                        Next i
                        If strCC <> "" And Left(strCC, 1) = ";" Then strCC = Mid(strCC, 2)
   '                     '2.副理（含）(44=代副理)以上的下班提醒通知，副本發給人事處(余佳叡和副理胡嘉渝)
   '                     If Val(GetStaffST20(strST01, 1)) <= 44 And Val(GetStaffST20(strST01, 1)) > 0 Then
   '                        strCC = Pub_GetSpecMan("人事室出缺勤電子簽核")
   '                     '1.一般同仁的下班提醒通知，副本發給其主任與經/副理級（含）(41=經理) 以下之主管
   '                     Else
   '                        varTemp = Split(strData, ",")
   '                        For i = 0 To UBound(varTemp)
   '                           If varTemp(i) <> "" Then
   '                              If Val(GetStaffST20(varTemp(i), 1)) >= 41 Then
   '                                 strCC = strCC & ";" & varTemp(i)
   '                              End If
   '                           End If
   '                        Next i
   '                        If strCC <> "" Then strCC = Mid(strCC, 2)
   '                     End If
                     'L02=法律所, 副本為簽核主管1
                     ElseIf PUB_GetST93(strST01) = "L02" Then
                        '1.台北所麗真副理、陳亮之、涂軼的下班提醒通知，副本發給何娜瑩執行長
                        '2.台中所洪鶯娟、楊世安的下班提醒通知，副本發給蔣所長
                        varTemp = Split(strData, ",")
                        strCC = varTemp(0)
                     End If
                  End If
                  '2025/5/28 END
                  '未填加班單,才寄信
                  If bolOvertime = False Then
                     'strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08) values ('QPGMR','" & strST01 & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss'),'" & strSubject & "','" & strContent & "')"
                     'cnnConnection.Execute strSql, intI
                     PUB_SendMail "QPGMR", strST01, "", ChangeWStringToTDateString(strDate) & strSubject & " (" & strST01 & GetPrjSalesNM(strST01) & " " & Format(Right("0000" & chkMinTime, 4), "00:00") & "~" & Format(Right("0000" & chkMaxTime, 4), "00:00") & ")", strContent, , , , , , strCC, , , , True, False
                  End If
               'Modify By Sindy 2025/8/9 +
               End If
               If bolOvertime = True Then wksaccrpt114.Range("f" & lngCounter).Value = "Y"
               '2025/8/9 END
               'Add By Sindy 2025/10/15 記錄ABS015下班逾30分鐘原因確認
               strSql = "insert into ABS015(b1501,b1502) values(" & CNULL(strST01) & "," & CNULL(strDate) & ")"
               cnnConnection.Execute strSql
               '2025/10/15 END
            End If
         End If
         '2025/4/25 END
ReadNext:
         rsTmp.MoveNext
      Loop
   Else
      WLog "資料庫中搜尋不到符合資料！"
      rsTmp.Close
      Set rsTmp = Nothing
      Screen.MousePointer = vbDefault
      Exit Function
   End If
   rsTmp.Close
   
   PUB_SendMailCache 'Add By Sindy 2025/4/25
   ChkWorkTime = True
   Screen.MousePointer = vbDefault
   Set rsTmp = Nothing
   Exit Function

ErrHand:
   Screen.MousePointer = vbDefault
   If bolGetCnn = True Then
      cnnConnection.RollbackTrans
   End If
   WLog "新增失敗！" & Err.Description & "；" & strSql
End Function

'Add By Sindy 2025/4/30
Private Sub SaveExcelFile(strKind As String, strDate As String)
   If m_strFileName <> "" And (m_strKind <> strKind Or m_strDate <> strDate) Then
      'xlsSalesPoint.Workbooks(1).SaveAs FileName:=strFileName
      If Val(xlsSalesPoint.Version) < 12 Then
         xlsSalesPoint.Workbooks(1).SaveAs FileName:=m_strFileName, FileFormat:=-4143
      Else
         xlsSalesPoint.Workbooks(1).SaveAs FileName:=m_strFileName, FileFormat:=56
      End If
      xlsSalesPoint.Workbooks.Close
      xlsSalesPoint.Quit
      DoEvents
      Set wksaccrpt114 = Nothing
      Set xlsSalesPoint = Nothing
      PUB_SendMail "QPGMR", Pub_GetSpecMan("人事室出缺勤電子簽核"), "", ChangeWStringToTDateString(strDate) & "加班關心提醒清單", "同主旨", , m_strFileName, , , , , , , , True, False
      m_strFileName = "": m_strDate = "": m_strKind = ""
   End If
End Sub

'Add By Sindy 2013/8/27
Private Function GetPollData(strDate As String, strUserId As String) As Boolean
Dim Rs As New ADODB.Recordset

On Error GoTo ErrHand
   
   GetPollData = False
   strSql = "select scd01,pr01,pr02 from pollrecord,staffcarddata" & _
             " where pr03=scd02(+) and pr01=" & DBDATE(strDate) & _
             " and scd01='" & strUserId & "'"
   Rs.CursorLocation = adUseClient
   Rs.Open strSql, cnnConnection
   If Rs.RecordCount > 0 Then
      GetPollData = True
   End If
   
   Rs.Close
   Set Rs = Nothing
   Exit Function
   
ErrHand:
   MsgBox Err.Description
End Function

'設定User Data至Session
Private Function ClsPDSetUserData(ByRef strUserNum As String, ByRef strUserName As String, ByRef strGroup As String) As Boolean
Dim lngRt As Long, strUser As String * 100, a As String
Dim strSql As String, rsRecordset As New ADODB.Recordset

On Error GoTo ErrHand
'lngRt = WNetGetUser("", strUser, 10)
'lngRt = 0
'If lngRt = 0 Then
   strUserNum = "QPGMR"
   'strUserNum = "74001"
   strSql = "select st04,st02,st11 from staff where upper(st01)=" + CNULL(strUserNum)
   rsRecordset.CursorLocation = adUseClient
   rsRecordset.Open strSql, cnnConnection
   If rsRecordset.RecordCount > 0 Then
      If rsRecordset.Fields(0) = "1" Then
         strSql = "begin " + _
            "select st02,st03,st05,st11 into user_data.user_name,user_data.user_department," + _
            "user_data.user_level,user_data.user_group from staff where upper(st01)=" + CNULL(strUserNum) + ";" + _
            "user_data.user_num:=" + CNULL(strUserNum) + ";" + _
            "end;"
         cnnConnection.Execute strSql
         strUserName = IIf(IsNull(rsRecordset.Fields(1)), "", rsRecordset.Fields(1))
         strGroup = IIf(IsNull(rsRecordset.Fields(2)), "", rsRecordset.Fields(2))
         ClsPDSetUserData = True
      Else
         ShowMsg MsgText(9165)
      End If
   Else
      ShowMsg MsgText(9166)
   End If
   rsRecordset.Close
'Else
'   ShowMsg MsgText(9167)
'End If
Exit Function
ErrHand:
   'edit by nickc 2007/02/02
   'ErrorLog
   MsgBox Err.Description
End Function

Function WLog(oStrLog As String)
   Dim ffa As Integer
   Dim strNow As String
   Dim stLogFolder As String, stLogFile As String 'Added by Morgan 2019/7/10
   
   'Modified by Morgan 2019/7/10 避免太大且需要維護,改放資料夾,檔名用年週,只保留1年
   stLogFolder = App.path & "\" & App.EXEName & "Log"
   If Dir(stLogFolder, vbDirectory) = "" Then
      MkDir stLogFolder
   End If
   
   'log保留一年(清除前一年的log)
   stLogFile = stLogFolder & "\" & (Format(Now, "yyyyww") - 100) & ".log"
   If Dir(stLogFile) <> "" Then
      Kill stLogFile
   End If
   stLogFile = stLogFolder & "\" & (Format(Now, "yyyyww")) & ".log"
   
   strNow = Trim(Now)
   '寫在畫面上
   'Modified by Morgan 2019/7/12 空白不列
   If oStrLog <> "" Then
      lstHistory.AddItem strNow & "  -->  " & oStrLog, 0
   End If
   '寫在文字檔
   ffa = FreeFile
   
   'Open App.path & "\HTAautolog.log" For Append As ffa
   Open stLogFile For Append As ffa
   
   'Modified by Morgan 2019/7/12 空白跳行就好
   If oStrLog = "" Then
      Print #ffa, ""
   Else
      Print #ffa, strNow & "  ==>  " & oStrLog
   End If
   Close ffa
End Function

'Added by Morgan 2019/7/10
'下載台銀當日結匯匯率CSV檔並匯入
Private Function ImportTWBankRate(Optional pErrMsg As String) As Boolean
   Dim strText As String, strErr As String
   
   WLog "下載 台銀當日結匯匯率CSV檔(" & strSrvDate(2) & ")"
   strText = PUB_GetTwBankRate(Inet1, strSrvDate(2), False, strErr)
   If strText = "" Then
      pErrMsg = strErr
      WLog strErr
      Exit Function
   Else
      WLog "匯入 台銀當日結匯匯率CSV檔"
      If PUB_ImportRate(strText, strSrvDate(2), False, strErr) = True Then
         WLog strSrvDate(2) & "匯率已更新!"
      Else
         pErrMsg = strErr
         WLog strErr
         Exit Function
      End If
   End If
   ImportTWBankRate = True
   
End Function
