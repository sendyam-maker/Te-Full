VERSION 5.00
Begin VB.Form frm084002 
   BorderStyle     =   1  '單線固定
   Caption         =   "法務委任案件統計表"
   ClientHeight    =   2190
   ClientLeft      =   795
   ClientTop       =   2115
   ClientWidth     =   5475
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   5475
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   3
      Left            =   1440
      MaxLength       =   1
      TabIndex        =   3
      Top             =   1608
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   2
      Left            =   1440
      MaxLength       =   1
      TabIndex        =   2
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   1
      Left            =   2760
      MaxLength       =   7
      TabIndex        =   1
      Top             =   648
      Width           =   975
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   4464
      TabIndex        =   5
      Top             =   144
      Width           =   800
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   400
      Left            =   3612
      TabIndex        =   4
      Top             =   144
      Width           =   800
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   0
      Left            =   1440
      MaxLength       =   7
      TabIndex        =   0
      Top             =   648
      Width           =   975
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
      Height          =   345
      Left            =   330
      TabIndex        =   10
      Top             =   180
      Width           =   3045
   End
   Begin VB.Label lblSysKind 
      Caption         =   "(1.法務 2.顧問 3.全部)"
      Height          =   252
      Left            =   2160
      TabIndex        =   9
      Top             =   1128
      Width           =   1932
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "報  表  別：                     (1.刑事 2.民事 3.強制執行 4.全部)"
      Height          =   180
      Index           =   2
      Left            =   360
      TabIndex        =   8
      Top             =   1608
      Width           =   4440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "系統類別︰"
      Height          =   180
      Index           =   1
      Left            =   360
      TabIndex        =   7
      Top             =   1128
      Width           =   900
   End
   Begin VB.Line Line1 
      X1              =   2520
      X2              =   2640
      Y1              =   768
      Y2              =   768
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "發文日期："
      Height          =   180
      Index           =   0
      Left            =   360
      TabIndex        =   6
      Top             =   648
      Width           =   900
   End
End
Attribute VB_Name = "frm084002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/08/31 法務系統的工作點數分配功能先上線(110/9/1)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/3 日期欄已修改
Option Explicit

Dim PLeft(0 To 12) As Integer
Dim m_print1 As Integer
Dim m_print2 As Integer
Dim m_print3 As Integer
'Add By Cheng 2002/09/10
Dim blnClkSure As Boolean

Private Sub cmdBack_Click()
   Unload Me
End Sub

Private Sub cmdPrint_Click()
   'Add By Cheng 2002/09/10
   blnClkSure = False

   m_print1 = 0
   m_print2 = 0
   m_print3 = 0
   
   If ChkRange(Text1(0), Text1(1), "日期") = False Then
      blnClkSure = True
      Me.Text1(0).SetFocus
      Text1_GotFocus 0
      Exit Sub
   End If
   'Add By Cheng 2002/03/22
   If PUB_CheckKeyInDate(Me.Text1(0)) = -1 Then
      Me.Text1(0).SetFocus
      Text1_GotFocus 0
      Exit Sub
   End If
   If PUB_CheckKeyInDate(Me.Text1(1)) = -1 Then
      Me.Text1(1).SetFocus
      Text1_GotFocus 1
      Exit Sub
   End If
   
   If Text1(2) = "" Then
      MsgBox "系統別不得為空值 !", vbCritical
      Text1(2).SetFocus
      Exit Sub
   End If
   If Text1(3) = "" Then
      MsgBox "報表別不得為空值 !", vbCritical
      Text1(3).SetFocus
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   GetPrintLeft
   Select Case Text1(3).Text
      Case "1"
         PrintCase1
      Case "2"
         PrintCase2
      Case "3"
         PrintCase3
      Case "4"
         PrintCase1
         PrintCase2
         PrintCase3
   End Select
   Screen.MousePointer = vbDefault
   Select Case Text1(3).Text
          Case "1"
                If m_print1 = 0 Then
                   MsgBox "列印結束!", vbInformation
                ElseIf m_print1 = 1 Then
                   MsgBox "查無資料!", vbInformation
                End If
          Case "2"
                If m_print2 = 0 Then
                   MsgBox "列印結束!", vbInformation
                ElseIf m_print2 = 1 Then
                   MsgBox "查無資料!", vbInformation
                End If
          Case "3"
                If m_print3 = 0 Then
                   MsgBox "列印結束!", vbInformation
                ElseIf m_print3 = 1 Then
                   MsgBox "查無資料!", vbInformation
                End If
                
          Case "4"
               If m_print1 = 0 Or m_print2 = 0 Or m_print3 = 0 Then
                  MsgBox "列印結束!", vbInformation
               ElseIf m_print1 = 1 And m_print2 = 1 And m_print3 = 1 Then
                  MsgBox "查無資料!", vbInformation
               End If
  End Select
End Sub

Private Sub PrintCase1()
 Dim i As Integer, j As Integer, St As String, Page As Integer, iPrint As Integer
 Dim TmpArea As String
 Dim StS(1 To 11) As String, StRng(0 To 9) As String
 Dim stTmp(1 To 12) As String, StN(1 To 4) As String
 'Dim Wo As DAO.Workspace, Db As DAO.Database, Rc As DAO.Recordset, Rc1 As DAO.Recordset 'Remove by Lydia 2020/04/10
 Dim Qty As String
 Dim strSqlq As String
 'Dim rsTmpQ As DAO.Recordset 'Remove by Lydia 2020/04/10
 Dim strSql As String
 Dim rsTmp As New ADODB.Recordset
 Dim strCP01 As String
 Dim strCP02 As String
 Dim strCP03 As String
 Dim strCP04 As String
 Dim strCP18 As String
 'Added by Lydia 2020/04/10 暫存檔的序號
 Dim mESeqNo As String
 Dim xRows As Integer
 Dim nCnt As Integer 'Added by Lydia 2020/04/17
 
On Error GoTo ErrHand

   'Modified by Lydia 2020/04/10 改用暫存檔Rdatafactory
'   If CreateDatabase = False Then
'      m_print1 = 3
'      MsgBox "無法建立暫存區，列印失敗 !", vbInformation
'      Exit Sub
'   End If
'   Set Wo = DBEngine.Workspaces(0)
'   Set Db = Wo.OpenDatabase(App.path & "\Case.mdb", False, False, ";PWD=taie")
'   Qty = "DELETE FROM TEMP"
'   Db.Execute Qty
'   If RsTemp.State = adStateOpen Then RsTemp.Close
'   Qty = "INSERT INTO TEMP (TMP01,TMP02,TMP03,TMP04,TMP05,TMP06,TMP07,TMP08,TMP09," & _
'      "TMP10,TMP11,TMP12,TMP13) VALUES ('刑事案件','專利','總點數','平均'," & _
'      "'商標','總點數','平均','著作權','總點數','平均','其他案件','總點數','平均')"
'   Db.Execute Qty
   intI = 1
   Qty = "select '刑事案件','專利','總點數','平均','商標','總點數','平均','著作權','總點數','平均','其他案件','總點數','平均'  from dual"
   Set RsTemp = ClsLawReadRstMsg(intI, Qty)
   Set rsTmp = PUB_CreateRecordset(RsTemp, , , , Me.Name, mESeqNo)
   xRows = xRows + 1
   rsTmp.Close
   'end 2020/04/10
   
   StS(1) = "告訴(取締)":  StS(2) = "告訴(不取締)": StS(3) = "地檢署辯護"
   StS(4) = "自訴(含辯護)": StS(5) = "一審(含辯護)": StS(6) = "二審(含辯護)"
   StS(7) = "三審(含辯護)": StS(8) = "再審": StS(9) = "非常上訴"
   StS(10) = "合計": StS(11) = "平均"
   StRng(0) = " AND NOT (CP10 IN ('2101','2102','2103','2104','2105','2106','2108'," & _
      "'2121','2122','2123','2124','2201','2202','2203','2204','2205','2206','2207'," & _
      "'2109','2123','2124','2137','2139','2140'))"
   StRng(1) = " CP10 IN ('2101','2102','2103','2104','2105','2106','2108') AND CP21='Y'"
   StRng(2) = " CP10 IN ('2101','2102','2103','2104','2105','2106','2108') AND CP21 IS NULL"
   StRng(3) = " CP10 IN ('2121','2122','2123','2124')"
   StRng(4) = " CP10 IN ('2201','2202','2203','2204','2205','2206','2207')"
   StRng(5) = " CP10 IN ('2109','2137') AND CP71='00001'"
   StRng(6) = " CP10 IN ('2109','2137') AND CP71='00002'"
   StRng(7) = " CP10 IN ('2109','2137') AND CP71='00003'"
   StRng(8) = " CP10 IN ('2139')"
   StRng(9) = " CP10 IN ('2140')"
   StN(1) = " AND CR05 IN ('P','FCP','CFP')) "
   StN(2) = " AND CR05 IN ('T','FCT','CFT','TF')) "
   StN(3) = " AND CR05 IN ('TC','CFC')) "
   StN(4) = " AND CR05 NOT IN ('P','FCP','CFP','T','FCT','CFT','TF','TC','CFC')) "

   For i = 1 To 9
'      strExc(1) = strExc(0) & StRng(i)
      'Modify By Cheng 2002/03/26
      '多加CP09<'C'控制
'       strSQL = "SELECT CP01,CP02,CP03,CP04,CP18 FROM CASEPROGRESS WHERE " & _
'                 StRng(i) & GetSQL
       'Modify by Morgan 2011/6/1 若有建點數分配資料時點數改分配點數(目前L會有分配) cp18->nvl(a0n03/1000,cp18)
       'Modified by Lydia 2015/06/02 +工作分配點數
       'strSql = "SELECT CP01,CP02,CP03,CP04,nvl(a0n03/1000,cp18) CP18 FROM CASEPROGRESS ,acc0n0 where a0n02(+)=cp09 and  CP09<'C' AND " & _
                 StRng(i) & GetSql
       strExc(4) = "SELECT CP01,CP02,CP03,CP04,cp09,nvl(a0n03/1000,cp18) CP18,sum(decode(p1.a1n01,null,decode(p3.a1n05,null,0,p3.a1n05),p1.a1n05)) a07 " & _
                   "FROM CASEPROGRESS ,acc0n0,acc1n0 p1,acc1n0 p3 where a0n02(+)=cp09 and  CP09<'C' AND " & _
                 StRng(i) & GetSql & " AND p1.a1n02(+)='3' AND p1.a1n03(+)=cp09 " & _
                 "AND p3.a1n02(+)='2' AND p3.a1n03(+)=cp09 AND p3.a1n01(+)=cp60 " & _
                 "group by CP01,CP02,CP03,CP04,cp09,nvl(a0n03/1000,cp18)"
       strSql = "SELECT CP01,CP02,CP03,CP04,decode(a07,0,cp18,a07) CP18 from (" & strExc(4) & ") "
        
       If rsTmp.State = adStateOpen Then rsTmp.Close  'Added by Lydia 2020/04/10
       rsTmp.Open strSql, cnnConnection
       strCP01 = ""
       strCP02 = ""
       strCP03 = ""
       strCP04 = ""
       strCP18 = ""
       'Move by Lydia 2020/04/10 從j = 1 To 4前面，移過來
       For j = 1 To 12
          stTmp(j) = "0"
       Next
       'end 2020/04/10
       If rsTmp.EOF = False Then
         Do While rsTmp.EOF = False
            If Not IsNull(rsTmp.Fields("CP01")) Then
               strCP01 = rsTmp.Fields("CP01")
            Else
               strCP01 = ""
            End If
            If Not IsNull(rsTmp.Fields("CP02")) Then
               strCP02 = rsTmp.Fields("CP02")
            Else
               strCP02 = ""
            End If
            If Not IsNull(rsTmp.Fields("CP03")) Then
               strCP03 = rsTmp.Fields("CP03")
            Else
               strCP03 = ""
            End If
            If Not IsNull(rsTmp.Fields("CP04")) Then
               strCP04 = rsTmp.Fields("CP04")
            Else
               strCP04 = ""
            End If
            If Not IsNull(rsTmp.Fields("CP18")) Then
               strCP18 = rsTmp.Fields("CP18")
            Else
               strCP18 = "0"
            End If
            
            nCnt = 0 'Added by Lydia 2020/04/17
            
            For j = 1 To 4
                'Modified by Lydia 2020/04/10 CASERELATION=>CASERELATION1
                'Added by Lydia 2020/04/17 相關案件都只設在最初收文的案號, 之後一審,二審的案號為L-xxx-1-00; 所以要另外抓
                If strCP03 <> "0" Then
                    strExc(2) = " select COUNT(*) from (SELECT DISTINCt cr01,cr02,cr03,cr04 FROM CASERELATION1 WHERE " & _
                            " CR01 ='" & strCP01 & "'" & _
                            " AND CR02 ='" & strCP02 & "'" & _
                            " AND CR03 ='0'" & _
                            " AND CR04 ='" & strCP04 & "'" & StN(j)
                Else
                'end 2020/04/17
                    strExc(2) = " select COUNT(*) from (SELECT DISTINCt cr01,cr02,cr03,cr04 FROM CASERELATION1 WHERE " & _
                            " CR01 ='" & strCP01 & "'" & _
                            " AND CR02 ='" & strCP02 & "'" & _
                            " AND CR03 ='" & strCP03 & "'" & _
                            " AND CR04 ='" & strCP04 & "'" & StN(j)
                End If 'Added by Lydia 2020/04/17
               If RsTemp.State = adStateOpen Then RsTemp.Close  'Added by Lydia 2020/04/10
               RsTemp.Open strExc(2), cnnConnection
               If RsTemp.Fields(0) > 0 Then
                  nCnt = nCnt + 1 'Added by Lydia 2020/04/17 記錄是否有相關案件記錄
                  'TMP01,TMP04,TMP07,TMP10
                  If Not IsNull(RsTemp.Fields(0)) Then
                     stTmp(3 * j - 2) = CInt(stTmp(3 * j - 2)) + RsTemp.Fields(0)
                  Else
                     stTmp(3 * j - 2) = CInt(stTmp(3 * j - 2)) + 0
                  End If
                  'TMP02,TMP05,TMP08,TMP11
                  If strCP18 <> "" Then
                     stTmp(3 * j - 1) = CInt(stTmp(3 * j - 1)) + CInt(strCP18)
                  Else
                     stTmp(3 * j - 1) = CInt(stTmp(3 * j - 1)) + 0
                  End If
                  'TMP03,TMP06,TMP09,TMP12
                  If Not IsNull(RsTemp.Fields(0)) Then
                     stTmp(3 * j) = CInt(strCP18) / RsTemp.Fields(0)
                  Else
                     stTmp(3 * j) = CInt(stTmp(3 * j)) + 0
                  End If
               'Added by Lydia 2020/04/17 找不到關聯,全部歸到其他
               ElseIf j = 4 And nCnt = 0 Then
                      stTmp(3 * j - 2) = CInt(stTmp(3 * j - 2)) + 1
                      stTmp(3 * j - 1) = CInt(stTmp(3 * j - 1)) + CInt(strCP18)
                      stTmp(3 * j) = stTmp(3 * j - 1) / stTmp(3 * j - 2)
               'end 2020/04/17
               End If
              RsTemp.Close
            Next
            rsTmp.MoveNext
         Loop
         rsTmp.Close
       Else
            'Remove by Lydia 2020/04/10
            'For j = 1 To 12
            '   stTmp(j) = "0"
            'Next
            'end 2020/04/10
            rsTmp.Close
       End If
      'Modified by Lydia 2020/04/10 改用暫存檔Rdatafactory
      'Qty = "INSERT INTO TEMP (TMP01,TMP02,TMP03,TMP04,TMP05,TMP06,TMP07,TMP08," & _
         "TMP09,TMP10,TMP11,TMP12,TMP13) VALUES ('" & StS(i) & "','" & stTmp(1) & "','" & _
         stTmp(2) & "','" & stTmp(3) & "','" & stTmp(4) & "','" & stTmp(5) & "','" & _
         stTmp(6) & "','" & stTmp(7) & "','" & stTmp(8) & "','" & stTmp(9) & "','" & _
         stTmp(10) & "','" & stTmp(11) & "','" & stTmp(12) & "')"
      'Db.Execute Qty
      xRows = xRows + 1
      Qty = "insert into rdatafactory(formname,id,seqno,rowseq,r001,r002,r003,r004,r005,r006,r007,r008,r009,r010,r011,r012,r013) " & _
               "values ('" & Me.Name & "', '" & strUserNum & "', '" & mESeqNo & "', " & xRows & ",'" & StS(i) & "','" & stTmp(1) & "','" & _
                stTmp(2) & "','" & stTmp(3) & "','" & stTmp(4) & "','" & stTmp(5) & "','" & stTmp(6) & "','" & stTmp(7) & "','" & stTmp(8) & "','" & stTmp(9) & "','" & _
                stTmp(10) & "','" & stTmp(11) & "','" & stTmp(12) & "')"
      cnnConnection.Execute Qty
      'end 2020/04/10
   Next
   'Modified by Lydia 2020/04/10 改用暫存檔Rdatafactory
   'Qty = "INSERT INTO TEMP (TMP01,TMP02,TMP03,TMP04,TMP05,TMP06,TMP07,TMP08," & _
         "TMP09,TMP10,TMP11,TMP12,TMP13) SELECT '" & StS(10) & "',SUM(VAL(TMP02)),SUM(VAL(TMP03)),SUM(VAL(TMP04)),SUM(VAL(TMP05))," & _
      "SUM(VAL(TMP06)),SUM(VAL(TMP07)),SUM(VAL(TMP08)),SUM(VAL(TMP09)),SUM(VAL(TMP10))," & _
      "SUM(VAL(TMP11)),SUM(VAL(TMP12)),SUM(VAL(TMP13)) FROM TEMP WHERE TMP01<>'刑事案件'"
   'Db.Execute Qty
   'Qty = "INSERT INTO TEMP (TMP01,TMP02,TMP03,TMP04,TMP05,TMP06,TMP07,TMP08," & _
      "TMP09,TMP10,TMP11,TMP12,TMP13) SELECT '" & StS(11) & "',FORMAT(TMP02/9,""0.0""),FORMAT(TMP03/9,""0.0""),FORMAT(TMP04/9,""0.0""),FORMAT(TMP05/9,""0.0"")," & _
      "FORMAT(TMP06/9,""0.0""),FORMAT(TMP07/9,""0.0""),FORMAT(TMP08/9,""0.0""),FORMAT(TMP09/9,""0.0""),FORMAT(TMP10/9,""0.0"")," & _
      "FORMAT(TMP11/9,""0.0""),FORMAT(TMP12/9,""0.0""),FORMAT(TMP13/9,""0.0"") FROM TEMP WHERE TMP01='" & StS(10) & "'"
   'Db.Execute Qty
      xRows = xRows + 1
      Qty = "insert into rdatafactory(formname,id,seqno,rowseq,r001,r002,r003,r004,r005,r006,r007,r008,r009,r010,r011,r012,r013) " & _
               " SELECT '" & Me.Name & "', '" & strUserNum & "', '" & mESeqNo & "', " & xRows & ",'" & StS(10) & "',SUM(R002) as R002,SUM(R003) as R003,SUM(R004) as R004,SUM(R005) as R005," & _
                "SUM(R006) as R006,SUM(R007) as R007,SUM(R008) as R008,SUM(R009) as R009,SUM(R010) as R010," & _
                "SUM(R011) as R011,SUM(R012) as R012,SUM(R013) as R013 FROM rdatafactory WHERE R001<>'刑事案件' " & _
                "AND formname='" & Me.Name & "' and id='" & strUserNum & "' and seqno='" & mESeqNo & "' "
      cnnConnection.Execute Qty
      xRows = xRows + 1
      Qty = "insert into rdatafactory(formname,id,seqno,rowseq,r001,r002,r003,r004,r005,r006,r007,r008,r009,r010,r011,r012,r013) " & _
               " SELECT '" & Me.Name & "', '" & strUserNum & "', '" & mESeqNo & "','" & xRows & "' ,'" & StS(11) & "',round((R002 / 9), 1) as r002,round((R003 / 9), 1) as r003,round((R004 / 9), 1) as r004,round((R005 / 9), 1) as r005," & _
               "round((R006 / 9), 1) as r006, round((R007 / 9), 1) as r007, round((R008 / 9), 1) as r008, round((R009 / 9), 1) as r009, round((R010 / 9), 1) as r010, round((R011 / 9), 1) as r011, round((R012 / 9), 1) as r012,round((R013 / 9), 1) as r013 " & _
               "FROM rdatafactory  WHERE R001='" & StS(10) & "' AND formname='" & Me.Name & "' and id='" & strUserNum & "' and seqno='" & mESeqNo & "' "
      cnnConnection.Execute Qty
   'end 2020/04/10
   
   'Modified by Lydia 2020/04/10 改用暫存檔Rdatafactory
   'strSqlq = "SELECT cint(iif(ISNULL(TMP02),0,TMP02))+cint(iif(ISNULL(TMP03),0,TMP03))+cint(iif(ISNULL(TMP04),0,TMP04))+cint(iif(ISNULL(TMP05),0,TMP05))" & _
             "+cint(iif(ISNULL(TMP06),0,TMP06))+cint(iif(ISNULL(TMP07),0,TMP07))+cint(iif(ISNULL(TMP08),0,TMP08))+cint(iif(ISNULL(TMP09),0,TMP09))+cint(iif(ISNULL(TMP06),0,TMP06))" & _
             "+cint(iif(ISNULL(TMP11),0,TMP11))+cint(iif(ISNULL(TMP12),0,TMP12))+cint(iif(ISNULL(TMP13),0,TMP13)) FROM TEMP WHERE TMP01='合計'"
   'Set rsTmpQ = Db.OpenRecordset(strSqlq)
   'If rsTmpQ.EOF = False Then
   strSqlq = "select R002+R003+R004+R005+R006+R007+R008+R009+R010+R011+R012+R013 FROM rdatafactory WHERE R001='合計' AND formname='" & Me.Name & "' and id='" & strUserNum & "' and seqno='" & mESeqNo & "' "
   intI = 1
   Set rsTmp = ClsLawReadRstMsg(intI, strSqlq)
   If intI = 1 Then
   'end 2020/04/10
      'Modified by Lydia 2020/04/10 rsTmpQ=>rsTmp
      If Not IsNull(rsTmp.Fields(0)) Then
         If rsTmp.Fields(0) = 0 Then
            m_print1 = 1
            rsTmp.Close
            Exit Sub
         End If
      Else
          m_print1 = 1
          rsTmp.Close
          Exit Sub
      End If
   End If
   
   i = 500
   Printer.KillDoc
   Printer.Orientation = 2
   Printer.Font.Size = 22
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.CurrentX = 6000:         Printer.CurrentY = i
   Printer.Print "刑事案件統計表"
   Printer.Font.Underline = False
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   Printer.CurrentX = 6200:         Printer.CurrentY = i + 500
   Printer.Print "發文日期 : " & ChangeTStringToTDateString(Text1(0)) & _
      " - " & ChangeTStringToTDateString(Text1(1))
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   Printer.CurrentX = 500:               Printer.CurrentY = i + 800
   Printer.Print "列印人 : " & strUserName
   Printer.CurrentX = 13000:             Printer.CurrentY = i + 800
   Printer.Print "列印日期 : " & ChangeTStringToTDateString(ChangeWStringToTString(GetTodayDate))
   Printer.CurrentX = 13000:             Printer.CurrentY = i + 1100
   Printer.Print "頁次 : 1"
   iPrint = i + 1400
   Printer.CurrentX = 500:               Printer.CurrentY = iPrint
   Printer.Print String(210, "-")
   iPrint = iPrint + 300
   
   'Modified by Lydia 2020/04/10 改用暫存檔Rdatafactory
   'Qty = "SELECT TMP01,TMP02,TMP03,TMP04,TMP05,TMP06,TMP07,TMP08,TMP09,TMP10," & _
         "TMP11,TMP12,TMP13 FROM TEMP"
   'Set Rc = Db.OpenRecordset(Qty)
   'With Rc
   Qty = "select r001,r002,r003,r004,r005,r006,r007,r008,r009,r010,r011,r012,r013 from rdatafactory WHERE  formname='" & Me.Name & "' and id='" & strUserNum & "' and seqno='" & mESeqNo & "' order by rowseq "
   intI = 1
   Set rsTmp = ClsLawReadRstMsg(intI, Qty)
   If intI = 1 Then
        With rsTmp
           .MoveFirst
   'end 2020/04/10
           Do While Not .EOF
              If .Fields(0) <> "平均" Then
                    Printer.CurrentX = PLeft(0):   Printer.CurrentY = iPrint
                    Printer.Print .Fields(0)
                    For i = 1 To 12
                       If .Fields(0) = StS(11) Then
                          Printer.CurrentX = PLeft(i) - Printer.TextWidth(Format(.Fields(i), "0.0")):   Printer.CurrentY = iPrint
                          Printer.Print Format(.Fields(i), "0.0")
                       Else
                          Select Case i
                             Case 1, 4, 7, 10
                                Printer.CurrentX = PLeft(i) - Printer.TextWidth(Format(.Fields(i), "0")):   Printer.CurrentY = iPrint
                                Printer.Print Format(.Fields(i), "0")
                             Case 2, 3, 5, 6, 8, 9, 11, 12
                                Printer.CurrentX = PLeft(i) - Printer.TextWidth(Format(.Fields(i), "0.0")):   Printer.CurrentY = iPrint
                                Printer.Print Format(.Fields(i), "0.0")
                          End Select
                       End If
                       
                    Next
                    iPrint = iPrint + 300
             End If
              .MoveNext
           Loop
        End With
        Printer.CurrentX = 500:          Printer.CurrentY = iPrint
        Printer.Print String(210, "-")
        Printer.EndDoc
   End If 'Added by Lydia 2020/04/10
   
   Exit Sub
ErrHand:
   MsgBox Err.Description
End Sub

Private Sub PrintCase2()
 Dim i As Integer, j As Integer, St As String, Page As Integer, iPrint As Integer
 Dim StS(1 To 6) As String, StRng(1 To 4) As String
 Dim stTmp(1 To 3) As String
 'Dim Wo As DAO.Workspace, Db As DAO.Database, Rc As DAO.Recordset, Rc1 As DAO.Recordset 'Remove by Lydia 2020/04/10
 Dim Qty As String
 'Added by Lydia 2020/04/10 暫存檔的序號
 Dim mESeqNo As String
 Dim xRows As Integer
 Dim rsQuery As New ADODB.Recordset
 
On Error GoTo ErrHand

   'Modified by Lydia 2020/04/10 改用暫存檔Rdatafactory
   'If CreateDatabase = False Then
   '   MsgBox "無法建立暫存區，列印失敗 !", vbInformation
   '   m_print2 = 3
   '   Exit Sub
   'End If
   'Set Wo = DBEngine.Workspaces(0)
   'Set Db = Wo.OpenDatabase(App.path & "\Case.mdb", False, False, ";PWD=taie")
   'Qty = "DELETE FROM TEMP"
   'Db.Execute Qty
   'If RsTemp.State = adStateOpen Then RsTemp.Close
   'Qty = "INSERT INTO TEMP (TMP01,TMP02,TMP03,TMP04) VALUES ('民事案件','件數','總點數','平均')"
   'Db.Execute Qty
   intI = 1
   Qty = "select '民事案件','件數','總點數','平均' from dual"
   Set RsTemp = ClsLawReadRstMsg(intI, Qty)
   Set rsQuery = PUB_CreateRecordset(RsTemp, , , , Me.Name, mESeqNo)
   xRows = xRows + 1
   rsQuery.Close
   'end 2020/04/10
   
   StS(1) = "一審":  StS(2) = "二審":  StS(3) = "三審"
   StS(4) = "再審":  StS(5) = "合計": StS(6) = "平均"
   StRng(1) = " AND CP10 IN ('1111','1112','1113','1115')" 'Memo by Lydia 2020/04/20 指定CP10與現有案件性質代號不同
   'StRng(1) = " AND CP10 IN ('1101','1102','1103','1105')" 'Memo by Lydia 2020/04/20 測試用
   StRng(2) = " AND CP10 IN ('1122','1124')"
   StRng(3) = " AND CP10 IN ('1132','1134')"
   StRng(4) = " AND CP10 IN ('1133')"
   'Modify By Cheng 2002/03/26
   '多加CP09<'C'控制
'   strExc(0) = "SELECT COUNT(*),SUM(CP18),SUM(CP18)/COUNT(*) FROM CASEPROGRESS WHERE " & Mid(GetSQL, 5)
   'Modify by Morgan 2011/6/1 若有建點數分配資料時點數改分配點數(目前L會有分配) cp18->nvl(a0n03/1000,cp18)
   'Modified by Lydia 2015/06/02
   'strExc(0) = "SELECT COUNT(*),SUM(nvl(a0n03/1000,cp18)),SUM(nvl(a0n03/1000,cp18))/COUNT(*) FROM CASEPROGRESS ,acc0n0 where a0n02(+)=cp09 and CP09<'C' AND " & Mid(GetSql, 5)
 
   For i = 1 To 4
     'Modified by Lydia 2015/06/02 +工作分配點數
     ' strExc(1) = strExc(0) & StRng(i)
      strExc(4) = "SELECT CP01,CP02,CP03,CP04,cp09,nvl(a0n03/1000,cp18) CP18,sum(decode(p1.a1n01,null,decode(p3.a1n05,null,0,p3.a1n05),p1.a1n05)) a07 " & _
                   "FROM CASEPROGRESS ,acc0n0,acc1n0 p1,acc1n0 p3 where a0n02(+)=cp09 and  CP09<'C' AND " & _
                 Mid(GetSql, 5) & StRng(i) & " AND p1.a1n02(+)='3' AND p1.a1n03(+)=cp09 " & _
                 "AND p3.a1n02(+)='2' AND p3.a1n03(+)=cp09 AND p3.a1n01(+)=cp60 " & _
                 "group by CP01,CP02,CP03,CP04,cp09,nvl(a0n03/1000,cp18)"
      strExc(1) = "SELECT COUNT(*),SUM(decode(a07,0,cp18,a07)),SUM(decode(a07,0,cp18,a07))/COUNT(*) from (" & strExc(4) & ") "
      'end 2015/06/02
      For j = 1 To 3
         stTmp(j) = "0"
      Next
      If RsTemp.State = adStateOpen Then RsTemp.Close  'Added by Lydia 2020/04/10
      RsTemp.Open strExc(1), cnnConnection
      If RsTemp.Fields(0) > 0 Then
         stTmp(1) = RsTemp.Fields(0)
         stTmp(2) = RsTemp.Fields(1)
         stTmp(3) = RsTemp.Fields(2)
      End If
      RsTemp.Close
      'Modified by Lydia 2020/04/10 改用暫存檔Rdatafactory
      'Qty = "INSERT INTO TEMP (TMP01,TMP02,TMP03,TMP04) VALUES ('" & StS(i) & _
         "','" & stTmp(1) & "','" & stTmp(2) & "','" & stTmp(3) & "')"
      'Db.Execute Qty
      xRows = xRows + 1
      Qty = "insert into rdatafactory(formname,id,seqno,rowseq,r001,r002,r003,r004) " & _
                "values ('" & Me.Name & "', '" & strUserNum & "', '" & mESeqNo & "', " & xRows & ",'" & StS(i) & "','" & stTmp(1) & "','" & stTmp(2) & "','" & stTmp(3) & "' )"
      cnnConnection.Execute Qty
      'end 2020/04/10
   Next
   'Modified by Lydia 2020/04/10 改用暫存檔Rdatafactory
   'Qty = "INSERT INTO TEMP (TMP01,TMP02,TMP03,TMP04) SELECT '" & StS(5) & _
      "',SUM(VAL(TMP02)),SUM(VAL(TMP03)),SUM(VAL(TMP04)) FROM TEMP WHERE TMP01<>'民事案件'"
   'Db.Execute Qty
   'Qty = "INSERT INTO TEMP (TMP01,TMP02,TMP03,TMP04) SELECT '" & StS(6) & _
      "',FORMAT(TMP02/4,""0.00""),FORMAT(TMP03/4,""0.00"")," & _
      "FORMAT(TMP04/4,""0.00"") FROM TEMP WHERE TMP01='" & StS(5) & "'"
   'Db.Execute Qty
    xRows = xRows + 1
    Qty = "insert into rdatafactory(formname,id,seqno,rowseq,r001,r002,r003,r004) " & _
              " SELECT '" & Me.Name & "', '" & strUserNum & "', '" & mESeqNo & "', " & xRows & ",'" & StS(5) & "',SUM(R002) as R002,SUM(R003) as R003,SUM(R004) as R004 from rdatafactory WHERE R001<>'民事案件' " & _
              "AND formname='" & Me.Name & "' and id='" & strUserNum & "' and seqno='" & mESeqNo & "' "
    cnnConnection.Execute Qty
    xRows = xRows + 1
    Qty = "insert into rdatafactory(formname,id,seqno,rowseq,r001,r002,r003,r004) " & _
             " SELECT '" & Me.Name & "', '" & strUserNum & "', '" & mESeqNo & "','" & xRows & "' ,'" & StS(6) & "',round((R002 / 4), 2) as r002,round((R003 / 4), 2) as r003,round((R004 / 4), 2) as r004 FROM rdatafactory WHERE R001='" & StS(5) & "' AND formname='" & Me.Name & "' and id='" & strUserNum & "' and seqno='" & mESeqNo & "' "
    cnnConnection.Execute Qty
   'end 2020/04/10
   
   'Modified by Lydia 2020/04/10 改用暫存檔Rdatafactory
   'Qty = "SELECT cint(iif(ISNULL(TMP02),0,TMP02))+cint(iif(ISNULL(TMP03),0,TMP03))+cint(iif(ISNULL(TMP04),0,TMP04)) FROM TEMP WHERE TMP01='合計'"
   'Set Rc = Db.OpenRecordset(Qty)
   'If Rc.EOF = False Then
   Qty = "select r002+r002+r004 from rdatafactory WHERE R001='合計' AND formname='" & Me.Name & "' and id='" & strUserNum & "' and seqno='" & mESeqNo & "' "
   intI = 1
   Set rsQuery = ClsLawReadRstMsg(intI, Qty)
   If intI = 1 Then
   'end 2020/04/10
         'Modified by Lydia 2020/04/10 Rc=>rsQuery
         If Not IsNull(rsQuery.Fields(0)) Then
            If rsQuery.Fields(0) = 0 Then
               m_print2 = 1
               rsQuery.Close
               Exit Sub
            End If
         Else
             m_print2 = 1
             rsQuery.Close
             Exit Sub
         End If
         'end 2020/04/10
   End If
   
   i = 500
   Printer.Orientation = 1
   Printer.Font.Size = 22
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.CurrentX = 4200:         Printer.CurrentY = i
   Printer.Print "民事案件統計表"
   Printer.Font.Underline = False
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   Printer.CurrentX = 4200:         Printer.CurrentY = i + 500
   Printer.Print "發文日期 : " & ChangeTStringToTDateString(Text1(0)) & _
      " - " & ChangeTStringToTDateString(Text1(1))
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   Printer.CurrentX = 500:               Printer.CurrentY = i + 800
   Printer.Print "列印人 : " & strUserName
   Printer.CurrentX = 9000:             Printer.CurrentY = i + 800
   Printer.Print "列印日期 : " & ChangeTStringToTDateString(ChangeWStringToTString(GetTodayDate))
   Printer.CurrentX = 9000:             Printer.CurrentY = i + 1100
   Printer.Print "頁次 : 1"
   iPrint = i + 1400
   Printer.CurrentX = 500:               Printer.CurrentY = iPrint
   'Modified by Lydia 2020/04/20 180=>125
   Printer.Print String(125, "-")
   iPrint = iPrint + 300
   
   'Modified by Lydia 2020/04/10 改用暫存檔Rdatafactory
   'Qty = "SELECT TMP01,TMP02,TMP03,TMP04 FROM TEMP"
   'Set Rc = Db.OpenRecordset(Qty)
   'With Rc
   Qty = "select r001,r002,r003,r004 from rdatafactory WHERE  formname='" & Me.Name & "' and id='" & strUserNum & "' and seqno='" & mESeqNo & "' order by rowseq "
   intI = 1
   Set rsQuery = ClsLawReadRstMsg(intI, Qty)
   If intI = 1 Then
        With rsQuery
           .MoveFirst
   'end 2020/04/10
           Do While Not .EOF
              If .Fields(0) <> "平均" Then
                    For i = 0 To 3
                     '  Printer.CurrentX = PLeft(2 * i) + 1000 - Printer.TextWidth(Format(.Fields(i), "@@@@")):  Printer.CurrentY = iPrint
                       Select Case i
                              Case 0
                                   Printer.CurrentX = 500: Printer.CurrentY = iPrint
                                   Printer.Print .Fields(i)
                              Case 1
                                   Printer.CurrentX = 3500: Printer.CurrentY = iPrint
                                   Printer.Print .Fields(i)
                              Case 2
                                   Printer.CurrentX = PLeft(2 * i) + 1000 - Printer.TextWidth(Format(.Fields(i), "0.0")): Printer.CurrentY = iPrint
                                   If .Fields(0) = "平均" Then
                                      Printer.Print Format(.Fields(i), "0.00")
                                   Else
                                      Printer.Print Format(.Fields(i), "0.0")
                                   End If
                              Case 3
                                   Printer.CurrentX = PLeft(2 * i) + 1000 - Printer.TextWidth(Format(.Fields(i), "0.00")): Printer.CurrentY = iPrint
                                   Printer.Print Format(.Fields(i), "0.00")
                       End Select
                    Next
                    iPrint = iPrint + 300
              End If
              .MoveNext
           Loop
        End With
   End If 'Added by Lydia 2020/04/10
   
   Printer.CurrentX = 500:          Printer.CurrentY = iPrint
   'Modified by Lydia 2020/04/20 180=>125
   Printer.Print String(125, "-")
   Printer.EndDoc
   Exit Sub
ErrHand:
   MsgBox Err.Description
End Sub

Private Sub PrintCase3()
 Dim i As Integer, j As Integer, St As String, Page As Integer, iPrint As Integer
 Dim StS(1 To 6) As String, StRng(1 To 4) As String
 Dim stTmp(1 To 3) As String
 'Dim Wo As DAO.Workspace, Db As DAO.Database, Rc As DAO.Recordset, Rc1 As DAO.Recordset 'Remove by Lydia 2020/04/10
 Dim Qty As String
 'Added by Lydia 2020/04/10 暫存檔的序號
 Dim mESeqNo As String
 Dim xRows As Integer
 Dim rsQuery As New ADODB.Recordset
 
On Error GoTo ErrHand

   'Modified by Lydia 2020/04/10 改用暫存檔Rdatafactory
   'If CreateDatabase = False Then
   '   m_print3 = 3
   '   MsgBox "無法建立暫存區，列印失敗 !", vbInformation
   '   Exit Sub
   'End If
   'Set Wo = DBEngine.Workspaces(0)
   'Set Db = Wo.OpenDatabase(App.path & "\Case.mdb", False, False, ";PWD=taie")
   'Qty = "DELETE FROM TEMP"
   'Db.Execute Qty
   'If RsTemp.State = adStateOpen Then RsTemp.Close
   'Qty = "INSERT INTO TEMP (TMP01,TMP02,TMP03,TMP04) VALUES ('強制執行','件數','總點數','平均')"
   'Db.Execute Qty
   intI = 1
   Qty = "select '強制執行','件數','總點數','平均' from dual"
   Set RsTemp = ClsLawReadRstMsg(intI, Qty)
   Set rsQuery = PUB_CreateRecordset(RsTemp, , , , Me.Name, mESeqNo)
   xRows = xRows + 1
   rsQuery.Close
   'end 2020/04/10
   
   StS(1) = "終局執行":  StS(2) = "假扣押執行":  StS(3) = "假處分執行"
   StS(4) = "假執行":  StS(5) = "合計": StS(6) = "平均"
   StRng(1) = " AND CP10 IN ('1301')"
   StRng(2) = " AND CP10 IN ('1311','1313')"
   StRng(3) = " AND CP10 IN ('1312','1314')"
   StRng(4) = " AND CP10 IN ('1304')"
   'Modify By Cheng 2002/03/26
   '多加CP09<'C'控制
'   strExc(0) = "SELECT COUNT(*),SUM(CP18),SUM(CP18)/COUNT(*) FROM CASEPROGRESS WHERE " & Mid(GetSQL, 5)
   'Modify by Morgan 2011/6/1 若有建點數分配資料時點數改分配點數(目前L會有分配) cp18->nvl(a0n03/1000,cp18)
   'Modified by Lydia 2015/06/02
   'strExc(0) = "SELECT COUNT(*),SUM(nvl(a0n03/1000,cp18)),SUM(nvl(a0n03/1000,cp18))/COUNT(*) FROM CASEPROGRESS ,acc0n0 where a0n02(+)=cp09 and CP09<'C' AND " & Mid(GetSql, 5)
   For i = 1 To 4
     'Modified by Lydia 2015/06/02 +工作分配點數
     ' strExc(1) = strExc(0) & StRng(i)
      strExc(4) = "SELECT CP01,CP02,CP03,CP04,cp09,nvl(a0n03/1000,cp18) CP18,sum(decode(p1.a1n01,null,decode(p3.a1n05,null,0,p3.a1n05),p1.a1n05)) a07 " & _
                   "FROM CASEPROGRESS ,acc0n0,acc1n0 p1,acc1n0 p3 where a0n02(+)=cp09 and  CP09<'C' AND " & _
                 Mid(GetSql, 5) & StRng(i) & " AND p1.a1n02(+)='3' AND p1.a1n03(+)=cp09 " & _
                 "AND p3.a1n02(+)='2' AND p3.a1n03(+)=cp09 AND p3.a1n01(+)=cp60 " & _
                 "group by CP01,CP02,CP03,CP04,cp09,nvl(a0n03/1000,cp18)"
      strExc(1) = "SELECT COUNT(*),SUM(decode(a07,0,cp18,a07)),SUM(decode(a07,0,cp18,a07))/COUNT(*) from (" & strExc(4) & ") "
      'end 2015/06/02
      For j = 1 To 3
         stTmp(j) = "0"
      Next
      If RsTemp.State = adStateOpen Then RsTemp.Close  'Added by Lydia 2020/04/10
      RsTemp.Open strExc(1), cnnConnection
      If RsTemp.Fields(0) > 0 Then
         stTmp(1) = RsTemp.Fields(0)
         stTmp(2) = RsTemp.Fields(1)
         stTmp(3) = RsTemp.Fields(2)
      End If
      RsTemp.Close
      'Modified by Lydia 2020/04/10 改用暫存檔Rdatafactory
      'Qty = "INSERT INTO TEMP (TMP01,TMP02,TMP03,TMP04) VALUES ('" & StS(i) & _
         "','" & stTmp(1) & "','" & stTmp(2) & "','" & stTmp(3) & "')"
      'Db.Execute Qty
      xRows = xRows + 1
      Qty = "insert into rdatafactory(formname,id,seqno,rowseq,r001,r002,r003,r004) " & _
                "values ('" & Me.Name & "', '" & strUserNum & "', '" & mESeqNo & "', " & xRows & ",'" & StS(i) & "','" & stTmp(1) & "','" & stTmp(2) & "','" & stTmp(3) & "' )"
      cnnConnection.Execute Qty
      'end 2020/04/10
   Next
   'Modified by Lydia 2020/04/10 改用暫存檔Rdatafactory
   'Qty = "INSERT INTO TEMP (TMP01,TMP02,TMP03,TMP04) SELECT '" & StS(5) & _
      "',SUM(VAL(TMP02)),SUM(VAL(TMP03)),SUM(VAL(TMP04)) FROM TEMP WHERE TMP01<>'民事案件'"
   'Db.Execute Qty
   'Qty = "INSERT INTO TEMP (TMP01,TMP02,TMP03,TMP04) SELECT '" & StS(6) & _
      "',FORMAT(TMP02/4,""0.00""),FORMAT(TMP03/4,""0.00"")," & _
      "FORMAT(TMP04/4,""0.00"") FROM TEMP WHERE TMP01='" & StS(5) & "'"
   'Db.Execute Qty
    xRows = xRows + 1
    Qty = "insert into rdatafactory(formname,id,seqno,rowseq,r001,r002,r003,r004) " & _
             " SELECT '" & Me.Name & "', '" & strUserNum & "', '" & mESeqNo & "', " & xRows & ",'" & StS(5) & "',SUM(R002) as R002,SUM(R003) as R003,SUM(R004) as R004 from rdatafactory WHERE R001<>'強制執行' " & _
              "AND formname='" & Me.Name & "' and id='" & strUserNum & "' and seqno='" & mESeqNo & "' "
    cnnConnection.Execute Qty
    xRows = xRows + 1
    Qty = "insert into rdatafactory(formname,id,seqno,rowseq,r001,r002,r003,r004) " & _
             " SELECT '" & Me.Name & "', '" & strUserNum & "', '" & mESeqNo & "','" & xRows & "' ,'" & StS(6) & "',round((R002 / 4), 2) as r002,round((R003 / 4), 2) as r003,round((R004 / 4), 2) as r004 FROM rdatafactory WHERE R001='" & StS(5) & "' AND formname='" & Me.Name & "' and id='" & strUserNum & "' and seqno='" & mESeqNo & "' "
    cnnConnection.Execute Qty
   'end 2020/04/10
   
   'Modified by Lydia 2020/04/10 改用暫存檔Rdatafactory
   'Qty = "SELECT cint(iif(ISNULL(TMP02),0,TMP02))+cint(iif(ISNULL(TMP03),0,TMP03))+cint(iif(ISNULL(TMP04),0,TMP04)) FROM TEMP WHERE TMP01='合計'"
   'Set Rc = Db.OpenRecordset(Qty)
   'If Rc.EOF = False Then
   Qty = "select r002+r002+r004 from rdatafactory WHERE R001='合計' AND formname='" & Me.Name & "' and id='" & strUserNum & "' and seqno='" & mESeqNo & "' "
   intI = 1
   Set rsQuery = ClsLawReadRstMsg(intI, Qty)
   If intI = 1 Then
   'end 2020/04/10
         'Modified by Lydia 2020/04/10 Rc=>rsQuery
         If Not IsNull(rsQuery.Fields(0)) Then
            If rsQuery.Fields(0) = 0 Then
               m_print3 = 1
               rsQuery.Close
               Exit Sub
            End If
         Else
             m_print3 = 1
             rsQuery.Close
             Exit Sub
         End If
         'end 2020/04/10
   End If
     
   i = 500
   Printer.Orientation = 1
   Printer.Font.Size = 22
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.CurrentX = 4000:         Printer.CurrentY = i
   Printer.Print "強制執行案件統計表"
   Printer.Font.Underline = False
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   Printer.CurrentX = 4200:         Printer.CurrentY = i + 500
   Printer.Print "發文日期 : " & ChangeTStringToTDateString(Text1(0)) & _
      " - " & ChangeTStringToTDateString(Text1(1))
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   Printer.CurrentX = 500:               Printer.CurrentY = i + 500
   Printer.Print "列印人 : " & strUserName
   Printer.CurrentX = 9000:             Printer.CurrentY = i + 500
   Printer.Print "列印日期 : " & ChangeTStringToTDateString(ChangeWStringToTString(GetTodayDate))
   Printer.CurrentX = 9000:             Printer.CurrentY = i + 800
   Printer.Print "頁次 : 1"
   iPrint = i + 1100
   Printer.CurrentX = 500:               Printer.CurrentY = iPrint
   'Modified by Lydia 2020/04/20 180=>125
   Printer.Print String(125, "-")
   iPrint = iPrint + 300
   
   'Modified by Lydia 2020/04/10 改用暫存檔Rdatafactory
   'Qty = "SELECT TMP01,TMP02,TMP03,TMP04 FROM TEMP"
   'Set Rc = Db.OpenRecordset(Qty)
   'j = 0
   'With Rc
   Qty = "select r001,r002,r003,r004 from rdatafactory WHERE  formname='" & Me.Name & "' and id='" & strUserNum & "' and seqno='" & mESeqNo & "' order by rowseq "
   intI = 1
   Set rsQuery = ClsLawReadRstMsg(intI, Qty)
   If intI = 1 Then
        With rsQuery
           .MoveFirst
   'end 2020/04/10
           Do While Not .EOF
              If .Fields(0) <> "平均" Then
                  For i = 0 To 3
                 '    Printer.CurrentX = PLeft(2 * i) + 1000 - Printer.TextWidth(Format(.Fields(i), "@@@@")):  Printer.CurrentY = iPrint
                     Select Case i
                            Case 0
                                 Printer.CurrentX = 500: Printer.CurrentY = iPrint
                                 Printer.Print .Fields(i)
                            Case 1
                                 Printer.CurrentX = 3500: Printer.CurrentY = iPrint
                                 Printer.Print .Fields(i)
                            Case 2
                                 Printer.CurrentX = PLeft(2 * i) + 1000 - Printer.TextWidth(Format(.Fields(i), "0.0")): Printer.CurrentY = iPrint
                                 If .Fields(0) = "平均" Then
                                    Printer.Print Format(.Fields(i), "0.00")
                                 Else
                                    Printer.Print Format(.Fields(i), "0.0")
                                 End If
                                 
                            Case 3
                                 Printer.CurrentX = PLeft(2 * i) + 1000 - Printer.TextWidth(Format(.Fields(i), "0.00")): Printer.CurrentY = iPrint
                                 Printer.Print Format(.Fields(i), "0.00")
                     End Select
                  Next
                  iPrint = iPrint + 300
            End If
              .MoveNext
           Loop
        End With
   End If 'Added by Lydia 2020/04/10
   Printer.CurrentX = 500:          Printer.CurrentY = iPrint
   'Modified by Lydia 2020/04/20 180=>125
   Printer.Print String(125, "-")
   Printer.EndDoc
   Exit Sub
ErrHand:
   MsgBox Err.Description
End Sub

Private Sub GetPrintLeft()
   Erase PLeft
   PLeft(0) = 500:     PLeft(1) = 2400
   PLeft(2) = 3600:    PLeft(3) = 4800
   PLeft(4) = 6200:    PLeft(5) = 7400
   PLeft(6) = 8600:     PLeft(7) = 9800
   PLeft(8) = 11000:    PLeft(9) = 12200
   PLeft(10) = 13400:    PLeft(11) = 14600
   PLeft(12) = 15800
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   
End Sub

Private Sub Form_Paint()
  If Me.Tag = 1 Then
     lblSysKind.Caption = "(1.FCL 2.CFL 3.全部)"
  ElseIf Me.Tag = 0 Then
     lblSysKind.Caption = "(1.法務 2.顧問 3.全部)"
  End If
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0, 1
         If (KeyAscii > 57 Or KeyAscii < 48) And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
         End If
      Case 2
         If (KeyAscii > 51 Or KeyAscii < 49) And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
         End If
      Case 3
         If (KeyAscii > 52 Or KeyAscii < 49) And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
         End If
   End Select
End Sub

Private Sub Text1_LostFocus(Index As Integer)
   Select Case Index
   Case 1 '發文日期
      'Add By Cheng 2002/09/10
      If blnClkSure = False Then
         If Me.Text1(0).Text <> "" And Me.Text1(1).Text <> "" Then
            If Val(Me.Text1(0).Text) > Val(Me.Text1(1).Text) Then
               MsgBox "發文日期範圍輸入錯誤!!!", vbExclamation + vbOKOnly
               Me.Text1(0).SetFocus
               Text1_GotFocus 0
               Exit Sub
            End If
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
      Case 0, 1
         If CheckIsTaiwanDate(Text1(Index)) = False Then Cancel = True
   End Select
   If Cancel Then TextInverse Text1(Index)
End Sub

Private Function GetSql() As String
   If Text1(0) = "" And Text1(1) <> "" Then
      strExc(1) = " AND CP27 <='" & ChangeTStringToWString(Text1(1)) + "'"
   ElseIf Text1(0) <> "" And Text1(1) <> "" Then
      strExc(1) = " AND CP27 BETWEEN '" & ChangeTStringToWString(Text1(0)) + _
         "' AND '" + ChangeTStringToWString(Text1(1)) + "'"
   End If
   If Me.Tag = 0 Then
      Select Case Text1(2).Text
         Case "1"
            strExc(1) = strExc(1) & " AND CP01='L'"
         Case "2"
            strExc(1) = strExc(1) & " AND CP01='LA'"
         Case "3"
            strExc(1) = strExc(1) & " AND CP01 IN ('L','LA')"
      End Select
   Else
      Select Case Text1(2).Text
         Case "1"
            'Modify By Sindy 2009/07/24 增加LIN系統類別
            strExc(1) = strExc(1) & " AND CP01 in ('FCL','LIN')"
         Case "2"
            strExc(1) = strExc(1) & " AND CP01='CFL'"
         Case "3"
            'Modify By Sindy 2009/07/24 增加LIN系統類別
            strExc(1) = strExc(1) & " AND CP01 IN ('FCL','CFL','LIN')"
      End Select
   End If
   GetSql = strExc(1) & " AND CP26 IS NULL AND CP57 IS NULL"
End Function

Private Sub Form_Unload(Cancel As Integer)
   Set frm084002 = Nothing
End Sub
