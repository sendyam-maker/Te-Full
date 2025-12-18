VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm084001 
   BorderStyle     =   1  '單線固定
   Caption         =   "收／發文統計表"
   ClientHeight    =   2610
   ClientLeft      =   1845
   ClientTop       =   1875
   ClientWidth     =   4965
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   4965
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   4
      Left            =   1620
      MaxLength       =   6
      TabIndex        =   4
      Top             =   2010
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   3
      Left            =   1620
      MaxLength       =   1
      TabIndex        =   3
      Top             =   1620
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   2
      Left            =   3120
      MaxLength       =   7
      TabIndex        =   2
      Top             =   1230
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   1
      Left            =   1620
      MaxLength       =   7
      TabIndex        =   1
      Top             =   1230
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   0
      Left            =   1620
      MaxLength       =   1
      TabIndex        =   0
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   3900
      TabIndex        =   6
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   400
      Left            =   3072
      TabIndex        =   5
      Top             =   70
      Width           =   800
   End
   Begin MSForms.Label Label2 
      Height          =   255
      Left            =   2400
      TabIndex        =   11
      Top             =   2040
      Width           =   1125
      Size            =   "1984;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "承  辦  人："
      Height          =   180
      Index           =   2
      Left            =   660
      TabIndex        =   10
      Top             =   2010
      Width           =   900
   End
   Begin VB.Line Line1 
      X1              =   2700
      X2              =   2964
      Y1              =   1350
      Y2              =   1350
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "系  統  別：                (1:FCL,2.CFL,3全部)"
      Height          =   180
      Index           =   2
      Left            =   660
      TabIndex        =   9
      Top             =   1650
      Width           =   3180
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "日        期："
      Height          =   180
      Index           =   1
      Left            =   660
      TabIndex        =   8
      Top             =   1260
      Width           =   900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "報表方式：              (1.收文  2.發文)"
      Height          =   180
      Index           =   0
      Left            =   660
      TabIndex        =   7
      Top             =   870
      Width           =   2880
   End
End
Attribute VB_Name = "frm084001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/01/16 外法(FCL,CFL)和內法(LA,L)改成Word列印; 原本內法2.L案改3.L民刑事，並且逐字檢查Unicode文字改以圖片方式列印
'Memo by Lydia 2023/01/16 改成Form2.0 ; Label2
'Memo by Lydia 2021/08/31 法務系統的工作點數分配功能先上線(110/9/1)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/3 日期欄已修改
Option Explicit

Dim PLeft(0 To 10) As Integer
Dim m_Kind As Integer
Dim m_print As Integer
Dim m_print1 As Integer
Dim m_Type As Integer
Dim blnClkSure As Boolean '判斷是否按下確定按鈕
Dim iPrint As Integer
'Added by Lydia 2023/01/16 開啟Word
Dim m_WordLeft As Long, m_WordTop As Long 'Word開啟位置
Dim bVisible As Boolean
Private Const iLimit As Integer = 38   '單頁最大列數
Dim iPage As Integer, intCounter As Integer
Dim strTemp(0 To 6) As String

Private Sub cmdBack_Click()
   Unload Me
End Sub

Private Sub cmdPrint_Click()
Dim strTempName As String
Dim douHeight As Double
   
   blnClkSure = False
   
   m_print = 0
   m_print1 = 0
   If Text1(0) = "" Then
      MsgBox "報表方式不得為空值 !", vbCritical
      Text1(0).SetFocus
      Exit Sub
   End If
   If ChkRange(Text1(1), Text1(2), "日期") = False Then
      blnClkSure = True
      Me.Text1(1).SetFocus
      Text1_GotFocus 1
      Exit Sub
   End If
   If PUB_CheckKeyInDate(Me.Text1(1)) = -1 Then
      Me.Text1(1).SetFocus
      Text1_GotFocus 1
      Exit Sub
   End If
   If PUB_CheckKeyInDate(Me.Text1(2)) = -1 Then
      Me.Text1(2).SetFocus
      Text1_GotFocus 2
      Exit Sub
   End If
   
   'Add By Sindy 2010/01/20
   '檢查承辦人
   If Me.Text1(4).Text <> "" Then
      If ClsPDGetStaffN(Text1(4), strTempName) Then
         Label2 = strTempName
      Else
         Label2 = ""
         Me.Text1(4).SetFocus
         Text1_GotFocus 4
         Exit Sub
      End If
   End If
   
   If Text1(3) = "" Then
      MsgBox "系統別不得為空值 !", vbCritical
      Text1(3).SetFocus
      Exit Sub
   End If
   
   'Added by ydia 2023/01/16 改成Word
   m_print1 = 0
   Screen.MousePointer = vbHourglass
   If Me.Tag = 1 Then '外法
      If PrintCaseWord(IIf(Text1(3) = "3", 2, Val(Text1(3)))) = True Then
          m_print1 = 1
      End If
   Else     '內法
      If Text1(3).Text = "3" Then 'L民刑事
         PrintCase2 500
      Else
         If PrintCaseWord(IIf(Text1(3) = "4", 2, Val(Text1(3)))) = True Then
            m_print1 = 1
         End If
         If Text1(3) = "4" Then
             PrintCase2 500
         End If
      End If
   End If
   Screen.MousePointer = vbDefault
   If m_print1 = 0 Then
       MsgBox "查無資料!", vbInformation
   Else
       MsgBox "列印結束!", vbInformation
   End If
   Exit Sub
   'end 2023/01/16
   
   Screen.MousePointer = vbHourglass
   GetPrintLeft
   'Modify By Sindy 2010/01/20
   If Me.Tag = 1 Then '外法
      Select Case Text1(3).Text
         Case "1"
            PrintCase3 500, "1"
         Case "2"
            PrintCase3 500, "2"
         Case "3"
            douHeight = PrintCase3(500, "1")
            If m_print <> 1 Then
               If douHeight > 6000 Then
                  PrintCase3 douHeight, "2"
               Else
                  PrintCase3 6000, "2"
               End If
            Else
               PrintCase3 500, "2"
            End If
      End Select
   '2010/01/20 End
   Else '內法
      Select Case Text1(3).Text
         Case "1"
            PrintCase1 0
         Case "2"
            PrintCase2 500
         Case "3"
            PrintCase1 1
            If m_print <> 1 Then
               PrintCase2 6000
            Else
               PrintCase2 500
            End If
      End Select
   End If
   Screen.MousePointer = vbDefault
   Select Case Text1(3).Text
          Case "1"
                If m_print = 0 Then
                   MsgBox "列印結束!", vbInformation
                ElseIf m_print = 1 Then
                   MsgBox "查無資料!", vbInformation
                End If
          Case "2"
                If m_print1 = 0 Then
                   MsgBox "列印結束!", vbInformation
                ElseIf m_print1 = 1 Then
                   MsgBox "查無資料!", vbInformation
                End If
          Case "3"
               If m_print = 0 Or m_print1 = 0 Then
                  If Me.Tag = 1 And iPrint <> 0 Then '外法
                     iPrint = iPrint + 600
                     Printer.CurrentX = 500:          Printer.CurrentY = iPrint
                     Printer.Print "PS.不含非個人點數"
                  End If
                  Printer.EndDoc
                  MsgBox "列印結束!", vbInformation
               ElseIf m_print = 1 And m_print1 = 1 Then
                  MsgBox "查無資料!", vbInformation
               End If
  End Select
End Sub

'內法--查詢
Private Sub PrintCase1(ByVal Situ As Integer)
 Dim i As Integer, St As String, Page As Integer
 Dim TmpArea As String
 Dim StS(1 To 6) As String, StRng(1 To 6) As String
 'Dim Wo As DAO.Workspace, Db As DAO.Database, Rc As DAO.Recordset, Rc1 As DAO.Recordset 'Remove by Lydia 2020/04/10
 Dim Qty As String
 Dim mdbR As Integer 'Added by Lydia 2015/06/05 無資料跳過
 'Added by Lydia 2020/04/10 暫存檔的序號
 Dim mESeqNo As String
 Dim xRows As Integer
 Dim rsQuery As New ADODB.Recordset
 
On Error GoTo ErrHand

   'Remove by Lydia 2020/04/10 改用暫存檔Rdatafactory
   'm_Kind = 1
   'If CreateDatabase = False Then
   '   MsgBox "無法建立暫存區，列印失敗 !", vbInformation
   '   m_print = 3
   '   Exit Sub
   'End If
   'Set Wo = DBEngine.Workspaces(0)
   'Set Db = Wo.OpenDatabase(App.path & "\Case.mdb", False, False, ";PWD=taie")
   'Qty = "DELETE FROM TEMP"
   'Db.Execute Qty
   'end 2020/04/10
   
   If RsTemp.State = adStateOpen Then RsTemp.Close
   If Text1(0).Text = "1" Then
      St = "收文"
   Else
      St = "發文"
   End If
   'Modified by Lydia 2015/06/05 點數分配另外顯示
   'Qty = "INSERT INTO TEMP (TMP01,TMP02,TMP03,TMP04,TMP05) VALUES ('所別','" & St & "','總點數','平均點數','備註')"
   'Modified by Lydia 2020/04/10 改用暫存檔Rdatafactory
   'Qty = "INSERT INTO TEMP (TMP01,TMP02,TMP03,TMP04,TMP05,TMP06,TMP07) VALUES ('所別','" & St & "','" & St & "點數" & "','平均點數','分配點數','總點數','備註')"
   'Db.Execute Qty
   intI = 1
   Qty = "select '所別','" & St & "','" & St & "點數" & "','平均點數','分配點數','總點數','備註' from dual"
   Set RsTemp = ClsLawReadRstMsg(intI, Qty)
   Set rsQuery = PUB_CreateRecordset(RsTemp, , , , Me.Name, mESeqNo)
   xRows = xRows + 1
   rsQuery.Close
   'end 2020/04/10
   
   StS(1) = "北所": StS(2) = "中所": StS(3) = "南所": StS(4) = "高所"
   StS(5) = "其他": StS(6) = "總計"
   StRng(1) = " AND SUBSTR(CP12,1,2)='S1' "
   StRng(2) = " AND SUBSTR(CP12,1,2)='S2' "
   StRng(3) = " AND SUBSTR(CP12,1,2)='S3' "
   StRng(4) = " AND SUBSTR(CP12,1,2)='S4' "
   StRng(5) = " AND SUBSTR(CP12,1,1)<>'S' "
   'Modify By Sindy 2010/5/10 點數改抓點數分配檔
'   strExc(0) = "SELECT COUNT(*),SUM(CP18),SUM(CP18)/COUNT(*) FROM CASEPROGRESS WHERE " & _
'      "CP26 IS NULL" & GetSQL1
   For i = 1 To 5
      'Modify by Morgan 2011/6/1 若有建點數分配資料時點數改分配點數(目前L會有分配) cp18->nvl(a0n03/1000,cp18)
      'Modified by Lydia 2015/05/08 + CFL分配點(acc0k0->acc1n0),原本只抓同業務收文部門點+同所屬部門人員
        ' strExc(1) = "SELECT COUNT(*),SUM(a1n05),SUM(a1n05)/COUNT(*) FROM ( " & _
                        "SELECT CP14,decode(substr(cp60,1,1),'X',decode(a1k25,null,a1n05,''),nvl(a0n03/1000,cp18)) as a1n05 From CASEPROGRESS, acc1n0,acc1k0 " & _
                        ",acc0n0 where a0n02(+)=cp09 and cp26 Is Null AND CP09<'C' AND CP57 IS NULL " & GetSQL1 & StRng(i) & _
                        " AND a1n01(+)=cp60 AND a1n02(+)='2' AND a1n03(+)=cp09 AND a1n04(+)=cp14 AND a1k01(+)=cp60 " & IIf(Trim(Text1(4)) <> "", " and CP14='" + Trim(Text1(4)) + "' ", "") & _
                        "Union All " & _
                        "SELECT a1n04 as CP14,decode(substr(cp60,1,1),'X',decode(a1k25,null,a1n05,''),nvl(a0n03/1000,cp18)) as a1n05 From CASEPROGRESS, acc1n0,acc1k0,staff a,staff b " & _
                        ",acc0n0 where a0n02(+)=cp09 and cp26 Is Null AND CP09<'C' AND CP57 IS NULL " & GetSQL1 & StRng(i) & _
                        " AND a1n01(+)=cp60 AND a1n02(+)='2' AND a1n03(+)=cp09 AND a1n04(+)<>cp14 AND a1n05>0 AND a1k01(+)=cp60 AND a1n04=a.st01(+) AND cp14=b.st01(+) AND substr(a.st15,1,2)=substr(b.st15,1,2) " & IIf(Trim(Text1(4)) <> "", " AND a1n04='" + Trim(Text1(4)) + "' ", "") & _
                        ") "
          'strExc(1) = strExc(0) & StRng(i)
        'Modified by 2015/06/02 +工作點數: 分成3段擷取資料,避免收文次數重複
               '1.未分配
            'Modified by Lydia 2015/06/05 +cnt,dot
            'Modified by Lydia 2022/12/21 判斷B類收文無請款點數,不算收文次數 => and (substr(CP09,1,1) = 'A' or nvl(nvl(a0n03/1000,cp18),0) > 0)
            strExc(4) = " SELECT CP14,cp60,cp09,nvl(a0n03/1000,cp18) as a1n05,1 cnt,0.000 dot From CASEPROGRESS,acc0n0" & _
                        " where a0n02(+)=cp09 and cp26 Is Null AND CP09<'C' AND CP57 IS NULL" & GetSQL1 & StRng(i) & " and (substr(CP09,1,1) = 'A' or nvl(nvl(a0n03/1000,cp18),0) > 0) " & _
                        " and not exists (select a1n03 from acc1n0,acc1k0,acc0k0 where a1n02 in ('2','3') AND a1n03=cp09 and a1n01=a1k01(+) and a1n01=a0k01(+) and a1k25 is null and a0k10 is null) "
                        
            '2.單一承辦人分配
            'Modified by Lydia 2015/06/05 +cnt,dot
            'Modified by Lydia 2022/12/21 判斷B類收文無請款點數,不算收文次數 => and (substr(CP09,1,1) = 'A' or nvl(nvl(p1.a1n05,p3.a1n05),0) > 0)
            strExc(5) = " SELECT nvl(p1.a1n04,p3.a1n04) CP14,cp60,cp09,nvl(p1.a1n05,p3.a1n05) a1n05,1 cnt,0.000 dot From CASEPROGRESS, acc1n0 p1, acc1n0 p3,acc1k0,acc0k0" & _
                        " where cp26 Is Null AND CP09<'C' AND CP57 IS NULL" & GetSQL1 & StRng(i) & " and (substr(CP09,1,1) = 'A' or nvl(nvl(p1.a1n05,p3.a1n05),0) > 0) and a1k01(+)=cp60 and a1k25 is null and a0k01(+)=cp60 and a0k10 is null " & _
                        " AND p1.a1n02(+)='3' AND p1.a1n03(+)=cp09 AND p1.a1n04(+)=cp14" & _
                        " AND p3.a1n02(+)='2' AND p3.a1n03(+)=cp09 AND p3.a1n01(+)=cp60 AND p3.a1n04(+)=cp14 and (p1.a1n05>0 or p3.a1n05>0) "
            '3.多人分配
            'Modified by Lydia 2015/06/05 +cnt,dot
            'Memo by Lydia 2015/06/11 因為法務部還處於整合狀態,有部門判斷會造成不同結果
            '去掉同部門判斷 " AND (substr(c.st15,1,2)=substr(b.st15,1,2) or b.st15 is null) and (substr(c.st15,1,2)=substr(a.st15,1,2) or a.st15 is null) "
            strExc(6) = " SELECT nvl(p2.a1n04,p4.a1n04) as CP14,cp60,cp09,0 as a1n05,0 cnt,nvl(p2.a1n05,p4.a1n05) as dot From CASEPROGRESS, acc1n0 p2,acc1n0 p4,acc1k0,acc0k0,staff a,staff b,staff c" & _
                        " where cp26 Is Null AND CP09<'C' AND CP57 IS NULL" & GetSQL1 & StRng(i) & "AND a1k01(+)=cp60 and a1k25 is null and a0k01(+)=cp60 and a0k10 is null" & _
                        " AND p2.a1n02(+)='3' AND p2.a1n03(+)=cp09 AND p2.a1n04(+)<>decode(cp14,null,'111',cp14) AND p2.a1n05(+)>0" & _
                        " AND p4.a1n01(+)=cp60 AND p4.a1n02(+)='2' AND p4.a1n03(+)=cp09 and p4.a1n04(+)<>decode(cp14,null,'111',cp14) AND p4.a1n05(+)>0 AND a1k01(+)=cp60" & _
                        " AND p2.a1n04=a.st01(+) AND p4.a1n04=b.st01(+) AND cp14=c.st01(+) and (p2.a1n05>0 or p4.a1n05>0)"
                        
          'end 2015/06/02
            strExc(4) = strExc(4) & " union all " & strExc(5) & " union all " & strExc(6)

         'Modified by Lydia 2015/06/05 點數分配另外顯示
        'strExc(7) = "SELECT COUNT(*),SUM(a1n05),SUM(a1n05)/COUNT(*) FROM (" & strExc(4) & ") " & IIf(Trim(Text1(4)) <> "", " where CP14='" + Trim(Text1(4)) + "'", "")
        strExc(7) = "SELECT nvl(sum(cnt),0) cnt ,nvl(SUM(a1n05),0) 收文點數,nvl(SUM(dot),0) 分配點數 FROM (" & strExc(4) & ") " & IIf(Trim(Text1(4)) <> "", " where CP14='" + Trim(Text1(4)) + "'", "")
       'end 2015/05/08
            
      If RsTemp.State = adStateOpen Then RsTemp.Close      'Added by Lydia 2020/04/10
      RsTemp.Open strExc(7), cnnConnection
      'Added by Lydia 2015/06/05 無資料跳過
      mdbR = mdbR + RsTemp.RecordCount
      If RsTemp.RecordCount > 0 Then
      'end 2015/06/05
            With RsTemp
               Do While Not .EOF
                  'Modified by Lydia 2015/06/05 點數分配另外顯示
                  '收/發文點數
                  If IsNull(.Fields("收文點數")) Then
                     strExc(2) = "0"
                  Else
                     strExc(2) = Format(.Fields("收文點數"), "0.0") 'Modified by Lydia 2015/06/08 統一小數點
                  End If
                  '分配點數
                  If IsNull(.Fields("分配點數")) Then
                     strExc(3) = "0"
                  Else
                     strExc(3) = Format(.Fields("分配點數"), "0.0") 'Modified by Lydia 2015/06/08 統一小數點
                  End If
                  '平均點數
                  If Val(.Fields("cnt")) = 0 Then
                     strExc(9) = "0"
                  Else
                     strExc(9) = Val(strExc(2)) / Val(.Fields("cnt"))
                  End If
                  '總點數
                  strExc(10) = Format(Val(strExc(2)) + Val(strExc(3)), "##0.0")
                  
                  'Modified by Lydia 2020/04/10 改用暫存檔Rdatafactory
                  '   Qty = "INSERT INTO TEMP (TMP01,TMP02,TMP03,TMP04,TMP05,TMP06) VALUES ('" & StS(i) & _
                        "','" & .Fields("cnt") & "','" & strExc(2) & "','" & strExc(9) & "','" & strExc(3) & "','" & strExc(10) & "')"
                  'Db.Execute Qty
                  xRows = xRows + 1
                  Qty = "insert into rdatafactory(formname,id,seqno,rowseq,r001,r002,r003,r004,r005,r006) " & _
                           "values ('" & Me.Name & "', '" & strUserNum & "', '" & mESeqNo & "', " & xRows & ",'" & StS(i) & _
                          "','" & .Fields("cnt") & "','" & strExc(2) & "','" & strExc(9) & "','" & strExc(3) & "','" & strExc(10) & "')"
                  cnnConnection.Execute Qty
                  'end 2020/04/10
                  Exit Do
               Loop
            End With
      End If 'Added by Lydia 2015/06/05
      RsTemp.Close
   Next
   
   If mdbR > 0 Then   'Added by Lydia 2015/06/05 無資料跳過
       'Modified by Lydia 2015/06/05 點數分配另外顯示
'        Qty = "INSERT INTO TEMP (TMP01,TMP02,TMP03,TMP04) SELECT '" & StS(6) & _
'           "',SUM(VAL(TMP02)),SUM(VAL(TMP03)),SUM(VAL(TMP03))/SUM(VAL(TMP02)) FROM TEMP WHERE TMP01<>'所別'"
        'Qty = "SELECT cint(iif(ISNULL(TMP02),0,TMP02))+cint(iif(ISNULL(TMP03),0,TMP03))+cint(iif(ISNULL(TMP04),0,TMP04)) FROM TEMP WHERE TMP01='總計'"
        'Modified by Lydia 2020/04/10 改用暫存檔Rdatafactory
        'Qty = "INSERT INTO TEMP (TMP01,TMP02,TMP03,TMP04,TMP05,TMP06) SELECT '" & StS(6) & _
        '   "',SUM(VAL(TMP02)),SUM(VAL(TMP03)),SUM(VAL(TMP03))/SUM(VAL(TMP02)),SUM(VAL(TMP05)),SUM(VAL(TMP06)) FROM TEMP WHERE TMP01<>'所別'"
        'Db.Execute Qty
        'Qty = "SELECT cint(TMP06) FROM TEMP WHERE TMP01='總計'"
        'Set Rc = Db.OpenRecordset(Qty)
        'If Rc.EOF = False Then
        'end 2015/06/05
        xRows = xRows + 1
        Qty = "insert into rdatafactory(formname,id,seqno,rowseq,r001,r002,r003,r004,r005,r006) " & _
                 " SELECT '" & Me.Name & "', '" & strUserNum & "', '" & mESeqNo & "', " & xRows & ",'" & StS(6) & _
                 "' ,SUM(R002) as R002,SUM(R003) as R003,SUM(R004) as R004,SUM(R005) as R005,SUM(R006) as R006 " & _
                 " FROM rdatafactory WHERE R001<>'所別' AND formname='" & Me.Name & "' and id='" & strUserNum & "' and seqno='" & mESeqNo & "' "
        cnnConnection.Execute Qty
        Qty = "SELECT R006 FROM rdatafactory where  r001='總計' and formname='" & Me.Name & "' and id='" & strUserNum & "' and seqno='" & mESeqNo & "' "
        intI = 1
        Set rsQuery = ClsLawReadRstMsg(intI, Qty)
        If intI = 1 Then
        'end 2020/04/10
           'Modified by Lydia 2020/04/10 Rc=>rsQuery
           If Not IsNull(rsQuery.Fields(0)) Then
              If rsQuery.Fields(0) = 0 Then
                 m_print = 1
                 rsQuery.Close
                 Exit Sub
              End If
           Else
               m_print = 1
               rsQuery.Close
               Exit Sub
           End If
           'end 2020/04/10
        End If
        
        i = 500
        Printer.Orientation = vbPRORPortrait
        Printer.Font.Size = 22
        Printer.Font.Bold = True
        Printer.Font.Underline = True
        Printer.CurrentX = 4000:         Printer.CurrentY = i
        If Me.Tag = 0 Then
           Printer.Print "LA " & St & "統計表"
        Else
           Printer.Print "FCL " & St & "統計表"
        End If
        Printer.Font.Underline = False
        Printer.Font.Size = 12
        Printer.Font.Bold = False
        Printer.CurrentX = 500:               Printer.CurrentY = i + 500
        Printer.Print "列印人 : " & strUserName
        Printer.CurrentX = 4000:             Printer.CurrentY = i + 500
        'Modify By Sindy 2010/01/20
        'Printer.Print "收文日期 : " & ChangeTStringToTDateString(Text1(1)) & _
        '" - " & ChangeTStringToTDateString(Text1(2))
        If Text1(0).Text = "1" Then
           Printer.Print "收文日期 : " & ChangeTStringToTDateString(Text1(1)) & _
           " - " & ChangeTStringToTDateString(Text1(2))
        Else
           Printer.Print "發文日期 : " & ChangeTStringToTDateString(Text1(1)) & _
           " - " & ChangeTStringToTDateString(Text1(2))
        End If
        '2010/01/20 End
        Printer.CurrentX = 9000:             Printer.CurrentY = i + 500
        Printer.Print "列印日期 : " & ChangeTStringToTDateString(ChangeWStringToTString(GetTodayDate))
        'Add By Sindy 2010/01/20
        If Trim(Text1(4)) <> "" Then
           Printer.CurrentX = 500:             Printer.CurrentY = i + 800
           Printer.Print "承辦人 : " & Trim(Label2.Caption)
        End If
        '2010/01/20 End
        Printer.CurrentX = 9000:             Printer.CurrentY = i + 800
        Printer.Print "頁次 : 1"
        iPrint = i + 1100
        Printer.CurrentX = 500:               Printer.CurrentY = iPrint
        Printer.Print String(140, "-")
        iPrint = iPrint + 300
        'Modified by Lydia 2015/06/05 + TMP06,TMP07
        'Modified by Lydia 2020/04/10 改用暫存檔Rdatafactory
        'Qty = "SELECT TMP01,TMP02,TMP03,TMP04,TMP05,TMP06,TMP07 FROM TEMP"
        'Set Rc = Db.OpenRecordset(Qty)
        'With Rc
        Qty = "SELECT R001,R002,R003,R004,R005,R006,R007 FROM rdatafactory where formname='" & Me.Name & "' and id='" & strUserNum & "' and seqno='" & mESeqNo & "' order by rowseq "
        intI = 1
        Set rsQuery = ClsLawReadRstMsg(intI, Qty)
        If intI = 1 Then
            With rsQuery
        'end 2020/04/10
               Do While Not .EOF
                  Printer.CurrentX = PLeft(0):      Printer.CurrentY = iPrint
                  Printer.Print .Fields("R001")
                  Printer.CurrentX = PLeft(1) - Printer.TextWidth(Format(.Fields("R002"))):    Printer.CurrentY = iPrint
                  Printer.Print .Fields("R002")
                  Printer.CurrentX = PLeft(2) - Printer.TextWidth(Format(.Fields("R003"), "0.0")):   Printer.CurrentY = iPrint
                  Printer.Print Format(.Fields("R003"), "0.0")
                  Printer.CurrentX = PLeft(3) - Printer.TextWidth(Format(.Fields("R004"), "0.0")):    Printer.CurrentY = iPrint
                  Printer.Print Format(.Fields("R004"), "0.0")
                  'Modified by Lydia 2015/06/05
                  'Printer.CurrentX = PLeft(4) - Printer.TextWidth(Format(.Fields("R005"))):    Printer.CurrentY = iPrint
                  'Printer.Print Format(.Fields("R005"))
                    Printer.CurrentX = PLeft(4) - Printer.TextWidth(Format(.Fields("R005"), "0.0")):    Printer.CurrentY = iPrint
                    Printer.Print Format(.Fields("R005"), "0.0")
                    Printer.CurrentX = PLeft(5) - Printer.TextWidth(Format(.Fields("R006"), "0.0")):    Printer.CurrentY = iPrint
                    Printer.Print Format(.Fields("R006"), "0.0")
                    Printer.CurrentX = PLeft(6) - Printer.TextWidth(Format(.Fields("R007"), "0.0")):    Printer.CurrentY = iPrint
                    Printer.Print Format(.Fields("R007"), "0.0")
                  iPrint = iPrint + 300
                  .MoveNext
               Loop
            End With
        End If 'Added by Lydia 2020/04/10

        Printer.CurrentX = 500:          Printer.CurrentY = iPrint
        Printer.Print String(140, "-")
        iPrint = iPrint + 600
        Printer.CurrentX = 500:          Printer.CurrentY = iPrint
        Printer.Print "PS.不含非個人點數"
        If Situ = 0 Then Printer.EndDoc
   End If 'Added by Lydia 2015/06/05 無資料跳過
   
   'Added by Lydia 2020/04/10
   Set RsTemp = Nothing
   Set rsQuery = Nothing
   'end 2020/04/10
   
   Exit Sub
ErrHand:
   MsgBox Err.Description
End Sub

'內法--查詢
Private Sub PrintCase2(ByVal intHei As Integer)
 Dim i As Integer, j As Integer, St As String, Page As Integer
 Dim TmpArea As String
 Dim StS(1 To 6) As String, StRng(1 To 4) As String, StN(1 To 4) As String
 Dim stTmp(1 To 4) As String
 'Dim Wo As DAO.Workspace, Db As DAO.Database, Rc As DAO.Recordset, Rc1 As DAO.Recordset 'Remove by Lydia 2020/04/10
 Dim Qty As String
 Dim strSql As String
 Dim Rs As New ADODB.Recordset
 Dim nCount As Integer
 Dim strCP01 As String
 Dim strCP02 As String
 Dim strCP03 As String
 Dim strCP04 As String
 Dim strKind1(1 To 4) As String
 Dim strKind2(1 To 4) As String
 Dim strKind3(1 To 4) As String
 Dim strkind4(1 To 4) As String
 'Added by Lydia 2020/04/10 暫存檔的序號
 Dim mESeqNo As String
 Dim xRows As Integer
 Dim rsQuery As New ADODB.Recordset
 Dim Xo As Long, Yo As Long 'Added by Lydia 2023/01/16
 
On Error GoTo ErrHand

   'Remove by Lydia 2020/04/10 改用暫存檔Rdatafactory
   'If CreateDatabase = False Then
   '   MsgBox "無法建立暫存區，列印失敗 !", vbInformation
   '   m_print1 = 3
   '   Exit Sub
   'End If
   'Set Wo = DBEngine.Workspaces(0)
   'Set Db = Wo.OpenDatabase(App.path & "\Case.mdb", False, False, ";PWD=taie")
   'Qty = "DELETE FROM TEMP"
   'Db.Execute Qty
   'end 2020/04/10
   
   GetPrintLeft 'Added by Lydia 2023/01/16
   
   If RsTemp.State = adStateOpen Then RsTemp.Close
   If Text1(0).Text = "1" Then
      St = "收文"
   Else
      St = "發文"
   End If
   'Modified by Lydia 2020/04/10 改用暫存檔Rdatafactory
   'Qty = "INSERT INTO TEMP (TMP01,TMP02,TMP03,TMP04,TMP05,TMP06) VALUES ('類別','刑事','民事','強制執行','雜文','備註')"
   'Db.Execute Qty
   intI = 1
   Qty = "select '類別','刑事','民事','強制執行','雜文','備註' from dual"
   Set RsTemp = ClsLawReadRstMsg(intI, Qty)
   Set rsQuery = PUB_CreateRecordset(RsTemp, , , , Me.Name, mESeqNo)
   xRows = xRows + 1
   rsQuery.Close
   'end 2020/04/10
   
   StS(1) = "專利": StS(2) = "商標": StS(3) = "著作權": StS(4) = "其他"
   StS(5) = "總計": StS(6) = "合計"
   StRng(1) = " CP10 IN ('2101','2102','2103','2104','2105','2106','2108'," & _
      "'2121','2122','2123','2124','2201','2202','2203','2204','2205','2206'," & _
      "'2207','2109','2137','2139','2140','2221','2222')"
   
   StRng(2) = " CP10 IN ('1101','1102','1103','1104','1105','1111','1112','1113','1114','1115','1121','1122','1123','1124','1131','1132','1133','1134')"
   
   StRng(3) = " CP10 IN ('1301','1311','1312','1313','1314')"
   
   StRng(4) = " CP10 NOT IN ('2101','2102','2103','2104','2105','2106','2108'," & _
      "'2121','2122','2123','2124','2201','2202','2203','2204','2205','2206'," & _
      "'2207','2109','2137','2139','2140','2221','2222'," & _
      "'1101','1102','1103','1104','1105','1111','1112','1113','1114','1115','1121','1122','1123','1124','1131','1132','1133','1134'," & _
      "'1301','1311','1312','1313','1314')"
   
   StN(1) = " AND CR05 IN ('P','FCP','CFP'))"
   StN(2) = " AND CR05 IN ('T','FCT','CFT','TF'))"
   StN(3) = " AND CR05 IN ('TC','CFC'))"
   StN(4) = " AND CR05 not IN ('P','FCP','CFP','T','FCT','CFT','TF','TC','CFC'))"
   For j = 1 To 4
      strKind1(j) = 0
      strKind2(j) = 0
      strKind3(j) = 0
      strkind4(j) = 0
   Next
   
   For i = 1 To 4
       strCP01 = ""
       strCP02 = ""
       strCP03 = ""
       strCP04 = ""
      strSql = "SELECT CP01,CP02,CP03,CP04 FROM CASEPROGRESS WHERE" & StRng(i) & GetSQL2
      Rs.Open strSql, cnnConnection
      If Rs.EOF = False Then
          
         Do While Rs.EOF = False
            m_Type = 0
            If Not IsNull(Rs.Fields("CP01")) Then
               strCP01 = Rs.Fields("CP01")
            Else
               strCP01 = ""
            End If
            If Not IsNull(Rs.Fields("CP02")) Then
               strCP02 = Rs.Fields("CP02")
            Else
               strCP02 = ""
            End If
            If Not IsNull(Rs.Fields("CP03")) Then
               strCP03 = Rs.Fields("CP03")
            Else
               strCP03 = ""
            End If
            If Not IsNull(Rs.Fields("CP04")) Then
               strCP04 = Rs.Fields("CP04")
            Else
               strCP04 = ""
            End If
        
            For j = 1 To 3
                'Modified by Lydia 2020/04/10 CASERELATION=>CASERELATION1
                strExc(2) = " select count(*) from (SELECT DISTINCt cr01,cr02,cr03,cr04 FROM CASERELATION1 WHERE " & _
                            " CR01 ='" & strCP01 & "'" & _
                            " AND CR02 ='" & strCP02 & "'" & _
                            " AND CR03 ='" & strCP03 & "'" & _
                            " AND CR04 ='" & strCP04 & "'" & StN(j)
                If RsTemp.State = adStateOpen Then RsTemp.Close      'Added by Lydia 2020/04/10
                RsTemp.Open strExc(2), cnnConnection
                If RsTemp.EOF = False Then
                   If Val(RsTemp.Fields(0)) <> 0 Then
                     m_Type = 1
                     If i = 1 Then
                        strKind1(j) = Val(strKind1(j)) + Val(RsTemp.Fields(0))
                     End If
                     If i = 2 Then
                        strKind2(j) = Val(strKind2(j)) + Val(RsTemp.Fields(0))
                     End If
                     If i = 3 Then
                        strKind3(j) = Val(strKind3(j)) + Val(RsTemp.Fields(0))
                     End If
                     If i = 4 Then
                        strkind4(j) = Val(strkind4(j)) + Val(RsTemp.Fields(0))
                     End If
                   End If
                       
                End If
                RsTemp.Close
             Next
                If m_Type <> 1 Then
                     If i = 1 Then
                        strKind1(4) = Val(strKind1(4)) + 1
                     End If
                     If i = 2 Then
                        strKind2(4) = Val(strKind2(4)) + 1
                     End If
                     If i = 3 Then
                        strKind3(4) = Val(strKind3(4)) + 1
                     End If
                     If i = 4 Then
                        strkind4(4) = Val(strkind4(4)) + 1
                     End If
                End If
                   
             Rs.MoveNext
             nCount = nCount + 1
          Loop
          Rs.Close
       Else
           For j = 1 To 4
                If i = 1 Then
                   strKind1(j) = 0
                End If
                If i = 2 Then
                   strKind2(j) = 0
                End If
                If i = 3 Then
                   strKind3(j) = 0
                End If
                If i = 4 Then
                   strkind4(j) = 0
                End If
           Next
           Rs.Close
       End If
   Next
   
   For i = 1 To 4
      'Modified by Lydia 2020/04/10 改用暫存檔Rdatafactory
      'Qty = "INSERT INTO TEMP (TMP01,TMP02,TMP03,TMP04,TMP05) VALUES ('" & StS(i) & _
         "','" & strKind1(i) & "','" & strKind2(i) & "','" & strKind3(i) & "','" & strkind4(i) & "')"
      'Db.Execute Qty
      xRows = xRows + 1
      Qty = "insert into rdatafactory(formname,id,seqno,rowseq,r001,r002,r003,r004,r005) " & _
                "values ('" & Me.Name & "', '" & strUserNum & "', '" & mESeqNo & "', " & xRows & ",'" & StS(i) & _
                "','" & strKind1(i) & "','" & strKind2(i) & "','" & strKind3(i) & "','" & strkind4(i) & "')"
      cnnConnection.Execute Qty
      'end 2020/04/10
   Next i
   
   'Modified by Lydia 2020/04/10 改用暫存檔Rdatafactory
   'Qty = "INSERT INTO TEMP (TMP01,TMP02,TMP03,TMP04,TMP05) SELECT '" & StS(5) & "',SUM(VAL(TMP02)),SUM(VAL(TMP03)),SUM(VAL(TMP04)),SUM(VAL(TMP05)) FROM TEMP WHERE TMP01<>'類別'"
   'Db.Execute Qty
   '
   'Qty = "SELECT cint(iif(ISNULL(TMP02),0,TMP02))+cint(iif(ISNULL(TMP03),0,TMP03))+cint(iif(ISNULL(TMP04),0,TMP04))++cint(iif(ISNULL(TMP05),0,TMP05)) FROM TEMP WHERE TMP01='總計'"
   'Set Rc = Db.OpenRecordset(Qty)
   'If Rc.EOF = False Then
    xRows = xRows + 1
    Qty = "insert into rdatafactory(formname,id,seqno,rowseq,r001,r002,r003,r004,r005) " & _
             " SELECT '" & Me.Name & "', '" & strUserNum & "', '" & mESeqNo & "', " & xRows & ",'" & StS(6) & _
             "' ,SUM(R002) as R002,SUM(R003) as R003,SUM(R004) as R004,SUM(R005) as R005 " & _
             " FROM rdatafactory WHERE R001<>'類別' AND formname='" & Me.Name & "' and id='" & strUserNum & "' and seqno='" & mESeqNo & "' "
    cnnConnection.Execute Qty
    'Modified by Ldyia 2023/01/16
    'Qty = "SELECT R002+R003+R004+R005 FROM rdatafactory where  r001='總計' and formname='" & Me.Name & "' and id='" & strUserNum & "' and seqno='" & mESeqNo & "' "
    Qty = "SELECT R002+R003+R004+R005 FROM rdatafactory where  r001='合計' and formname='" & Me.Name & "' and id='" & strUserNum & "' and seqno='" & mESeqNo & "' "
    intI = 1
    Set rsQuery = ClsLawReadRstMsg(intI, Qty)
    If intI = 1 Then
   'end 2020/04/10
      'Modified by Lydia 2020/04/10 Rc=>rsQuery
      If Val("" & rsQuery.Fields(0)) = 0 Then
            'Modified by Lydia 2023/01/16
            'm_print1 = 1
            m_print1 = 0
            rsQuery.Close
            Exit Sub
      End If
      'end 2020/04/10
   'Added by Lydia 2023/01/16
   Else
      m_print1 = 0
   'end 2023/01/16
   End If
   'Rc.Close 'Remove by Lydia 2020/04/10
   
   m_print1 = 1  'Added by Lydia 2023/01/16
   i = intHei
   'If intHei = 500 Then Printer.Orientation = vbPRORPortrait 'Mark by Lydia 2023/01/16
   Printer.Font.Size = 22
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.CurrentX = 4000:         Printer.CurrentY = i
   If Me.Tag = 0 Then
      Printer.Print "L " & St & "統計表"
   Else
      Printer.Print "CFL " & St & "統計表"
   End If
   Printer.Font.Underline = False
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   If intHei = 500 Then
      Printer.CurrentX = 500:               Printer.CurrentY = i + 500
      'Modified by Lydia 2023/01/16 逐字檢查Unicode文字改以圖片方式列印
      'Printer.Print "列印人 : " & strUserName
      Xo = Printer.CurrentX
      Yo = Printer.CurrentY
      PUB_PrintUnicodeText "列印人 : " & strUserName, Xo, Yo, 0
      'end 2023/01/16
      Printer.CurrentX = 4000:             Printer.CurrentY = i + 500
      'Modify By Sindy 2010/01/20
      'Printer.Print "收文日期 : " & ChangeTStringToTDateString(Text1(1)) & _
      '" - " & ChangeTStringToTDateString(Text1(2))
      If Text1(0).Text = "1" Then
         Printer.Print "收文日期 : " & ChangeTStringToTDateString(Text1(1)) & _
         " - " & ChangeTStringToTDateString(Text1(2))
      Else
         Printer.Print "發文日期 : " & ChangeTStringToTDateString(Text1(1)) & _
         " - " & ChangeTStringToTDateString(Text1(2))
      End If
      '2010/01/20 End
      Printer.CurrentX = 9000:             Printer.CurrentY = i + 500
      Printer.Print "列印日期 : " & ChangeTStringToTDateString(ChangeWStringToTString(GetTodayDate))
      'Add By Sindy 2010/01/20
      If Trim(Text1(4)) <> "" Then
         Printer.CurrentX = 500:             Printer.CurrentY = i + 800
         'Modified by Lydia 2023/01/16 逐字檢查Unicode文字改以圖片方式列印
         'Printer.Print "承辦人 : " & Trim(Label2.Caption)
         Xo = Printer.CurrentX
         Yo = Printer.CurrentY
         PUB_PrintUnicodeText "承辦人 : " & Trim(Label2.Caption), Xo, Yo, 0
         'end 2023/01/16
      End If
      '2010/01/20 End
      Printer.CurrentX = 9000:             Printer.CurrentY = i + 800
      'Modified by Lydia 2023/01/16 +Word頁數
      'Printer.Print "頁次 : 1"
      Printer.Print "頁次 : " & iPage + 1
      iPrint = i + 1100
   Else
      iPrint = i + 500
   End If
   Printer.CurrentX = 500:               Printer.CurrentY = iPrint
   Printer.Print String(140, "-")
   iPrint = iPrint + 300
   
   'Modified by Lydia 2020/04/10 改用暫存檔Rdatafactory
   'Qty = "SELECT TMP01,TMP02,TMP03,TMP04,TMP05,TMP06 FROM TEMP"
   'Set Rc = Db.OpenRecordset(Qty)
   '
   'With Rc
    Qty = "SELECT R001,R002,R003,R004,R005,R006 FROM rdatafactory where formname='" & Me.Name & "' and id='" & strUserNum & "' and seqno='" & mESeqNo & "' order by rowseq "
    intI = 1
    Set rsQuery = ClsLawReadRstMsg(intI, Qty)
    If intI = 1 Then
        With rsQuery
    'end 2020/04/10
           Do While Not .EOF
              Printer.CurrentX = PLeft(0):      Printer.CurrentY = iPrint
              Printer.Print .Fields(0)
              Printer.CurrentX = PLeft(1) - Printer.TextWidth(Format(.Fields(1))):    Printer.CurrentY = iPrint
              Printer.Print .Fields(1)
              Printer.CurrentX = PLeft(2) - Printer.TextWidth(Format(.Fields(2))):    Printer.CurrentY = iPrint
              Printer.Print .Fields(2)
              Printer.CurrentX = PLeft(3) - Printer.TextWidth(Format(.Fields(3))):    Printer.CurrentY = iPrint
              Printer.Print .Fields(3)
              Printer.CurrentX = PLeft(4) - Printer.TextWidth(Format(.Fields(4))):    Printer.CurrentY = iPrint
              Printer.Print .Fields(4)
              Printer.CurrentX = PLeft(5) - Printer.TextWidth(Format(.Fields(5)))::    Printer.CurrentY = iPrint
              Printer.Print Format(.Fields(5))
              iPrint = iPrint + 300
              .MoveNext
           Loop
        End With
   End If 'Added by Lydia 2020/04/10
   
   Printer.CurrentX = 500:          Printer.CurrentY = iPrint
   Printer.Print String(140, "-")
   'Modified by Lydia 2023/01/16
   'If Text1(3).Text = "2" Then
   '   Printer.EndDoc
   'End If
   Printer.EndDoc
   'end 2023/01/16
   
   'Added by Lydia 2020/04/10
   Set RsTemp = Nothing
   Set rsQuery = Nothing
   'end 2020/04/10
   
   Exit Sub
ErrHand:
   MsgBox Err.Description
End Sub

'Add By Sindy 2010/01/20 外法--查詢
Private Function PrintCase3(ByVal intHei As Integer, ByVal strSysID As String) As Double
 Dim i As Integer, St As String, Page As Integer
 Dim TmpArea As String
 'Dim Wo As DAO.Workspace, Db As DAO.Database, Rc As DAO.Recordset, Rc1 As DAO.Recordset 'Remove by Lydia 2020/04/10
 Dim Qty As String
 Dim mdbR As Integer 'Added by Lydia 2015/06/05 無資料跳過
 'Added by Lydia 2020/04/10 暫存檔的序號
 Dim mESeqNo As String
 Dim xRows As Integer
 Dim rsQuery As New ADODB.Recordset

On Error GoTo ErrHand

   'Remove by Lydia 2020/04/10 改用暫存檔Rdatafactory
   'm_Kind = 1
   'If CreateDatabase = False Then
   '   MsgBox "無法建立暫存區，列印失敗 !", vbInformation
   '   m_print = 3
   '   Exit Function
   'End If
   'Set Wo = DBEngine.Workspaces(0)
   'Set Db = Wo.OpenDatabase(App.path & "\Case.mdb", False, False, ";PWD=taie")
   'Qty = "DELETE FROM TEMP"
   'Db.Execute Qty
   'end 2020/04/10

   If RsTemp.State = adStateOpen Then RsTemp.Close
   If Text1(0).Text = "1" Then
      St = "收文"
   Else
      St = "發文"
   End If
   'Modified by Lydia 2015/05/08 國外收據點數分配另外顯示
   'Qty = "INSERT INTO TEMP (TMP01,TMP02,TMP03,TMP04,TMP05) VALUES ('承辦人','" & St & "','總點數','平均點數','備註')"
   'Modified by Lydia 2020/04/10 改用暫存檔Rdatafactory
   'Qty = "INSERT INTO TEMP (TMP01,TMP02,TMP03,TMP04,TMP05,TMP06,TMP07) VALUES ('承辦人','" & St & "','" & St & "點數" & "','平均點數','分配點數','總點數','備註')"
   'Db.Execute Qty
   intI = 1
   'Mofieid by Lydia 2023/01/07
   'Qty = "select '承辦人','" & St & "','" & St & "點數" & "','平均點數','分配點數','總點數','備註' from dual"
   Qty = "select '承辦人','" & St & "','請款點數" & "','平均點數','分配點數','總點數','工作點數','備註' from dual"
   Set RsTemp = ClsLawReadRstMsg(intI, Qty)
   Set rsQuery = PUB_CreateRecordset(RsTemp, , , , Me.Name, mESeqNo)
   xRows = xRows + 1
   rsQuery.Close
   'end 2020/04/10

   'Modify By Sindy 2010/5/10 點數改抓點數分配檔
'   strExc(0) = "SELECT CP14,COUNT(*),SUM(CP18),SUM(CP18)/COUNT(*) FROM CASEPROGRESS WHERE " & _
'      "CP26 IS NULL" & GetSQL3(strSysID)
   'Modify by Morgan 2011/6/1 若有建點數分配資料時點數改分配點數(目前L會有分配) cp18->nvl(a0n03/1000,cp18)
   'Modified by Lydia 2015/05/08 + CFL分配點(acc0k0->acc1n0),原本只抓同業務收文部門點+同所屬部門人員
  ' strExc(0) = "SELECT CP14,COUNT(*),SUM(a1n05),SUM(a1n05)/COUNT(*) FROM ( " & _
                     "SELECT CP14,decode(substr(cp60,1,1),'X',decode(a1k25,null,a1n05,''),nvl(a0n03/1000,cp18)) as a1n05 From CASEPROGRESS, acc1n0,acc1k0 " & _
                     ",acc0n0 where a0n02(+)=cp09 and cp26 Is Null AND CP09<'C' AND CP57 IS NULL " & GetSQL3(strSysID) & _
                     " AND a1n01(+)=cp60 AND a1n02(+)='2' AND a1n03(+)=cp09 AND a1n04(+)=cp14 AND a1k01(+)=cp60 " & IIf(Trim(Text1(4)) <> "", " and CP14='" + Trim(Text1(4)) + "' ", "") & _
                     "Union All " & _
                     "SELECT a1n04 as CP14,decode(substr(cp60,1,1),'X',decode(a1k25,null,a1n05,''),nvl(a0n03/1000,cp18)) as a1n05 From CASEPROGRESS, acc1n0,acc1k0,staff a,staff b " & _
                     ",acc0n0 where a0n02(+)=cp09 and cp26 Is Null AND CP09<'C' AND CP57 IS NULL " & GetSQL3(strSysID) & _
                     " AND a1n01(+)=cp60 AND a1n02(+)='2' AND a1n03(+)=cp09 AND a1n04(+)<>cp14 AND a1n05>0 AND a1k01(+)=cp60 AND a1n04=a.st01(+) AND cp14=b.st01(+) AND substr(a.st15,1,2)=substr(b.st15,1,2) " & IIf(Trim(Text1(4)) <> "", " AND a1n04='" + Trim(Text1(4)) + "' ", "") & _
                     ") GROUP BY CP14 ORDER BY CP14 "
   'Modifiedby Lydia 2015/06/02 分成3段擷取資料,避免收文次數重複
            '承辦人無點數,FCL會double?
       'Modified by 2015/06/02 +工作點數
'Modifed by Lydia 2023/01/07 收文點數改成請款點數(未請款+a1n02='2'),分配點數dot=非CP14分配到的點數,
                                            '總點數=請款點數+分配點數,另外列出工作點數區分; 收文次數cnt: A類收文+B類收文有點數才算
'            '1.未分配
'            'Modified by Lydia 2022/12/21 判斷B類收文無請款點數,不算收文次數 => and (substr(CP09,1,1) = 'A' or nvl(nvl(a0n03/1000,cp18),0) > 0)
'            strExc(4) = " SELECT CP14,cp60,cp09,nvl(a0n03/1000,cp18) as a1n05,1 cnt,0.000 dot From CASEPROGRESS,acc0n0" & _
'                        " where a0n02(+)=cp09 and cp26 Is Null AND CP09<'C' AND CP57 IS NULL" & GetSQL3(strSysID) & " and (substr(CP09,1,1) = 'A' or nvl(nvl(a0n03/1000,cp18),0) > 0) " & _
'                        " and not exists (select a1n03 from acc1n0,acc1k0,acc0k0 where a1n02 in ('2','3') AND a1n03=cp09 and a1n01=a1k01(+) and a1n01=a0k01(+) and a1k25 is null and a0k10 is null) "
'            '2.單一承辦人分配
'            'Modified by Lydia 2022/12/21 判斷B類收文無請款點數,不算收文次數 => and (substr(CP09,1,1) = 'A' or nvl(nvl(p1.a1n05,p3.a1n05),0) > 0)
'            strExc(5) = " SELECT nvl(p1.a1n04,p3.a1n04) CP14,cp60,cp09,nvl(p1.a1n05,p3.a1n05) a1n05,1 cnt,0.000 dot From CASEPROGRESS, acc1n0 p1, acc1n0 p3,acc1k0,acc0k0" & _
'                        " where cp26 Is Null AND CP09<'C' AND CP57 IS NULL" & GetSQL3(strSysID) & " and (substr(CP09,1,1) = 'A' or nvl(nvl(p1.a1n05,p3.a1n05),0) > 0) and a1k01(+)=cp60 and a1k25 is null and a0k01(+)=cp60 and a0k10 is null " & _
'                        " AND p1.a1n02(+)='3' AND p1.a1n03(+)=cp09 AND p1.a1n04(+)=cp14" & _
'                        " AND p3.a1n02(+)='2' AND p3.a1n03(+)=cp09 AND p3.a1n01(+)=cp60 AND p3.a1n04(+)=cp14 and (p1.a1n05>0 or p3.a1n05>0) "
'            '3.多人分配
'            'Memo by Lydia 2015/06/11 因為法務部還處於整合狀態,有部門判斷會造成不同結果
'            '去掉同部門判斷 " AND (substr(c.st15,1,2)=substr(b.st15,1,2) or b.st15 is null) and (substr(c.st15,1,2)=substr(a.st15,1,2) or a.st15 is null) "
'            strExc(6) = " SELECT nvl(p2.a1n04,p4.a1n04) as CP14,cp60,cp09,0 as a1n05,0 cnt,nvl(p2.a1n05,p4.a1n05) as dot From CASEPROGRESS, acc1n0 p2,acc1n0 p4,acc1k0,acc0k0,staff a,staff b,staff c" & _
'                        " where cp26 Is Null AND CP09<'C' AND CP57 IS NULL" & GetSQL3(strSysID) & "AND a1k01(+)=cp60 and a1k25 is null and a0k01(+)=cp60 and a0k10 is null" & _
'                        " AND p2.a1n02(+)='3' AND p2.a1n03(+)=cp09 AND p2.a1n04(+)<>decode(cp14,null,'111',cp14) AND p2.a1n05(+)>0" & _
'                        " AND p4.a1n01(+)=cp60 AND p4.a1n02(+)='2' AND p4.a1n03(+)=cp09 AND p4.a1n04(+)<>decode(cp14,null,'111',cp14) AND p4.a1n05(+)>0 AND a1k01(+)=cp60" & _
'                        " AND p2.a1n04=a.st01(+) AND p4.a1n04=b.st01(+) AND cp14=c.st01(+) and (p2.a1n05>0 or p4.a1n05>0)"
'
'            strExc(4) = strExc(4) & " union all " & strExc(5) & " union all " & strExc(6)
'            strExc(0) = "SELECT CP14,sum(cnt) cnt ,SUM(a1n05) 收文點數,SUM(dot) 分配點數 FROM (" & strExc(4) & ") " & _
'                        IIf(Trim(Text1(4)) <> "", " where CP14='" + Trim(Text1(4)) + "'", "") & _
'                        " GROUP BY CP14 ORDER BY CP14 "
'       'end 2015/06/02
       '1.未請款
       strExc(4) = "SELECT cp01,cp02,cp03,cp04,CP14,cp60,cp09,nvl(a0n03/1000,cp18) as 請款點數,decode(substr(cp09,1,1),'A',1,decode(nvl(nvl(a0n03,cp18),0),0,1)) as cnt,0.000 dot,0.000 as 工作點數 " & _
                        "From CASEPROGRESS,acc0n0 where a0n02(+)=cp09 and cp26 Is Null AND CP09<'C' AND CP159=0 and cp60 is null " & GetSQL3(strSysID) & " and (substr(CP09,1,1) = 'A' or nvl(nvl(a0n03/1000,cp18),0) > 0) "
       '2.已請款:承辦人
       strExc(5) = "SELECT cp01,cp02,cp03,cp04,CP14,cp60,cp09,nvl(p3.a1n05,0) as 請款點數,decode(substr(cp09,1,1),'A',1,decode(nvl(nvl(p1.a1n05,p3.a1n05),0),0,1)) as cnt,0.000 dot,nvl(p1.a1n05,0) as 工作點數 " & _
                        "From CASEPROGRESS, acc1n0 p1, acc1n0 p3,acc1k0,acc0k0 where cp26 Is Null AND CP09<'C' AND CP159=0 " & GetSQL3(strSysID) & " and (substr(CP09,1,1) = 'A' or nvl(nvl(p1.a1n05,p3.a1n05),0) > 0) " & _
                        "and a1k01(+)=cp60 and a1k25 is null and a0k01(+)=cp60 and a0k10 is null " & _
                        "AND p1.a1n02(+)='3' AND p1.a1n03(+)=cp09 AND p1.a1n04(+)=cp14 " & _
                        "AND p3.a1n02(+)='2' AND p3.a1n01(+)=cp60 AND p3.a1n04(+)=cp14 "
       '2.已請款:非承辦人
        strExc(6) = "SELECT cp01,cp02,cp03,cp04,p2.a1n04 as CP14,cp60,cp09,0 as 請款點數,0 cnt,0 as dot,nvl(p2.a1n05,0) as 工作點數 " & _
                        "From CASEPROGRESS, acc1n0 p2,acc1k0,acc0k0 where cp26 Is Null AND CP09<'C' AND CP159=0 " & GetSQL3(strSysID) & _
                        "AND a1k01(+)=cp60 and a1k25 is null and a0k01(+)=cp60 and a0k10 is null " & _
                        "AND p2.a1n02(+)='3' AND p2.a1n03(+)=cp09 AND p2.a1n04(+)<>decode(cp14,null,'111',cp14) AND nvl(p2.a1n05,0) > 0 " & _
                        "Union All SELECT cp01,cp02,cp03,cp04,p4.a1n04 as CP14,cp60,cp09,0 as 請款點數,0 cnt,nvl(p4.a1n05,0) as dot,0 as 工作點數 " & _
                        "From CASEPROGRESS, acc1n0 p4,acc1k0,acc0k0 c where cp26 Is Null AND CP09<'C' AND CP159=0 " & GetSQL3(strSysID) & _
                        "AND a1k01(+)=cp60 and a1k25 is null and a0k01(+)=cp60 and a0k10 is null " & _
                        "AND p4.a1n02(+)='2' AND p4.a1n01(+)=cp60 AND p4.a1n04(+)<>decode(cp14,null,'111',cp14) AND nvl(p4.a1n05,0) > 0"
        strExc(4) = strExc(4) & " union all " & strExc(5) & " union all " & strExc(6)
        strExc(0) = "SELECT CP14,sum(cnt) cnt,SUM(nvl(請款點數,0)) 請款點數,SUM(dot) 分配點數,SUM(工作點數) 工作點數 " & _
                         "FROM (" & strExc(4) & ") " & IIf(Trim(Text1(4)) <> "", " where CP14='" + Trim(Text1(4)) + "'", "") & _
                        " GROUP BY CP14 ORDER BY CP14 "
'end ----- 'Modifed by Lydia 2023/01/07 收文點數改成請款點數

   If RsTemp.State = adStateOpen Then RsTemp.Close      'Added by Lydia 2020/04/10
   RsTemp.Open strExc(0), cnnConnection
   mdbR = RsTemp.RecordCount 'Added by Lydia 2015/06/05 無資料跳過

   If RsTemp.RecordCount > 0 Then
      With RsTemp
         .MoveFirst
         Do While Not .EOF
            'Modified by Lydia 2023/01/07
            '收/發文點數
            'If IsNull(.Fields("收文點數")) Then
            '   strExc(2) = "0"
            'Else
            '   strExc(2) = Format(.Fields("收文點數"), "0.0") 'Modified by Lydia 2015/06/08 統一小數點
            'End If
            '請款點數
            If IsNull(.Fields("請款點數")) Then
               strExc(2) = "0"
            Else
               strExc(2) = Format(.Fields("請款點數"), "0.0") 'Modified by Lydia 2015/06/08 統一小數點
            End If
            'end 2023/01/07
            '分配點數
            If IsNull(.Fields("分配點數")) Then
               strExc(3) = "0"
            Else
               strExc(3) = Format(.Fields("分配點數"), "0.0") 'Modified by Lydia 2015/06/08 統一小數點
            End If
            'Modified by Lydia 2015/05/08 國外收據點數分配另外顯示
            '平均點數
            If Val(.Fields("cnt")) = 0 Then
               strExc(9) = "0"
            Else
               'Modified by Lydia 2015/06/05 小數點差異
               'strExc(9) = Format(Val(strExc(2)) / Val(.Fields("cnt")), "##0.0")
               strExc(9) = Val(strExc(2)) / Val(.Fields("cnt"))
            End If
            '總點數
            strExc(10) = Format(Val(strExc(2)) + Val(strExc(3)), "##0.0")
            'Added by Lydia 2023/01/07 工作點數
            If IsNull(.Fields("工作點數")) Then
               strExc(1) = "0"
            Else
               strExc(1) = Format(.Fields("工作點數"), "0.0")
            End If
            'end 2023/01/07
             'Modified by Lydia 2020/04/10 改用暫存檔Rdatafactory
            'If "" & .Fields("cp14") = "" Then
            '   Qty = "INSERT INTO TEMP (TMP01,TMP02,TMP03,TMP04,TMP05,TMP06) VALUES (' ','" & _
                  .Fields("cnt") & "','" & strExc(2) & "','" & strExc(9) & "','" & strExc(3) & "','" & strExc(10) & "')"
            'Else
            '   Qty = "INSERT INTO TEMP (TMP01,TMP02,TMP03,TMP04,TMP05,TMP06) VALUES ('" & GetPrjSalesNM(.Fields("cp14")) & _
                  "','" & .Fields("cnt") & "','" & strExc(2) & "','" & strExc(9) & "','" & strExc(3) & "','" & strExc(10) & "')"
            'End If
            'end 2015/05/08
            'Db.Execute Qty
            xRows = xRows + 1
            If "" & .Fields("cp14") = "" Then
                strExc(0) = " "
            Else
                strExc(0) = GetPrjSalesNM(.Fields("cp14"))
                '工作點數分配人員非法律所人員請姓名後加*
                If PUB_ChkLCompStaff(.Fields("cp14")) = False Then
                    strExc(0) = strExc(0) & "*"
                End If
            End If
            'Modified by Lydia 2023/01/07 +R007=strExc(1)=工作點數
            Qty = "insert into rdatafactory(formname,id,seqno,rowseq,r001,r002,r003,r004,r005,r006,R007) " & _
                     "values ('" & Me.Name & "', '" & strUserNum & "', '" & mESeqNo & "', " & xRows & ",'" & strExc(0) & "', '" & _
                     .Fields("cnt") & "','" & strExc(2) & "','" & strExc(9) & "','" & strExc(3) & "','" & strExc(10) & "','" & strExc(1) & "')"
            cnnConnection.Execute Qty
            'end 2020/04/10
            .MoveNext
         Loop
      End With
   End If
   RsTemp.Close
   'Remove by Lydia 2015/05/08 國外收據點數分配另外顯示
'   Qty = "INSERT INTO TEMP (TMP01,TMP02,TMP03,TMP04) SELECT '" & "總計" & _
'      "',SUM(VAL(TMP02)),SUM(VAL(TMP03)),SUM(VAL(TMP03))/SUM(VAL(TMP02)) FROM TEMP WHERE TMP01<>'承辦人'"
'   Db.Execute Qty
'
'   Qty = "SELECT cint(iif(ISNULL(TMP02),0,TMP02))+cint(iif(ISNULL(TMP03),0,TMP03))+cint(iif(ISNULL(TMP04),0,TMP04)) FROM TEMP WHERE TMP01='總計'"
   'end 2015/05/08
   If mdbR > 0 Then 'Added by Lydia 2015/06/05 無資料跳過

       'Modified by Lydia 2020/04/10 改用暫存檔Rdatafactory
       'Qty = "INSERT INTO TEMP (TMP01,TMP02,TMP03,TMP04,TMP05,TMP06) SELECT '" & "總計" & _
          "',SUM(VAL(TMP02)),SUM(VAL(TMP03)),SUM(VAL(TMP03))/SUM(VAL(TMP02)),SUM(VAL(TMP05)),SUM(VAL(TMP06)) FROM TEMP WHERE TMP01<>'承辦人'"
       'Db.Execute Qty
       'Qty = "SELECT cint(TMP06) FROM TEMP WHERE TMP01='總計'"
       'Set Rc = Db.OpenRecordset(Qty)
       'If Rc.EOF = False Then
       xRows = xRows + 1
       'Modified by Lydia 2023/01/07 +R007=工作點數
       Qty = "insert into rdatafactory(formname,id,seqno,rowseq,r001,r002,r003,r004,r005,r006,R007) " & _
                " SELECT '" & Me.Name & "', '" & strUserNum & "', '" & mESeqNo & "', " & xRows & ",'" & "總計" & _
                "' ,SUM(R002) as R002,SUM(R003) as R003,SUM(R004) as R004,SUM(R005) as R005,SUM(R006) as R006,SUM(R007) as R007 " & _
                " FROM rdatafactory WHERE R001<>'承辦人' AND formname='" & Me.Name & "' and id='" & strUserNum & "' and seqno='" & mESeqNo & "' "
       cnnConnection.Execute Qty
       'Modifed by Lydia 2023/01/07 +R007
       Qty = "SELECT R006,R007 FROM rdatafactory where  r001='總計' and formname='" & Me.Name & "' and id='" & strUserNum & "' and seqno='" & mESeqNo & "' "
       intI = 1
       Set rsQuery = ClsLawReadRstMsg(intI, Qty)
       If intI = 1 Then
       'end 2020/04/10
          'Modified by Lydia 2020/04/10 Rc=>rsQuery
          'Modified by Lydia 2023/01/07
          'If Not IsNull(rsQuery.Fields(0)) Then
          '    If rsQuery.Fields(0) = 0 Then
          '      m_print = 1
          '      rsQuery.Close
          '      Exit Function
          '   End If
          'Else
          If Val("" & rsQuery.Fields("R006")) + Val("" & rsQuery.Fields("R007")) = 0 Then
              m_print = 1
              rsQuery.Close
              Exit Function
          End If
          'end 2020/04/10
       End If

       i = intHei
       If intHei = 500 Then Printer.Orientation = vbPRORPortrait
       Printer.Font.Size = 22
       Printer.Font.Bold = True
       Printer.Font.Underline = True
       Printer.CurrentX = 4000:         Printer.CurrentY = i
       If strSysID = "1" Then
          Printer.Print "FCL " & St & "統計表"
       Else
          Printer.Print "CFL " & St & "統計表"
       End If
       Printer.Font.Underline = False
       Printer.Font.Size = 12
       Printer.Font.Bold = False
       If intHei = 500 Then
          Printer.CurrentX = 500:               Printer.CurrentY = i + 500
          Printer.Print "列印人 : " & strUserName
          Printer.CurrentX = 4000:             Printer.CurrentY = i + 500
          'Modify By Sindy 2010/01/20
          'Printer.Print "收文日期 : " & ChangeTStringToTDateString(Text1(1)) & _
          '" - " & ChangeTStringToTDateString(Text1(2))
          If Text1(0).Text = "1" Then
             Printer.Print "收文日期 : " & ChangeTStringToTDateString(Text1(1)) & _
             " - " & ChangeTStringToTDateString(Text1(2))
          Else
             Printer.Print "發文日期 : " & ChangeTStringToTDateString(Text1(1)) & _
             " - " & ChangeTStringToTDateString(Text1(2))
          End If
          '2010/01/20 End
          Printer.CurrentX = 9000:             Printer.CurrentY = i + 500
          Printer.Print "列印日期 : " & ChangeTStringToTDateString(ChangeWStringToTString(GetTodayDate))
          Printer.CurrentX = 9000:             Printer.CurrentY = i + 800
          Printer.Print "頁次 : 1"
          iPrint = i + 1100
       Else
          iPrint = i + 500
       End If
       Printer.CurrentX = 500:               Printer.CurrentY = iPrint
       Printer.Print String(140, "-")
       iPrint = iPrint + 300
       'Modified by Lydia 2015/05/08 + TMP06,TMP07
       'Modified by Lydia 2020/04/10 改用暫存檔Rdatafactory
       'Qty = "SELECT TMP01,TMP02,TMP03,TMP04,TMP05,TMP06,TMP07 FROM TEMP"
       'Set Rc = Db.OpenRecordset(Qty)
       'With Rc
       'Mofieid by Lydia 2023/01/07 +R008
       Qty = "SELECT R001,R002,R003,R004,R005,R006,R007,R008 FROM rdatafactory where formname='" & Me.Name & "' and id='" & strUserNum & "' and seqno='" & mESeqNo & "' order by rowseq "
       intI = 1
       Set rsQuery = ClsLawReadRstMsg(intI, Qty)
       If intI = 1 Then
           With rsQuery
       'end 2020/04/10
              Do While Not .EOF
                 '承辦人
                 Printer.CurrentX = PLeft(0):      Printer.CurrentY = iPrint
                 Printer.Print .Fields("R001")
                 'Modified by Lydia 2015/05/08
                 '次數
                 Printer.CurrentX = PLeft(1) - Printer.TextWidth(Format(.Fields("R002"))):    Printer.CurrentY = iPrint
                 Printer.Print .Fields("R002")
                 '請款點數
                 Printer.CurrentX = PLeft(2) - Printer.TextWidth(Format(.Fields("R003"), "0.0")):   Printer.CurrentY = iPrint
                 Printer.Print Format(.Fields("R003"), "0.0")
                 '平均點數
                 Printer.CurrentX = PLeft(3) - Printer.TextWidth(Format(.Fields("R004"), "0.0")):    Printer.CurrentY = iPrint
                 Printer.Print Format(.Fields("R004"), "0.0") '
                 'Modified by Lydia 2015/05/08
        '         Printer.CurrentX = PLeft(4) - Printer.TextWidth(Format(.Fields("R005"))):    Printer.CurrentY = iPrint
        '         Printer.Print Format(.Fields("R005"))
                 '分配點數
                 Printer.CurrentX = PLeft(4) - Printer.TextWidth(Format(.Fields("R005"), "0.0")):    Printer.CurrentY = iPrint
                 Printer.Print Format(.Fields("R005"), "0.0")
                 '總點數
                 Printer.CurrentX = PLeft(5) - Printer.TextWidth(Format(.Fields("R006"), "0.0")):    Printer.CurrentY = iPrint
                 Printer.Print Format(.Fields("R006"), "0.0")
                 '工作點數
                 Printer.CurrentX = PLeft(6) - Printer.TextWidth(Format(.Fields("R007"), "0.0")):    Printer.CurrentY = iPrint
                 Printer.Print Format(.Fields("R007"), "0.0")
                 'Added by Lydia 2023/01/07 備註(往後移)
                 Printer.CurrentX = PLeft(7):      Printer.CurrentY = iPrint
                 Printer.Print "" & .Fields("R008")
                 'end 2023/01/07
                 'Added by Lydia 2023/01/07
                 iPrint = iPrint + 300
                 .MoveNext
              Loop
           End With
       End If 'Added by Lydia 2020/04/10

       Printer.CurrentX = 500:          Printer.CurrentY = iPrint
       Printer.Print String(140, "-")
       If Text1(3).Text = "1" Or Text1(3).Text = "2" Then
          iPrint = iPrint + 600
          Printer.CurrentX = 500:          Printer.CurrentY = iPrint
          Printer.Print "PS.不含非個人點數"
          Printer.EndDoc
       End If
       PrintCase3 = iPrint + 600
   End If 'Added by Lydia 2015/06/05 無資料跳過

   'Added by Lydia 2020/04/10
   Set RsTemp = Nothing
   Set rsQuery = Nothing
   'end 2020/04/10

   Exit Function
ErrHand:
   MsgBox Err.Description
End Function

Private Sub GetPrintLeft()
   Erase PLeft

   PLeft(0) = 500:     PLeft(1) = 2200
   PLeft(2) = 3400:    PLeft(3) = 4800
   PLeft(4) = 6000:    PLeft(5) = 7200
   'Added by Lydia 2015/05/08
   PLeft(6) = 8400:    PLeft(7) = 9600
End Sub

Private Sub Form_Activate()
   'Add By Sindy 2010/01/20
   If Me.Tag = 0 Then '內法
      Label1(2).Visible = False
      Text1(4).Visible = False
      Label2.Visible = False
   End If
   '2010/01/20 End
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0
         If (KeyAscii > 50 Or KeyAscii < 49) And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
         End If
      Case 3
         'Added by Lydia 2023/01/16 內法1-4
         If Me.Tag = 0 Then
            If (KeyAscii > 52 Or KeyAscii < 49) And KeyAscii <> 8 Then
               KeyAscii = 0
               Beep
            End If
         Else
         'end 2023/01/16
            If (KeyAscii > 51 Or KeyAscii < 49) And KeyAscii <> 8 Then
               KeyAscii = 0
               Beep
            End If
         End If 'Added by Lydia 2023/01/16
      'Add By Sindy 2010/01/20
      Case 4
         KeyAscii = UpperCase(KeyAscii)
   End Select
End Sub

Private Sub Text1_LostFocus(Index As Integer)
   Select Case Index
   Case 2 '日期
      'Add By Cheng 2002/09/10
      If blnClkSure = False Then
         If Me.Text1(1).Text <> "" And Me.Text1(2).Text <> "" Then
            If Val(Me.Text1(1).Text) > Val(Me.Text1(2).Text) Then
               MsgBox "日期範圍輸入錯誤!!!", vbExclamation + vbOKOnly
               Me.Text1(1).SetFocus
               Text1_GotFocus 1
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
   
   'Add By Sindy 2010/01/20
   If Index = 4 Then
      If Trim(Text1(Index)) = "" Then
         Label2 = ""
         Exit Sub
      End If
   End If
   '2010/01/20 End
   If Text1(Index) = "" Then Exit Sub
   Select Case Index
      Case 1, 2
         If CheckIsTaiwanDate(Text1(Index)) = False Then Cancel = True
      'Add By Sindy 2010/01/20
      Case 4
         If ClsPDGetStaffN(Text1(Index), strTempName) Then
            Label2 = strTempName
         Else
            Label2 = ""
            Cancel = True
         End If
   End Select
   If Cancel Then TextInverse Text1(Index)
End Sub

Private Function GetSQL1() As String
Dim i As Integer, St As String
   If Text1(0).Text = "1" Then
      St = "CP05"
   Else
      St = "CP27"
   End If
   If Text1(1) = "" And Text1(2) <> "" Then
      strExc(1) = " AND " & St & " <='" & ChangeTStringToWString(Text1(2)) + "' "
   ElseIf Text1(1) <> "" And Text1(2) <> "" Then
      strExc(1) = " AND " & St & " BETWEEN '" & ChangeTStringToWString(Text1(1)) + _
         "' AND '" + ChangeTStringToWString(Text1(2)) + "' "
   End If
   If Me.Tag = 0 Then
      strExc(1) = strExc(1) & " AND CP01='LA' "
   Else
      'Modify By Sindy 2009/07/24 增加LIN系統類別
      strExc(1) = strExc(1) & " AND CP01 in ('FCL','LIN') "
   End If
   'Modify By Sindy 2010/5/10
   GetSQL1 = strExc(1)
'   'Add By Sindy 2010/01/20
'   If Trim(Text1(4)) <> "" Then strExc(1) = strExc(1) & " and CP14='" + Trim(Text1(4)) + "'"
'   '2010/01/20 End
'   GetSQL1 = strExc(1) & " AND CP57 IS NULL AND CP09<'C' "
End Function

Private Function GetSQL2() As String
 Dim i As Integer, St As String
   If Text1(0).Text = "1" Then
      St = "CP05"
   Else
      St = "CP27"
   End If
   If Text1(1) = "" And Text1(2) <> "" Then
      strExc(1) = " AND " & St & " <='" & ChangeTStringToWString(Text1(2)) + "'"
   ElseIf Text1(1) <> "" And Text1(2) <> "" Then
      strExc(1) = " AND " & St & " BETWEEN '" & ChangeTStringToWString(Text1(1)) + _
         "' AND '" + ChangeTStringToWString(Text1(2)) + "'"
   End If
   If Me.Tag = 0 Then
      strExc(1) = strExc(1) & " AND CP01='L'"
   Else
      strExc(1) = strExc(1) & " AND CP01='CFL'"
   End If
   'Add By Sindy 2010/01/20
   If Trim(Text1(4)) <> "" Then strExc(1) = strExc(1) & " and CP14='" + Trim(Text1(4)) + "'"
   '2010/01/20 End
   GetSQL2 = strExc(1) & " AND CP57 IS NULL AND CP26 IS NULL AND CP09<'C' "
End Function

'Add By Sindy 2010/01/20
Private Function GetSQL3(ByVal strSysID As String) As String
Dim i As Integer, St As String
   If Text1(0).Text = "1" Then
      St = "CP05"
   Else
      St = "CP27"
   End If
   If Text1(1) = "" And Text1(2) <> "" Then
      strExc(1) = " AND " & St & " <='" & ChangeTStringToWString(Text1(2)) + "' "
   ElseIf Text1(1) <> "" And Text1(2) <> "" Then
      strExc(1) = " AND " & St & " BETWEEN '" & ChangeTStringToWString(Text1(1)) + _
         "' AND '" + ChangeTStringToWString(Text1(2)) + "' "
   End If
   If strSysID = "1" Then
      strExc(1) = strExc(1) & " AND CP01 in ('FCL','LIN') "
   Else
      strExc(1) = strExc(1) & " AND CP01='CFL' "
   End If
   'Modify By Sindy 2010/5/10
   GetSQL3 = strExc(1)
   'If Trim(Text1(4)) <> "" Then strExc(1) = strExc(1) & " and CP14='" + Trim(Text1(4)) + "'"
   'GetSQL3 = strExc(1) & " AND CP57 IS NULL AND CP09<'C' GROUP BY CP14 ORDER BY CP14 "
End Function

Private Sub Form_Unload(Cancel As Integer)
   Set frm084001 = Nothing
End Sub

'Added by Lydia 2023/01/16 外法: 新格式Word
Private Function PrintCaseWord(ByVal iCall As Integer) As Boolean
 Dim rsQuery As New ADODB.Recordset
 Dim strGrp As String, strTitle As String
 Dim iRound  As Integer, intQ As Integer
 Dim strCon As String, strCon2 As String
 Dim m_TempPDF As String
 
On Error GoTo ErrHand
   
   PrintCaseWord = False
   
   iPage = 0: intCounter = 0
   For iRound = IIf(Text1(3) > "2", 1, Val(Text1(3))) To iCall
      strCon = "": strCon2 = ""
      If Text1(1) = "" And Text1(2) <> "" Then
         strCon = " AND " & IIf(Text1(0) = "1", "CP05", "CP27") & " <='" & ChangeTStringToWString(Text1(2)) + "' "
      ElseIf Text1(1) <> "" And Text1(2) <> "" Then
         strCon = " AND " & IIf(Text1(0) = "1", "CP05", "CP27") & " BETWEEN '" & ChangeTStringToWString(Text1(1)) + _
            "' AND '" + ChangeTStringToWString(Text1(2)) + "' "
      End If
      strCon = "(Cp09<'B' Or (Substr(Cp09,1,1)='B' And Cp60 Is Not Null))" & strCon
      If iRound = 1 Then
         If Me.Tag = "1" Then '外法FCL
             strCon2 = " AND CP01 in ('FCL','LIN') "
             strTitle = "FCL"
         Else                          '內法LA
             strCon2 = " AND CP01 in ('LA') "
             strTitle = "LA"
         End If
      ElseIf iRound = 2 Then
         If Me.Tag = "1" Then '外法CFL
             strCon2 = " AND CP01 in ('CFL') "
             strTitle = "CFL"
         Else             '內法L
             strCon2 = " AND CP01 in ('L') "
             strTitle = "L"
         End If
      End If
      '收/發文次數, 點數
      strSql = "Select Decode(Cp60,Null,'2',Null) Type,Cp09,Cp60,1 Cnt,((nvl(cp16,0)-nvl(a1u07,0)-nvl(a1u09,0))-(nvl(cp17,0)-nvl(a1u09,0)))/1000 Cp18, " & _
                   "Cp14 A1n04,0 A1n05,0 A1n05d,St02,0 A1n05b,0 A1n05c From Caseprogress,Staff, " & _
                  "(SELECT a1u03,sum(a1u04) a1u04,sum(a1u05) a1u05,sum(a1u07) a1u07,sum(a1u08) a1u08,sum(a1u09) a1u09,sum(a1u10) a1u10 " & _
                  "From Acc1u0 Where A1u03 In " & _
                  "(Select Cp09 From Caseprogress Where " & strCon & " And Instr(Cp01,'L')>0 And Cp159=0 And Cp60 Is Not Null) " & _
                  "GROUP BY a1u03) " & _
                 "Where " & strCon & strCon2 & " and cp159=0 And Cp14=St01(+) And Cp09=A1u03(+) "
      '請款點數
      strSql = strSql & "Union Select '' Type,cp09,Cp60,0 Cnt,0 Cp18,A1n04,A1n05,0 a1n05d,St02,0 a1n05b,0 a1n05c From Caseprogress,Acc1n0,Staff " & _
                  "Where " & strCon & strCon2 & " And Cp60=A1n01(+) And A1n02='2' and A1n04=St01(+) and substr(st15,1,1)='L' "
      '分配智慧所點數
      strSql = strSql & "Union Select '' Type,cp09,Cp60,0 Cnt,0 Cp18,A1n04,0 A1n05,a1n05 a1n05d,St02,0 a1n05b,0 a1n05c From Caseprogress,Acc1n0,Staff " & _
                  "Where " & strCon & strCon2 & " And Cp60=A1n01(+) And A1n02='2' And A1n04=St01(+) and substr(st15,1,1)<>'L' "
      '工作點數
      strSql = strSql & "Union Select '' Type,cp09,Cp60,0 Cnt,0 Cp18,A1n04,0 A1n05,0 A1n05d,St02,a1n05 a1n05b,0 a1n05c From Caseprogress,Acc1n0,Staff " & _
                  "Where " & strCon & strCon2 & " And a1n03(+)=cp09 And A1n02='3' And A1n04=St01(+) "
      '智慧所國外案分配點數
      If InStr(strCon2, "FCL") > 0 Then '有FCL案才抓
         strSql = strSql & "Union Select '' Type,Cp09,Cp60,0 Cnt,0 Cp18,A1n04,0 A1n05,0 a1n05d,St02,0 A1n05b,a1n05 a1n05c From Caseprogress,Acc1n0,Staff " & _
                     "Where " & strCon & " and Instr(Cp01,'L')=0 And A1n03(+)=Cp09 And A1n02='2' And A1n04=St01(+) and st03 like 'L%' "
      End If
      '合併
      strSql = "Select Decode(Type,'2','未請款',Type) 請款,A1n04,St02 承辦人,Sum(Cnt) 收文次數,Sum(Cp18) 收文點數," & _
                  "Sum(A1n05) 請款點數,Sum(A1n05d) 分配智慧所點數,Sum(A1n05+A1n05d) 請款總點數, " & _
                  "sum(a1n05b) 工作點數,sum(a1n05c) 智慧所國外案分配點數 From (" & strSql & ") " & _
                  "Group By decode(Type,'2','未請款',type),a1n04,st02 order by 1 desc,2 "
      intQ = 1
      Set rsQuery = ClsLawReadRstMsg(intQ, strSql)
      If intQ = 1 Then
          If strGrp <> strTitle Then
               If strGrp = "" Then
                   If NewListPage(strTitle) = False Then GoTo ErrHand
               Else
                   With g_WordAp
                      .Selection.SelectRow
                      .Selection.Borders(wdBorderBottom).LineStyle = wdLineStyleSingle '用下部框線當做分隔線
                      .Selection.Collapse Direction:=wdCollapseStart
                      .Selection.MoveRight Unit:=wdCharacter, Count:=1
                      .Selection.TypeText Text:="總　計"
                      .Selection.MoveRight Unit:=wdCharacter, Count:=1
                      For intI = 0 To 6
                         .Selection.TypeText Text:=Format(strTemp(intI), "##0.0")
                         .Selection.MoveRight Unit:=wdCharacter, Count:=1
                      Next intI
                      .Selection.MoveRight Unit:=wdCharacter, Count:=2
                      .Selection.TypeParagraph
                      .Selection.TypeParagraph
                   End With
                   intCounter = intCounter + 3
                   If NewListPage(strTitle) = False Then GoTo ErrHand
               End If
          End If
          strTemp(0) = "": strTemp(1) = "": strTemp(2) = "": strTemp(3) = "": strTemp(4) = "": strTemp(5) = "": strTemp(6) = ""
          rsQuery.MoveFirst
          Do While Not rsQuery.EOF
               With g_WordAp.Application
                  .Selection.TypeText Text:="" & rsQuery.Fields("請款")
                  .Selection.MoveRight Unit:=wdCharacter, Count:=1
                  strExc(1) = ""
                  If "" & rsQuery.Fields("a1n04") <> "" Then
                     If PUB_ChkLCompStaff("" & rsQuery.Fields("a1n04")) = False Then
                         strExc(1) = "*"
                     End If
                  End If
                  .Selection.TypeText Text:="" & rsQuery.Fields("承辦人") & strExc(1)
                  .Selection.MoveRight Unit:=wdCharacter, Count:=1
                  strTemp(0) = Val(strTemp(0)) + Val(Format("" & rsQuery.Fields("收文次數"), "##0"))
                  .Selection.TypeText Text:=Format("" & rsQuery.Fields("收文次數"), "##0")
                  .Selection.MoveRight Unit:=wdCharacter, Count:=1
                  strTemp(1) = Val(strTemp(1)) + Val(Format("" & rsQuery.Fields("收文點數"), "##0.0"))
                  .Selection.TypeText Text:=Format("" & rsQuery.Fields("收文點數"), "##0.0")
                  .Selection.MoveRight Unit:=wdCharacter, Count:=1
                  strTemp(2) = Val(strTemp(2)) + Val(Format("" & rsQuery.Fields("請款點數"), "##0.0"))
                  .Selection.TypeText Text:=Format("" & rsQuery.Fields("請款點數"), "##0.0")
                  .Selection.MoveRight Unit:=wdCharacter, Count:=1
                  strTemp(3) = Val(strTemp(3)) + Val(Format("" & rsQuery.Fields("分配智慧所點數"), "##0.0"))
                  .Selection.TypeText Text:=Format("" & rsQuery.Fields("分配智慧所點數"), "##0.0")
                  .Selection.MoveRight Unit:=wdCharacter, Count:=1
                  strTemp(4) = Val(strTemp(4)) + Val(Format("" & rsQuery.Fields("請款總點數"), "##0.0"))
                  .Selection.TypeText Text:=Format("" & rsQuery.Fields("請款總點數"), "##0.0")
                  .Selection.MoveRight Unit:=wdCharacter, Count:=1
                  strTemp(5) = Val(strTemp(5)) + Val(Format("" & rsQuery.Fields("工作點數"), "##0.0"))
                  .Selection.TypeText Text:=Format("" & rsQuery.Fields("工作點數"), "##0.0")
                  .Selection.MoveRight Unit:=wdCharacter, Count:=1
                  strTemp(6) = Val(strTemp(6)) + Val(Format("" & rsQuery.Fields("智慧所國外案分配點數"), "##0.0"))
                  .Selection.TypeText Text:=Format("" & rsQuery.Fields("智慧所國外案分配點數"), "##0.0")
                  .Selection.MoveRight Unit:=wdCharacter, Count:=1
                  .Selection.InsertRows
                  .Selection.Collapse Direction:=wdCollapseStart
               End With
               intCounter = intCounter + 1
               If intCounter > iLimit Then
                  If NewListPage(strTitle, "1") = False Then GoTo ErrHand
               End If
              rsQuery.MoveNext
          Loop

          strGrp = strTitle
      End If
   Next iRound
   
   If intCounter > 0 Then
      '列印總計
      With g_WordAp.Application
          .Selection.SelectRow
          .Selection.Borders(wdBorderBottom).LineStyle = wdLineStyleSingle '用下部框線當做分隔線
          .Selection.Collapse Direction:=wdCollapseStart
          .Selection.MoveRight Unit:=wdCharacter, Count:=1
          .Selection.TypeText Text:="總　計"
          .Selection.MoveRight Unit:=wdCharacter, Count:=1
          For intI = 0 To 6
             .Selection.TypeText Text:=Format(strTemp(intI), "##0.0")
             .Selection.MoveRight Unit:=wdCharacter, Count:=1
          Next intI
          .ActiveDocument.PrintOut Background:=False, Copies:=1, Collate:=True
      End With

      Pub_RePosWord g_WordAp, bVisible, m_WordLeft, m_WordTop
      g_WordAp.Quit wdDoNotSaveChanges
      Set g_WordAp = Nothing '避免快速開啟Word,程式出錯
      
      PrintCaseWord = True
   End If
   
   Set rsQuery = Nothing
   Exit Function
   
ErrHand:
   MsgBox Err.Description
End Function

'Added by Lydia 2023/01/16
Private Function NewListPage(ByVal iTitle As String, Optional ByVal iType As String = "0") As Boolean
'iType: 1=跳頁
    NewListPage = False
    '開啟Word檔
    If iPage = 0 Then
        If Pub_NewWordDoc(g_WordAp, bVisible, m_WordLeft, m_WordTop) = False Then Exit Function
        
        With g_WordAp.Application
           '邊界
           .Selection.PageSetup.LeftMargin = .CentimetersToPoints(1)
           .Selection.PageSetup.RightMargin = .CentimetersToPoints(1)
           .Selection.PageSetup.TopMargin = .CentimetersToPoints(1)
           .Selection.PageSetup.BottomMargin = .CentimetersToPoints(1)
        End With
        iPage = iPage + 1
    End If
    '跳頁
    If iType = "1" Then
        '刪除原有表格
        g_WordAp.Selection.SelectRow
        g_WordAp.Selection.Cells.Delete
        g_WordAp.Selection.SelectRow
        g_WordAp.Selection.Cells.Delete
        g_WordAp.Selection.InsertBreak Type:=wdPageBreak
        g_WordAp.Selection.GoTo what:=wdGoToPage, which:=wdGoToNext, Count:=1
        iPage = iPage + 1
        intCounter = 0
    End If

    '列印表頭
    With g_WordAp.Application
         '新增表格(1*3)
        .Selection.Tables.add Range:=.Selection.Range, NumRows:=1, NumColumns:=3
        intCounter = intCounter + 1
        With .Selection.Tables(1)
          .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
          .Borders(wdBorderRight).LineStyle = wdLineStyleNone
          .Borders(wdBorderTop).LineStyle = wdLineStyleNone
          .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
          .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
          .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
          .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
          .Borders.Shadow = False
        End With
        .Selection.SelectRow
        .Selection.Cells.VerticalAlignment = wdAlignVerticalCenter
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Selection.Cells.SetHeight RowHeight:=16, HeightRule:=wdRowHeightAuto  '自動列高
        .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(5), RulerStyle:=wdAdjustProportional
        .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(9.5), RulerStyle:=wdAdjustProportional
        
        .Selection.InsertRows 1
        .Selection.Collapse Direction:=wdCollapseStart
        .Selection.SelectRow
        .Selection.Cells.Merge
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Selection.Cells.SetHeight RowHeight:=36, HeightRule:=wdRowHeightExactly '固定列高
        .Selection.Font.Size = 24
        .Selection.Font.Bold = True
        .Selection.Font.Underline = True
        .Selection.TypeText Text:=iTitle & IIf(Text1(0) = "1", "收文", "發文") & "統計表"
        .Selection.MoveRight Unit:=wdCharacter, Count:=2
        .Selection.SelectRow
        .Selection.Font.Size = 12
        .Selection.Font.Bold = False
        .Selection.Font.Underline = False
        .Selection.InsertRows 1
        intCounter = intCounter + 1
        
        If intCounter < 5 Then '排除第二份報表
            .Selection.Collapse Direction:=wdCollapseStart
            .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
            .Selection.TypeText Text:="列印人員：" & strUserName
            .Selection.MoveRight Unit:=wdCharacter, Count:=1
            strExc(1) = ""
            If Text1(1) <> "" Or Text1(2) <> "" Then
                strExc(1) = IIf(Text1(0) = "1", "收文", "發文") & "日期：" & IIf(Text1(1) = "", "   ", ChangeTStringToTDateString(Text1(1))) & " - " & IIf(Text1(2) = "", "   ", ChangeTStringToTDateString(Text1(2)))
            End If
            .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
            .Selection.TypeText Text:=strExc(1)
            .Selection.MoveRight Unit:=wdCharacter, Count:=1
            .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
            .Selection.TypeText Text:="列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
            .Selection.MoveRight Unit:=wdCharacter, Count:=2
            intCounter = intCounter + 1
            .Selection.MoveRight Unit:=wdCharacter, Count:=2
            .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
            .Selection.TypeText Text:="頁　　數：" & iPage
            .Selection.MoveRight Unit:=wdCharacter, Count:=2
            intCounter = intCounter + 1
            .Selection.InsertRows
        End If
        .Selection.SelectRow
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Selection.Cells.Merge
        .Selection.Cells.Split 1, 9
        .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(1.7), RulerStyle:=wdAdjustProportional
        .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(2.5), RulerStyle:=wdAdjustProportional
        .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(2), RulerStyle:=wdAdjustProportional
        .Selection.Cells(4).SetWidth ColumnWidth:=.CentimetersToPoints(2), RulerStyle:=wdAdjustProportional
        .Selection.Cells(5).SetWidth ColumnWidth:=.CentimetersToPoints(2), RulerStyle:=wdAdjustProportional
        .Selection.Cells(6).SetWidth ColumnWidth:=.CentimetersToPoints(2), RulerStyle:=wdAdjustProportional
        .Selection.Cells(7).SetWidth ColumnWidth:=.CentimetersToPoints(2), RulerStyle:=wdAdjustProportional
        .Selection.Cells(8).SetWidth ColumnWidth:=.CentimetersToPoints(2), RulerStyle:=wdAdjustProportional
         
        .Selection.Borders(wdBorderTop).LineStyle = wdLineStyleSingle '用上部框線當做分隔線
        .Selection.Font.Size = 11
        .Selection.Collapse Direction:=wdCollapseStart
        '欄位抬頭
        .Selection.TypeText Text:="請款"
        .Selection.MoveRight Unit:=wdCharacter, Count:=1
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
        .Selection.TypeText Text:="承辦人"
        .Selection.MoveRight Unit:=wdCharacter, Count:=1
        .Selection.TypeText Text:=IIf(Text1(0) = "1", "收文", "發文") & "次數"
        .Selection.MoveRight Unit:=wdCharacter, Count:=1
        .Selection.TypeText Text:=IIf(Text1(0) = "1", "收文", "發文") & "點數"
        .Selection.MoveRight Unit:=wdCharacter, Count:=1
        .Selection.TypeText Text:="請款點數"
        .Selection.MoveRight Unit:=wdCharacter, Count:=1
        .Selection.TypeText Text:="分配智慧" & vbCrLf & "所點數"
        .Selection.MoveRight Unit:=wdCharacter, Count:=1
        .Selection.TypeText Text:="請款" & vbCrLf & " 總點數"
        .Selection.MoveRight Unit:=wdCharacter, Count:=1
        .Selection.TypeText Text:="工作點數"
        .Selection.MoveRight Unit:=wdCharacter, Count:=1
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
        .Selection.TypeText Text:="智慧所國外" & vbCrLf & "案分配點數"
        .Selection.MoveRight Unit:=wdCharacter, Count:=1
        intCounter = intCounter + 1
        .Selection.InsertRows 1
        .Selection.Borders(wdBorderTop).LineStyle = wdLineStyleNone
        .Selection.SelectRow
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
        .Selection.Collapse Direction:=wdCollapseStart
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Selection.MoveRight Unit:=wdCharacter, Count:=1
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
        .Selection.SelectRow
        .Selection.Collapse Direction:=wdCollapseStart
    End With
    NewListPage = True
End Function
