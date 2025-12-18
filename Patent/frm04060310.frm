VERSION 5.00
Begin VB.Form frm04060310 
   BorderStyle     =   1  '單線固定
   Caption         =   "專利公報國內各區同業排名"
   ClientHeight    =   2310
   ClientLeft      =   2790
   ClientTop       =   3945
   ClientWidth     =   4650
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   4650
   Begin VB.TextBox txtDate 
      Height          =   345
      Index           =   1
      Left            =   2640
      MaxLength       =   5
      TabIndex        =   1
      Top             =   840
      Width           =   795
   End
   Begin VB.TextBox txtDate 
      Height          =   345
      Index           =   0
      Left            =   1650
      MaxLength       =   5
      TabIndex        =   0
      Top             =   840
      Width           =   795
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   3750
      TabIndex        =   3
      Top             =   90
      Width           =   756
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Excel(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2700
      TabIndex        =   2
      Top             =   90
      Width           =   990
   End
   Begin VB.Label Label2 
      Caption         =   "PS：程式會與前一年同時段做比較"
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   600
      TabIndex        =   6
      Top             =   1440
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "公報年月："
      Height          =   210
      Index           =   1
      Left            =   600
      TabIndex        =   5
      Top             =   900
      Width           =   1245
   End
   Begin VB.Label Label3 
      Caption         =   "PS2：Excel儲存於"
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   510
      TabIndex        =   4
      Top             =   1920
      Width           =   3195
   End
   Begin VB.Line Line1 
      X1              =   2160
      X2              =   2850
      Y1              =   1035
      Y2              =   1035
   End
End
Attribute VB_Name = "frm04060310"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2021/11/17 Form2.0已檢查 (無需修改的物件)
'Create By Lydia 2017/01/05 專利公報國內各區同業排名
Option Explicit

Private Sub cmdOK_Click(Index As Integer)
Dim Cancel As Boolean
   
   Select Case Index
      Case 0

         For intI = 0 To 1
             If Trim(txtDate(intI)) = "" Then
                MsgBox IIf(intI = 0, "起始", "截止") & "公報年月不可空白！", vbInformation, "輸入錯誤！"
                txtDate(intI).SetFocus
                Exit Sub
             End If
             ChkTxtDate intI, Cancel
             If Cancel = True Then
                txtDate(intI).SetFocus
             End If
         Next
         If Val(txtDate(1)) < Val(txtDate(0)) Then
            MsgBox "截止年月必須大於起始年月！", vbInformation, "輸入錯誤！"
            txtDate(1).SetFocus
            Exit Sub
         End If
         Screen.MousePointer = vbHourglass
         If StrMenu = False Then
         End If
         Screen.MousePointer = vbDefault
         
      Case 1
         Unload Me
   End Select
End Sub

Private Function StrMenu() As Boolean
Dim startX As String '行-起始位置
Dim endX As String '台一與聖島比較的國內合計
Dim m_rs As New ADODB.Recordset
Dim strVol1_S As String, strVol1_E As String '日期範圍
Dim strVol2_S As String, strVol2_E As String '前一年日期範圍

Dim m_TopNum As Integer '前X名
Dim strMid As String, strMid2 As String
Dim xlsSalesPoint As New Excel.Application
Dim wks4630 As New Worksheet
Dim xRows As Integer '目前列位置
Dim midR As Integer '台一與聖島比較的抬頭位置

Dim strTemp As String, strTemp2 As String
Dim strPath As String, strTempFile As String
Dim iCall As Integer  '切換國家檔台灣的各區,國內合計
Dim inJ As Integer
Dim mArea As Integer '各區計數
Dim mCall As Integer '切換今年度和前一年
Dim mESeqNo As String '暫存檔序號

   StrMenu = False

   strVol1_S = TransDate(txtDate(0) & "01", 2)
   strVol1_E = GetLastDay(TransDate(txtDate(1) & "01", 2))
   strVol2_S = CompDate(0, -1, strVol1_S)
   strVol2_E = CompDate(0, -1, strVol1_E)
   m_TopNum = 10 '預設前10名
   
   'Added By Lydia 2022/01/13 查詢印表記錄檔欄位
   ClearQueryLog (Me.Name)
   pub_QL05 = pub_QL05 & ";" & Label1(1).Caption & strVol1_S & "-" & strVol1_E
   pub_QL05 = pub_QL05 & ";前期:" & strVol2_S & "-" & strVol2_E
  'end 2022/01/13
  
   strMid = "SELECT TPB01,TPB02,TPB03,TPB04,TPB05,decode(tpb06,'020','C0',substr(na02,1,2))||TPB06 TPB06,TPB07,TPB08,'' TPB09,PA01,PA02,PA03,PA04 " & _
            "From TPBulletin, nation, PATENT where na01(+)=tpb06 and pa11(+)=tpb01 and pa23(+)='1'"
            
   '暫存檔- 無代理人且為本所案件,若最後發文的AB類程序為不出名則算台一案件
   strSql = "SELECT TPB01,TPB02,TPB03,TPB04,TPB05,TPB06,TPB07,CP01,CP02,CP03,CP04,substr(max(cp27||cp22),9) A1 FROM CASEPROGRESS, " & _
            "(" & strMid & " and TPB03 >=" & strVol2_S & " and TPB03 <=" & strVol1_E & " and tpb07 is null and pa01 is not null) WHERE PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) " & _
            " and cp09<'C' and cp27<TPB03 GROUP BY TPB01,TPB02,TPB03,TPB04,TPB05,TPB06,TPB07,CP01,CP02,CP03,CP04 "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      Set m_rs = PUB_CreateRecordset(RsTemp, , , , Me.Name, mESeqNo)
      '刪除不符合的資料
      strSql = "Delete from rdatafactory where id = '" & strUserNum & "' and formname = '" & Me.Name & "' and seqno = '" & mESeqNo & "' and nvl(r012,'A')='A' "
      cnnConnection.Execute strSql, intI
   End If
   
   '各區和國內合計的次數
   strSql = "SELECT COUNT(*) CNT FROM NATION WHERE NA01>'000' AND NA01<='010' AND NA02 LIKE 'A%'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then mArea = RsTemp.Fields("cnt")
   
   '排名資料
   '-------------切換今年度和前一年
   For mCall = 1 To 2
       For iCall = 1 To mArea + 1
           If m_rs.State = 1 Then m_rs.Close
           m_rs.CursorLocation = adUseClient
           If mCall = 1 Then '今年度
              strMid2 = " and TPB03 >=" & CNULL(strVol1_S) & " and TPB03 <=" & CNULL(strVol1_E)
           Else              '前一年
              strMid2 = " and TPB03 >=" & CNULL(strVol2_S) & " and TPB03 <=" & CNULL(strVol2_E)
           End If
           '大於mArea=國內合計
           If iCall > mArea Then
                strSql = "select '國內' title,tpb08,count(*) cnt from nation,(" & strMid & strMid2 & _
                    ") where substr(tpb06,1,1)='A' and substr(tpb06,3)=na01(+) and nvl(tpb08,' ') <> ' '" & _
                     "group by tpb08 order by 3 desc"
           Else
                strSql = "select substr(na03,1,4) title,tpb08,count(*) cnt from nation,(" & strMid & strMid2 & _
                         ") where substr(tpb06,1,1)='A' and substr(tpb06,3)=na01(+) and nvl(tpb08,' ') <> ' ' and substr(tpb06,5,1)='" & iCall & "' " & _
                          "group by substr(na03,1,4),tpb08 order by 3 desc"
           End If
           
           m_rs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
           If mCall = 1 And iCall = 1 And m_rs.RecordCount = 0 Then
              MsgBox "資料庫無資料！", vbExclamation
              InsertQueryLog (0) 'Added by Lydia 2022/01/13
              Exit Function
           End If
           If Not m_rs.EOF And Not m_rs.BOF Then
           '---------------開啟xls檔
              If mCall = 1 And iCall = 1 Then
                  InsertQueryLog (m_rs.RecordCount) 'Added by Lydia 2022/01/13
                  StrMenu = True
                  startX = "c"
                  xRows = 1
                  strTempFile = Me.Caption & txtDate(0) & "至" & txtDate(1) & "-" & ACDate(ServerDate) & ServerTime & MsgText(43)
                  strPath = strExcelPath & strTempFile
                  
                  If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = "" Then
                     MkDir strExcelPath
                  End If
                  If Dir(strPath) <> "" Then
                     Kill strPath
                  End If
                  
                  xlsSalesPoint.SheetsInNewWorkbook = 1 'Added by Lydia 2019/03/13 預設工作表數量
                  xlsSalesPoint.Workbooks.add
                  Set wks4630 = xlsSalesPoint.Worksheets(1)
                  wks4630.PageSetup.Orientation = xlPortrait '直印
                  '抬頭
                   wks4630.PageSetup.PrintTitleRows = "$1:$3"
                   
                   wks4630.Columns("a:a").ColumnWidth = 10 '項目
                   wks4630.Columns("b:b").ColumnWidth = 10
                   For intI = Asc(startX) To Asc(startX) + m_TopNum '多年度一欄
                       wks4630.Columns(Chr(intI) & ":" & Chr(intI)).ColumnWidth = 13
                       wks4630.Columns(Chr(intI) & ":" & Chr(intI)).HorizontalAlignment = xlCenter
                       wks4630.Columns(Chr(intI) & ":" & Chr(intI)).VerticalAlignment = xlBottom
                   Next
                   
                   strExc(0) = Mid(txtDate(0), 1, 3) & "/" & Mid(txtDate(0), 4, 2) & "-" & Mid(txtDate(1), 1, 3) & "/" & Mid(txtDate(1), 4, 2)
                   strExc(0) = strExc(0) & "和" & Mid(Val(txtDate(0)) - 100, 1, 3) & "/" & Mid(Val(txtDate(0)) - 100, 4, 2) & "-" & Mid(Val(txtDate(1)) - 100, 1, 3) & "/" & Mid(Val(txtDate(1)) - 100, 4, 2)
                   wks4630.Range("a1").Value = strExc(0) & " " & Me.Caption
                   With wks4630.Range("a1:" & Chr(Asc(startX) + m_TopNum - 1) & "1")
                     .WrapText = False
                     .MergeCells = True
                     .HorizontalAlignment = xlCenter
                     .VerticalAlignment = xlBottom
                  End With
            
                  xRows = 3
                  '排名編號
                  For intI = 1 To m_TopNum
                     wks4630.Range(Chr(Asc(startX) + intI - 1) & xRows).Value = intI
                  Next
                  xRows = xRows + 1
                  
              '重設前一年度的起始列
              ElseIf mCall = 2 And iCall = 1 Then
                  xRows = 6
                  strTemp = ""
              End If
           '---------------
              
              With m_rs
                  .MoveFirst
                  inJ = 1
                    '切換左邊標題(a=區域,b=年度)
                    If strTemp <> .Fields("title") Then
                       If strTemp <> "" Then xRows = xRows + 4 '跳過今年的資料列
                       If mCall = 1 Then wks4630.Range("a" & xRows).Value = "" & .Fields("title") '區域
                       If wks4630.Range("b" & xRows).Value = "" Then
                          wks4630.Range("b" & xRows).Value = IIf(mCall = 1, Mid(txtDate(0), 1, 3) & "年", Val(Mid(txtDate(0), 1, 3)) - 1 & "年")  '年度
                          wks4630.Range("b" & xRows + 1).Value = "筆數"
                       End If
                       strTemp = .Fields("title")
                    End If
                  '讀取前X名資料
                  Do While inJ <= m_TopNum
                     '事務所名稱
                     wks4630.Range(Chr(Asc(startX) + inJ - 1) & xRows).Value = "" & .Fields("tpb08")
                     '筆數
                     wks4630.Range(Chr(Asc(startX) + inJ - 1) & xRows + 1).Value = Val("" & .Fields("cnt"))
                     '無代理人且為本所案件,若最後發文的AB類程序為不出名則算台一案件
                     If "" & .Fields("tpb08") = "台一國際" Then
                         strExc(7) = Replace(UCase(strMid2), "TPB", "R0")
                         If iCall > mArea Then '國內合計
                            strExc(7) = strExc(7) & "and substr(R006,1,1)='A' "
                         Else
                            strExc(7) = strExc(7) & "and substr(R006,1,1)='A' and substr(R006,5,1)='" & iCall & "' "
                         End If
                         strSql = "SELECT COUNT(*) Add1 FROM rdatafactory " & _
                                  "where id = '" & strUserNum & "' and formname = '" & Me.Name & "' and seqno = '" & mESeqNo & "' " & strExc(7)
                         intI = 1
                         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                         If intI = 1 Then
                            If Val("" & RsTemp.Fields("add1")) > 0 Then
                                wks4630.Range(Chr(Asc(startX) + inJ - 1) & xRows + 1).Value = Val(wks4630.Range(Chr(Asc(startX) + inJ - 1) & xRows + 1).Value) + Val("" & RsTemp.Fields("add1"))
                            End If
                         End If
                     End If
                     
                     inJ = inJ + 1
                     .MoveNext
                     If .EOF = True Then Exit Do
                  Loop
              End With
           End If
       Next iCall
       
       '跳台一和聖島比較
       If mCall = 2 Then midR = xRows + 4
   Next mCall
   '----------------切換今年度和前一年
   
   '抓台一和聖島比較
   strExc(0) = Mid(txtDate(0), 1, 3) & "/" & Mid(txtDate(0), 4, 2) & "-" & Mid(txtDate(1), 1, 3) & "/" & Mid(txtDate(1), 4, 2)
   strExc(0) = strExc(0) & "和" & Mid(Val(txtDate(0)) - 100, 1, 3) & "/" & Mid(Val(txtDate(0)) - 100, 4, 2) & "-" & Mid(Val(txtDate(1)) - 100, 1, 3) & "/" & Mid(Val(txtDate(1)) - 100, 4, 2)
   wks4630.Range("a" & midR).Value = strExc(0) & " " & "國內各區與聖島比較"
   With wks4630.Range("a" & midR & ":j" & midR)
     .WrapText = False
     .MergeCells = True
     .HorizontalAlignment = xlCenter
     .VerticalAlignment = xlBottom
   End With
   
   xRows = midR + 1
   strTemp = ""
   If m_rs.State = 1 Then m_rs.Close
   m_rs.CursorLocation = adUseClient
   
   strSql = "select '" & Trim(Mid(txtDate(0), 1, 3)) & "' t_yy,substr(tpb06,5,1) ano,substr(na03,1,4) title,tpb08,count(*) cnt from nation, " & _
            "(" & strMid & " and TPB03 >=" & strVol1_S & " and TPB03 <=" & strVol1_E & _
            ") where substr(tpb06,1,1)='A' and tpb08 in ('聖島國際','台一國際') and substr(tpb06,3)=na01(+) group by substr(tpb06,5,1),substr(na03,1,4),tpb08 "
   strSql = strSql & "union all select '" & Trim(Val(Mid(txtDate(0), 1, 3)) - 1) & "' t_yy,substr(tpb06,5,1) ano,substr(na03,1,4) title,tpb08,count(*) cnt from nation, " & _
            "(" & strMid & " and TPB03 >=" & strVol2_S & " and TPB03 <=" & strVol2_E & _
            ") where substr(tpb06,1,1)='A' and tpb08 in ('聖島國際','台一國際') and substr(tpb06,3)=na01(+) group by substr(tpb06,5,1),substr(na03,1,4),tpb08 " & _
            "order by ano, 1 desc,tpb08 "
   m_rs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If Not m_rs.EOF And Not m_rs.BOF Then
      With m_rs
         .MoveFirst
         endX = startX
         Do While Not .EOF
            '切換標題(上方:區域,左邊:年度,事務所,差距)
            If strTemp <> .Fields("ano") Then
               '上方:區域
               If strTemp <> "" Then endX = Chr(Asc(endX) + 1)
               wks4630.Range(endX & midR + 1).Value = "" & .Fields("title")
               xRows = midR + 2
               
               '左邊:年度,事務所,差距
               If wks4630.Range("a" & xRows).Value = "" Then
                  wks4630.Range("a" & xRows).Value = Mid(txtDate(0), 1, 3) & "年"
                  wks4630.Range("b" & xRows).Value = "台一國際"
                  wks4630.Range("b" & xRows + 1).Value = "聖島國際"
                  wks4630.Range("b" & xRows + 2).Value = "差距"
                  wks4630.Range("a" & xRows + 3).Value = Val(Mid(txtDate(0), 1, 3)) - 1 & "年"
                  wks4630.Range("b" & xRows + 3).Value = "台一國際"
                  wks4630.Range("b" & xRows + 4).Value = "聖島國際"
                  wks4630.Range("b" & xRows + 5).Value = "差距"
               End If
               strTemp = "" & .Fields("ano")
               
               'Added by Lydia 2017/09/05 2017/09/05 台一若當期無案件,需要跳行
               If "" & .Fields("tpb08") <> "台一國際" Then
                  wks4630.Range(endX & xRows).Value = "0"
                  xRows = xRows + 1
               End If
               'end 2017/09/05
            Else '跳到前一年
                If strTemp2 <> "" & .Fields("t_yy") And strTemp2 <> "" Then
                   xRows = midR + 5
                  'Added by Lydia 2017/09/05 2017/09/05 台一若當期無案件,需要跳行
                  If "" & .Fields("tpb08") <> "台一國際" Then
                     wks4630.Range(endX & xRows).Value = "0"
                     xRows = xRows + 1
                  End If
                  'end 2017/09/05
                End If
            End If

            '筆數
            'Remove by Lydia 2017/09/05 有可能當期無案件,需要跳行
            wks4630.Range(endX & xRows).Value = "" & .Fields("cnt")
            '無代理人且為本所案件,若最後發文的AB類程序為不出名則算台一案件
            If "" & .Fields("tpb08") = "台一國際" Then
                 If "" & .Fields("t_yy") = Trim(Mid(txtDate(0), 1, 3)) Then '今年度
                    strExc(7) = " and R003 >=" & CNULL(strVol1_S) & " and R003 <=" & CNULL(strVol1_E)
                 Else              '前一年
                    strExc(7) = " and R003 >=" & CNULL(strVol2_S) & " and R003 <=" & CNULL(strVol2_E)
                 End If
                 '各區
                 strExc(7) = strExc(7) & " and substr(R006,1,1)='A' and substr(R006,5,1)='" & .Fields("ano") & "' "

                 strSql = "SELECT COUNT(*) Add1 FROM rdatafactory " & _
                          "where id = '" & strUserNum & "' and formname = '" & Me.Name & "' and seqno = '" & mESeqNo & "' " & strExc(7)
                intI = 1
                Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                If intI = 1 Then
                   If Val("" & RsTemp.Fields("add1")) > 0 Then
                       wks4630.Range(endX & xRows).Value = Val(wks4630.Range(endX & xRows).Value) + Val("" & RsTemp.Fields("add1"))
                   End If
                End If
            End If
            
            '差距
            If InStr("" & .Fields("tpb08"), "聖島") > 0 Then
               wks4630.Range(endX & xRows + 1).Value = wks4630.Range(endX & xRows - 1).Value - wks4630.Range(endX & xRows).Value
            End If
            
            strTemp2 = "" & .Fields("t_yy")
            xRows = xRows + 1
            .MoveNext
         Loop
      End With
      
      '---國內合計
      wks4630.Range(Chr(Asc(endX) + 1) & midR + 1).Value = "國內"
      For inJ = midR + 2 To xRows
          wks4630.Range(Chr(Asc(endX) + 1) & inJ).Value = "=SUM(" & startX & inJ & ":" & endX & inJ & ")"
      Next inJ
   End If

    '判斷若版本2007以上改變存檔格式
    If Val(xlsSalesPoint.Version) < 12 Then
    xlsSalesPoint.Workbooks(1).SaveAs FileName:=strPath, FileFormat:=-4143
    Else
    xlsSalesPoint.Workbooks(1).SaveAs FileName:=strPath, FileFormat:=56
    End If
    xlsSalesPoint.Workbooks.Close
    xlsSalesPoint.Quit
    'Modify by Amy 2021/06/22 原:strPath 改中文字顯示
    MsgBox "檔案已產生！" & vbCrLf & "檔案存於 " & strExcelPathN & " " & strTempFile, vbInformation

End Function

Private Sub Form_Load()
   MoveFormToCenter Me
   
   txtDate(0) = Left(strSrvDate(2), 5)
   txtDate(1) = Left(strSrvDate(2), 5)
   Label3.Caption = Label3 & strExcelPathN 'Add by Amy 2021/06/22
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm04060310 = Nothing
End Sub

Private Sub txtDate_GotFocus(Index As Integer)
   InverseTextBox txtDate(Index)
   CloseIme
End Sub

Private Sub txtDate_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub ChkTxtDate(Index As Integer, Cancel As Boolean)
Dim rsQuery As ADODB.Recordset
Dim stSQL As String, intR As Integer
Dim intCnt As Integer
   
   If txtDate(Index) <> "" Then
      If ChkDate(txtDate(Index) & "01") = False Then
          txtDate_GotFocus Index
          Cancel = True
          Exit Sub
      End If
      '檢查資料是否已存在
      stSQL = "select count(*) cnt from tpbulletin where TPB03 >=" & TransDate(Trim(txtDate(Index)) & "01", 2) & " and TPB03 <=" & GetLastDay(TransDate(Trim(txtDate(Index)) & "01", 2))
      intR = 1
      Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
      If intR = 1 Then
         intCnt = Val("" & rsQuery.Fields(0))
      End If
      rsQuery.Close
      
      If intCnt = 0 Then
         MsgBox txtDate(Index) & "此月份尚無公報資料!!"
         txtDate_GotFocus (Index)
         Cancel = True
         Exit Sub
      End If
   End If
   
   Set rsQuery = Nothing
End Sub
