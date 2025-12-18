VERSION 5.00
Begin VB.Form frm04060307 
   BorderStyle     =   1  '單線固定
   Caption         =   "專利公報市場排名"
   ClientHeight    =   2190
   ClientLeft      =   2790
   ClientTop       =   3945
   ClientWidth     =   4650
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   4650
   Begin VB.CheckBox Check1 
      Caption         =   "顯示占有率%"
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   1320
      Width           =   1575
   End
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
   Begin VB.Label Label3 
      Caption         =   "PS：Excel儲存於"
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   570
      TabIndex        =   5
      Top             =   1800
      Width           =   3195
   End
   Begin VB.Line Line1 
      X1              =   2220
      X2              =   2910
      Y1              =   1020
      Y2              =   1020
   End
   Begin VB.Label Label1 
      Caption         =   "公報年月："
      Height          =   210
      Left            =   690
      TabIndex        =   4
      Top             =   930
      Width           =   1125
   End
End
Attribute VB_Name = "frm04060307"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2021/11/17 Form2.0已檢查 (無需修改的物件)
'Create By Lydia 2017/01/05 專利公報市場排名
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
Dim m_rs As New ADODB.Recordset
Dim strVol1_S As String, strVol1_E As String '日期範圍
Dim m_Tot1 As Double, m_Tot2 As Double, m_Tot3 As Double '國內、大陸、國外總計(算占有率)
Dim m_TopNum As Integer '前X名
Dim strMid As String
Dim xlsSalesPoint As New Excel.Application
Dim wks4637 As New Worksheet
Dim xRows As Integer '目前列位置
Dim strTemp As String
Dim strPath As String, strTempFile As String
Dim iCall As Integer  '切換國內、大陸、國外
Dim inJ As Integer
Dim mESeqNo As String '暫存檔序號

   StrMenu = False

   'Added By Lydia 2022/01/13 查詢印表記錄檔欄位
   ClearQueryLog (Me.Name)
   pub_QL05 = pub_QL05 & ";" & Label1.Caption & strVol1_S & "-" & strVol1_E
   If Check1.Value = 1 Then
      pub_QL05 = pub_QL05 & ";顯示占有率%"
   End If
  'end 2022/01/13
  
   strVol1_S = TransDate(txtDate(0) & "01", 2)
   strVol1_E = GetLastDay(TransDate(txtDate(1) & "01", 2))
   m_TopNum = 10 '預設前10名
   
   strMid = "SELECT TPB01,TPB02,TPB03,TPB04,TPB05,decode(tpb06,'020','C0',substr(na02,1,2))||TPB06 TPB06,TPB07,TPB08,'' TPB09,PA01,PA02,PA03,PA04 " & _
            "From TPBulletin, nation, PATENT where na01(+)=tpb06 and pa11(+)=tpb01 and pa23(+)='1' and TPB03 >=" & strVol1_S & " and TPB03 <=" & strVol1_E
       
   '暫存檔- 無代理人且為本所案件,若最後發文的AB類程序為不出名則算台一案件
   strSql = "SELECT TPB01,TPB02,TPB03,TPB04,TPB05,TPB06,TPB07,CP01,CP02,CP03,CP04,substr(max(cp27||cp22),9) A1 FROM CASEPROGRESS, " & _
            "(" & strMid & " and tpb07 is null and pa01 is not null) WHERE PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) " & _
            " and cp09<'C' and cp27<TPB03 GROUP BY TPB01,TPB02,TPB03,TPB04,TPB05,TPB06,TPB07,CP01,CP02,CP03,CP04 "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      Set m_rs = PUB_CreateRecordset(RsTemp, , , , Me.Name, mESeqNo)
      '刪除不符合的資料
      strSql = "Delete from rdatafactory where id = '" & strUserNum & "' and formname = '" & Me.Name & "' and seqno = '" & mESeqNo & "' and nvl(r012,'A')='A' "
      cnnConnection.Execute strSql, intI
   End If
   
   '抓全部公告總件數
   If Check1.Value = 1 Then
      strSql = "select decode(substr(tpb06,1,1),'A','國內',decode(tpb06,'C0020','大陸','國外')) title,count(*) cnt from (" & _
               strMid & " ) group by decode(substr(tpb06,1,1),'A','國內',decode(tpb06,'C0020','大陸','國外')) " & _
               "order by decode(substr(tpb06,1,1),'A','國內',decode(tpb06,'C0020','大陸','國外')) "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         RsTemp.MoveFirst
         Do While Not RsTemp.EOF
            Select Case "" & RsTemp.Fields("title")
                Case "國內": m_Tot1 = Val("" & RsTemp.Fields("cnt"))
                Case "大陸": m_Tot2 = Val("" & RsTemp.Fields("cnt"))
                Case "國外": m_Tot3 = Val("" & RsTemp.Fields("cnt"))
            End Select
            RsTemp.MoveNext
         Loop
      End If
   End If

   '----------
   
   For iCall = 1 To 3
       If m_rs.State = 1 Then m_rs.Close
       m_rs.CursorLocation = adUseClient
       strSql = "select decode(substr(tpb06,1,1),'A','1',decode(tpb06,'C0020','2','3')) ord1," & _
                 "decode(substr(tpb06,1,1),'A','國內',decode(tpb06,'C0020','大陸','國外')) title,tpb08,count(*) cnt from (" & _
                 strMid & " ) where decode(substr(tpb06,1,1),'A','1',decode(tpb06,'C0020','2','3')) = " & CNULL(Format(iCall, "0")) & " and nvl(tpb08,' ') <> ' ' " & _
                 " group by decode(substr(tpb06,1,1),'A','1',decode(tpb06,'C0020','2','3'))," & _
                 "decode(substr(tpb06,1,1),'A','國內',decode(tpb06,'C0020','大陸','國外')),tpb08 " & _
                 "order by 4 desc,tpb08 "
       
       m_rs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
       If iCall = 1 And m_rs.RecordCount = 0 Then
          MsgBox "資料庫無資料！", vbExclamation
          InsertQueryLog (0) 'Added by Lydia 2022/01/13
          Exit Function
       End If
       
       If Not m_rs.EOF And Not m_rs.BOF Then
       '---------------開啟xls檔
          If iCall = 1 Then
              InsertQueryLog (m_rs.RecordCount) 'Added by Lydia 2022/01/13
              StrMenu = True
              startX = "b"
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
              Set wks4637 = xlsSalesPoint.Worksheets(1)
              wks4637.PageSetup.Orientation = xlPortrait '直印
              '抬頭
               wks4637.PageSetup.PrintTitleRows = "$1:$3"
               
               wks4637.Columns("a:a").ColumnWidth = 10 '項目
               For intI = Asc(startX) To Asc(startX) + m_TopNum - 1
                   wks4637.Columns(Chr(intI) & ":" & Chr(intI)).ColumnWidth = 13
                   wks4637.Columns(Chr(intI) & ":" & Chr(intI)).HorizontalAlignment = xlCenter
                   wks4637.Columns(Chr(intI) & ":" & Chr(intI)).VerticalAlignment = xlBottom
               Next
               strExc(0) = Mid(txtDate(0), 1, 3) & "/" & Mid(txtDate(0), 4, 2) & "至" & Mid(txtDate(1), 1, 3) & "/" & Mid(txtDate(1), 4, 2)
               wks4637.Range("a1").Value = strExc(0) & " " & Me.Caption
               With wks4637.Range("a1:" & Chr(Asc(startX) + m_TopNum - 1) & "1")
                 .WrapText = False
                 .MergeCells = True
                 .HorizontalAlignment = xlCenter
                 .VerticalAlignment = xlBottom
              End With
        
              xRows = 3
              '排名編號
              For intI = 1 To m_TopNum
                 wks4637.Range(Chr(Asc(startX) + intI - 1) & xRows).Value = intI
              Next
              xRows = xRows + 1
          End If
       '---------------
          
          With m_rs
              .MoveFirst
              inJ = 1
                '切換國內、大陸、國外的左邊標題
                If strTemp <> .Fields("title") Then
                   If strTemp <> "" Then
                      If Check1.Value = 1 Then
                         xRows = xRows + 3
                      Else
                         xRows = xRows + 2
                      End If
                   End If
                   wks4637.Range("a" & xRows).Value = "" & .Fields("title")
                   wks4637.Range("a" & xRows + 1).Value = "筆數"
                   If Check1.Value = 1 Then wks4637.Range("a" & xRows + 2).Value = "占有率"
                   strTemp = .Fields("title")
                End If
              '讀取前X名資料
              Do While inJ <= m_TopNum
                 '事務所名稱
                 strExc(1) = "" & .Fields("tpb08")
                 wks4637.Range(Chr(Asc(startX) + inJ - 1) & xRows).Value = "" & .Fields("tpb08")
                 '筆數
                 wks4637.Range(Chr(Asc(startX) + inJ - 1) & xRows + 1).Value = Val("" & .Fields("cnt"))
                 '無代理人且為本所案件,若最後發文的AB類程序為不出名則算台一案件
                 If "" & .Fields("tpb08") = "台一國際" Then
                     Select Case "" & .Fields("ord1")
                        Case "1": strExc(7) = "and substr(R006,1,1)='A' "
                        Case "2": strExc(7) = "and R006='C0020' "
                        Case "3": strExc(7) = "and substr(R006,1,1)<>'A' and R006<>'C0020' "
                     End Select
                     strSql = "SELECT COUNT(*) Add1 FROM rdatafactory " & _
                              "where id = '" & strUserNum & "' and formname = '" & Me.Name & "' and seqno = '" & mESeqNo & "' " & strExc(7)
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                     If intI = 1 Then
                        If Val("" & RsTemp.Fields("add1")) > 0 Then
                            wks4637.Range(Chr(Asc(startX) + inJ - 1) & xRows + 1).Value = Val(wks4637.Range(Chr(Asc(startX) + inJ - 1) & xRows + 1).Value) + Val("" & RsTemp.Fields("add1"))
                        End If
                     End If
                 End If
                 
                 '占有率
                 If Check1.Value = 1 Then
                    wks4637.Range(Chr(Asc(startX) + inJ - 1) & xRows + 2).Value = "=$" & Chr(Asc(startX) + inJ - 1) & xRows + 1 & " / " & IIf(iCall = 1, m_Tot1, IIf(iCall = 2, m_Tot2, m_Tot3))
                    wks4637.Range(Chr(Asc(startX) + inJ - 1) & xRows + 2).NumberFormatLocal = "##0.00%"
                 End If
                 inJ = inJ + 1
                 .MoveNext
                 If .EOF = True Then Exit Do
              Loop
          End With
       End If
   Next iCall
      
   '合計
   If Check1.Value = 1 Then
      xRows = xRows + 3
   Else
      xRows = xRows + 2
   End If
   wks4637.Range("a" & xRows).Value = "合計"
   wks4637.Range("a" & xRows + 1).Value = "筆數"
   If Check1.Value = 1 Then wks4637.Range("a" & xRows + 2).Value = "占有率"
   
   strSql = "select tpb08,count(*) cnt from (" & strMid & ") where nvl(tpb08,' ') <> ' ' group by tpb08 order by cnt desc,tpb08 "
   intI = 1
   Set m_rs = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      With m_rs
        .MoveFirst
        For inJ = 1 To m_TopNum
          '事務所名稱
          wks4637.Range(Chr(Asc(startX) + inJ - 1) & xRows).Value = "" & .Fields("tpb08")
          '筆數
          wks4637.Range(Chr(Asc(startX) + inJ - 1) & xRows + 1).Value = Val("" & .Fields("cnt"))
          '無代理人且為本所案件,若最後發文的AB類程序為不出名則算台一案件
          If "" & .Fields("tpb08") = "台一國際" Then
              strSql = "SELECT COUNT(*) Add1 FROM rdatafactory " & _
                       "where id = '" & strUserNum & "' and formname = '" & Me.Name & "' and seqno = '" & mESeqNo & "' "
              intI = 1
              Set RsTemp = ClsLawReadRstMsg(intI, strSql)
              If intI = 1 Then
                 If Val("" & RsTemp.Fields("add1")) > 0 Then
                     wks4637.Range(Chr(Asc(startX) + inJ - 1) & xRows + 1).Value = Val(wks4637.Range(Chr(Asc(startX) + inJ - 1) & xRows + 1).Value) + Val("" & RsTemp.Fields("add1"))
                 End If
              End If
          End If
          '占有率
          If Check1.Value = 1 Then
             wks4637.Range(Chr(Asc(startX) + inJ - 1) & xRows + 2).Value = "=$" & Chr(Asc(startX) + inJ - 1) & xRows + 1 & " / " & m_Tot1 + m_Tot2 + m_Tot3
             wks4637.Range(Chr(Asc(startX) + inJ - 1) & xRows + 2).NumberFormatLocal = "##0.00%"
          End If
          .MoveNext
        Next inJ
      End With
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
   Set frm04060307 = Nothing
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
