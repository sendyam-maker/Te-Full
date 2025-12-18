VERSION 5.00
Begin VB.Form frm04060311 
   BorderStyle     =   1  '單線固定
   Caption         =   "專利公報國外同業排名"
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
Attribute VB_Name = "frm04060311"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2021/11/17 Form2.0已檢查 (無需修改的物件)
'Create By Lydia 2017/01/05 專利公報國外同業排名
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
Dim strMid As String, strMid2 As String, strMid3 As String
Dim xlsSalesPoint As New Excel.Application
Dim wks4631 As New Worksheet
Dim xRows As Integer '目前列位置
Dim midR As Integer '台一與聖島比較的抬頭位置
Dim midR2 As Integer 'FCP與聖島比較的抬頭位置

Dim strTemp As String
Dim strPath As String, strTempFile As String
Dim iCall As Integer  '切換大陸,美,日和各洲
Dim inJ As Integer
Dim mArea As Integer '大陸,美,日和各洲計數
Dim mCall As Integer '切換今年度和前一年
Dim tmpTot As Double '暫存洲別的加總
Dim mESeqNo As String '暫存檔序號
Dim strAt As String, strAd As String
Dim passAt As String '已出現過的洲別
Dim mQty As Integer '是否計件
Dim tmpArr As Variant

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
  
   '國籍和洲別的行位置
    strAt = "美國,日本,亞洲,美洲,歐洲,大洋洲,非洲,小計"
    strAd = "美國c,日本d,C0e,C1f,C2g,C4h,C3i,T0j"
    tmpArr = Split(strAt, ",")
    
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
   
   mArea = 8 '大陸,美,日和各洲(亞洲、美洲、歐洲、大洋洲、非洲)
   
   '排名資料
   '-------------切換今年度和前一年
   For mCall = 1 To 2
       For iCall = 1 To mArea
           If m_rs.State = 1 Then m_rs.Close
           m_rs.CursorLocation = adUseClient
           If mCall = 1 Then '今年度
              strMid2 = " and TPB03 >=" & CNULL(strVol1_S) & " and TPB03 <=" & CNULL(strVol1_E)
           Else              '前一年
              strMid2 = " and TPB03 >=" & CNULL(strVol2_S) & " and TPB03 <=" & CNULL(strVol2_E)
           End If
            Select Case iCall
                '大陸
                Case 1
                    strMid3 = " and tpb06 ='C0020' "
                    strExc(3) = "'大陸'"
                '美國
                Case 2
                    strMid3 = " and tpb06 ='C1101' "
                    strExc(3) = "'美國'"
                '日本
                Case 3
                    strMid3 = " and tpb06 ='C0011' "
                    strExc(3) = "'日本'"
                '各洲
                Case Else:
                   '2017/1/12 不含台灣,大陸
                   'If iCall = 4 Then '亞洲含台灣,大陸
                   '   strMid3 = "and (substr(tpb06,1,2) ='C0' or substr(tpb06,1,1) ='A') "
                   'Else
                   If iCall >= 4 And iCall <= 6 Then
                      strExc(4) = "and substr(tpb06,1,2) ='C" & iCall - 4 & "' "
                   Else
                      If iCall = 7 Then '大洋洲排在非洲前面
                         strExc(4) = "and substr(tpb06,1,2) ='C4' "
                      Else
                         strExc(4) = "and substr(tpb06,1,2) ='C3' "
                      End If
                   End If
                   strMid3 = strExc(4) & "and substr(tpb06,1,1)<>'A' and tpb06<>'C0020' "
                   'End If
                   strExc(3) = "decode(substr(tpb06,1,1),'A','0亞洲',decode(substr(tpb06,1,2),'C0','0亞洲','C1','1美洲','C2','2歐洲','C3','4非洲','C4','3大洋洲',tpb06))"
            End Select
            strSql = "select " & strExc(3) & " title,tpb08,count(*) cnt from nation,(" & strMid & strMid2 & _
                     ") where substr(tpb06,3)=na01(+) and nvl(tpb08,' ') <> ' ' " & strMid3 & _
                      "group by " & strExc(3) & ",tpb08 order by 1, 3 desc"

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
                  Set wks4631 = xlsSalesPoint.Worksheets(1)
                  wks4631.PageSetup.Orientation = xlPortrait '直印
                  '抬頭
                   wks4631.PageSetup.PrintTitleRows = "$1:$3"
                   
                   'Modified by Lydia 2017/02/02 項目從10放大到12
                   wks4631.Columns("a:a").ColumnWidth = 12 '項目
                   For intI = Asc(startX) To Asc(startX) + m_TopNum '多年度一欄
                       wks4631.Columns(Chr(intI) & ":" & Chr(intI)).ColumnWidth = 13
                       wks4631.Columns(Chr(intI) & ":" & Chr(intI)).HorizontalAlignment = xlCenter
                       wks4631.Columns(Chr(intI) & ":" & Chr(intI)).VerticalAlignment = xlBottom
                   Next
                   
                   strExc(0) = Mid(txtDate(0), 1, 3) & "/" & Mid(txtDate(0), 4, 2) & "-" & Mid(txtDate(1), 1, 3) & "/" & Mid(txtDate(1), 4, 2)
                   strExc(0) = strExc(0) & "和" & Mid(Val(txtDate(0)) - 100, 1, 3) & "/" & Mid(Val(txtDate(0)) - 100, 4, 2) & "-" & Mid(Val(txtDate(1)) - 100, 1, 3) & "/" & Mid(Val(txtDate(1)) - 100, 4, 2)
                   wks4631.Range("a1").Value = strExc(0) & " " & Me.Caption
                   With wks4631.Range("a1:" & Chr(Asc(startX) + m_TopNum - 1) & "1")
                     .WrapText = False
                     .MergeCells = True
                     .HorizontalAlignment = xlCenter
                     .VerticalAlignment = xlBottom
                   End With
                   
                   'Modified by Lydia 2017/02/02 +亞洲地區
                   wks4631.Range("a2").Value = "(亞洲地區不含台灣、大陸)"
                   With wks4631.Range("a2:" & Chr(Asc(startX) + m_TopNum - 1) & "2")
                     .WrapText = False
                     .MergeCells = True
                     .HorizontalAlignment = xlCenter
                     .VerticalAlignment = xlBottom
                   End With
                   
                  xRows = 3
                  '排名編號
                  For intI = 1 To m_TopNum
                     wks4631.Range(Chr(Asc(startX) + intI - 1) & xRows).Value = intI
                  Next
                  xRows = xRows + 1
                  
              '重設前一年度的起始列
              ElseIf mCall = 2 And iCall = 1 Then
                  xRows = 7
                  strTemp = ""
              End If
           '---------------
              
              With m_rs
                  .MoveFirst
                  inJ = 1
                    '切換左邊標題(a=區域,b=年度)
                    If strTemp <> .Fields("title") Then
                       If strTemp <> "" Then xRows = xRows + 6 '跳過今年的資料列
                       '各洲抬頭
                       If mCall = 1 Or (mCall = 2 And InStr(passAt, "" & .Fields("title")) = 0) Then
                          If mCall = 2 And InStr(passAt, "" & .Fields("title")) = 0 Then
                             xRows = xRows - 3
                          End If
                          wks4631.Range("a" & xRows).Value = "" & IIf(InStr("0,1,2,3,4,5,6,7,8,9", Mid(.Fields("title"), 1, 1)) > 0, Mid(.Fields("title"), 2), .Fields("title"))
                          passAt = passAt & .Fields("title")
                          If iCall = 4 Then
                             strExc(9) = wks4631.Range("a" & xRows).Value
                             wks4631.Range("a" & xRows & ":a" & xRows + 2).MergeCells = True
                             wks4631.Range("a" & xRows).Value = strExc(9) & vbCrLf & "(不含臺灣、" & vbCrLf & "大陸)"
                          End If
                       End If
                       If wks4631.Range("b" & xRows).Value = "" Then
                          wks4631.Range("b" & xRows).Value = IIf(mCall = 1, Mid(txtDate(0), 1, 3) & "年", Val(Mid(txtDate(0), 1, 3)) - 1 & "年")  '年度
                          wks4631.Range("b" & xRows + 1).Value = "筆數"
                          wks4631.Range("b" & xRows + 2).Value = "占有率"
                       End If
                       strTemp = .Fields("title")
                       '各洲總計
                        strSql = "select count(*) cnt from (" & strMid & IIf(mCall = 1, " and TPB03 >=" & strVol1_S & " and TPB03 <=" & strVol1_E, " and TPB03 >=" & strVol2_S & " and TPB03 <=" & strVol2_E) & " ) " & _
                                 "where 1=1 " & strMid3
                        intI = 1
                        Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                        If intI = 1 Then
                           tmpTot = Val("" & RsTemp(0))
                        Else
                           tmpTot = 0
                        End If
                    End If
                  '讀取前X名資料
                  Do While inJ <= m_TopNum
                     '事務所名稱
                     wks4631.Range(Chr(Asc(startX) + inJ - 1) & xRows).Value = "" & .Fields("tpb08")
                     '筆數
                     wks4631.Range(Chr(Asc(startX) + inJ - 1) & xRows + 1).Value = Val("" & .Fields("cnt"))
                     '無代理人且為本所案件,若最後發文的AB類程序為不出名則算台一案件
                     If "" & .Fields("tpb08") = "台一國際" Then
                         strExc(6) = Replace(UCase(strMid2 & " " & strMid3), "TPB", "R0")
                         strSql = "SELECT COUNT(*) Add1 FROM rdatafactory " & _
                                  "where id = '" & strUserNum & "' and formname = '" & Me.Name & "' and seqno = '" & mESeqNo & "' " & strExc(6)
                         intI = 1
                         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                         If intI = 1 Then
                            If Val("" & RsTemp.Fields("add1")) > 0 Then
                                wks4631.Range(Chr(Asc(startX) + inJ - 1) & xRows + 1).Value = Val(wks4631.Range(Chr(Asc(startX) + inJ - 1) & xRows + 1).Value) + Val("" & RsTemp.Fields("add1"))
                            End If
                         End If
                     End If
                     
                     '占有率
                     wks4631.Range(Chr(Asc(startX) + inJ - 1) & xRows + 2).Value = "=$" & Chr(Asc(startX) + inJ - 1) & xRows + 1 & " / " & tmpTot
                     wks4631.Range(Chr(Asc(startX) + inJ - 1) & xRows + 2).NumberFormatLocal = "##0.00%"
                     inJ = inJ + 1
                     .MoveNext
                     If .EOF = True Then Exit Do
                  Loop
              End With
           End If
       Next iCall
       
   Next mCall
   '----------------切換今年度和前一年
   
   '國外合計(不含台灣,大陸)
   xRows = xRows + 3
   strTemp = ""
   For mCall = 1 To 2
        If m_rs.State = 1 Then m_rs.Close
        m_rs.CursorLocation = adUseClient
        If mCall = 1 Then '今年度
           strMid2 = " and TPB03 >=" & CNULL(strVol1_S) & " and TPB03 <=" & CNULL(strVol1_E)
           wks4631.Range("a" & xRows & ":a" & xRows + 2).MergeCells = True
           wks4631.Range("a" & xRows).Value = "合計" & vbCrLf & "(不含臺灣、" & vbCrLf & "大陸)"
        Else              '前一年
           strMid2 = " and TPB03 >=" & CNULL(strVol2_S) & " and TPB03 <=" & CNULL(strVol2_E)
           xRows = xRows + 3
        End If
        If wks4631.Range("b" & xRows).Value = "" Then
           wks4631.Range("b" & xRows).Value = IIf(mCall = 1, Mid(txtDate(0), 1, 3) & "年", Val(Mid(txtDate(0), 1, 3)) - 1 & "年")  '年度
           wks4631.Range("b" & xRows + 1).Value = "筆數"
           wks4631.Range("b" & xRows + 2).Value = "占有率"
        End If
        strMid3 = " and substr(tpb06,1,1)<>'A' and tpb06<>'C0020' "
        
       '各洲總計
        strSql = "select count(*) cnt from (" & strMid & strMid2 & strMid3 & " ) " & _
                 "where 1=1 " & strMid3
        intI = 1
        Set RsTemp = ClsLawReadRstMsg(intI, strSql)
        If intI = 1 Then
           tmpTot = Val("" & RsTemp(0))
        Else
           tmpTot = 0
        End If
        strSql = "select '合計' title,tpb08,count(*) cnt from nation,(" & strMid & strMid2 & _
            ") where substr(tpb06,3)=na01(+) and nvl(tpb08,' ')<>' ' " & strMid3 & _
             "group by tpb08 order by 3 desc"
        m_rs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If Not m_rs.EOF And Not m_rs.BOF Then
           With m_rs
               .MoveFirst
               inJ = 1
               '讀取前X名資料
               Do While inJ <= m_TopNum
                  '事務所名稱
                  wks4631.Range(Chr(Asc(startX) + inJ - 1) & xRows).Value = "" & .Fields("tpb08")
                  '筆數
                  wks4631.Range(Chr(Asc(startX) + inJ - 1) & xRows + 1).Value = Val("" & .Fields("cnt"))
                  '無代理人且為本所案件,若最後發文的AB類程序為不出名則算台一案件
                  If "" & .Fields("tpb08") = "台一國際" Then
                      strExc(6) = Replace(UCase(strMid2 & " " & strMid3), "TPB", "R0")
                      strSql = "SELECT COUNT(*) Add1 FROM rdatafactory " & _
                               "where id = '" & strUserNum & "' and formname = '" & Me.Name & "' and seqno = '" & mESeqNo & "' " & strExc(6)
                      intI = 1
                      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                      If intI = 1 Then
                         If Val("" & RsTemp.Fields("add1")) > 0 Then
                             wks4631.Range(Chr(Asc(startX) + inJ - 1) & xRows + 1).Value = Val(wks4631.Range(Chr(Asc(startX) + inJ - 1) & xRows + 1).Value) + Val("" & RsTemp.Fields("add1"))
                         End If
                      End If
                  End If
                    
                  '占有率
                  wks4631.Range(Chr(Asc(startX) + inJ - 1) & xRows + 2).Value = "=$" & Chr(Asc(startX) + inJ - 1) & xRows + 1 & " / " & tmpTot
                  wks4631.Range(Chr(Asc(startX) + inJ - 1) & xRows + 2).NumberFormatLocal = "##0.00%"
                  inJ = inJ + 1
                  .MoveNext
                  If .EOF = True Then Exit Do
               Loop
           End With
        End If
   Next mCall
   
   midR = xRows + 4
   
   '抓台一和聖島比較
   strExc(0) = Mid(txtDate(0), 1, 3) & "/" & Mid(txtDate(0), 4, 2) & "-" & Mid(txtDate(1), 1, 3) & "/" & Mid(txtDate(1), 4, 2)
   strExc(0) = strExc(0) & "和" & Mid(Val(txtDate(0)) - 100, 1, 3) & "/" & Mid(Val(txtDate(0)) - 100, 4, 2) & "-" & Mid(Val(txtDate(1)) - 100, 1, 3) & "/" & Mid(Val(txtDate(1)) - 100, 4, 2)
   wks4631.Range("a" & midR).Value = strExc(0) & " " & "國外各洲與聖島比較"

   With wks4631.Range("a" & midR & ":j" & midR)
     .WrapText = False
     .MergeCells = True
     .HorizontalAlignment = xlCenter
     .VerticalAlignment = xlBottom
   End With
   
   xRows = midR + 1
   strTemp = ""
   
   '設定區域抬頭
   For inJ = 0 To UBound(tmpArr)
      If Trim(tmpArr(inJ)) <> "" Then
          wks4631.Range(Chr(Asc(startX) + inJ) & xRows).Value = Trim(tmpArr(inJ))
          endX = Chr(Asc(startX) + inJ)
      End If
   Next inJ
   xRows = xRows + 1
   '左邊:年度,事務所,差距
    wks4631.Range("a" & xRows).Value = Mid(txtDate(0), 1, 3) & "年"
    wks4631.Range("b" & xRows).Value = "台一國際"
    wks4631.Range("b" & xRows + 1).Value = "聖島國際"
    wks4631.Range("b" & xRows + 2).Value = "差距"
    wks4631.Range("a" & xRows + 3).Value = Val(Mid(txtDate(0), 1, 3)) - 1 & "年"
    wks4631.Range("b" & xRows + 3).Value = "台一國際"
    wks4631.Range("b" & xRows + 4).Value = "聖島國際"
    wks4631.Range("b" & xRows + 5).Value = "差距"

   '----------------切換今年度和前一年
   For mCall = 1 To 2
      If mCall = 1 Then '今年度
         strMid2 = " and TPB03 >=" & CNULL(strVol1_S) & " and TPB03 <=" & CNULL(strVol1_E)
      Else              '前一年
         strMid2 = " and TPB03 >=" & CNULL(strVol2_S) & " and TPB03 <=" & CNULL(strVol2_E)
      End If
      '美國
      strSql = "select '1' ord1, '" & IIf(mCall = 1, Trim(Mid(txtDate(0), 1, 3)), Trim(Val(Mid(txtDate(0), 1, 3))) - 1) & "' t_yy," & _
               " 'C1101' ano,substr(na03,1,2) title,tpb08,count(*) cnt from nation,(" & _
               strMid & strMid2 & ") where tpb06='C1101' and tpb08 in ('聖島國際','台一國際') and substr(tpb06,3)=na01(+) group by substr(na03,1,2) ,tpb08"
      '日本
      strSql = strSql & " Union all select '2' ord1, '" & IIf(mCall = 1, Trim(Mid(txtDate(0), 1, 3)), Trim(Val(Mid(txtDate(0), 1, 3))) - 1) & "' t_yy," & _
               " 'C0011' ano,substr(na03,1,2) title,tpb08,count(*) cnt from nation,(" & _
               strMid & strMid2 & ") where tpb06='C0011' and tpb08 in ('聖島國際','台一國際') and substr(tpb06,3)=na01(+) group by substr(na03,1,2) ,tpb08"
      '各洲
      strSql = strSql & " Union all select '3' ord1, '" & IIf(mCall = 1, Trim(Mid(txtDate(0), 1, 3)), Trim(Val(Mid(txtDate(0), 1, 3))) - 1) & "' t_yy," & _
               " substr(tpb06,1,2) ano,decode(substr(tpb06,1,1),'A','0亞洲',decode(substr(tpb06,1,2),'C0','0亞洲','C1','1美洲','C2','2歐洲','C3','4非洲','C4','3大洋洲',tpb06)) title,tpb08,count(*) cnt from nation,(" & _
               strMid & strMid2 & ") where substr(tpb06,1,1)<>'A' and tpb06<>'C0020' and tpb08 in ('聖島國際','台一國際') and substr(tpb06,3)=na01(+) group by substr(tpb06,1,2),decode(substr(tpb06,1,1),'A','0亞洲',decode(substr(tpb06,1,2),'C0','0亞洲','C1','1美洲','C2','2歐洲','C3','4非洲','C4','3大洋洲',tpb06)),tpb08"
      
      strSql = strSql & " order by ord1,title,tpb08 "

      If m_rs.State = 1 Then m_rs.Close
      m_rs.CursorLocation = adUseClient
      m_rs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If Not m_rs.EOF And Not m_rs.BOF Then
         With m_rs
            .MoveFirst
            strTemp = ""
            Do While Not .EOF
               '切換列
               If Mid(txtDate(0), 1, 3) = "" & .Fields("t_yy") Then
                  xRows = midR + 2
               Else
                  xRows = midR + 5
               End If
               If InStr("" & .Fields("tpb08"), "聖島") > 0 Then
                  xRows = xRows + 1
               End If
               
               If "" & .Fields("ano") = "C1101" Then '美國
                   strExc(0) = startX
               ElseIf "" & .Fields("ano") = "C0011" Then '日本
                   strExc(0) = Chr(Asc(startX) + 1)
               Else                                '各洲
                   strExc(0) = Mid(strAd, InStr(strAd, "" & .Fields("ano")) + Len("" & .Fields("ano")), 1)
               End If
               
               If InStr("" & .Fields("tpb08"), "聖島") = 0 Then
                  '筆數
                  wks4631.Range(strExc(0) & xRows).Value = "" & .Fields("cnt")
                  '無代理人且為本所案件,若最後發文的AB類程序為不出名則算台一案件
                  If "" & .Fields("tpb08") = "台一國際" Then
                      Select Case "" & .Fields("ord1")
                         Case "1": strMid3 = " and R006='C1101' "
                         Case "2": strMid3 = " and R006='C0011' "
                         Case "3": strMid3 = " and substr(R006,1,1)<>'A' and R006<>'C0020' and substr(R006,1,2)='" & .Fields("ano") & "' "
                      End Select
                      
                      strSql = "SELECT COUNT(*) Add1 FROM rdatafactory " & _
                               "where id = '" & strUserNum & "' and formname = '" & Me.Name & "' and seqno = '" & mESeqNo & "' " & _
                               Replace(UCase(strMid2 & strMid3), "TPB", "R0")
                      intI = 1
                      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                      If intI = 1 Then
                         If Val("" & RsTemp.Fields("add1")) > 0 Then
                            wks4631.Range(strExc(0) & xRows).Value = Val(wks4631.Range(strExc(0) & xRows).Value) + Val("" & RsTemp.Fields("add1"))
                         End If
                      End If
                  End If
               Else
                 '筆數
                 wks4631.Range(strExc(0) & xRows).Value = "" & .Fields("cnt")
                 '差距
                 wks4631.Range(strExc(0) & xRows + 1).Value = wks4631.Range(strExc(0) & xRows - 1).Value - wks4631.Range(strExc(0) & xRows).Value
               End If
               .MoveNext
            Loop
         End With
      End If
   Next mCall
   
   For inJ = 2 To 7
       wks4631.Range(endX & midR + inJ).Value = "=SUM(e" & midR + inJ & ":" & Chr(Asc(endX) - 1) & midR + inJ & ")"
       For intI = Asc(startX) To Asc(endX)
           If Val(wks4631.Range(Chr(intI) & midR + inJ).Value) = 0 Then
              wks4631.Range(Chr(intI) & midR + inJ).Value = "0"
           End If
       Next intI
   Next inJ
     
   '抓FCP與聖島的差異
   midR2 = midR + 7 + 4
   strExc(0) = Mid(txtDate(0), 1, 3) & "/" & Mid(txtDate(0), 4, 2) & "-" & Mid(txtDate(1), 1, 3) & "/" & Mid(txtDate(1), 4, 2)
   strExc(0) = strExc(0) & "和" & Mid(Val(txtDate(0)) - 100, 1, 3) & "/" & Mid(Val(txtDate(0)) - 100, 4, 2) & "-" & Mid(Val(txtDate(1)) - 100, 1, 3) & "/" & Mid(Val(txtDate(1)) - 100, 4, 2)
   wks4631.Range("a" & midR2).Value = strExc(0) & " " & "國外各洲與聖島比較"

   With wks4631.Range("a" & midR2 & ":j" & midR2)
     .WrapText = False
     .MergeCells = True
     .HorizontalAlignment = xlCenter
     .VerticalAlignment = xlBottom
   End With
   
   wks4631.Range("a" & midR2 + 1).Value = "(不含無新申請案進度案件)"
   wks4631.Range("a" & midR2 + 1).Font.Color = &HFF&
   With wks4631.Range("a" & midR2 + 1 & ":j" & midR2 + 1)
     .WrapText = False
     .MergeCells = True
     .HorizontalAlignment = xlCenter
     .VerticalAlignment = xlBottom
   End With
   
   xRows = midR2 + 2
   '設定區域抬頭
   For inJ = 0 To UBound(tmpArr)
      If Trim(tmpArr(inJ)) <> "" Then
          wks4631.Range(Chr(Asc(startX) + inJ) & xRows).Value = Trim(tmpArr(inJ))
      End If
   Next inJ

   '----------------Start of FCP
   For mCall = 1 To 2
      If mCall = 1 Then '今年度
         strMid2 = " and TPB03 >=" & strVol1_S & " and TPB03 <=" & strVol1_E
         xRows = xRows + 1
         wks4631.Range("a" & xRows).Value = Mid(txtDate(0), 1, 3) & "年"
      Else              '前一年
         strMid2 = " and TPB03 >=" & strVol2_S & " and TPB03 <=" & strVol2_E
         wks4631.Range(endX & xRows).Value = "=SUM(e" & xRows & ":" & Chr(Asc(endX) - 1) & xRows & " )"  '今年度的小計
         xRows = xRows + 3
        wks4631.Range("a" & xRows).Value = Trim(Val(Mid(txtDate(0), 1, 3)) - 1) & "年"
      End If
      wks4631.Range("b" & xRows).Value = "FCP"
      wks4631.Range("b" & xRows + 1).Value = "聖島國際"
      wks4631.Range("b" & xRows + 2).Value = "差距"
         
      strSql = "select tpb03,tpb06,tpb08,pa01,pa02,pa03,pa04 from (SELECT TPB01,TPB02,TPB03,TPB04,TPB05,decode(tpb06,'020','C0',substr(na02,1,2))||TPB06 TPB06,TPB07,TPB08,'' TPB09,PA01,PA02,PA03,PA04" & _
            " From TPBulletin, nation, PATENT where na01(+)=tpb06 and pa11(+)=tpb01 and pa23(+)='1' " & _
            " and (tpb08='台一國際' or (tpb07 is null and pa01 is not null)) " & strMid2 & _
            ") where substr(tpb06,1,1) <> 'A' and tpb06 <> 'C0020' "

      If m_rs.State = 1 Then m_rs.Close
      m_rs.CursorLocation = adUseClient
      m_rs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If Not m_rs.EOF And Not m_rs.BOF Then
         With m_rs
            .MoveFirst
            Do While Not .EOF
                '無代理人且為本所案件,若最後發文的AB類程序為不出名則算台一案件
                If Trim("" & .Fields("tpb08")) = "" Then
                   mQty = 0
                   strSql = "select substr(max(cp27||cp22),9) from caseprogress where cp01='" & .Fields("pa01") & "' and cp02='" & .Fields("pa02") & "' and cp03='" & .Fields("pa03") & "' and cp04='" & .Fields("pa04") & "' and cp09<'C' and cp27<" & .Fields("tpb03")
                   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                   If intI = 1 Then
                      If "" & RsTemp(0) = "N" Then
                         mQty = 1
                      End If
                   End If
                Else '出名
                   mQty = 1
                End If
                If mQty = 0 Then GoTo JumpNextRec2
                
                If "" & .Fields("tpb06") = "C1101" Then '美國
                    strExc(0) = startX
                ElseIf "" & .Fields("tpb06") = "C0011" Then '日本
                    strExc(0) = Chr(Asc(startX) + 1)
                Else  '各洲
                    strExc(0) = Mid(strAd, InStr(strAd, Mid("" & .Fields("tpb06"), 1, 2)) + 2, 1)
                End If
                  
                strSql = "SELECT DECODE(SUBSTR(CP12,1,1),'S','1'||CP12,'F','2'||CP12,'3其他') DEPTNO, 1 AS A1 FROM CASEPROGRESS " & _
                         "WHERE CP01='" & .Fields("PA01") & "' AND CP02='" & .Fields("PA02") & "' AND CP03='" & .Fields("PA03") & "' AND CP04='" & .Fields("PA04") & "' " & _
                         "AND CP10 IN ('101','102','103','104','105','109','110','112','113','114','115','118','120','122','125','307') " & _
                         "GROUP BY DECODE(SUBSTR(CP12,1,1),'S','1'||CP12,'F','2'||CP12,'3其他') "
                intI = 1
                Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                If intI = 1 Then
                   If Mid("" & RsTemp.Fields("deptno"), 2, 1) = "F" Then
                      wks4631.Range(strExc(0) & xRows).Value = Val(wks4631.Range(strExc(0) & xRows).Value) + mQty
                      If "" & .Fields("tpb06") = "C1101" Then '美國->美洲
                          wks4631.Range(Chr(Asc(strExc(0)) + 3) & xRows).Value = Val(wks4631.Range(Chr(Asc(strExc(0)) + 3) & xRows).Value) + mQty
                      ElseIf "" & .Fields("tpb06") = "C0011" Then '日本->亞洲
                          wks4631.Range(Chr(Asc(strExc(0)) + 1) & xRows).Value = Val(wks4631.Range(Chr(Asc(strExc(0)) + 1) & xRows).Value) + mQty
                      End If
                   End If
                End If
JumpNextRec2:
                .MoveNext
            Loop
         End With
      End If
   Next mCall
   
    wks4631.Range(endX & xRows).Value = "=SUM(e" & xRows & ":" & Chr(Asc(endX) - 1) & xRows & " )"  '前一年的小計
   
   '設定聖島和FCP的差距
   For inJ = Asc(startX) To Asc(endX)
      '今年度-聖島
      wks4631.Range(Chr(inJ) & midR2 + 4).Value = wks4631.Range(Chr(inJ) & midR + 3).Value
      '差距
      wks4631.Range(Chr(inJ) & midR2 + 5).Value = Val(wks4631.Range(Chr(inJ) & midR2 + 3).Value) - Val(wks4631.Range(Chr(inJ) & midR2 + 4).Value)
      '前一年-聖島
      wks4631.Range(Chr(inJ) & midR2 + 7).Value = wks4631.Range(Chr(inJ) & midR + 6).Value
      '差距
      wks4631.Range(Chr(inJ) & midR2 + 8).Value = Val(wks4631.Range(Chr(inJ) & midR2 + 6).Value) - Val(wks4631.Range(Chr(inJ) & midR2 + 7).Value)
      For intI = 3 To 8
         If Val(wks4631.Range(Chr(inJ) & midR2 + intI).Value) = 0 Then
            wks4631.Range(Chr(inJ) & midR2 + intI).Value = "0"
         End If
      Next intI
   Next inJ
   '--------------End of FCP
   
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
   Set frm04060311 = Nothing
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
