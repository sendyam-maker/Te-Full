VERSION 5.00
Begin VB.Form frm030621 
   BorderStyle     =   1  '單線固定
   Caption         =   "申請人國籍及洲別統計(含同業)"
   ClientHeight    =   3675
   ClientLeft      =   2790
   ClientTop       =   3945
   ClientWidth     =   4650
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   4650
   Begin VB.TextBox Text2 
      Height          =   345
      Left            =   2760
      MaxLength       =   5
      TabIndex        =   1
      Top             =   720
      Width           =   795
   End
   Begin VB.TextBox txt1 
      Height          =   345
      Index           =   1
      Left            =   1680
      TabIndex        =   3
      Top             =   2960
      Width           =   2295
   End
   Begin VB.TextBox txt1 
      Height          =   345
      Index           =   0
      Left            =   1680
      TabIndex        =   2
      Top             =   2520
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Left            =   1800
      MaxLength       =   5
      TabIndex        =   0
      Top             =   720
      Width           =   795
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   3750
      TabIndex        =   5
      Top             =   90
      Width           =   756
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Excel(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2700
      TabIndex        =   4
      Top             =   90
      Width           =   990
   End
   Begin VB.Label Label3 
      Caption         =   "PS：Excel儲存於"
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   780
      TabIndex        =   14
      Top             =   3360
      Width           =   3195
   End
   Begin VB.Line Line1 
      X1              =   2400
      X2              =   3090
      Y1              =   880
      Y2              =   880
   End
   Begin VB.Label Label4 
      Caption         =   "6."
      Height          =   255
      Index           =   6
      Left            =   1440
      TabIndex        =   6
      Top             =   3000
      Width           =   795
   End
   Begin VB.Label Label4 
      Caption         =   "5."
      Height          =   255
      Index           =   5
      Left            =   1440
      TabIndex        =   13
      Top             =   2565
      Width           =   795
   End
   Begin VB.Label Label4 
      Caption         =   "4.台灣國際"
      Height          =   255
      Index           =   4
      Left            =   1440
      TabIndex        =   12
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "3.理律法律"
      Height          =   255
      Index           =   3
      Left            =   1440
      TabIndex        =   11
      Top             =   1995
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "2.聖島國際"
      Height          =   255
      Index           =   2
      Left            =   1440
      TabIndex        =   10
      Top             =   1725
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "1.台一國際"
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   9
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "事務所或代理人名稱："
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   8
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "公報年月："
      Height          =   210
      Left            =   930
      TabIndex        =   7
      Top             =   810
      Width           =   1005
   End
End
Attribute VB_Name = "frm030621"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create By Lydia 2015/12/08 申請人國籍及洲別統計(含同業)
Option Explicit

Private Sub cmdOK_Click(Index As Integer)
Dim Cancel As Boolean
   
   Select Case Index
      Case 0
         If Trim(Text1) = "" Then
            MsgBox "起始公報年月不可空白！", vbInformation, "輸入錯誤！"
            Text1.SetFocus
            Exit Sub
         End If
         If Trim(Text2) = "" Then
            MsgBox "截止公報年月不可空白！", vbInformation, "輸入錯誤！"
            Text2.SetFocus
            Exit Sub
         End If
         Text1_Validate Cancel
         If Cancel = True Then
            Text1.SetFocus
            Exit Sub
         End If
         Text2_Validate Cancel
         If Cancel = True Then
            Text2.SetFocus
            Exit Sub
         End If
         If Val(Text2) < Val(Text1) Then
            MsgBox "截止年月必須大於起始年月！", vbInformation, "輸入錯誤！"
            Text2.SetFocus
            Exit Sub
         End If
         For intI = 0 To 1
             txt1_Validate intI, Cancel
             If Cancel = True Then Exit Sub
         Next
         
         StrMenu
      Case 1
         Unload Me
   End Select
End Sub

Private Function StrMenu() As Boolean
Dim m_rs As New ADODB.Recordset
Dim i As Integer
Dim strVol1_S As String, strVol2_S As String
Dim strVol1_E As String, strVol2_E As String
Dim xlsSalesPoint As New Excel.Application
Dim wks621 As New Worksheet
Dim xRows As Integer '目前列位置
Dim strR As Integer
Dim strAd As String
Dim strTemp As String
Dim strPath As String, strTempFile As String

   StrMenu = False
   
   '報表樣式:
   '申請人國籍           FCT件數    FCT類別數  T件數      T類別數
   '-------------------- ---------- ---------- ---------- ----------
   '日本                         17         40          1          1
   '韓國                          1          1          1          1
   '香港                          0          0         10         10
   '.
   '.
   '.
   Call Pub_ChgDateToTMBM07(Text1, strVol1_S, strVol2_S)
   Call Pub_ChgDateToTMBM07(Text2, strVol1_E, strVol2_E)
   
   'Added By Lydia 2022/01/13 查詢印表記錄檔欄位
   ClearQueryLog (Me.Name)
   pub_QL05 = pub_QL05 & ";" & Label1.Caption & Text1 & "-" & Text2
   pub_QL05 = pub_QL05 & ";" & Label4(0).Caption & Label4(1).Caption & "、" & Label4(2).Caption & "、" & Label4(3).Caption & "、" & Label4(4).Caption
   If txt1(0) <> "" Then pub_QL05 = pub_QL05 & "、5." & txt1(0)
   If txt1(1) <> "" Then pub_QL05 = pub_QL05 & "、6." & txt1(1)
  'end 2022/01/13
  
   '類別依公報為準
   strSql = "select decode(substr(na01,1,1),'A','台灣','B','大陸',TMBM05) V1," & _
            "sum(nvl(FCTcnt,0)) V2,sum(nvl(FCTclass,0)) V3,sum(nvl(Tcnt,0)) V4,sum(nvl(Tclass,0)) V5" & _
            " from nation,(" & _
            " select TMBM05,count(*) FCTcnt,sum(counting(tmbm08)) FCTclass,0 Tcnt,0 Tclass" & _
            " From tmbulletin, Trademark" & _
            " Where tmbm07>=" & strVol1_S & " And tmbm07<=" & strVol2_E & _
            " and tmbm01=tm15(+) and tm15 is not null" & _
            " and tm16='1' and tm01='FCT' and tmbm06='林晉章'" & _
            " group by TMBM05" & _
            " Union select TMBM05,0 FCTcnt,0 FCTclass,count(*) Tcnt,sum(counting(tmbm08)) Tclass" & _
            " From tmbulletin, Trademark" & _
            " Where tmbm07>=" & strVol1_S & " And tmbm07<=" & strVol2_E & _
            " and tmbm01=tm15(+) and tm15 is not null" & _
            " and tm16='1' and tm01='T' and tmbm06='林晉章'" & _
            " group by TMBM05" & _
            ") where TMBM05=na03(+)" & _
            " group by decode(substr(na01,1,1),'A','000','B','002',na01),decode(substr(na01,1,1),'A','台灣','B','大陸',TMBM05)" & _
            " order by decode(substr(na01,1,1),'A','000','B','002',na01) asc"
   If m_rs.State = 1 Then m_rs.Close
   m_rs.CursorLocation = adUseClient
   m_rs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If Not m_rs.EOF And Not m_rs.BOF Then
      StrMenu = True
      InsertQueryLog (m_rs.RecordCount)  'Added by Lydia 2022/01/13
      Screen.MousePointer = vbHourglass
    
      xRows = 1
      strTempFile = Me.Caption & Text1 & "至" & Text2 & "-" & ACDate(ServerDate) & ServerTime & MsgText(43)
      strPath = strExcelPath & strTempFile
      If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = "" Then
         MkDir strExcelPath
      End If
      If Dir(strPath) <> "" Then
         Kill strPath
      End If
      
      xlsSalesPoint.SheetsInNewWorkbook = 1 'Added by Lydia 2019/03/13 預設工作表數量
      xlsSalesPoint.Workbooks.add
      Set wks621 = xlsSalesPoint.Worksheets(1)
      wks621.PageSetup.Orientation = xlPortrait '直印
      '抬頭
       wks621.PageSetup.PrintTitleRows = "$1:$3"
       wks621.Columns("a:a").ColumnWidth = 15
       wks621.Columns("b:b").ColumnWidth = 11
       wks621.Columns("c:c").ColumnWidth = 11
       wks621.Columns("d:d").ColumnWidth = 11
       wks621.Columns("e:e").ColumnWidth = 11
       wks621.Columns("f:f").ColumnWidth = 11
       wks621.Columns("g:g").ColumnWidth = 11
       wks621.Columns("h:h").ColumnWidth = 11
       wks621.Columns("i:i").ColumnWidth = 11
       wks621.Range("a1").Value = Me.Caption & " " & Mid(ChangeTStringToTDateString(Text1 & "01"), 1, 6) & "至" & Mid(ChangeTStringToTDateString(Text2 & "01"), 1, 6)
       wks621.Range("a1:i1").Select
       With wks621.Range("a1:e1")
         .HorizontalAlignment = xlCenter
         .VerticalAlignment = xlBottom
         .WrapText = False
         .Orientation = 0
         .AddIndent = False
         .ShrinkToFit = False
         .MergeCells = True
      End With
      wks621.Range("a2").Value = "列印日期:"
      wks621.Range("b2").Value = ChangeTStringToTDateString(strSrvDate(2))
      wks621.Range("a3").Value = "申請人國籍"
      wks621.Range("b3").Value = "FCT件數"
      wks621.Range("c3").Value = "FCT類別數"
      wks621.Range("d3").Value = "T件數"
      wks621.Range("e3").Value = "T類別數"
      With wks621.Range("a3:e3")
         .HorizontalAlignment = xlCenter
         .VerticalAlignment = xlBottom
      End With
      
      xRows = 3: strR = xRows + 1
      With m_rs
          m_rs.MoveFirst
          Do While Not m_rs.EOF
             xRows = xRows + 1
             wks621.Range("a" & xRows).Value = m_rs.Fields("V1")
             wks621.Range("b" & xRows).Value = Trim(m_rs.Fields("V2"))
             wks621.Range("c" & xRows).Value = Trim(m_rs.Fields("V3"))
             wks621.Range("d" & xRows).Value = Trim(m_rs.Fields("V4"))
             wks621.Range("e" & xRows).Value = Trim(m_rs.Fields("V5"))
             m_rs.MoveNext
          Loop
      End With
      xRows = xRows + 1
      '台一合計
      wks621.Range("a" & xRows).Value = "合計"
      wks621.Range("b" & xRows).Formula = "=sum($b" & strR & ":$b" & xRows - 1 & ")"
      wks621.Range("c" & xRows).Formula = "=sum($c" & strR & ":$c" & xRows - 1 & ")"
      
      wks621.Range("d" & xRows).Formula = "=sum($d" & strR & ":$d" & xRows - 1 & ")"
      wks621.Range("e" & xRows).Formula = "=sum($e" & strR & ":$e" & xRows - 1 & ")"
      wks621.Range("b" & strR & ":e" & xRows).Select
      wks621.Range("b" & strR & ":e" & xRows).NumberFormatLocal = "##0"
      
      '各洲統計
      strAd = "101B,011C,C00D,C10E,C20F,C40G,C30H,小計I"
      strExc(1) = ""
      '下表不含台灣,大陸
      If txt1(0) <> "" Then
           'Modified by Lydia 2018/04/25 'T'=TA01 -> 'T'=TA01(+)
           'Memo by Lydia 2021/01/11 這段修改應該源自於2018年幫"葉特助抓商標公報統計資料->調整商標公報國外地區的代理人(TMBM06)空白也要抓到資料"；
           'Modified by Lydia 2021/01/11 'T'=TA01(+) -> 'T'=TA01 ;  遇到針對出名代理人(事務所)會造成SQL執行時間過長，所以對出名代理人(事務所)的查詢一定要抓到TMBM06。
           strExc(1) = strExc(1) & " union SELECT '50' ord1,TMBM01,DECODE(TA04,NULL,TMBM06,TA04) AS TA04,NA01,NA02,TMBM08 FROM TMBULLETIN, TAGENT, NATION" & _
                    " WHERE TMBM05=NA03(+) AND length(na01)=3 AND TMBM06=TA03(+) AND 'T'=TA01 and tmbm07>=" & strVol1_S & " And tmbm07<=" & strVol2_E & _
                    " and ta04 in (" & Pub_GetTA04(txt1(0)) & ") and substrb(na02,1,1) <> 'A' and substrb(na02,1,1) <> 'B'"
      End If
      If txt1(1) <> "" Then
            'Modified by Lydia 2018/04/25 'T'=TA01 -> 'T'=TA01(+)
            'Modified by Lydia 2021/01/11 'T'=TA01(+) -> 'T'=TA01 ;
           strExc(1) = strExc(1) & " union SELECT '60' ord1,TMBM01,DECODE(TA04,NULL,TMBM06,TA04) AS TA04,NA01,NA02,TMBM08 FROM TMBULLETIN, TAGENT, NATION" & _
                    " WHERE TMBM05=NA03(+) AND length(na01)=3 AND TMBM06=TA03(+) AND 'T'=TA01 and tmbm07>=" & strVol1_S & " And tmbm07<=" & strVol2_E & _
                    " and ta04 in (" & Pub_GetTA04(txt1(1)) & ") and substrb(na02,1,1) <> 'A' and substrb(na02,1,1) <> 'B'"
      End If
       'Modified by Lydia 2018/04/25 'T'=TA01 -> 'T'=TA01(+)
       'Modified by Lydia 2021/01/11 'T'=TA01(+) -> 'T'=TA01 ;
      strSql = "select ord1,TA04,na01,decode(substr(na02,1,1),'A','C00','B','C00',na02) na02,sum(counting(tmbm08)) cnt " & _
               "from (SELECT decode(substrb(cp12,1,1),'S','01',decode(substrb(cp12,1,2),'P2','02','F1','03','10')) ord1," & _
               "TMBM01,DECODE(TA04,NULL,TMBM06,TA04) AS TA04,NA01,NA02,TMBM08 " & _
               "From TMBULLETIN, TAGENT, NATION,trademark,caseprogress " & _
               "WHERE TMBM05=NA03(+) AND length(na01)=3 AND TMBM06=TA03(+) AND 'T'= TA01 and tmbm07>=" & strVol1_S & " And tmbm07<=" & strVol2_E & _
               " and tmbm06 ='林晉章' and tmbm01=tm15(+) and tm15 is not null and tm16='1' and tm01=cp01(+) and tm02=cp02(+) and tm03=cp03(+) and tm04=cp04(+) and cp10='101' " & _
               "and substrb(na02,1,1) <> 'A' and substrb(na02,1,1) <> 'B' " & _
               "union SELECT decode(ta04,'聖島國際','20','理律法律','30','台灣國際','40') ord1," & _
               "TMBM01,DECODE(TA04,NULL,TMBM06,TA04) AS TA04,NA01,NA02,TMBM08 FROM TMBULLETIN, TAGENT, NATION " & _
               "WHERE TMBM05=NA03(+) AND length(na01)=3 AND TMBM06=TA03(+) AND 'T'=TA01 and tmbm07>=" & strVol1_S & " And tmbm07<=" & strVol2_E & _
               " and ta04 in ('聖島國際','理律法律','台灣國際') and substrb(na02,1,1) <> 'A' and substrb(na02,1,1) <> 'B'"
      strSql = strSql & strExc(1) & ") group by ord1,ta04,na01,decode(substr(na02,1,1),'A','C00','B','C00',na02) order by ord1 "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
          xRows = xRows + 3
          wks621.Range("a" & xRows).Value = "各洲商標公告件數統計(不含台灣、大陸)"
          wks621.Range("a" & xRows & ":i" & xRows).MergeCells = True
          wks621.Range("a" & xRows & ":i" & xRows).HorizontalAlignment = xlCenter
          wks621.Range("a" & xRows & ":i" & xRows).VerticalAlignment = xlBottom
          xRows = xRows + 1
          wks621.Range("a" & xRows).Value = "(以類計)"
          wks621.Range("a" & xRows & ":i" & xRows).MergeCells = True
          wks621.Range("a" & xRows & ":i" & xRows).HorizontalAlignment = xlCenter
          wks621.Range("a" & xRows & ":i" & xRows).VerticalAlignment = xlBottom
          xRows = xRows + 1
          wks621.Range("a" & xRows & ":i" & xRows).HorizontalAlignment = xlCenter
          wks621.Range("a" & xRows & ":i" & xRows).VerticalAlignment = xlBottom
          wks621.Range("a" & xRows).Value = "部門／事務所"
          wks621.Range("b" & xRows).Value = "美國"
          wks621.Range("c" & xRows).Value = "日本"
          wks621.Range("d" & xRows).Value = "亞洲"
          wks621.Range("e" & xRows).Value = "美洲"
          wks621.Range("f" & xRows).Value = "歐洲"
          wks621.Range("g" & xRows).Value = "大洋洲"
          wks621.Range("h" & xRows).Value = "非洲"
          wks621.Range("i" & xRows).Value = "小計"
          strR = xRows + 1
          RsTemp.MoveFirst
          With RsTemp
              Do While Not RsTemp.EOF
                  i = InStr(strAd, Trim(.Fields("na01")))
                  If i = 0 Then i = InStr(strAd, Trim(.Fields("na02")))
                  If i = 0 Then GoTo NoShow
                  
                  If strTemp <> .Fields("ord1") Then
                     xRows = xRows + 1
                     '中間新增：台一小計
                     If .Fields("ord1") = "20" Then
                         wks621.Range("a" & xRows).Value = "台一小計"
                         For intI = Asc("b") To Asc("h")
                             wks621.Range(Chr(intI) & xRows).Formula = "=sum(" & Chr(intI) & strR & ":" & Chr(intI) & xRows - 1 & ")"
                         Next
                         xRows = xRows + 1
                     End If
                     If Val(.Fields("ord1")) < 20 Then
                        Select Case .Fields("ord1")
                            Case "01": wks621.Range("a" & xRows).Value = "智權部"
                            Case "02": wks621.Range("a" & xRows).Value = "商標處"
                            Case "03": wks621.Range("a" & xRows).Value = "外商"
                            Case Else: wks621.Range("a" & xRows).Value = "其他"
                        End Select
                     Else
                        If wks621.Range("a" & xRows).Value = "" Then
                            wks621.Range("a" & xRows).Value = Trim(.Fields("ta04"))
                        End If
                     End If
                  Else
                     If wks621.Range("a" & xRows).Value <> "" And Trim(.Fields("ta04")) <> "" Then
                         strExc(6) = Trim(wks621.Range("a" & xRows).Value)
                         If InStr(strExc(6), Trim(.Fields("ta04"))) = 0 Then
                            wks621.Range("a" & xRows).Value = strExc(6) & "、" & Trim(.Fields("ta04"))
                         End If
                     End If
                  End If
                  wks621.Range(Mid(strAd, i + 3, 1) & xRows).Value = Trim(Val(wks621.Range(Mid(strAd, i + 3, 1) & xRows).Value) + .Fields("cnt"))
NoShow:
                  strTemp = .Fields("ord1")
                  RsTemp.MoveNext
              Loop
          End With

          For intI = strR To xRows
              If InStr(Trim(wks621.Range("a" & intI).Value), "小計") = 0 Then
                 '美國+美洲
                 wks621.Range("e" & intI).Value = Trim(Val(wks621.Range("b" & intI).Value) + Val(wks621.Range("e" & intI).Value))
                 '日本+亞洲
                 wks621.Range("d" & intI).Value = Trim(Val(wks621.Range("c" & intI).Value) + Val(wks621.Range("d" & intI).Value))
              End If
             '部門小計
              wks621.Range("i" & intI).Formula = "=sum(d" & intI & ":h" & intI & ")"
          Next
          '美國、日本加框線
            wks621.Range("b" & strR - 1 & ":c" & xRows).Select
            xlsSalesPoint.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
            xlsSalesPoint.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
            xlsSalesPoint.Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
            xlsSalesPoint.Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
            xlsSalesPoint.Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
            xlsSalesPoint.Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
      End If
      
       wks621.Range("c1").Select
      '判斷若版本2007以上改變存格式
      If Val(xlsSalesPoint.Version) < 12 Then
          xlsSalesPoint.Workbooks(1).SaveAs FileName:=strPath, FileFormat:=-4143
      Else
          xlsSalesPoint.Workbooks(1).SaveAs FileName:=strPath, FileFormat:=56
      End If
      xlsSalesPoint.Workbooks.Close
      xlsSalesPoint.Quit
      'Modify by Amy 2021/06/21 原:strPath 改中文字顯示
      MsgBox "檔案已產生！" & vbCrLf & "檔案存於 " & strExcelPathN & " " & strTempFile, vbInformation
'---------------------------------------------------
   Else
      MsgBox "查詢無資料！", vbExclamation + vbOKOnly
      InsertQueryLog (0) 'Added by Lydia 2022/01/13
      Exit Function
   End If
   Screen.MousePointer = vbDefault
End Function


Private Sub Form_Load()
   MoveFormToCenter Me
   
   Text1 = Left(strSrvDate(2), 5)
   Text2 = Left(strSrvDate(2), 5)
   Label3.Caption = Label3 & strExcelPathN 'Modify by Amy 2021/06/21
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set frm030621 = Nothing
End Sub


Private Sub Text1_GotFocus()
   InverseTextBox Text1
   CloseIme
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
Dim rsQuery As ADODB.Recordset
Dim stSQL As String, intR As Integer
Dim strVol1 As String, strVol2 As String
Dim intCnt As Integer
Dim intC2 As Integer
   If Text1 <> "" Then
      If ChkDate(Text1 & "01") = False Then
         Call Text1_GotFocus
         Cancel = True
         Exit Sub
      End If
      '檢查資料是否已存在,每月有2期公報
      Call Pub_ChgDateToTMBM07(Text1, strVol1, strVol2)
      stSQL = "select tmbm07 from tmbulletin where tmbm07>='" & strVol1 & "' and tmbm07<='" & strVol2 & "' group by tmbm07"
      intR = 1
      Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
      If intR = 1 Then
         intCnt = Val("" & rsQuery.Fields(0))
      End If
      
      intC2 = rsQuery.RecordCount
      rsQuery.Close
      If intCnt = 0 Then
         MsgBox Text1 & "此月份尚無公報資料!!"
         Call Text1_GotFocus
         Cancel = True
         Exit Sub
      ElseIf intC2 < 2 Then
         MsgBox Text1 & "此月份公報資料尚不足!!"
         Call Text1_GotFocus
         Cancel = True
         Exit Sub
      End If
   End If
   
   Set rsQuery = Nothing
End Sub

Private Sub Text2_GotFocus()
   InverseTextBox Text2
   CloseIme
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
Dim rsQuery As ADODB.Recordset
Dim stSQL As String, intR As Integer
Dim strVol1 As String, strVol2 As String
Dim intCnt As Integer
Dim intC2 As Integer
   If Text2 <> "" Then
      If ChkDate(Text2 & "01") = False Then
          Call Text2_GotFocus
          Cancel = True
          Exit Sub
      End If
      '檢查資料是否已存在,每月有2期公報
      Call Pub_ChgDateToTMBM07(Text2, strVol1, strVol2)
      stSQL = "select tmbm07 from tmbulletin where tmbm07>='" & strVol1 & "' and tmbm07<='" & strVol2 & "' group by tmbm07"
      intR = 1
      Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
      If intR = 1 Then
         intCnt = Val("" & rsQuery.Fields(0))
      End If
      
      intC2 = rsQuery.RecordCount
      rsQuery.Close
      If intCnt = 0 Then
         MsgBox Text2 & "此月份尚無公報資料!!"
         Call Text2_GotFocus
         Cancel = True
         Exit Sub
      ElseIf intC2 < 2 Then
         MsgBox Text2 & "此月份公報資料尚不足!!"
         Call Text2_GotFocus
         Cancel = True
         Exit Sub
      End If
   End If
   Set rsQuery = Nothing

End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
    If txt1(Index) <> "" Then
       If InStr(Label4(1).Caption & ";" & Label4(2).Caption & ";" & Label4(3).Caption & ";" & Label4(4).Caption & ";", txt1(Index)) > 0 Then
          MsgBox "請勿輸入1-4的事務所名稱!", vbCritical
          txt1(Index).SetFocus
          txt1_GotFocus Index
          Cancel = True
       End If
    End If
End Sub
Private Sub txt1_GotFocus(Index As Integer)
    TextInverse txt1(Index)
    OpenIme
End Sub
