VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm04060312 
   BorderStyle     =   1  '單線固定
   Caption         =   "國籍及洲別統計(含同業)"
   ClientHeight    =   4170
   ClientLeft      =   2790
   ClientTop       =   3945
   ClientWidth     =   4830
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   4830
   Begin VB.TextBox txtDate 
      Height          =   345
      Index           =   1
      Left            =   2760
      MaxLength       =   5
      TabIndex        =   1
      Top             =   720
      Width           =   795
   End
   Begin VB.TextBox txtDate 
      Height          =   345
      Index           =   0
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
   Begin MSForms.TextBox txt1 
      Height          =   345
      Index           =   1
      Left            =   1680
      TabIndex        =   3
      Top             =   2970
      Width           =   2295
      VariousPropertyBits=   671105051
      Size            =   "4048;609"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   345
      Index           =   0
      Left            =   1680
      TabIndex        =   2
      Top             =   2550
      Width           =   2295
      VariousPropertyBits=   671105051
      Size            =   "4048;609"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label2 
      Caption         =   "PS：本所不含無新申請案進度案件(例:FCP-050893)"
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   360
      TabIndex        =   15
      Top             =   3480
      Width           =   4095
   End
   Begin VB.Label Label3 
      Caption         =   "PS2：Excel儲存於"
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   360
      TabIndex        =   14
      Top             =   3840
      Width           =   3200
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
Attribute VB_Name = "frm04060312"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/10/08 改成Form2.0 ; txt1(index)
'Create By Lydia 2017/01/05 申請人國籍及洲別統計(含同業)
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
         
         For intI = 0 To 1
             txt1_Validate intI, Cancel
             If Cancel = True Then Exit Sub
         Next
         
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
Dim endX As String   '小計-位置
Dim m_rs As New ADODB.Recordset
Dim strVol1_S As String, strVol1_E As String '日期範圍

Dim mESeqNo As String

Dim strMid As String, strMid2 As String
Dim xlsSalesPoint As New Excel.Application
Dim wks4632 As New Worksheet
Dim xRows As Integer '目前列位置

Dim strTemp As String
Dim strPath As String, strTempFile As String
Dim iCall As Integer  '切換各單位和指定事務所
Dim inJ As Integer
Dim mArea As Integer '各部門和事務所的計算
Dim mQty As Integer  '計件
Dim strRC As String, strAd As String

   StrMenu = False

   strVol1_S = TransDate(txtDate(0) & "01", 2)
   strVol1_E = GetLastDay(TransDate(txtDate(1) & "01", 2))
   'Added By Lydia 2022/01/13 查詢印表記錄檔欄位
   ClearQueryLog (Me.Name)
   pub_QL05 = pub_QL05 & ";" & Label1.Caption & strVol1_S & "-" & strVol1_E
   pub_QL05 = pub_QL05 & ";" & Label4(0).Caption & Label4(1).Caption & "、" & Label4(2).Caption & "、" & Label4(3).Caption & "、" & Label4(4).Caption
   If txt1(0) <> "" Then pub_QL05 = pub_QL05 & "、" & txt1(0)
   If txt1(1) <> "" Then pub_QL05 = pub_QL05 & "、" & txt1(1)
  'end 2022/01/13
  
   '智權部(A)、FCP(B)、其他(C)的列位置
    strRC = "A04,B05,C06,T07"
    '國籍和洲別的行位置
    strAd = "C0d,C1e,C2f,C4g,C3h,T0i"
    
   strMid = "SELECT TPB01,TPB02,TPB03,TPB04,TPB05,decode(tpb06,'020','C0',substr(na02,1,2))||TPB06 TPB06,TPB07,TPB08,'' TPB09,PA01,PA02,PA03,PA04 " & _
            "From TPBulletin, nation, PATENT where na01(+)=tpb06 and pa11(+)=tpb01 and pa23(+)='1' and tpb03>=" & strVol1_S & " and tpb03<=" & strVol1_E
   
   mArea = 6 + IIf(Trim(txt1(0).Text) <> "", 1, 0) + IIf(Trim(txt1(1).Text) <> "", 1, 0)
   
   For iCall = 1 To mArea
       '台一各單位
       If iCall <= 3 Then
           Select Case iCall
               Case 1 '美國
                    strSql = "SELECT '0' ORD1,'美國' TITLE,X1.* FROM (" & strMid & " AND (TPB08='台一國際' OR (TPB07 IS NULL AND PA01 IS NOT NULL))) X1 " & _
                             "WHERE TPB06='C1101' "
               Case 2 '日本
                    strSql = "SELECT '1' ORD1,'日本' TITLE,X1.* FROM (" & strMid & " AND (TPB08='台一國際' OR (TPB07 IS NULL AND PA01 IS NOT NULL))) X1 " & _
                             "WHERE TPB06='C0011' "
               Case 3 '各洲
                    strSql = "SELECT '2' ORD1,'各洲' TITLE,X1.* FROM (" & strMid & " AND (TPB08='台一國際' OR (TPB07 IS NULL AND PA01 IS NOT NULL))) X1 " & _
                             "WHERE SUBSTR(TPB06,1,1)<>'A' AND TPB06<>'C0020' "
           End Select
       '指定事務所
       Else
           Select Case iCall
               Case 4: strMid2 = " and tpb08='聖島國際' "
               Case 5: strMid2 = " and tpb08='理律法律' "
               Case 6: strMid2 = " and tpb08='台灣國際' "
               Case 7: strMid2 = " and tpb08 like '%" & IIf(Trim(txt1(0)) = "", Trim(txt1(1)), Trim(txt1(0))) & "%' "
               Case 8: strMid2 = " and tpb08 like '%" & Trim(txt1(1)) & "%' "
           End Select
            '美國
            strSql = "select '0' ord1,substr(na03,1,2) title,tpb08 deptno,count(*) cnt from nation,(" & _
                     strMid & ") where tpb06='C1101' " & strMid2 & " and substr(tpb06,3)=na01(+) group by substr(na03,1,2) ,tpb08 "
            '日本
            strSql = strSql & "Union all select '1' ord1,substr(na03,1,2) title,tpb08 deptno,count(*) cnt from nation,(" & _
                     strMid & ") where tpb06='C0011' " & strMid2 & " and substr(tpb06,3)=na01(+) group by substr(na03,1,2) ,tpb08 "
            '各洲
            strSql = strSql & " Union all select '2' ord1,decode(substr(tpb06,1,1),'A','0亞洲',decode(substr(tpb06,1,2),'C0','0亞洲','C1','1美洲','C2','2歐洲','C3','4非洲','C4','3大洋洲',tpb06)) title,tpb08 deptno,count(*) cnt from nation,(" & _
                     strMid & ") where substr(tpb06,3)=na01(+) " & strMid2 & " and substr(tpb06,1,1)<>'A' and tpb06<>'C0020' group by decode(substr(tpb06,1,1),'A','0亞洲',decode(substr(tpb06,1,2),'C0','0亞洲','C1','1美洲','C2','2歐洲','C3','4非洲','C4','3大洋洲',tpb06)),tpb08"
                     
            strSql = strSql & " order by 1,2,3 "
       End If
       
       If m_rs.State = 1 Then m_rs.Close
       m_rs.CursorLocation = adUseClient
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
             Set wks4632 = xlsSalesPoint.Worksheets(1)
             wks4632.PageSetup.Orientation = xlPortrait '直印
             '抬頭
             wks4632.PageSetup.PrintTitleRows = "$1:$3"
             
             wks4632.Columns("a:a").ColumnWidth = 10 '項目
             For intI = Asc(startX) To Asc(startX) + 8
                 wks4632.Columns(Chr(intI) & ":" & Chr(intI)).ColumnWidth = 8
                 wks4632.Columns(Chr(intI) & ":" & Chr(intI)).HorizontalAlignment = xlCenter
                 wks4632.Columns(Chr(intI) & ":" & Chr(intI)).VerticalAlignment = xlBottom
             Next
             
             strExc(0) = Mid(txtDate(0), 1, 3) & "/" & Mid(txtDate(0), 4, 2) & "至" & Mid(txtDate(1), 1, 3) & "/" & Mid(txtDate(1), 4, 2)
             wks4632.Range("a1").Value = strExc(0) & " " & Me.Caption
             With wks4632.Range("a1:i1")
               .WrapText = False
               .MergeCells = True
               .HorizontalAlignment = xlCenter
               .VerticalAlignment = xlBottom
             End With
             
             wks4632.Range("a2").Value = "(不含台灣、大陸;本所不含無新申請案進度案件)"
             wks4632.Range("a2").Font.Color = &HFF&
             With wks4632.Range("a2:i2")
               .WrapText = False
               .MergeCells = True
               .HorizontalAlignment = xlCenter
               .VerticalAlignment = xlBottom
             End With
             
             xRows = 3
             wks4632.Range("b3").Value = "美國"
             wks4632.Range("c3").Value = "日本"
             wks4632.Range("d3").Value = "亞洲"
             wks4632.Range("e3").Value = "美洲"
             wks4632.Range("f3").Value = "歐洲"
             wks4632.Range("g3").Value = "大洋洲"
             wks4632.Range("h3").Value = "非洲"
             wks4632.Range("i3").Value = "小計"
             endX = "i"
             xRows = xRows + 1
             
             wks4632.Range("a" & xRows).Value = "智權部"
             wks4632.Range("a" & xRows + 1).Value = "FCP"
             wks4632.Range("a" & xRows + 2).Value = "其他"
             wks4632.Range("a" & xRows + 3).Value = "台一小計"
          End If
          '---------------
          
          With m_rs
             .MoveFirst
             If iCall = 4 Then xRows = 7 '跳到台一小計
             Do While Not .EOF
                If iCall <= 3 Then '台一各單位
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
                    If mQty = 0 Then
                       GoTo JumpNextRec '不計件,就跳下一筆資料
                    End If
                  
                    strSql = "SELECT DECODE(SUBSTR(CP12,1,1),'S','1'||CP12,'F','2'||CP12,'3其他') DEPTNO, 1 AS A1 FROM CASEPROGRESS " & _
                             "WHERE CP01='" & .Fields("PA01") & "' AND CP02='" & .Fields("PA02") & "' AND CP03='" & .Fields("PA03") & "' AND CP04='" & .Fields("PA04") & "' " & _
                             "AND CP10 IN ('101','102','103','104','105','109','110','112','113','114','115','118','120','122','125','307') " & _
                             "GROUP BY DECODE(SUBSTR(CP12,1,1),'S','1'||CP12,'F','2'||CP12,'3其他') "
                    intI = 1
                    Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                    If intI = 1 Then
                       '計算各部門資料
                       Select Case Mid(RsTemp.Fields("deptno"), 1, 1)
                           Case "1": xRows = Val(Mid(strRC, InStr(strRC, "A") + 1, 2))
                           Case "2": xRows = Val(Mid(strRC, InStr(strRC, "B") + 1, 2))
                           Case "3": xRows = Val(Mid(strRC, InStr(strRC, "C") + 1, 2))
                       End Select
                       '抓各洲的行位置
                        If "" & .Fields("ord1") = "0" Then
                            strExc(1) = "b"
                        ElseIf "" & .Fields("ord1") = "1" Then
                            strExc(1) = "c"
                        ElseIf "" & .Fields("ord1") = "2" Then
                            strExc(1) = Mid(strAd, InStr(strAd, Mid("" & .Fields("TPB06"), 1, 2)) + 2, 1)
                        End If
                       
                         wks4632.Range(strExc(1) & xRows).Value = Val(wks4632.Range(strExc(1) & xRows).Value) + mQty
                    End If
                Else
                   '其他事務所
                   If strTemp <> "" & .Fields("deptno") Then
                      xRows = xRows + 1
                      If Trim(wks4632.Range("a" & xRows).Value) = "" Then
                         wks4632.Range("a" & xRows).Value = "" & .Fields("deptno")
                      End If
                   End If
                   
                   '各洲順序
                   If InStr("0,1,2,3,4,5,6,7,8,9", Mid(.Fields("title"), 1, 1)) > 0 Then
                      wks4632.Range(Chr(Asc(startX) + Val(.Fields("ord1")) + Val(Mid(.Fields("title"), 1, 1))) & xRows).Value = "" & .Fields("cnt")
                   Else
                      wks4632.Range(Chr(Asc(startX) + Val(.Fields("ord1"))) & xRows).Value = "" & .Fields("cnt")
                   End If
                End If
                
                If iCall >= 4 Then strTemp = "" & .Fields("deptno")
JumpNextRec:
                .MoveNext
             Loop
          End With
       End If
   Next iCall
   
   '設定台一小計
   For inJ = Asc(startX) To Asc(endX) - 1
       wks4632.Range(Chr(inJ) & "7").Value = "=SUM(" & Chr(inJ) & "4:" & Chr(inJ) & "6)"
   Next inJ
   
   '設定各單位和事務所小計
   For inJ = 4 To xRows
       '美國、日本不計入
       wks4632.Range(endX & inJ).Value = "=SUM(" & Chr(Asc(startX) + 2) & inJ & ":" & Chr(Asc(endX) - 1) & inJ & ")"
   Next inJ
   
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
   Set frm04060312 = Nothing
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


