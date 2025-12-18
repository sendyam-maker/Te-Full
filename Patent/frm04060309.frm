VERSION 5.00
Begin VB.Form frm04060309 
   BorderStyle     =   1  '單線固定
   Caption         =   "各單位專利公報件數統計"
   ClientHeight    =   2295
   ClientLeft      =   2790
   ClientTop       =   3945
   ClientWidth     =   4650
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   4650
   Begin VB.TextBox txtDate 
      Height          =   345
      Index           =   1
      Left            =   2520
      MaxLength       =   5
      TabIndex        =   1
      Top             =   1080
      Width           =   795
   End
   Begin VB.TextBox txtDate 
      Height          =   345
      Index           =   0
      Left            =   1530
      MaxLength       =   5
      TabIndex        =   0
      Top             =   1080
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
      Caption         =   "PS：不含無新申請案進度案件(例:FCP-050893)"
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   600
      TabIndex        =   6
      Top             =   1560
      Width           =   3615
   End
   Begin VB.Label Label3 
      Caption         =   "PS：Excel儲存於"
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   600
      TabIndex        =   5
      Top             =   1920
      Width           =   3200
   End
   Begin VB.Line Line1 
      X1              =   2100
      X2              =   2790
      Y1              =   1260
      Y2              =   1260
   End
   Begin VB.Label Label1 
      Caption         =   "公報年月："
      Height          =   210
      Left            =   570
      TabIndex        =   4
      Top             =   1170
      Width           =   900
   End
End
Attribute VB_Name = "frm04060309"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2021/11/17 Form2.0已檢查 (無需修改的物件)
'Create By Lydia 2017/01/05 各單位專利公報件數統計
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
Dim inR1 As Integer
Dim m_rs As New ADODB.Recordset
Dim strVol1_S As String, strVol1_E As String
Dim xlsSalesPoint As New Excel.Application
Dim wks4639 As New Worksheet
Dim xRows As Integer '目前列位置
Dim strR As Integer
Dim strAd As String, strAt As String
Dim strRC As String
Dim tmpArr As Variant
Dim strTemp As String
Dim strPath As String, strTempFile As String
Dim strA1 As String, strA2 As String
Dim strMid As String
Dim mQty As Integer  '計件

   StrMenu = False
   
   strVol1_S = TransDate(txtDate(0) & "01", 2)
   strVol1_E = GetLastDay(TransDate(txtDate(1) & "01", 2))
   'Added By Lydia 2022/01/13 查詢印表記錄檔欄位
   ClearQueryLog (Me.Name)
   pub_QL05 = pub_QL05 & ";" & Label1.Caption & strVol1_S & "-" & strVol1_E
  'end 2022/01/13
  
    '各部門欄位置
    strAt = "北一,北三,北四,北五,中一,中二,中三,南所,高所,智權部,FCP,其他,小計"
    strAd = "S11B,S13C,S14D,S15E,S21F,S22G,S23H,S31I,S41J,SXXK,FXXL,OXXM,小計N"
    '國內A、大陸B、國外C的列位置
    strRC = "A04,B06,C08,T10"
   
    strSql = "SELECT TPB01,TPB02,TPB03,TPB04,TPB05,decode(tpb06,'020','C0',substr(na02,1,2))||TPB06 TPB06,TPB07,TPB08,'' TPB09,PA01,PA02,PA03,PA04" & _
            " From TPBulletin, nation, PATENT where na01(+)=tpb06 and pa11(+)=tpb01 and pa23(+)='1' and TPB03 >=" & strVol1_S & " and TPB03 <=" & strVol1_E & _
            " and (tpb08='台一國際' or (tpb07 is null and pa01 is not null)) "
    strSql = "select tpb03,tpb06,tpb08,pa01,pa02,pa03,pa04 from (" & strSql & ") "
   If m_rs.State = 1 Then m_rs.Close
   m_rs.CursorLocation = adUseClient
   m_rs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If Not m_rs.EOF And Not m_rs.BOF Then
      StrMenu = True
      InsertQueryLog (m_rs.RecordCount) 'Added by Lydia 2022/01/13
      With m_rs
          .MoveFirst
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
          Set wks4639 = xlsSalesPoint.Worksheets(1)
          wks4639.PageSetup.Orientation = xlPortrait '直印
          '抬頭
           wks4639.PageSetup.PrintTitleRows = "$1:$3"
           
           wks4639.Columns("a:a").ColumnWidth = 7 '項目
           For intI = Asc("b") To Asc("n")
               wks4639.Columns(Chr(intI) & ":" & Chr(intI)).ColumnWidth = 8
           Next
           
           strExc(0) = Mid(txtDate(0), 1, 3) & "/" & Mid(txtDate(0), 4, 2) & "至" & Mid(txtDate(1), 1, 3) & "/" & Mid(txtDate(1), 4, 2)
           wks4639.Range("a1").Value = strExc(0) & " " & Me.Caption
           With wks4639.Range("a1:n1")
             .HorizontalAlignment = xlCenter
             .VerticalAlignment = xlBottom
             .WrapText = False
             .MergeCells = True
          End With
          wks4639.Range("a2").Value = "(不含無新申請案進度案件)"
          wks4639.Range("a2").Font.Color = &HFF&
           With wks4639.Range("a2:n2")
             .HorizontalAlignment = xlCenter
             .VerticalAlignment = xlBottom
             .WrapText = False
             .MergeCells = True
          End With
          
          xRows = 3
          tmpArr = Empty
          tmpArr = Split(strAt, ",")
          inR1 = Asc("b")
          wks4639.Range("a" & xRows).Value = "項目"
          For intI = 0 To UBound(tmpArr)
              If tmpArr(intI) <> "" Then
                 wks4639.Range(Chr(inR1 + intI) & xRows).Value = Trim(tmpArr(intI))
              End If
          Next
          With wks4639.Range("a" & xRows & ":" & Chr(inR1 + intI) & xRows)
             .HorizontalAlignment = xlCenter
             .VerticalAlignment = xlBottom
          End With
          
          strR = xRows + 1
          wks4639.Range("a" & Val(Mid(strRC, InStr(strRC, "A") + 1, 2))).Value = "國內"
          wks4639.Range("a" & Val(Mid(strRC, InStr(strRC, "A") + 1, 2)) + 1).Value = "比例"
          '小計
          wks4639.Range("n" & Val(Mid(strRC, InStr(strRC, "A") + 1, 2))).Value = "=SUM(k" & Val(Mid(strRC, InStr(strRC, "A") + 1, 2)) & ":m" & Val(Mid(strRC, InStr(strRC, "A") + 1, 2)) & ")"
          
          wks4639.Range("a" & Val(Mid(strRC, InStr(strRC, "B") + 1, 2))).Value = "大陸"
          wks4639.Range("a" & Val(Mid(strRC, InStr(strRC, "B") + 1, 2)) + 1).Value = "比例"
          '小計
          wks4639.Range("n" & Val(Mid(strRC, InStr(strRC, "B") + 1, 2))).Value = "=SUM(k" & Val(Mid(strRC, InStr(strRC, "B") + 1, 2)) & ":m" & Val(Mid(strRC, InStr(strRC, "B") + 1, 2)) & ")"
          
          wks4639.Range("a" & Val(Mid(strRC, InStr(strRC, "C") + 1, 2))).Value = "國外"
          wks4639.Range("a" & Val(Mid(strRC, InStr(strRC, "C") + 1, 2)) + 1).Value = "比例"
          '小計
          wks4639.Range("n" & Val(Mid(strRC, InStr(strRC, "C") + 1, 2))).Value = "=SUM(k" & Val(Mid(strRC, InStr(strRC, "C") + 1, 2)) & ":m" & Val(Mid(strRC, InStr(strRC, "C") + 1, 2)) & ")"
                    
          wks4639.Range("a" & Val(Mid(strRC, InStr(strRC, "T") + 1, 2))).Value = "合計"
          wks4639.Range("a" & Val(Mid(strRC, InStr(strRC, "T") + 1, 2)) + 1).Value = "比例"
          '小計
          wks4639.Range("n" & Val(Mid(strRC, InStr(strRC, "T") + 1, 2))).Value = "=SUM(k" & Val(Mid(strRC, InStr(strRC, "T") + 1, 2)) & ":m" & Val(Mid(strRC, InStr(strRC, "T") + 1, 2)) & ")"
                    
                    
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
            If mQty = 0 Then
               GoTo JumpNextRec '不計件,就跳下一筆資料
            End If
            
            If Mid("" & .Fields("tpb06"), 1, 1) = "A" Then '國內
               xRows = Val(Mid(strRC, InStr(strRC, "A") + 1, 2))
            ElseIf "" & .Fields("tpb06") = "C0020" Then    '大陸
               xRows = Val(Mid(strRC, InStr(strRC, "B") + 1, 2))
            Else                                           '國外
               xRows = Val(Mid(strRC, InStr(strRC, "C") + 1, 2))
            End If
            
            strSql = "SELECT DECODE(SUBSTR(CP12,1,1),'S','1'||CP12,'F','2'||CP12,'3其他') DEPTNO, 1 AS A1 FROM CASEPROGRESS " & _
                     "WHERE CP01='" & .Fields("PA01") & "' AND CP02='" & .Fields("PA02") & "' AND CP03='" & .Fields("PA03") & "' AND CP04='" & .Fields("PA04") & "' " & _
                     "AND CP10 IN ('101','102','103','104','105','109','110','112','113','114','115','118','120','122','125','307') " & _
                     "GROUP BY DECODE(SUBSTR(CP12,1,1),'S','1'||CP12,'F','2'||CP12,'3其他') "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
             '計算各部門資料
             Select Case Mid(RsTemp.Fields("deptno"), 2, 1)
                Case "S"
                      inR1 = InStr(strAd, Mid(RsTemp.Fields("deptno"), 2))
                      If inR1 > 0 Then
                          wks4639.Range(Mid(strAd, inR1 + 3, 1) & xRows).Value = Val(wks4639.Range(Mid(strAd, inR1 + 3, 1) & xRows).Value) + mQty
                      Else
                         GoTo JumpOtherDept
                      End If
                      wks4639.Range(Mid(strAd, InStr(strAd, "SXX") + 3, 1) & xRows).Value = Val(wks4639.Range(Mid(strAd, InStr(strAd, "SXX") + 3, 1) & xRows).Value) + mQty
                Case "F"
                      wks4639.Range(Mid(strAd, InStr(strAd, "FXX") + 3, 1) & xRows).Value = Val(wks4639.Range(Mid(strAd, InStr(strAd, "FXX") + 3, 1) & xRows).Value) + mQty
                Case Else
JumpOtherDept:
                      wks4639.Range(Mid(strAd, InStr(strAd, "OXX") + 3, 1) & xRows).Value = Val(wks4639.Range(Mid(strAd, InStr(strAd, "OXX") + 3, 1) & xRows).Value) + mQty
             End Select
            End If

JumpNextRec:
             .MoveNext
          Loop
      End With
      
      '合計
      xRows = Val(Mid(strRC, InStr(strRC, "T") + 1, 2))
      wks4639.Range("a" & xRows).Value = "合計"
      For intI = Asc("B") To Asc("M")
         wks4639.Range(Chr(intI) & xRows).Formula = "=" & Chr(intI) & Val(Mid(strRC, InStr(strRC, "A") + 1, 2)) & "+" & Chr(intI) & Val(Mid(strRC, InStr(strRC, "B") + 1, 2)) & "+" & Chr(intI) & Val(Mid(strRC, InStr(strRC, "C") + 1, 2))
         wks4639.Range(Chr(intI) & xRows).NumberFormatLocal = "###"
         '計算比例
         If intI >= Asc("K") And Val(wks4639.Range(Chr(intI) & xRows).Value) <> 0 Then
            wks4639.Range(Chr(intI) & xRows + 1).Formula = "=" & Chr(intI) & xRows & "/" & "N" & xRows
            wks4639.Range(Chr(intI) & xRows + 1).NumberFormatLocal = "##0.00%"
            '小計的比例
            For inR1 = Asc("A") To Asc("C")
                strExc(1) = Val(Mid(strRC, InStr(strRC, Chr(inR1)) + 1, 2))
                If Val(wks4639.Range(Chr(intI) & strExc(1))) <> 0 Then
                    wks4639.Range(Chr(intI) & Val(strExc(1)) + 1).Formula = "=" & Chr(intI) & Val(strExc(1)) & "/" & "N" & Val(strExc(1))
                    wks4639.Range(Chr(intI) & Val(strExc(1)) + 1).NumberFormatLocal = "##0.00%"
                End If
            Next inR1
         End If
      Next

        '判斷若版本2007以上改變存格式
        If Val(xlsSalesPoint.Version) < 12 Then
        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strPath, FileFormat:=-4143
        Else
        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strPath, FileFormat:=56
        End If
        xlsSalesPoint.Workbooks.Close
        xlsSalesPoint.Quit
        'Modify by Amy 2021/06/22 原:strPath 改中文字顯示
        MsgBox "檔案已產生！" & vbCrLf & "檔案存於 " & strExcelPathN & " " & strTempFile, vbInformation
        Exit Function
   Else
      MsgBox "查詢無資料！", vbExclamation + vbOKOnly
      InsertQueryLog (0) 'Added by Lydia 2022/01/13
      Exit Function
   End If

End Function

Private Sub Form_Load()
   MoveFormToCenter Me
   
   txtDate(0) = Left(strSrvDate(2), 5)
   txtDate(1) = Left(strSrvDate(2), 5)
   Label3.Caption = Label3 & strExcelPathN 'Add by Amy 2021/06/22
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm04060309 = Nothing
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
