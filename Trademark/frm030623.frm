VERSION 5.00
Begin VB.Form frm030623 
   BorderStyle     =   1  '單線固定
   Caption         =   "各單位公報類別數統計"
   ClientHeight    =   2010
   ClientLeft      =   2790
   ClientTop       =   3945
   ClientWidth     =   4650
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2010
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
   Begin VB.Label Label3 
      Caption         =   "PS：Excel儲存於"
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   510
      TabIndex        =   5
      Top             =   1560
      Width           =   3195
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
Attribute VB_Name = "frm030623"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create By Lydia 2015/12/10 各單位公報類別數統計
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
             txtDate_Validate intI, Cancel
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
Dim strVol1_S As String, strVol2_S As String
Dim strVol1_E As String, strVol2_E As String
Dim xlsSalesPoint As New Excel.Application
Dim wks623 As New Worksheet
Dim xRows As Integer '目前列位置
Dim strR As Integer
Dim strAd As String, strAt As String
Dim tmpArr As Variant
Dim strTemp As String
Dim strPath As String, strTempFile As String
Dim strA1 As String, strA2 As String

   StrMenu = False
   
   'Added By Lydia 2022/01/13 查詢印表記錄檔欄位
   ClearQueryLog (Me.Name)
   pub_QL05 = pub_QL05 & ";" & Label1.Caption & txtDate(0) & "-" & txtDate(1)
  'end 2022/01/13
  
   Call Pub_ChgDateToTMBM07(txtDate(0), strVol1_S, strVol2_S)
   Call Pub_ChgDateToTMBM07(txtDate(1), strVol1_E, strVol2_E)
   'Modified by Lydia 2018/04/25 'T'=TA01 -> 'T'=TA01(+)
   'Memo by Lydia 2021/01/11 這段修改應該源自於2018年幫"葉特助抓商標公報統計資料->調整商標公報國外地區的代理人(TMBM06)空白也要抓到資料"；
   'Modified by Lydia 2021/01/11 'T'=TA01(+) -> 'T'=TA01 ;  遇到針對出名代理人(事務所)會造成SQL執行時間過長，所以對出名代理人(事務所)的查詢一定要抓到TMBM06。
   strSql = "SELECT CP12,SUBSTR(NA02,1,1) NA00,SUM(COUNTING(TMBM08)) CNT FROM TMBULLETIN, TAGENT, NATION,TRADEMARK,CASEPROGRESS " & _
            "WHERE TMBM05=NA03(+) AND LENGTH(NA01)=3 AND TMBM06=TA03(+) AND 'T'=TA01 AND TMBM07>=" & strVol1_S & " AND TMBM07<=" & strVol2_E & _
            "AND TMBM06 ='林晉章' AND TMBM01=TM15(+) AND TM15 IS NOT NULL AND TM16='1' AND TM01=CP01(+) AND TM02=CP02(+) AND TM03=CP03(+) AND TM04=CP04(+) AND CP10='101' " & _
            "GROUP BY CP12,SUBSTR(NA02,1,1) ORDER BY 2,1"
   If m_rs.State = 1 Then m_rs.Close
   m_rs.CursorLocation = adUseClient
   m_rs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If Not m_rs.EOF And Not m_rs.BOF Then
      StrMenu = True
      InsertQueryLog (m_rs.RecordCount)  'Added by Lydia 2022/01/13
      '各部門欄位置
      strAt = "北一,北三,北四,北五,中一,中二,中三,南所,高所,智權部,商標處,外商,其他,小計"
      strAd = "S11B,S13C,S14D,S15E,S21F,S22G,S23H,S31I,S41J,SXXK,P2XL,F1XM,OXXN,小計O"
      strExc(1) = ""
      m_rs.MoveFirst
      xRows = 1
      strTempFile = Me.Caption & txtDate(0) & "至" & txtDate(1) & "-" & ACDate(ServerDate) & ServerTime & MsgText(43)
      strPath = strExcelPath & strTempFile
      
      If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = "" Then
         MkDir strExcelPath
      End If
      If Dir(strPath) <> "" Then
         Kill strPath
      End If
      
      xlsSalesPoint.SheetsInNewWorkbook = 3 'Modfiy by Amy 2021/06/21 Added by Lydia 2019/03/13 預設工作表數量
      xlsSalesPoint.Workbooks.add
      Set wks623 = xlsSalesPoint.Worksheets(1)
      wks623.PageSetup.Orientation = xlPortrait '直印
      '抬頭
       wks623.PageSetup.PrintTitleRows = "$1:$3"
       
       wks623.Columns("a:a").ColumnWidth = 6 '項目
       For intI = Asc("b") To Asc("o")
           wks623.Columns(Chr(intI) & ":" & Chr(intI)).ColumnWidth = 6
       Next

       wks623.Range("a1").Value = Mid(ChangeTStringToTDateString(txtDate(0) & "01"), 1, 6) & "至" & Mid(ChangeTStringToTDateString(txtDate(1) & "01"), 1, 6) & " " & Me.Caption
       wks623.Range("a1:o1").Select
       With wks623.Range("a1:o1")
         .HorizontalAlignment = xlCenter
         .VerticalAlignment = xlBottom
         .WrapText = False
         .MergeCells = True
      End With
      wks623.Range("a2").Value = "(以類計)"
       With wks623.Range("a2:o2")
         .HorizontalAlignment = xlCenter
         .VerticalAlignment = xlBottom
         .WrapText = False
         .MergeCells = True
      End With
      xRows = 3
      tmpArr = Empty
      tmpArr = Split(strAt, ",")
      inR1 = Asc("b")
      wks623.Range("a" & xRows).Value = "項目"
      For intI = 0 To UBound(tmpArr)
          If tmpArr(intI) <> "" Then
             wks623.Range(Chr(inR1 + intI) & xRows).Value = Trim(tmpArr(intI))
          End If
      Next
      With wks623.Range("a" & xRows & ":" & Chr(inR1 + intI) & xRows)
         .HorizontalAlignment = xlCenter
         .VerticalAlignment = xlBottom
      End With
      
      strR = xRows + 1
      Do While Not m_rs.EOF
         If strTemp <> m_rs.Fields("na00") Then
            If strTemp <> "" Then
               strA1 = Mid(strAd, InStr(strAd, "SXX") + 3, 1)
               strA2 = Right(strAd, 1)
               '小計
               wks623.Range(strA2 & xRows).Formula = "=sum($" & strA1 & xRows & ":$" & Chr(Asc(strA2) - 1) & xRows & ")"
               '計算比例
               For intI = Asc(strA1) To Asc(strA2) - 1
                   If Val(wks623.Range(Chr(intI) & xRows).Value) <> 0 Then
                      wks623.Range(Chr(intI) & xRows + 1).Formula = "=" & Chr(intI) & xRows & "/" & strA2 & xRows
                      wks623.Range(Chr(intI) & xRows + 1).NumberFormatLocal = "##0.00%"
                   End If
               Next
               xRows = xRows + 1
            End If
            xRows = xRows + 1
            Select Case m_rs.Fields("na00")
               Case "A": strExc(1) = "國內"
               Case "B": strExc(1) = "大陸"
               Case "C": strExc(1) = "國外"
               Case Else: strExc(1) = ""
            End Select
            wks623.Range("a" & xRows).Value = strExc(1)
            wks623.Range("a" & xRows + 1).Value = "比例"
         End If
         Select Case Left(m_rs.Fields("cp12"), 1)
            Case "S"
                  inR1 = InStr(strAd, m_rs.Fields("cp12"))
                  If inR1 > 0 Then
                      wks623.Range(Mid(strAd, inR1 + 3, 1) & xRows).Value = Trim(Val(wks623.Range(Mid(strAd, inR1 + 3, 1) & xRows).Value) + m_rs.Fields("cnt"))
                  Else
                     GoTo JumpOtherDept
                  End If
                  wks623.Range(Mid(strAd, InStr(strAd, "SXX") + 3, 1) & xRows).Value = Trim(Val(wks623.Range(Mid(strAd, InStr(strAd, "SXX") + 3, 1) & xRows).Value) + m_rs.Fields("cnt"))
            Case "P"
                  If Left(m_rs.Fields("cp12"), 2) = "P2" Then
                     wks623.Range(Mid(strAd, InStr(strAd, "P2X") + 3, 1) & xRows).Value = Trim(Val(wks623.Range(Mid(strAd, InStr(strAd, "P2X") + 3, 1) & xRows).Value) + m_rs.Fields("cnt"))
                  Else
                     GoTo JumpOtherDept
                  End If
            Case "F"
                  If Left(m_rs.Fields("cp12"), 2) = "F1" Then
                     wks623.Range(Mid(strAd, InStr(strAd, "F1X") + 3, 1) & xRows).Value = Trim(Val(wks623.Range(Mid(strAd, InStr(strAd, "F1X") + 3, 1) & xRows).Value) + m_rs.Fields("cnt"))
                  Else
                     GoTo JumpOtherDept
                  End If
            Case Else
JumpOtherDept:
                  wks623.Range(Mid(strAd, InStr(strAd, "OXX") + 3, 1) & xRows).Value = Trim(Val(wks623.Range(Mid(strAd, InStr(strAd, "OXX") + 3, 1) & xRows).Value) + m_rs.Fields("cnt"))

         End Select
         
         strTemp = m_rs.Fields("na00")
         m_rs.MoveNext
      Loop
      
      '小計
      wks623.Range(strA2 & xRows).Formula = "=sum($" & strA1 & xRows & ":$" & Chr(Asc(strA2) - 1) & xRows & ")"
      '計算比例
      For intI = Asc(strA1) To Asc(strA2) - 1
          If Val(wks623.Range(Chr(intI) & xRows).Value) <> 0 Then
            wks623.Range(Chr(intI) & xRows + 1).Formula = "=" & Chr(intI) & xRows & "/" & strA2 & xRows
            wks623.Range(Chr(intI) & xRows + 1).NumberFormatLocal = "##0.00%"
          End If
      Next
      
      '合計
      xRows = xRows + 2
      wks623.Range("a" & xRows).Value = "合計"
      For intI = Asc("B") To Asc(strA2)
         wks623.Range(Chr(intI) & xRows).Formula = "=" & Chr(intI) & strR & "+" & Chr(intI) & strR + 2 & "+" & Chr(intI) & strR + 4
         wks623.Range(Chr(intI) & xRows).NumberFormatLocal = "###"
      Next
      
      wks623.Range("a" & xRows + 1).Value = "比例"

      '計算比例
      For intI = Asc(strA1) To Asc(strA2) - 1
          If Val(wks623.Range(Chr(intI) & xRows).Value) <> 0 Then
            wks623.Range(Chr(intI) & xRows + 1).Formula = "=" & Chr(intI) & xRows & "/" & strA2 & xRows
            wks623.Range(Chr(intI) & xRows + 1).NumberFormatLocal = "##0.00%"
          End If
      Next
      xRows = xRows + 1
      wks623.Range("c" & xRows).Select
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
   Label3.Caption = Label3 & strExcelPathN 'Modify by Amy 2021/06/21
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set frm030623 = Nothing
End Sub

Private Sub txtDate_GotFocus(Index As Integer)
   InverseTextBox txtDate(Index)
   CloseIme
End Sub

Private Sub txtDate_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub txtDate_Validate(Index As Integer, Cancel As Boolean)
Dim rsQuery As ADODB.Recordset
Dim stSQL As String, intR As Integer
Dim strVol1 As String, strVol2 As String
Dim intCnt As Integer, intC2 As Integer
   
   If txtDate(Index) <> "" Then
      If ChkDate(txtDate(Index) & "01") = False Then
          txtDate_GotFocus Index
          Cancel = True
          Exit Sub
      End If
      '檢查資料是否已存在,每月有2期公報
      Call Pub_ChgDateToTMBM07(txtDate(Index), strVol1, strVol2)
      stSQL = "select tmbm07 from tmbulletin where tmbm07>='" & strVol1 & "' and tmbm07<='" & strVol2 & "' group by tmbm07"
      intR = 1
      Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
      If intR = 1 Then
         intCnt = Val("" & rsQuery.Fields(0))
      End If
      intC2 = rsQuery.RecordCount
      rsQuery.Close
      
      If intCnt = 0 Then
         MsgBox txtDate(Index) & "此月份尚無公報資料!!"
         txtDate_GotFocus (Index)
         Cancel = True
         Exit Sub
      ElseIf intC2 < 2 Then
         MsgBox txtDate(Index) & "此月份公報資料尚不足!!"
         txtDate_GotFocus (Index)
         Cancel = True
         Exit Sub
      End If
   End If
   
   Set rsQuery = Nothing
End Sub
