VERSION 5.00
Begin VB.Form frm030622 
   BorderStyle     =   1  '單線固定
   Caption         =   "三部門案件來源比較"
   ClientHeight    =   2355
   ClientLeft      =   2790
   ClientTop       =   3945
   ClientWidth     =   4650
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   4650
   Begin VB.TextBox txtDate 
      Height          =   345
      Index           =   5
      Left            =   2880
      MaxLength       =   5
      TabIndex        =   5
      Top             =   1600
      Width           =   795
   End
   Begin VB.TextBox txtDate 
      Height          =   345
      Index           =   4
      Left            =   1890
      MaxLength       =   5
      TabIndex        =   4
      Top             =   1600
      Width           =   795
   End
   Begin VB.TextBox txtDate 
      Height          =   345
      Index           =   3
      Left            =   2880
      MaxLength       =   5
      TabIndex        =   3
      Top             =   1160
      Width           =   795
   End
   Begin VB.TextBox txtDate 
      Height          =   345
      Index           =   2
      Left            =   1890
      MaxLength       =   5
      TabIndex        =   2
      Top             =   1160
      Width           =   795
   End
   Begin VB.TextBox txtDate 
      Height          =   345
      Index           =   1
      Left            =   2880
      MaxLength       =   5
      TabIndex        =   1
      Top             =   720
      Width           =   795
   End
   Begin VB.TextBox txtDate 
      Height          =   345
      Index           =   0
      Left            =   1890
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
      TabIndex        =   7
      Top             =   90
      Width           =   756
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Excel(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2700
      TabIndex        =   6
      Top             =   90
      Width           =   990
   End
   Begin VB.Label Label4 
      Caption         =   "PS：Excel儲存於"
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   570
      TabIndex        =   11
      Top             =   2040
      Width           =   3195
   End
   Begin VB.Label Label3 
      Caption         =   "統計期間   三："
      Height          =   210
      Left            =   600
      TabIndex        =   10
      Top             =   1667
      Width           =   1245
   End
   Begin VB.Line Line3 
      X1              =   2520
      X2              =   3210
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label Label2 
      Caption         =   "統計期間   二："
      Height          =   210
      Left            =   600
      TabIndex        =   9
      Top             =   1227
      Width           =   1245
   End
   Begin VB.Line Line2 
      X1              =   2490
      X2              =   3180
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line1 
      X1              =   2460
      X2              =   3150
      Y1              =   900
      Y2              =   900
   End
   Begin VB.Label Label1 
      Caption         =   "比較基礎期間："
      Height          =   210
      Left            =   570
      TabIndex        =   8
      Top             =   810
      Width           =   1365
   End
End
Attribute VB_Name = "frm030622"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create By Lydia 2015/12/10 三部門案件來源比較
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
         If txtDate(2) & txtDate(3) <> "" Then
            For intI = 2 To 3
                If Trim(txtDate(intI)) = "" Then
                   MsgBox "統計期間 二 " & IIf(intI / 2 = 0, "起始", "截止") & "公報年月不可空白！", vbInformation, "輸入錯誤！"
                   txtDate(intI).SetFocus
                   Exit Sub
                End If
                txtDate_Validate intI, Cancel
                If Cancel = True Then
                   txtDate(intI).SetFocus
                End If
            Next
            If Val(txtDate(3)) < Val(txtDate(2)) Then
               MsgBox "統計期間 二 截止年月必須大於起始年月！", vbInformation, "輸入錯誤！"
               txtDate(1).SetFocus
               Exit Sub
            End If
         End If
         If txtDate(4) & txtDate(5) <> "" Then
            For intI = 4 To 5
                If Trim(txtDate(intI)) = "" Then
                   MsgBox "統計期間 三 " & IIf(intI / 2 = 0, "起始", "截止") & "公報年月不可空白！", vbInformation, "輸入錯誤！"
                   txtDate(intI).SetFocus
                   Exit Sub
                End If
                txtDate_Validate intI, Cancel
                If Cancel = True Then
                   txtDate(intI).SetFocus
                End If
            Next
            If Val(txtDate(5)) < Val(txtDate(4)) Then
               MsgBox "統計期間 三 截止年月必須大於起始年月！", vbInformation, "輸入錯誤！"
               txtDate(1).SetFocus
               Exit Sub
            End If
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
Dim inR1 As Integer, inXr As Integer
Dim m_rs As New ADODB.Recordset
Dim strVol1_S As String, strVol2_S As String
Dim strVol1_E As String, strVol2_E As String
Dim xlsSalesPoint As New Excel.Application
Dim wks622 As New Worksheet
Dim xRows As Integer '目前列位置
Dim strR As Integer
Dim strAd As String
Dim tmpArr As Variant
Dim strTemp As String
Dim strPath As String, strTempFile As String
Dim strA1 As String, strA2 As String
Dim cRange  As Integer
Dim cCol(1 To 3) As Integer '期間的起始欄位

   StrMenu = False
  
   'Added By Lydia 2022/01/13 查詢印表記錄檔欄位
   ClearQueryLog (Me.Name)
   pub_QL05 = pub_QL05 & ";" & Label1.Caption & txtDate(0) & "-" & txtDate(1)
   If txtDate(2) & txtDate(3) <> "" Then pub_QL05 = pub_QL05 & ";" & Label2.Caption & txtDate(2) & "-" & txtDate(3)
   If txtDate(4) & txtDate(5) <> "" Then pub_QL05 = pub_QL05 & ";" & Label3.Caption & txtDate(4) & "-" & txtDate(5)
  'end 2022/01/13
  
   strExc(1) = ""
   '3個期間
   If txtDate(2) <> "" And txtDate(4) <> "" Then
       cRange = 3
        For intI = 2 To 4 Step 2
            Call Pub_ChgDateToTMBM07(txtDate(intI), strVol1_S, strVol2_S)
            Call Pub_ChgDateToTMBM07(txtDate(intI + 1), strVol1_E, strVol2_E)
            'Modified by Lydia 2018/04/25 'T'=TA01 -> 'T'=TA01(+)
            'Memo by Lydia 2021/01/11 這段修改應該源自於2018年幫"葉特助抓商標公報統計資料->調整商標公報國外地區的代理人(TMBM06)空白也要抓到資料"；
            'Modified by Lydia 2021/01/11 'T'=TA01(+) -> 'T'=TA01 ;  遇到針對出名代理人(事務所)會造成SQL執行時間過長，所以對出名代理人(事務所)的查詢一定要抓到TMBM06。
            strExc(1) = strExc(1) & " UNION SELECT '" & IIf(intI > 2, "3", "2") & "' DTYPE,DECODE(SUBSTR(CP12,1,1),'S','01',DECODE(SUBSTR(CP12,1,2),'P2','02','F1','03','10')) ord1," & _
                 "SUBSTR(NA02,1,1) NA00,COUNT(TMBM01) VC,SUM(COUNTING(TMBM08)) CNT FROM TMBULLETIN, TAGENT, NATION,TRADEMARK,CASEPROGRESS " & _
                 "WHERE TMBM05=NA03(+) AND LENGTH(NA01)=3 AND TMBM06=TA03(+) AND 'T'=TA01 AND TMBM07>=" & strVol1_S & " AND TMBM07<=" & strVol2_E & _
                 " AND TMBM06 ='林晉章' AND TMBM01=TM15(+) AND TM15 IS NOT NULL AND TM16='1' AND TM01=CP01(+) AND TM02=CP02(+) AND TM03=CP03(+) AND TM04=CP04(+) AND CP10='101' " & _
                 "GROUP BY DECODE(SUBSTR(CP12,1,1),'S','01',DECODE(SUBSTR(CP12,1,2),'P2','02','F1','03','10')) ,SUBSTR(NA02,1,1) "
        Next intI
   '2個期間
   ElseIf txtDate(2) <> "" Or txtDate(4) <> "" Then
       cRange = 2
       inR1 = 2
       If txtDate(4) <> "" Then inR1 = 4
       Call Pub_ChgDateToTMBM07(txtDate(inR1), strVol1_S, strVol2_S)
       Call Pub_ChgDateToTMBM07(txtDate(inR1 + 1), strVol1_E, strVol2_E)
        'Modified by Lydia 2018/04/25 'T'=TA01 -> 'T'=TA01(+)
        'Modified by Lydia 2021/01/11 'T'=TA01(+) -> 'T'=TA01 ;
       strExc(1) = strExc(1) & " UNION SELECT '2' DTYPE,DECODE(SUBSTR(CP12,1,1),'S','01',DECODE(SUBSTR(CP12,1,2),'P2','02','F1','03','10')) ord1," & _
            "SUBSTR(NA02,1,1) NA00,COUNT(TMBM01) VC,SUM(COUNTING(TMBM08)) CNT FROM TMBULLETIN, TAGENT, NATION,TRADEMARK,CASEPROGRESS " & _
            "WHERE TMBM05=NA03(+) AND LENGTH(NA01)=3 AND TMBM06=TA03(+) AND 'T'=TA01 AND TMBM07>=" & strVol1_S & " AND TMBM07<=" & strVol2_E & _
            " AND TMBM06 ='林晉章' AND TMBM01=TM15(+) AND TM15 IS NOT NULL AND TM16='1' AND TM01=CP01(+) AND TM02=CP02(+) AND TM03=CP03(+) AND TM04=CP04(+) AND CP10='101' " & _
            "GROUP BY DECODE(SUBSTR(CP12,1,1),'S','01',DECODE(SUBSTR(CP12,1,2),'P2','02','F1','03','10')) ,SUBSTR(NA02,1,1) "
   Else
       cRange = 1
   End If
   Call Pub_ChgDateToTMBM07(txtDate(0), strVol1_S, strVol2_S)
   Call Pub_ChgDateToTMBM07(txtDate(1), strVol1_E, strVol2_E)
    'Modified by Lydia 2018/04/25 'T'=TA01 -> 'T'=TA01(+)
    'Modified by Lydia 2021/01/11 'T'=TA01(+) -> 'T'=TA01 ;
   strExc(1) = "SELECT '1' DTYPE,DECODE(SUBSTR(CP12,1,1),'S','01',DECODE(SUBSTR(CP12,1,2),'P2','02','F1','03','10')) ord1," & _
            "SUBSTR(NA02,1,1) NA00,COUNT(TMBM01) VC,SUM(COUNTING(TMBM08)) CNT FROM TMBULLETIN, TAGENT, NATION,TRADEMARK,CASEPROGRESS " & _
            "WHERE TMBM05=NA03(+) AND LENGTH(NA01)=3 AND TMBM06=TA03(+) AND 'T'=TA01 AND TMBM07>=" & strVol1_S & " AND TMBM07<=" & strVol2_E & _
            " AND TMBM06 ='林晉章' AND TMBM01=TM15(+) AND TM15 IS NOT NULL AND TM16='1' AND TM01=CP01(+) AND TM02=CP02(+) AND TM03=CP03(+) AND TM04=CP04(+) AND CP10='101' " & _
            "GROUP BY DECODE(SUBSTR(CP12,1,1),'S','01',DECODE(SUBSTR(CP12,1,2),'P2','02','F1','03','10')) ,SUBSTR(NA02,1,1) " & strExc(1)
   strSql = strExc(1) & " ORDER BY 1,2,3"

   If m_rs.State = 1 Then m_rs.Close
   m_rs.CursorLocation = adUseClient
   m_rs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If Not m_rs.EOF And Not m_rs.BOF Then
      StrMenu = True
      InsertQueryLog (m_rs.RecordCount)  'Added by Lydia 2022/01/13
      '統計期間
      cCol(1) = Asc("b")
      cCol(2) = cCol(1) + 2 '基礎期間只有件,類
      cCol(3) = cCol(2) + 4 '統計期間多出成長(以類計),成長率
      
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
      
      xlsSalesPoint.SheetsInNewWorkbook = 3 'Modify by Amy 2021/06/21 Added by Lydia 2019/03/13 預設工作表數量
      xlsSalesPoint.Workbooks.add
      Set wks622 = xlsSalesPoint.Worksheets(1)
      wks622.PageSetup.Orientation = xlPortrait '直印
      '抬頭
       wks622.PageSetup.PrintTitleRows = "$1:$3"
       wks622.Columns("a:a").ColumnWidth = 14 '部門
       For intI = cCol(1) To cCol(cRange) + IIf(cRange > 1, 3, 1)
          wks622.Columns(Chr(intI) & ":" & Chr(intI)).ColumnWidth = 8
       Next

       wks622.Range("a1").Value = Me.Caption
       wks622.Range("a1:" & Chr(cCol(cRange) + IIf(cRange > 1, 3, 1)) & "1").MergeCells = True

      wks622.Range("a2").Value = "比較基礎期間 " & Mid(ChangeTStringToTDateString(txtDate(0) & "01"), 1, 6) & "至" & Mid(ChangeTStringToTDateString(txtDate(1) & "01"), 1, 6)
      wks622.Range("a2:" & Chr(cCol(2) - 1) & "2").MergeCells = True
      If cRange >= 2 Then
         wks622.Range(Chr(cCol(2)) & "2").Value = "統計期間 二 " & Mid(ChangeTStringToTDateString(txtDate(2) & "01"), 1, 6) & "至" & Mid(ChangeTStringToTDateString(txtDate(3) & "01"), 1, 6)
         wks622.Range(Chr(cCol(2)) & "2:" & Chr(cCol(3) - 1) & "2").MergeCells = True
      End If
      If cRange >= 3 Then
         wks622.Range(Chr(cCol(3)) & "2").Value = "統計期間 三 " & Mid(ChangeTStringToTDateString(txtDate(4) & "01"), 1, 6) & "至" & Mid(ChangeTStringToTDateString(txtDate(5) & "01"), 1, 6)
         wks622.Range(Chr(cCol(3)) & "2:" & Chr(cCol(3) + 3) & "2").MergeCells = True
      End If
      wks622.Range("a3").Value = "部門\項目"
      For intI = 1 To cRange
          wks622.Range(Chr(cCol(intI)) & "3").Value = "件"
          wks622.Range(Chr(cCol(intI) + 1) & "3").Value = "類"
          If intI > 1 Then
             wks622.Range(Chr(cCol(intI) + 2) & "3").Value = "成長(類)"
             wks622.Range(Chr(cCol(intI) + 3) & "3").Value = "成長率"
          End If
      Next
           
      wks622.Range("a1:" & Chr(cCol(cRange) + IIf(cRange > 1, 3, 1)) & "3").HorizontalAlignment = xlCenter
      wks622.Range("a1:" & Chr(cCol(cRange) + IIf(cRange > 1, 3, 1)) & "3").VerticalAlignment = xlBottom
      
      xRows = 4:      strR = xRows
      '總計=>置頂
      wks622.Range("a" & strR).Value = "總    計"
      
     
      '各部門列位置
      '依序為智權部(SXX=>01),商標處(P2X=>02),外商(F1X=>03),其他(OXX=>10)
      'A=國內,B=大陸,C=國外,T=成長(以類計),P=成長率
      strAd = "01A01,01B02,01C03,01T04,01P05,02A06,02B07,02C08,02T09,02P10,03A11,03B12,03C13,03T14,03P15,10A16,10B17,10C18,10T19,10P20"
      tmpArr = Empty
      tmpArr = Split(strAd, ",")
      For intI = 0 To UBound(tmpArr)
          If Len(tmpArr(intI)) = 5 Then
             strExc(1) = Mid(tmpArr(intI), 1, 2)
             strExc(2) = Mid(tmpArr(intI), 3, 1)
             wks622.Range("a" & Trim(strR + Val(Mid(tmpArr(intI), 4, 2)))).Value = GetRowTitle(strExc(1), strExc(2))
             If strExc(2) = "T" Then
                For inR1 = 1 To cRange
                   '小計sum()
                   wks622.Range(Trim(Chr(cCol(inR1))) & Trim(strR + Val(Mid(tmpArr(intI), 4, 2)))).Formula = "=sum(" & Trim(Chr(cCol(inR1))) & Trim(strR + Val(Mid(tmpArr(intI), 4, 2)) - 3) & ":" & Trim(Chr(cCol(inR1))) & Trim(strR + Val(Mid(tmpArr(intI), 4, 2)) - 1) & ")"
                   wks622.Range(Trim(Chr(cCol(inR1) + 1)) & Trim(strR + Val(Mid(tmpArr(intI), 4, 2)))).Formula = "=sum(" & Trim(Chr(cCol(inR1) + 1)) & Trim(strR + Val(Mid(tmpArr(intI), 4, 2)) - 3) & ":" & Trim(Chr(cCol(inR1) + 1)) & Trim(strR + Val(Mid(tmpArr(intI), 4, 2)) - 1) & ")"
                   wks622.Range(Trim(Chr(cCol(inR1))) & Trim(strR + Val(Mid(tmpArr(intI), 4, 2)))).NumberFormat = "###"
                   wks622.Range(Trim(Chr(cCol(inR1) + 1)) & Trim(strR + Val(Mid(tmpArr(intI), 4, 2)))).NumberFormat = "###"
                Next inR1
             End If
          End If
      Next intI

      '先將數值填入excel
      Do While Not m_rs.EOF
         '抓資料的列位置
         inR1 = InStr(strAd, Trim(m_rs.Fields("ord1")) & Trim(m_rs.Fields("na00")))
         If inR1 > 0 Then
            inXr = strR + Val(Mid(strAd, inR1 + 3, 2))
            '寫入件數,類別數
            wks622.Range(Chr(cCol(Val(m_rs.Fields("dtype")))) & inXr).Value = Trim(m_rs.Fields("vc"))
            wks622.Range(Chr(cCol(Val(m_rs.Fields("dtype"))) + 1) & inXr).Value = Trim(m_rs.Fields("cnt"))
         End If
         m_rs.MoveNext
      Loop
      

      '總計和成長率
      For intI = 1 To cRange
          strExc(3) = "": strExc(4) = ""
          For inR1 = 0 To UBound(tmpArr)
              If Len(tmpArr(inR1)) = 5 Then
                 Select Case Mid(tmpArr(inR1), 3, 1)
                     Case "T"
                         '總計(暫存)
                         strExc(3) = strExc(3) & "+" & Trim(Chr(cCol(intI))) & Trim(strR + Val(Mid(tmpArr(inR1), 4, 2)))
                         strExc(4) = strExc(4) & "+" & Trim(Chr(cCol(intI) + 1)) & Trim(strR + Val(Mid(tmpArr(inR1), 4, 2)))
                         GoTo JumpSubTotal
                     Case "P"
                     '=IF(B4>0,B8/B4,0)
                         wks622.Range(Trim(Chr(cCol(intI))) & Trim(strR + Val(Mid(tmpArr(inR1), 4, 2)))).Formula = "=IF(" & Trim(Chr(cCol(intI))) & strR & " > 0, " & Trim(Chr(cCol(intI))) & Trim(strR + Val(Mid(tmpArr(inR1), 4, 2) - 1)) & "/" & Trim(Chr(cCol(intI))) & strR & ", 0)"
                         wks622.Range(Trim(Chr(cCol(intI) + 1)) & Trim(strR + Val(Mid(tmpArr(inR1), 4, 2)))).Formula = "=IF(" & Trim(Chr(cCol(intI) + 1)) & strR & " > 0, " & Trim(Chr(cCol(intI) + 1)) & Trim(strR + Val(Mid(tmpArr(inR1), 4, 2) - 1)) & "/" & Trim(Chr(cCol(intI) + 1)) & strR & ", 0)"
                         wks622.Range(Trim(Chr(cCol(intI))) & Trim(strR + Val(Mid(tmpArr(inR1), 4, 2)))).NumberFormatLocal = "###.00%"
                         wks622.Range(Trim(Chr(cCol(intI) + 1)) & Trim(strR + Val(Mid(tmpArr(inR1), 4, 2)))).NumberFormatLocal = "###.00%"
                     Case Else '國內,大陸,國外=>成長類,成長率
JumpSubTotal:
                         If intI > 1 Then
                            '=IF(E5>0,F5/E5,0)
                            wks622.Range(Trim(Chr(cCol(intI) + 2)) & Trim(strR + Val(Mid(tmpArr(inR1), 4, 2)))).Formula = "=" & Trim(Chr(cCol(1) + 1)) & Trim(strR + Val(Mid(tmpArr(inR1), 4, 2))) & "-" & Trim(Chr(cCol(intI) + 1)) & Trim(strR + Val(Mid(tmpArr(inR1), 4, 2)))
                            wks622.Range(Trim(Chr(cCol(intI) + 3)) & Trim(strR + Val(Mid(tmpArr(inR1), 4, 2)))).Formula = "=IF(" & Trim(Chr(cCol(intI) + 1)) & Trim(strR + Val(Mid(tmpArr(inR1), 4, 2))) & " > 0, " & Trim(Chr(cCol(intI) + 2)) & Trim(strR + Val(Mid(tmpArr(inR1), 4, 2))) & "/" & Trim(Chr(cCol(intI) + 1)) & Trim(strR + Val(Mid(tmpArr(inR1), 4, 2))) & ", 0)"
                            wks622.Range(Trim(Chr(cCol(intI) + 2)) & Trim(strR + Val(Mid(tmpArr(inR1), 4, 2)))).NumberFormatLocal = "###"
                            'If wks622.Range(Trim(Chr(cCol(intI) + 1)) & Trim(strR + Val(Mid(TmpArr(inR1), 4, 2)))).Value > 0 Then
                               wks622.Range(Trim(Chr(cCol(intI) + 3)) & Trim(strR + Val(Mid(tmpArr(inR1), 4, 2)))).NumberFormatLocal = "###.00%"
                            'End If
                         End If
                 End Select
              End If
          Next inR1
          wks622.Range(Trim(Chr(cCol(intI))) & strR).Formula = "=" & Mid(strExc(3), 2)
          wks622.Range(Trim(Chr(cCol(intI) + 1)) & strR).Formula = "=" & Mid(strExc(4), 2)
          If intI > 1 Then
              wks622.Range(Trim(Chr(cCol(intI) + 2)) & strR).Formula = "=" & Trim(Chr(cCol(1) + 1)) & strR & "-" & Trim(Chr(cCol(intI) + 1)) & strR
              wks622.Range(Trim(Chr(cCol(intI) + 3)) & strR).Formula = "=IF(" & Trim(Chr(cCol(intI) + 1)) & strR & " > 0, " & Trim(Chr(cCol(intI) + 2)) & strR & "/" & Trim(Chr(cCol(intI) + 1)) & strR & ", 0)"
              wks622.Range(Trim(Chr(cCol(intI) + 3)) & strR).NumberFormatLocal = "###.00%"
          End If
      Next intI
      
      
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
      
   Else
      MsgBox "查詢無資料！", vbExclamation + vbOKOnly
      InsertQueryLog (0) 'Added by Lydia 2022/01/13
      Exit Function
   End If
   
   Exit Function
   
End Function

Private Sub Form_Load()
   MoveFormToCenter Me
   
   txtDate(0) = Left(strSrvDate(2), 5)
   txtDate(1) = Left(strSrvDate(2), 5)
   Label4.Caption = Label4 & strExcelPathN 'Modify by Amy 2021/06/21
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm030622 = Nothing
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
      
      '期間不可重疊
      If Index > 1 And txtDate(0) <> "" And txtDate(1) <> "" Then
         If Val(txtDate(Index)) >= Val(txtDate(0)) And Val(txtDate(Index)) <= Val(txtDate(1)) Then
            MsgBox txtDate(Index) & "此月份與基礎期間重疊!!"
            txtDate_GotFocus (Index)
            Cancel = True
            Exit Sub
         End If
      End If
   End If
   
   Set rsQuery = Nothing
End Sub
Private Function GetRowTitle(ByVal pStr01 As String, Optional pStr02 As String) As String
Dim midStr As String

   GetRowTitle = ""
   Select Case pStr01
      Case "01": midStr = "智權部 "
      Case "02": midStr = "商標處 "
      Case "03": midStr = "外商 "
      Case "10": midStr = "其他 "
      Case "20": midStr = "總計 "
   End Select
   
   GetRowTitle = midStr
   If pStr01 <> "" Then
      midStr = ""
      Select Case pStr02
         Case "A": midStr = "國內"
         Case "B": midStr = "大陸"
         Case "C": midStr = "國外"
         Case "T": midStr = "小計"
         Case "P": midStr = "占有率"
      End Select
   End If
   
   If midStr <> "" Then GetRowTitle = GetRowTitle & midStr
   
End Function
