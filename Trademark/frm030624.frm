VERSION 5.00
Begin VB.Form frm030624 
   BorderStyle     =   1  '單線固定
   Caption         =   "同業案件來源比較"
   ClientHeight    =   4500
   ClientLeft      =   2790
   ClientTop       =   3945
   ClientWidth     =   4650
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   4650
   Begin VB.TextBox txt1 
      Height          =   345
      Index           =   1
      Left            =   1680
      TabIndex        =   7
      Top             =   3800
      Width           =   2175
   End
   Begin VB.TextBox txt1 
      Height          =   345
      Index           =   0
      Left            =   1680
      TabIndex        =   6
      Top             =   3360
      Width           =   2175
   End
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
      TabIndex        =   9
      Top             =   90
      Width           =   756
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Excel(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2700
      TabIndex        =   8
      Top             =   90
      Width           =   990
   End
   Begin VB.Label Label5 
      Caption         =   "PS：Excel儲存於"
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   720
      TabIndex        =   20
      Top             =   4200
      Width           =   3195
   End
   Begin VB.Label Label4 
      Caption         =   "6."
      Height          =   255
      Index           =   6
      Left            =   1440
      TabIndex        =   19
      Top             =   3845
      Width           =   795
   End
   Begin VB.Label Label4 
      Caption         =   "5."
      Height          =   255
      Index           =   5
      Left            =   1440
      TabIndex        =   18
      Top             =   3405
      Width           =   795
   End
   Begin VB.Label Label4 
      Caption         =   "4.台灣國際"
      Height          =   255
      Index           =   4
      Left            =   1440
      TabIndex        =   17
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "3.理律法律"
      Height          =   255
      Index           =   3
      Left            =   1440
      TabIndex        =   16
      Top             =   2840
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "2.聖島國際"
      Height          =   255
      Index           =   2
      Left            =   1440
      TabIndex        =   15
      Top             =   2560
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "1.台一國際"
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   14
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "事務所或代理人名稱："
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   13
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "統計期間   三："
      Height          =   210
      Left            =   600
      TabIndex        =   12
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
      TabIndex        =   11
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
      TabIndex        =   10
      Top             =   810
      Width           =   1365
   End
End
Attribute VB_Name = "frm030624"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create By Lydia 2015/12/14 同業案件來源比較
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
         'Added by Lydia 2018/02/27 必要輸入統計期間
         If Trim(txtDate(2) & txtDate(3) & txtDate(4) & txtDate(5)) = "" Or (Trim(txtDate(2) & txtDate(3)) = "" And Trim(txtDate(4) & txtDate(5)) <> "") Then
            MsgBox "請輸入統計期間 二 ！", vbInformation, "輸入錯誤！"
            txtDate(2).SetFocus
            Exit Sub
         End If
         'end 2018/02/27
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
Dim inR1 As Integer, inXr As Integer
Dim m_rs As New ADODB.Recordset
Dim strVol1_S As String, strVol2_S As String
Dim strVol1_E As String, strVol2_E As String
Dim xlsSalesPoint As New Excel.Application
Dim wks624 As New Worksheet
Dim xRows As Integer '目前列位置
Dim strR As Integer
Dim strAd As String
Dim tmpArr As Variant
Dim strTemp As String
Dim strPath As String, strTempFile As String
Dim cRange  As Integer
Dim cCol(1 To 3) As Integer '期間的起始欄位

   StrMenu = False

      strExc(1) = ""
     
   'Added By Lydia 2022/01/13 查詢印表記錄檔欄位
   ClearQueryLog (Me.Name)
   pub_QL05 = pub_QL05 & ";" & Label1.Caption & txtDate(0) & "-" & txtDate(1)
   If txtDate(2) & txtDate(3) <> "" Then pub_QL05 = pub_QL05 & ";" & Label2.Caption & txtDate(2) & "-" & txtDate(3)
   If txtDate(4) & txtDate(5) <> "" Then pub_QL05 = pub_QL05 & ";" & Label3.Caption & txtDate(4) & "-" & txtDate(5)
   pub_QL05 = pub_QL05 & ";" & Label4(0).Caption & Label4(1).Caption & "、" & Label4(2).Caption & "、" & Label4(3).Caption & "、" & Label4(4).Caption
   If txt1(0) <> "" Then pub_QL05 = pub_QL05 & "、5." & txt1(0)
   If txt1(1) <> "" Then pub_QL05 = pub_QL05 & "、6." & txt1(1)
  'end 2022/01/13
  
   Call Pub_ChgDateToTMBM07(txtDate(0), strVol1_S, strVol2_S)
   Call Pub_ChgDateToTMBM07(txtDate(1), strVol1_E, strVol2_E)
      '自訂同業5,6
      If txt1(0) <> "" Then
           'Modified by Lydia 2018/04/25 'T'=TA01 -> 'T'=TA01(+)
           'Memo by Lydia 2021/01/11 這段修改應該源自於2018年幫"葉特助抓商標公報統計資料->調整商標公報國外地區的代理人(TMBM06)空白也要抓到資料"；
           'Modified by Lydia 2021/01/11 'T'=TA01(+) -> 'T'=TA01 ;  遇到針對出名代理人(事務所)會造成SQL執行時間過長，所以對出名代理人(事務所)的查詢一定要抓到TMBM06。
           strExc(1) = strExc(1) & " union SELECT '50' ord1, '10' d2d,TMBM01,DECODE(TA04,NULL,TMBM06,TA04) AS TA04,NA01,NA02,TMBM08 FROM TMBULLETIN, TAGENT, NATION" & _
                    " WHERE TMBM05=NA03(+) AND length(na01)=3 AND TMBM06=TA03(+) AND 'T'=TA01 and tmbm07>=" & strVol1_S & " And tmbm07<=" & strVol2_E & _
                    " and ta04 in (" & Pub_GetTA04(txt1(0)) & ") "
      End If
      If txt1(1) <> "" Then
           'Modified by Lydia 2018/04/25 'T'=TA01 -> 'T'=TA01(+)
           'Modified by Lydia 2021/01/11 'T'=TA01(+) -> 'T'=TA01 ;
           strExc(1) = strExc(1) & " union SELECT '60' ord1, '10' d2d,TMBM01,DECODE(TA04,NULL,TMBM06,TA04) AS TA04,NA01,NA02,TMBM08 FROM TMBULLETIN, TAGENT, NATION" & _
                    " WHERE TMBM05=NA03(+) AND length(na01)=3 AND TMBM06=TA03(+) AND 'T'=TA01 and tmbm07>=" & strVol1_S & " And tmbm07<=" & strVol2_E & _
                    " and ta04 in (" & Pub_GetTA04(txt1(1)) & ") "
      End If
   'Modified by Lydia 2018/04/25 'T'=TA01 -> 'T'=TA01(+)
   'Modified by Lydia 2021/01/11 'T'=TA01(+) -> 'T'=TA01 ;
   strExc(1) = "SELECT '10' ord1, '10' d2d,TMBM01,DECODE(TA04,NULL,TMBM06,TA04) AS TA04,NA01,NA02,TMBM08 " & _
               "From TMBULLETIN, TAGENT, NATION,trademark,caseprogress " & _
               "WHERE TMBM05=NA03(+) AND length(na01)=3 AND TMBM06=TA03(+) AND 'T'=TA01 and tmbm07>=" & strVol1_S & " And tmbm07<=" & strVol2_E & _
               " and tmbm06 ='林晉章' and tmbm01=tm15(+) and tm15 is not null and tm16='1' and tm01=cp01(+) and tm02=cp02(+) and tm03=cp03(+) and tm04=cp04(+) and cp10='101' " & _
               "union SELECT decode(ta04,'聖島國際','20','理律法律','30','台灣國際','40') ord1, '10' d2d," & _
               "TMBM01,DECODE(TA04,NULL,TMBM06,TA04) AS TA04,NA01,NA02,TMBM08 FROM TMBULLETIN, TAGENT, NATION " & _
               "WHERE TMBM05=NA03(+) AND length(na01)=3 AND TMBM06=TA03(+) AND 'T'=TA01 and tmbm07>=" & strVol1_S & " And tmbm07<=" & strVol2_E & _
               " and ta04 in ('聖島國際','理律法律','台灣國際') " & strExc(1)
   strTemp = "tmbm07>=" & strVol1_S & " And tmbm07<=" & strVol2_E
   
   '不同期間
   strExc(3) = strExc(1)
   '3個
   If txtDate(2) <> "" And txtDate(4) <> "" Then
       cRange = 3
        For intI = 2 To 4 Step 2
            Call Pub_ChgDateToTMBM07(txtDate(intI), strVol1_S, strVol2_S)
            Call Pub_ChgDateToTMBM07(txtDate(intI + 1), strVol1_E, strVol2_E)
            strExc(2) = Replace(strExc(3), "'10' d2d", "'" & IIf(intI > 2, "3", "2") & "0' as d2d")
            strExc(2) = Replace(strExc(2), strTemp, "tmbm07>=" & strVol1_S & " And tmbm07<=" & strVol2_E)
            strExc(1) = strExc(1) & " UNION " & strExc(2)
        Next
   '2個
   ElseIf txtDate(2) <> "" Or txtDate(4) <> "" Then
       cRange = 2
       inR1 = 2
       If txtDate(4) <> "" Then inR1 = 4
       Call Pub_ChgDateToTMBM07(txtDate(inR1), strVol1_S, strVol2_S)
       Call Pub_ChgDateToTMBM07(txtDate(inR1 + 1), strVol1_E, strVol2_E)
       strExc(2) = Replace(strExc(3), "'10' d2d", "'20' as d2d")
       strExc(2) = Replace(strExc(2), strTemp, "tmbm07>=" & strVol1_S & " And tmbm07<=" & strVol2_E)
       strExc(1) = strExc(1) & " UNION " & strExc(2)
   End If
   
   strSql = "select ord1,d2d,TA04,substrb(na02,1,1) NA00,sum(counting(tmbm08)) vc from (" & _
             strExc(1) & ") group by ord1,d2d,ta04,substrb(na02,1,1) order by 1,2,4"
      
   If m_rs.State = 1 Then m_rs.Close
   m_rs.CursorLocation = adUseClient
   m_rs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If Not m_rs.EOF And Not m_rs.BOF Then
      StrMenu = True
      InsertQueryLog (m_rs.RecordCount)  'Added by Lydia 2022/01/13
      '統計期間
      cCol(1) = Asc("b")
      cCol(2) = cCol(1) + 4
      cCol(3) = cCol(2) + 8 '統計期間多出成長(以類計)
      
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
      
      xlsSalesPoint.SheetsInNewWorkbook = 1 'Added by Lydia 2019/03/13 預設工作表數量
      xlsSalesPoint.Workbooks.add
      Set wks624 = xlsSalesPoint.Worksheets(1)
      wks624.PageSetup.Orientation = xlPortrait '直印
      '抬頭
       wks624.PageSetup.PrintTitleRows = "$1:$4"
       wks624.Columns("a:a").ColumnWidth = 10 '事務所
       For intI = cCol(1) To cCol(cRange) + IIf(cRange > 1, 7, 3)
          wks624.Columns(Chr(intI) & ":" & Chr(intI)).ColumnWidth = 8
       Next

       wks624.Range("a1").Value = Me.Caption
       wks624.Range("a1:" & Chr(cCol(cRange) + IIf(cRange < 2, 3, 7)) & "1").MergeCells = True

       wks624.Range("a2").Value = "(以類計)"
       wks624.Range("a2:" & Chr(cCol(cRange) + IIf(cRange < 2, 3, 7)) & "2").MergeCells = True
       
      wks624.Range(Chr(cCol(1)) & "3").Value = "比較基礎期間 " & Mid(ChangeTStringToTDateString(txtDate(0) & "01"), 1, 6) & "至" & Mid(ChangeTStringToTDateString(txtDate(1) & "01"), 1, 6)
      wks624.Range(Chr(cCol(1)) & "3:" & Chr(cCol(2) - 1) & "3").MergeCells = True
      If cRange >= 2 Then
         wks624.Range(Chr(cCol(2)) & "3").Value = "統計期間 二 " & Mid(ChangeTStringToTDateString(txtDate(2) & "01"), 1, 6) & "至" & Mid(ChangeTStringToTDateString(txtDate(3) & "01"), 1, 6)
         wks624.Range(Chr(cCol(2)) & "3:" & Chr(cCol(2) + 3) & "3").MergeCells = True
         wks624.Range(Chr(cCol(2) + 4) & "3").Value = "統計期間 二 成長"
         wks624.Range(Chr(cCol(2) + 4) & "3:" & Chr(cCol(2) + 7) & "3").MergeCells = True
      End If
      If cRange >= 3 Then
         wks624.Range(Chr(cCol(3)) & "3").Value = "統計期間 三 " & Mid(ChangeTStringToTDateString(txtDate(4) & "01"), 1, 6) & "至" & Mid(ChangeTStringToTDateString(txtDate(5) & "01"), 1, 6)
         wks624.Range(Chr(cCol(3)) & "3:" & Chr(cCol(3) + 3) & "3").MergeCells = True
         wks624.Range(Chr(cCol(3) + 4) & "3").Value = "統計期間 三 成長"
         wks624.Range(Chr(cCol(3) + 4) & "3:" & Chr(cCol(3) + 7) & "3").MergeCells = True
      End If
      wks624.Range("a4").Value = "事務所"
      
      xRows = 4:      strR = xRows
      If txt1(0) <> "" And txt1(1) <> "" Then
         inXr = 6
      ElseIf txt1(0) <> "" Or txt1(1) <> "" Then
           inXr = 5
      Else
           inXr = 4
      End If
      For intI = 1 To cRange
          wks624.Range(Chr(cCol(intI)) & "4").Value = "國內"
          wks624.Range(Chr(cCol(intI) + 1) & "4").Value = "大陸"
          wks624.Range(Chr(cCol(intI) + 2) & "4").Value = "國外"
          wks624.Range(Chr(cCol(intI) + 3) & "4").Value = "合計"
          For inR1 = 1 To inXr
             wks624.Range(Chr(cCol(intI) + 3) & strR + inR1).Formula = "=SUM(" & Chr(cCol(intI)) & strR + inR1 & ":" & Chr(cCol(intI) + 2) & strR + inR1 & ")"
          Next inR1
          If intI > 1 Then
              wks624.Range(Chr(cCol(intI) + 4) & "4").Value = "國內"
              wks624.Range(Chr(cCol(intI) + 5) & "4").Value = "大陸"
              wks624.Range(Chr(cCol(intI) + 6) & "4").Value = "國外"
              wks624.Range(Chr(cCol(intI) + 7) & "4").Value = "合計"
              For inR1 = 1 To inXr
                 wks624.Range(Chr(cCol(intI) + 4) & strR + inR1).Formula = "=" & Chr(cCol(1)) & strR + inR1 & "-" & Chr(cCol(intI)) & strR + inR1
                 wks624.Range(Chr(cCol(intI) + 5) & strR + inR1).Formula = "=" & Chr(cCol(1) + 1) & strR + inR1 & "-" & Chr(cCol(intI) + 1) & strR + inR1
                 wks624.Range(Chr(cCol(intI) + 6) & strR + inR1).Formula = "=" & Chr(cCol(1) + 2) & strR + inR1 & "-" & Chr(cCol(intI) + 2) & strR + inR1
                 wks624.Range(Chr(cCol(intI) + 7) & strR + inR1).Formula = "=SUM(" & Chr(cCol(intI) + 4) & strR + inR1 & ":" & Chr(cCol(intI) + 6) & strR + inR1 & ")"
              Next inR1
          End If
      Next intI
      
      wks624.Range("a1:" & Chr(cCol(cRange) + IIf(cRange < 2, 3, 7)) & "3").HorizontalAlignment = xlCenter
      wks624.Range("a1:" & Chr(cCol(cRange) + IIf(cRange < 2, 3, 7)) & "3").VerticalAlignment = xlBottom
      
     
      '各部門列位置
      '10=比較基礎期間,20=統計期間2,21=統計期間2的成長,30=統計期間3,31=統計期間3的成長
      'A=國內,B=大陸,C=國外,T=成長(以類計)
      strAd = "10A01,10B02,10C03,10T04,20A05,20B06,20C07,20T08,21A09,21B10,21C11,21T12,30A13,30B14,30C15,30T16,31A17,31B18,31C19,31T20"

      Do While Not m_rs.EOF
         '抓資料的明細欄位置
         inR1 = InStr(strAd, Trim(m_rs.Fields("d2d")) & Trim(m_rs.Fields("na00")))
         If inR1 > 0 Then
            If strTemp <> m_rs.Fields("ord1") Then
               xRows = xRows + 1
            End If
            inXr = cCol(1) + Val(Mid(strAd, inR1 + 3, 2)) - 1
            '事務所 抬頭
            If wks624.Range(Chr(cCol(1) - 1) & xRows).Value = "" Then
               wks624.Range(Chr(cCol(1) - 1) & xRows).Value = Trim(m_rs.Fields("ta04"))
            ElseIf Trim(m_rs.Fields("ta04")) <> "" Then
                   strExc(6) = Trim(wks624.Range(Chr(cCol(1) - 1) & xRows).Value)
                   If InStr(strExc(6), Trim(m_rs.Fields("ta04"))) = 0 Then
                      wks624.Range(Chr(cCol(1) - 1) & xRows).Value = strExc(6) & "、" & Trim(m_rs.Fields("ta04"))
                   End If
            End If

            wks624.Range(Chr(cCol(1) + Val(Mid(strAd, inR1 + 3, 2)) - 1) & xRows).Value = Trim(m_rs.Fields("vc"))

         End If
         strTemp = m_rs.Fields("ord1")
         m_rs.MoveNext
      Loop
            
      wks624.Range(Chr(cCol(1)) & strR + 1 & ":" & Chr(cCol(cRange) + IIf(cRange > 1, 7, 3)) & xRows).NumberFormat = "##0"
      
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
   End If
   Exit Function

End Function

Private Sub Form_Load()
   MoveFormToCenter Me
   
   txtDate(0) = Left(strSrvDate(2), 5)
   txtDate(1) = Left(strSrvDate(2), 5)
   Label5.Caption = Label5 & strExcelPathN 'Modify by Amy 2021/06/21
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set frm030624 = Nothing
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

