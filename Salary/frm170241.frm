VERSION 5.00
Begin VB.Form frm170241 
   BorderStyle     =   1  '單線固定
   Caption         =   "員工年度所得統計 (非申報數)"
   ClientHeight    =   3360
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5328
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   5328
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   1
      Left            =   2940
      MaxLength       =   3
      TabIndex        =   1
      Top             =   1050
      Width           =   585
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   2
      Left            =   2130
      MaxLength       =   3
      TabIndex        =   2
      Top             =   1410
      Width           =   585
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   3
      Left            =   2940
      MaxLength       =   3
      TabIndex        =   3
      Top             =   1410
      Width           =   585
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   4
      Left            =   2130
      MaxLength       =   6
      TabIndex        =   4
      Top             =   1785
      Width           =   780
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   5
      Left            =   3120
      MaxLength       =   6
      TabIndex        =   5
      Top             =   1785
      Width           =   780
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   180
      TabIndex        =   9
      Top             =   2700
      Visible         =   0   'False
      Width           =   4875
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   705
         Style           =   2  '單純下拉式
         TabIndex        =   8
         Top             =   180
         Width           =   3870
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   315
         Index           =   1
         Left            =   75
         TabIndex        =   10
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   0
      Left            =   2130
      MaxLength       =   3
      TabIndex        =   0
      Top             =   1050
      Width           =   585
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "離開(&X)"
      Height          =   435
      Index           =   1
      Left            =   4215
      TabIndex        =   7
      Top             =   120
      Width           =   915
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "產生Excel檔(&E)"
      Default         =   -1  'True
      Height          =   435
      Index           =   0
      Left            =   2700
      TabIndex        =   6
      Top             =   120
      Width           =   1455
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   2670
      X2              =   2985
      Y1              =   1170
      Y2              =   1170
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   2670
      X2              =   2985
      Y1              =   1530
      Y2              =   1530
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   2850
      X2              =   3165
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "部　　門："
      Height          =   180
      Index           =   2
      Left            =   1140
      TabIndex        =   13
      Top             =   1470
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "員工編號："
      Height          =   180
      Index           =   0
      Left            =   1140
      TabIndex        =   12
      Top             =   1830
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "年　　度："
      Height          =   180
      Left            =   1140
      TabIndex        =   11
      Top             =   1110
      Width           =   900
   End
End
Attribute VB_Name = "frm170241"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2024/1/31 新部門修改
'Create By Sindy 2021/2/3
Option Explicit

Dim xlsSalesPoint As New Excel.Application
Dim wksaccrpt114 As New Worksheet
Dim i As Integer, j As Integer
Dim strFileName As String
Dim intSheets As Integer, strSheetsVal(0 To 12) As String

Private Sub cmdok_Click(Index As Integer)
   Select Case Index
      Case 0
         If Text1(0) = "" And Text1(1) = "" And _
            Text1(2) = "" And Text1(3) = "" And _
            Text1(4) = "" And Text1(5) = "" Then
            MsgBox "請至少輸入一項查詢條件！", vbInformation, "操作錯誤！"
            Text1(0).SetFocus
            Exit Sub
         End If
         
         strExc(0) = "select count(*) from yearbonus where yb01=" & Val(Text1(1)) + 1911
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp.Fields(0) = 0 Then
               MsgBox "該年度(" & Val(Text1(1)) & ")尚無年終獎金資料！", vbExclamation + vbOKOnly
               Text1(0).SetFocus
               Exit Sub
            End If
         End If
         
         Call Pub_ChkExcelPath 'Added by Lydia 2021/07/01 檢查xls資料夾的模組
         
         Screen.MousePointer = vbHourglass
         Call ExcelSave
         Screen.MousePointer = vbDefault
      Case 1
         Unload Me
   End Select
End Sub

'*************************************************
'轉成Excel檔案
'
'*************************************************
Private Sub ExcelSave()
Dim adoRs As New ADODB.Recordset
Dim adoRs2 As New ADODB.Recordset
Dim intRow As Integer, strSort As String
Dim strYear As String
Dim strKey As String
Dim strConSM As String, strConMB As String, strConOB As String, strConYB As String
Dim intSht As Integer
Dim bolNewDep As Boolean 'Added by Morgan 2024/1/31

On Error GoTo flgErr

   'Added by Morgan 2024/1/31
   strExc(0) = Left(新部門啟用日, 4) - 1911
   If Text1(1) = "" Or Val(Text1(1)) >= Val(strExc(0)) Then
      bolNewDep = True
   Else
      bolNewDep = False
   End If
   If Text1(2) <> "" Or Text1(3) <> "" Then
      If Text1(0) = "" And Text1(1) = "" Then
         MsgBox "有輸入部門條件時，年度起訖不可皆空白！", vbCritical, "新部門檢查"
         Exit Sub
      End If
      If Val(Text1(0)) < Val(strExc(0)) And (Text1(1) = "" Or Val(Text1(1)) >= Val(strExc(0))) Then
         MsgBox "有輸入部門條件時，年度起訖不可跨" & strExc(0) & "年！", vbCritical, "新部門檢查"
         Exit Sub
      End If
   End If
   'end 2024/1/3
   
   strKey = IIf(Val(Text1(0)) > 0, "-" & Val(Text1(0)), "") + _
            IIf(Val(Text1(1)) > 0 And Val(Text1(0)) <> Val(Text1(1)), "-" & Val(Text1(1)), "") + _
            IIf(Text1(2) <> "", "-" & Text1(2), "") + _
            IIf(Text1(3) <> "" And Text1(2) <> Text1(3), "-" & Text1(3), "") + _
            IIf(Trim(Text1(4)) <> "", "-" & Trim(Text1(4)), "") + _
            IIf(Trim(Text1(5)) <> "" And Trim(Text1(4)) <> Trim(Text1(5)), "-" & Trim(Text1(5)), "")
   strFileName = strExcelPath & IIf(strKey <> "", Mid(strKey, 2), "") & "員工年度所得統計(非申報數).xls"
   If Dir(strFileName) <> MsgText(601) Then
      Kill strFileName
   End If
   
   '查詢條件:
   '年度
   If Len(Text1(0)) > 0 Or Len(Text1(1)) > 0 Then
      strConSM = strConSM & " AND SM02>=" & Text1(0) + 1911 & "01 AND SM02<=" & Text1(1) + 1911 & "12"
      strConMB = strConMB & " AND MB01>=" & Text1(0) + 1911 & "0101 AND MB01<=" & Text1(1) + 1911 & "1231"
      strConOB = strConOB & " AND OB01>=" & Text1(0) + 1911 & "01 AND OB01<=" & Text1(1) + 1911 & "12"
      strConYB = strConYB & " AND YB01>=" & Text1(0) + 1911 & " AND YB01<=" & Text1(1) + 1911
   End If
   '部門
   If Len(Text1(2)) > 0 Or Len(Text1(3)) > 0 Then
      strConSM = strConSM & " AND SM03>='" & Text1(2) & "' AND SM03<='" & Text1(3) & "'"
      'Modified by Morgan 2024/1/31 +新部門判斷
      If bolNewDep Then
         strConMB = strConMB & " AND ST93>='" & Text1(2) & "' AND ST93<='" & Text1(3) & "'"
         strConOB = strConOB & " AND ST93>='" & Text1(2) & "' AND ST93<='" & Text1(3) & "'"
      Else
         strConMB = strConMB & " AND ST03>='" & Text1(2) & "' AND ST03<='" & Text1(3) & "'"
         strConOB = strConOB & " AND ST03>='" & Text1(2) & "' AND ST03<='" & Text1(3) & "'"
      End If
      strConYB = strConYB & " AND YB03>='" & Text1(2) & "' AND YB03<='" & Text1(3) & "'"
   End If
   '員工編號
   If Len(Text1(4)) > 0 Or Len(Text1(5)) > 0 Then
      strConSM = strConSM & " AND SM01>='" & Text1(4) & "' AND SM01<='" & Text1(5) & "'"
      strConMB = strConMB & " AND MB02>='" & Text1(4) & "' AND MB02<='" & Text1(5) & "'"
      strConOB = strConOB & " AND OB03>='" & Text1(4) & "' AND OB03<='" & Text1(5) & "'"
      strConYB = strConYB & " AND YB02>='" & Text1(4) & "' AND YB02<='" & Text1(5) & "'"
   End If
   
   '檢查年度
   strExc(0) = "SELECT distinct substr(sm02,1,4) FROM STAFF,SALARYMONTH" & _
               " WHERE ST01=SM01(+) " & strConSM & " GROUP BY SM02 ORDER BY substr(sm02,1,4) asc"
   intI = 1
   Set adoRs = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      intSheets = adoRs.RecordCount
   End If
   xlsSalesPoint.SheetsInNewWorkbook = IIf(intSheets > 0, intSheets, 1) 'Office2013建立excel檔案的工作表不一定存在,一開始預設工作表數量
   xlsSalesPoint.Workbooks.add
   adoRs.MoveFirst
   For intSht = 1 To intSheets
      strYear = "" & adoRs.Fields(0)
      Set wksaccrpt114 = xlsSalesPoint.Worksheets(intSht)
      xlsSalesPoint.Sheets(intSht).Activate
      wksaccrpt114.Name = Val(strYear) - 1911 & "年"
      '標題
      wksaccrpt114.Columns("a:a").ColumnWidth = 9: wksaccrpt114.Range("a1").Value = "部門"
      wksaccrpt114.Columns("b:b").ColumnWidth = 8: wksaccrpt114.Range("b1").Value = "編號"
      wksaccrpt114.Columns("c:c").ColumnWidth = 8: wksaccrpt114.Range("c1").Value = "姓名"
      wksaccrpt114.Columns("d:d").ColumnWidth = 8: wksaccrpt114.Range("d1").Value = "離職日"
      wksaccrpt114.Columns("e:e").ColumnWidth = 9: wksaccrpt114.Range("e1").Value = "年度薪資"
      wksaccrpt114.Columns("f:f").ColumnWidth = 9: wksaccrpt114.Range("f1").Value = "年度獎金"
      wksaccrpt114.Columns("g:g").ColumnWidth = 9: wksaccrpt114.Range("g1").Value = "年終獎金"
      wksaccrpt114.Columns("h:h").ColumnWidth = 9: wksaccrpt114.Range("h1").Value = "特殊功績獎金"
      wksaccrpt114.Columns("i:i").ColumnWidth = 9: wksaccrpt114.Range("i1").Value = "紅利"
      wksaccrpt114.Columns("j:j").ColumnWidth = 9: wksaccrpt114.Range("j1").Value = "合計"
      wksaccrpt114.Columns("k:k").ColumnWidth = 9: wksaccrpt114.Range("k1").Value = "加班費"
      wksaccrpt114.Columns("L:L").ColumnWidth = 9: wksaccrpt114.Range("L1").Value = "未休假代金"

      '明細資料查詢條件:
      strConSM = "": strConMB = "": strConOB = "": strConYB = ""
      '年度
      If Len(strYear) > 0 Then
         strConSM = strConSM & " AND SM02>=" & strYear & "01 AND SM02<=" & strYear & "12"
         strConMB = strConMB & " AND MB01>=" & strYear & "0101 AND MB01<=" & strYear & "1231"
         strConOB = strConOB & " AND OB01>=" & strYear & "01 AND OB01<=" & strYear & "12"
         strConYB = strConYB & " AND YB01>=" & strYear & " AND YB01<=" & strYear
      End If
      '部門
      If Len(Text1(2)) > 0 Or Len(Text1(3)) > 0 Then
         strConSM = strConSM & " AND SM03>='" & Text1(2) & "' AND SM03<='" & Text1(3) & "'"
         'Modified by Morgan 2024/1/31 +新部門判斷
         If bolNewDep Then
            strConMB = strConMB & " AND ST93>='" & Text1(2) & "' AND ST93<='" & Text1(3) & "'"
            strConOB = strConOB & " AND ST93>='" & Text1(2) & "' AND ST93<='" & Text1(3) & "'"
         Else
            strConMB = strConMB & " AND ST03>='" & Text1(2) & "' AND ST03<='" & Text1(3) & "'"
            strConOB = strConOB & " AND ST03>='" & Text1(2) & "' AND ST03<='" & Text1(3) & "'"
         End If
         strConYB = strConYB & " AND YB03>='" & Text1(2) & "' AND YB03<='" & Text1(3) & "'"
      End If
      '員工編號
      If Len(Text1(4)) > 0 Or Len(Text1(5)) > 0 Then
         strConSM = strConSM & " AND SM01>='" & Text1(4) & "' AND SM01<='" & Text1(5) & "'"
         strConMB = strConMB & " AND MB02>='" & Text1(4) & "' AND MB02<='" & Text1(5) & "'"
         strConOB = strConOB & " AND OB03>='" & Text1(4) & "' AND OB03<='" & Text1(5) & "'"
         strConYB = strConYB & " AND YB02>='" & Text1(4) & "' AND YB02<='" & Text1(5) & "'"
      End If
      
      '明細資料
      'Modified by Morgan 2024/1/31 +新部門判斷ACC090NEW
      strExc(0) = "SELECT " & IIf(strYear >= Left(新部門啟用日, 4), "ST93 ST03,A0922", "ST03,A0902") & " 部門,ST01 編號,ST02 姓名,SQLDATET(ST51) 離職日,NVL(SM04,0) 年度薪資,NVL(MB03,0)+NVL(OB05,0) 年度獎金,NVL(YB05,0) 年終獎金," & _
                  "NVL(YB06,0) 特殊功績獎金,NVL(YB26,0) 紅利,NVL(SM04,0)+NVL(MB03,0)+NVL(OB05,0)+NVL(YB05,0)+NVL(YB06,0)+NVL(YB26,0) 合計," & _
                  "NVL(SM12,0) 加班費,NVL(YB08,0) 未休假代金 FROM STAFF,ACC090,ACC090NEW," & _
                  "(SELECT SM01,SUM(NVL(SM04,0)+NVL(SM05,0)+NVL(SM07,0)) SM04,SUM(NVL(SM12,0)) SM12 FROM SALARYMONTH,STAFF" & _
                  " WHERE ST01=SM01(+) " & strConSM & " GROUP BY SM01)," & _
                  "(SELECT MB02,SUM(NVL(MB03,0)) MB03 FROM MONTHBONUS,STAFF WHERE ST01=MB02(+) " & strConMB & " GROUP BY MB02)," & _
                  "(SELECT OB03,SUM(NVL(OB05,0)) OB05 FROM OHBONUS,STAFF WHERE ST01=OB03(+) " & strConOB & " GROUP BY OB03)," & _
                  "(SELECT YB02,NVL(YB05,0) YB05,NVL(YB06,0) YB06,NVL(YB26,0) YB26,NVL(YB08,0) YB08 FROM YEARBONUS,STAFF WHERE ST01=YB02(+) " & strConYB & ")" & _
                  " WHERE ST01=SM01(+) AND ST03=A0901(+) AND ST93=A0921(+)" & _
                  " AND ST01=MB02(+) AND ST01=OB03(+) AND ST01=YB02(+)" & _
                  " AND NVL(SM04,0)+NVL(MB03,0)+NVL(OB05,0)+NVL(YB05,0)+NVL(YB06,0)+NVL(YB26,0)+NVL(SM12,0)+NVL(YB08,0)>0" & _
                  " ORDER BY 1,ST01"
      intI = 1
      Set adoRs2 = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         adoRs2.MoveFirst
         intRow = 1
         Do While Not adoRs2.EOF
            intRow = intRow + 1
            '明細資料
            wksaccrpt114.Range("a" & intRow).Value = adoRs2.Fields(1)
            wksaccrpt114.Range("b" & intRow).Value = adoRs2.Fields(2)
            wksaccrpt114.Range("c" & intRow).Value = adoRs2.Fields(3)
            wksaccrpt114.Range("d" & intRow).Value = "" & adoRs2.Fields(4)
            wksaccrpt114.Range("e" & intRow).Value = Val("" & adoRs2.Fields(5))
            wksaccrpt114.Range("f" & intRow).Value = Val("" & adoRs2.Fields(6))
            wksaccrpt114.Range("g" & intRow).Value = Val("" & adoRs2.Fields(7))
            wksaccrpt114.Range("h" & intRow).Value = Val("" & adoRs2.Fields(8))
            wksaccrpt114.Range("i" & intRow).Value = Val("" & adoRs2.Fields(9))
            'wksaccrpt114.Range("j" & intRow).Value = Val("" & adoRs2.Fields(10)) '合計
            wksaccrpt114.Range("j" & intRow).Formula = "=E" & intRow & "+F" & intRow & "+G" & intRow & "+H" & intRow & "+I" & intRow '合計
            wksaccrpt114.Range("k" & intRow).Value = Val("" & adoRs2.Fields(11))
            wksaccrpt114.Range("L" & intRow).Value = Val("" & adoRs2.Fields(12))
            
            adoRs2.MoveNext
         Loop
         '格式化
         wksaccrpt114.Range("e2:L" & intRow).Select
         wksaccrpt114.Range("e2:L" & intRow).NumberFormatLocal = "#,##0_ "
      End If
      
      adoRs.MoveNext
   Next intSht
   
   xlsSalesPoint.Sheets(1).Activate
   If Val(xlsSalesPoint.Version) < 12 Then
      xlsSalesPoint.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=-4143
   Else
      xlsSalesPoint.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=56
   End If
   xlsSalesPoint.Workbooks.Close
   xlsSalesPoint.Quit
   'Modify by Amy 2021/06/22 路徑改中文字顯示
   MsgBox "檔案已產生！電子檔位置：" & strExcelPathN & Replace(strFileName, strExcelPath, "")
   
flgErr:
   Set wksaccrpt114 = Nothing
   Set xlsSalesPoint = Nothing
   Set adoRs = Nothing
   Set adoRs2 = Nothing
   If Err.Number <> 0 Then
      MsgBox Err.Description
   End If
End Sub

Private Sub Form_Load()
Dim strSystemKind As String
   
   MoveFormToCenter Me
   
   strSystemKind = GetSystemKindByNick
   PUB_SetPrinter Me.Name, Combo1
   
   Text1(0) = Left(strSrvDate(2), 3) - 1
   Text1(1) = Left(strSrvDate(2), 3) - 1
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm170241 = Nothing
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   InverseTextBox Text1(Index)
   CloseIme
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0, 1
         KeyAscii = Pub_NumAscii(KeyAscii)
      Case 2, 3, 4, 5
         KeyAscii = UpperCase(KeyAscii)
      Case Else
   End Select
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 0, 1
         If Index = 0 Then
            If Text1(Index) <> "" And Text1(Index + 1) = "" Then
               Text1(Index + 1) = Text1(Index)
            End If
         Else
            If RunNick(Text1(Index - 1), Text1(Index)) Then
               Cancel = True
            End If
         End If
      Case 2, 3
'         If Text1(Index) <> "" Then
'            If ChkDate(Text1(Index) & "01") = False Then
'               Cancel = True
'            End If
'         End If
         If Index = 2 Then
            If Text1(Index) <> "" And Text1(Index + 1) = "" Then
               Text1(Index + 1) = Text1(Index)
            End If
         Else
            If RunNick(Text1(Index - 1), Text1(Index)) Then
               Cancel = True
            End If
         End If
      Case 4, 5
         If Index = 4 Then
            If Text1(Index) <> "" And Text1(Index + 1) = "" Then
               Text1(Index + 1) = Text1(Index)
            End If
         Else
            If RunNick(Text1(Index - 1), Text1(Index)) Then
               Cancel = True
            End If
         End If
   End Select
End Sub
