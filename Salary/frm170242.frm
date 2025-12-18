VERSION 5.00
Begin VB.Form frm170242 
   BorderStyle     =   1  '單線固定
   Caption         =   "每月薪資異動金額明細"
   ClientHeight    =   1848
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   3996
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1848
   ScaleWidth      =   3996
   Begin VB.CommandButton cmdok 
      Caption         =   "產生Excel檔(&E)"
      Default         =   -1  'True
      Height          =   435
      Index           =   0
      Left            =   1464
      TabIndex        =   4
      Top             =   96
      Width           =   1455
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "離開(&X)"
      Height          =   435
      Index           =   1
      Left            =   2976
      TabIndex        =   3
      Top             =   96
      Width           =   915
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   1728
      MaxLength       =   3
      TabIndex        =   1
      Text            =   "96"
      Top             =   888
      Width           =   435
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   2472
      MaxLength       =   2
      TabIndex        =   0
      Text            =   "5"
      Top             =   888
      Width           =   285
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "薪資月份：            年         月"
      Height          =   180
      Left            =   792
      TabIndex        =   2
      Top             =   948
      Width           =   2208
   End
End
Attribute VB_Name = "frm170242"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Created By Morgan 2023/6/15
Option Explicit

Private Sub ExcelSave()
   Dim strYM As String, strYM2 As String
   Dim xlsReport As Excel.Application
   Dim wksReport As Excel.Worksheet
   Dim ii As Integer
   Dim strFileName As String
   
On Error GoTo ErrHnd

   strYM = 100 * (Val(Text1(1)) + 1911) + Val(Text1(2))
   strYM2 = Left(CompDate(1, -1, strYM & "01"), 6)
   
   'Modified by Morgan 2023/7/4 排序只需用員工號--婉莘
   'Modified by Morgan 2023/12/27 +acc090new, st03-->sm03
   strExc(0) = "select " & IIf(strYM >= Left(新部門啟用日, 6), "a0922", "a0902") & " 部門,st01 員工編號,st02 員工姓名,decode(st04,'1',null,'2','離職',st04) 離職" & _
      ",M1 當月,M2 前月,sqldatet(sl02) 上次薪資異動日期,decode(CMP,'2','',a0820) 公司別,F1,F2" & _
      " from (select nvl(a.sm03,b.sm03) sm03,st01,st02,st04,sl02,a.sm37 CMP" & _
      ",nvl(a.sm04,0)+nvl(a.sm05,0)+nvl(a.sm06,0)+nvl(a.sm07,0)+nvl(a.sm08,0)+nvl(a.sm09,0)+nvl(a.sm45,0) M1,a.sm10 F1" & _
      ",nvl(b.sm04,0)+nvl(b.sm05,0)+nvl(b.sm06,0)+nvl(b.sm07,0)+nvl(b.sm08,0)+nvl(b.sm09,0)+nvl(b.sm45,0) M2,b.sm10 F2" & _
      " from staff,(select sm01,sm02,sm03,sm04,sm05,sm06,sm07,sm08,sm09,sm45,sm10,sm37 from salarymonth where sm02=" & strYM & ") a" & _
      ",(select sm01,sm02,sm03,sm04,sm05,sm06,sm07,sm08,sm09,sm45,sm10 from salarymonth where sm02=" & strYM2 & ") b" & _
      ",(select sl01,max(sl02) sl02 from salarylog group by sl01) c" & _
      " where st03 not in ('F51','F52') and a.sm01(+)=st01 and b.sm01(+)=st01 and sl01(+)=st01) x,acc090,acc090new,acc080" & _
      " where a0901(+)=sm03 and a0921(+)=sm03 and (nvl(M1,0)<>nvl(M2,0) or nvl(F1,0)<>nvl(F2,0)) and a0801(+)=CMP" & _
      " order by st01"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI <> 1 Then
      Exit Sub
   End If
         
   strFileName = strExcelPath & Val(Text1(1)) & "年" & Val(Text1(2)) & "月薪資異動明細.xlsx"
   If Dir(strFileName) <> "" Then Kill strFileName
   
   
   Set xlsReport = CreateObject("Excel.Application")
   'xlsReport.Visible = True
   
   xlsReport.SheetsInNewWorkbook = 1
   xlsReport.Workbooks.add
   Set wksReport = xlsReport.Worksheets(1)
   wksReport.Name = "薪資異動明細"
   wksReport.Range("A1") = "部門"
   wksReport.Range("B1") = "員工編號"
   wksReport.Range("C1") = "員工姓名"
   wksReport.Range("D1") = "離職"
   wksReport.Range("E1") = Val(Right(strYM2, 2)) & "月"
   wksReport.Range("F1") = Val(Right(strYM, 2)) & "月"
   wksReport.Range("G1") = "變動金額"
   wksReport.Range("H1") = "上次薪資異動日期"
   wksReport.Range("I1") = "公司別"
   wksReport.Range("J1") = Val(Right(strYM2, 2)) & "月特支費"
   wksReport.Range("K1") = Val(Right(strYM, 2)) & "月特支費"
   wksReport.Range("L1") = "特支費異動金額"
   wksReport.Range("A1:L1").Font.Bold = True
   wksReport.Range("A1:L1").HorizontalAlignment = xlCenter
   
   With RsTemp
   ii = 1
   Do While Not .EOF
      ii = ii + 1
      wksReport.Range("A" & ii) = "" & .Fields("部門")
      wksReport.Range("B" & ii).NumberFormatLocal = "@"
      wksReport.Range("B" & ii) = "" & .Fields("員工編號")
      wksReport.Range("C" & ii) = "" & .Fields("員工姓名")
      wksReport.Range("D" & ii) = "" & .Fields("離職")
      wksReport.Range("E" & ii).NumberFormatLocal = "#,##0"
      wksReport.Range("E" & ii) = "" & .Fields("前月")
      wksReport.Range("F" & ii).NumberFormatLocal = "#,##0"
      wksReport.Range("F" & ii) = "" & .Fields("當月")
      wksReport.Range("G" & ii).NumberFormatLocal = "#,##0;-#,##0"
      wksReport.Range("G" & ii).Formula = "=F" & ii & "-E" & ii '變動金額
      If wksReport.Range("G" & ii) < 0 Then
         wksReport.Range("G" & ii).Font.Color = vbRed
      End If
      wksReport.Range("H" & ii) = "" & .Fields("上次薪資異動日期")
      wksReport.Range("I" & ii) = "" & .Fields("公司別")
      wksReport.Range("J" & ii).NumberFormatLocal = "#,##0"
      If .Fields("F2") > 0 Then
         wksReport.Range("J" & ii) = "" & .Fields("F2") '前月特支費"
      End If
      wksReport.Range("K" & ii).NumberFormatLocal = "#,##0"
      If .Fields("F1") > 0 Then
         wksReport.Range("K" & ii) = "" & .Fields("F1") '當月特支費"
      End If
      If .Fields("F2") > 0 Or .Fields("F1") > 0 Then
         wksReport.Range("L" & ii).NumberFormatLocal = "#,##0;-#,##0"
         wksReport.Range("L" & ii).Formula = "=K" & ii & "-J" & ii '特支費異動金額
         If wksReport.Range("L" & ii) < 0 Then
            wksReport.Range("L" & ii).Font.Color = vbRed
         End If
      End If
      .MoveNext
   Loop
   End With
   '自動設定欄寬
   For ii = Asc("A") To Asc("L")
      wksReport.Columns(Chr(ii)).EntireColumn.Font.Name = "Arial"
      wksReport.Columns(Chr(ii)).EntireColumn.Font.Size = "10"
      wksReport.Columns(Chr(ii)).EntireColumn.AutoFit
   Next
   wksReport.Columns("B:B").HorizontalAlignment = xlCenter
   wksReport.Columns("C:C").HorizontalAlignment = xlCenter
   wksReport.Columns("D:D").HorizontalAlignment = xlCenter
   wksReport.Columns("H:H").HorizontalAlignment = xlCenter
   wksReport.Columns("I:I").HorizontalAlignment = xlCenter
   
   '凍結窗格(頂端列)
   xlsReport.ActiveWindow.SplitColumn = 0
   xlsReport.ActiveWindow.SplitRow = 1
   xlsReport.ActiveWindow.FreezePanes = True
                
   xlsReport.Workbooks(1).SaveAs strFileName
   If MsgBox("檔案已產生！電子檔位置：" & strExcelPathN & Replace(strFileName, strExcelPath, "") & vbCrLf & vbCrLf & "是否開啟？", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
      xlsReport.Visible = True
   Else
      xlsReport.Workbooks.Close
      xlsReport.Quit
   End If
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical

   Set xlsReport = Nothing
   Set wksReport = Nothing
End Sub

Private Sub cmdok_Click(Index As Integer)
   If Index = 0 Then
   
      If Text1(1) = "" Then
         MsgBox "請輸入年度！", vbInformation, "操作錯誤！"
         Text1(1).SetFocus
         Exit Sub
         
      ElseIf Val(Text1(1)) > Val(Left(strSrvDate(1), 3)) Or Val(Text1(1)) < 97 Then
         MsgBox "月份輸入錯誤！", vbInformation, "操作錯誤！"
         Text1_GotFocus 1
         Text1(1).SetFocus
         Exit Sub
         
      ElseIf Text1(2) = "" Then
         MsgBox "請輸入月份！", vbInformation, "操作錯誤！"
         Text1(2).SetFocus
         Exit Sub
      ElseIf Val(Text1(2)) > 12 Or Val(Text1(2)) < 1 Then
         MsgBox "月份輸入錯誤！", vbInformation, "操作錯誤！"
         Text1_GotFocus 2
         Text1(2).SetFocus
         Exit Sub
      ElseIf 100 * Val(Text1(1)) + Val(Text1(2)) > Val(Left(strSrvDate(2), 5)) Then
         MsgBox "輸入年月不可大於當月！", vbInformation, "操作錯誤！"
         Text1_GotFocus 2
         Text1(2).SetFocus
         Exit Sub
      End If
      
      strExc(0) = "select * from salarymonth where sm02=" & (100 * (Val(Text1(1)) + 1911) + Val(Text1(2))) & " and rownum<2"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         Call Pub_ChkExcelPath
         Screen.MousePointer = vbHourglass
         Call ExcelSave
         Screen.MousePointer = vbDefault
      Else
         MsgBox "該月份尚無月薪資資料！", vbExclamation + vbOKOnly
         Text1_GotFocus 2
         Text1(2).SetFocus
         Exit Sub
      End If
   ElseIf Index = 1 Then
      Unload Me
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   strExc(0) = Left(strSrvDate(2), 5)
   If Right(strExc(0), 2) = "01" Then
      strExc(0) = strExc(0) - 101
   Else
      strExc(0) = strExc(0) - 1
   End If
   Text1(1) = Left(strExc(0), 3)
   Text1(2) = Mid(strExc(0), 4)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm170242 = Nothing
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii <> 8 Then
      If Not IsNumeric(Chr(KeyAscii)) Then
         KeyAscii = 0
         Beep
      End If
   End If
End Sub
