VERSION 5.00
Begin VB.Form frm090643 
   BorderStyle     =   1  '單線固定
   Caption         =   "支援次數統計"
   ClientHeight    =   2256
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4284
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2256
   ScaleWidth      =   4284
   Begin VB.CheckBox Check1 
      Caption         =   "依案號合計"
      Height          =   276
      Left            =   1560
      TabIndex        =   5
      Top             =   1224
      Width           =   1788
   End
   Begin VB.CommandButton cmdExcel 
      BackColor       =   &H00C0FFC0&
      Caption         =   "產生 EXCEL (&O)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   1104
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   1632
      Width           =   2292
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   0
      Left            =   3384
      TabIndex        =   3
      Top             =   96
      Width           =   852
   End
   Begin VB.TextBox txtDate 
      Height          =   345
      Index           =   0
      Left            =   1536
      MaxLength       =   7
      TabIndex        =   0
      Top             =   732
      Width           =   795
   End
   Begin VB.TextBox txtDate 
      Height          =   345
      Index           =   1
      Left            =   2532
      MaxLength       =   7
      TabIndex        =   1
      Top             =   732
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "支援日期區間："
      Height          =   216
      Left            =   216
      TabIndex        =   4
      Top             =   816
      Width           =   1368
   End
   Begin VB.Line Line1 
      X1              =   2112
      X2              =   2802
      Y1              =   912
      Y2              =   912
   End
End
Attribute VB_Name = "frm090643"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Created by Morgan 2025/4/22
Option Explicit
Dim m_ESeq As String
Dim strFileName As String

Private Sub cmdExit_Click(Index As Integer)
   Unload Me
End Sub

Private Sub cmdExcel_Click()
Dim intErr As Integer
Dim bolTmp As Boolean
   
   If txtDate(0) = "" Then
      MsgBox "請輸入支援日期起日！", vbCritical
      txtDate(0).SetFocus
      Exit Sub
   End If
   
   If txtDate(1) = "" Then
      MsgBox "請輸入支援日期迄日！", vbCritical
      txtDate(1).SetFocus
      Exit Sub
   End If
   
   If Val(txtDate(0)) > Val(txtDate(1)) Then
      MsgBox "支援日期起日不可大於迄日！", vbCritical
      txtDate(0).SetFocus
      Exit Sub
   End If

   strFileName = strExcelPath & txtDate(0) & Me.Caption & IIf(Check1.Value = vbChecked, "(" & Check1.Caption & ")", "") & MsgText(43)
   RidFile strFileName
   
   Screen.MousePointer = vbHourglass
   cmdExcel.Enabled = False
   'Added by Morgan 2025/5/9
   If Check1.Value = vbChecked Then
      If Process2 = False Then
         strFileName = ""
      End If
   Else
   'end 2025/5/9
      If Process = False Then
         strFileName = ""
      End If
   End If
   cmdExcel.Enabled = True
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()

   MoveFormToCenter Me
   
   If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
       MkDir strExcelPath
   End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   Set frm090643 = Nothing
End Sub

Private Sub txtDate_GotFocus(Index As Integer)
   TextInverse txtDate(Index)
End Sub

Private Sub txtDate_Validate(Index As Integer, Cancel As Boolean)
Dim strTemp1 As String

   If txtDate(Index) <> "" Then
      strTemp1 = txtDate(Index)
      If CheckIsTaiwanDate(strTemp1) = False Then
         MsgBox "請輸入民國日期!", vbCritical
         txtDate(Index).SetFocus
         txtDate_GotFocus Index
         Cancel = True
         Exit Sub
      End If
   End If
End Sub

Private Function Process() As Boolean
Dim strDate(0 To 1) As String
   
   ClearQueryLog (Me.Name)
   strDate(0) = DBDATE(txtDate(0))
   strDate(1) = DBDATE(txtDate(1))
   pub_QL05 = pub_QL05 & ";統計區間：" & txtDate(0) & "-" & txtDate(1)
   
On Error GoTo ErrHandle

   strSql = "select sqldatet(sh01) SDate,s1.st02 SEng,s2.st02 SSal" & _
      ",sh06||'-'||sh07||decode(sh08||sh09,'000','','-'||sh08||'-'||sh09) CaseNo" & _
      ",sh05 Hr,x06 Cnt from Supporthour,staff s1,staff s2" & _
      ",(select s1.sh01 x01,s1.sh06 x02,s1.sh07 x03,s1.sh08 x04,s1.sh09 x05,count(distinct s2.sh01) x06" & _
      " from Supporthour s1,Supporthour s2 where s1.sh20 is not null and s1.sh01>=" & strDate(0) & " and s1.sh01<=" & strDate(1) & _
      " and s2.sh06(+)=s1.sh06 and s2.sh07(+)=s1.sh07 and s2.sh08(+)=s1.sh08 and s2.sh09(+)=s1.sh09 and s2.sh20(+)=s1.sh20 and s2.sh01<=s1.sh01" & _
      " group by s1.sh01,s1.sh06,s1.sh07,s1.sh08,s1.sh09) X" & _
      " Where sh20 Is Not Null and sh01>=" & strDate(0) & " and sh01<=" & strDate(1) & _
      " and s1.st01(+)=sh02 and s2.st01(+)=sh03" & _
      " and x01(+)=sh01 and x02(+)=sh06 and x03(+)=sh07 and x04(+)=sh08 and x05(+)=sh09" & _
      " order by sh01,sh02,sh03"
   intI = 0
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      InsertQueryLog (RsTemp.RecordCount)
      If ProcSaveExcel = True Then
         Process = True
      End If
   Else
      InsertQueryLog (0)
   End If
   Exit Function
   
ErrHandle:
   If Err.Number <> 0 Then
      MsgBox "執行失敗：" & Err.Description, vbCritical
   End If
End Function

Private Function ProcSaveExcel() As Boolean
   Dim xlsPoint As New Excel.Application
   Dim WksPoint As New Worksheet
   Dim bolOpenxlsPoint As Boolean
   Dim intB1 As Integer, intB2 As Integer
   Dim xRow As Integer, tmpArr As Variant

   '-------預設Excel
   xlsPoint.SheetsInNewWorkbook = 1
   xlsPoint.Workbooks.add
   xlsPoint.Application.Visible = False
   xlsPoint.Worksheets(1).Name = txtDate(0) & "-" & txtDate(1)
   Set WksPoint = xlsPoint.Worksheets(1)
   bolOpenxlsPoint = True
   
   WksPoint.Cells.Font.Name = "Arial"
   WksPoint.Cells.Font.Size = 10
   
   xRow = 1
   WksPoint.Range("A" & xRow).Value = "日期"
   WksPoint.Range("B" & xRow).Value = "工程師"
   WksPoint.Range("C" & xRow).Value = "智權人員"
   WksPoint.Range("D" & xRow).Value = "本所案號"
   WksPoint.Range("E" & xRow).Value = "支援時數"
   WksPoint.Range("F" & xRow).Value = "累積支援次數"
   WksPoint.Range("A" & xRow, "F" & xRow).Font.Bold = True
   WksPoint.Range("A1:F" & xRow).Columns.AutoFit
   RsTemp.MoveFirst
   Do While Not RsTemp.EOF
      xRow = xRow + 1
      For intI = 0 To 5
         WksPoint.Range(Chr(65 + intI) & xRow).Value = RsTemp.Fields(intI)
      Next
      RsTemp.MoveNext
   Loop
   xRow = xRow + 1
   WksPoint.Range("D" & xRow).Value = "合計:"
   WksPoint.Range("E" & xRow).Formula = "=sum($E$1:$E" & (xRow - 1) & ")"
   WksPoint.Range("D" & xRow, "E" & xRow).Font.Bold = True
   WksPoint.Range("A1:F" & xRow).Columns.AutoFit
   xlsPoint.ActiveWindow.SplitColumn = 0
   xlsPoint.ActiveWindow.SplitRow = 1
   xlsPoint.ActiveWindow.FreezePanes = True
   xlsPoint.Sheets(1).Select '選擇工作表

   '判斷版本
   If Val(xlsPoint.Version) < 12 Then
        xlsPoint.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=-4143
   Else
        xlsPoint.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=56
   End If
   
   
   If MsgBox("Excel檔案產生完成！檔案位置：" & strExcelPathN & vbCrLf & vbCrLf & "是否要開啟？", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
      xlsPoint.Visible = True
   Else
      xlsPoint.Workbooks.Close
      xlsPoint.Quit
   End If
   
   
   
   ProcSaveExcel = True
   Exit Function
   
ErrHnd1:
    If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
    If bolOpenxlsPoint = True Then
        xlsPoint.Workbooks(1).Close xlDoNotSaveChanges
        xlsPoint.Quit
    End If
End Function
'Added by Morgan 2025/5/9
Private Function Process2() As Boolean
Dim strDate(0 To 1) As String
   
   ClearQueryLog (Me.Name)
   strDate(0) = DBDATE(txtDate(0))
   strDate(1) = DBDATE(txtDate(1))
   pub_QL05 = pub_QL05 & ";統計區間：" & txtDate(0) & "-" & txtDate(1)
   pub_QL05 = pub_QL05 & ";依案號合計"
   
On Error GoTo ErrHandle

   strSql = "select max(s1.st02) SEng,max(s2.st02) SSal" & _
      ",sh06||'-'||sh07||decode(sh08||sh09,'000','','-'||sh08||'-'||sh09) CaseNo" & _
      ",sum(sh05) Hr,count(*) Cnt from Supporthour,staff s1,staff s2" & _
      " Where sh20 Is Not Null and sh01>=" & strDate(0) & " and sh01<=" & strDate(1) & _
      " and s1.st01(+)=sh02 and s2.st01(+)=sh03" & _
      " group by sh02,sh03,sh06,sh07,sh08,sh09 order by 3,1,2"
   intI = 0
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      InsertQueryLog (RsTemp.RecordCount)
      If ProcSaveExcel2 = True Then
         Process2 = True
      End If
   Else
      InsertQueryLog (0)
   End If
   Exit Function
   
ErrHandle:
   If Err.Number <> 0 Then
      MsgBox "執行失敗：" & Err.Description, vbCritical
   End If
End Function
'Added by Morgan 2025/5/9
Private Function ProcSaveExcel2() As Boolean
   Dim xlsPoint As New Excel.Application
   Dim WksPoint As New Worksheet
   Dim bolOpenxlsPoint As Boolean
   Dim intB1 As Integer, intB2 As Integer
   Dim xRow As Integer, tmpArr As Variant

   '-------預設Excel
   xlsPoint.SheetsInNewWorkbook = 1
   xlsPoint.Workbooks.add
   xlsPoint.Application.Visible = False
   xlsPoint.Worksheets(1).Name = txtDate(0) & "-" & txtDate(1)
   Set WksPoint = xlsPoint.Worksheets(1)
   bolOpenxlsPoint = True
   
   WksPoint.Cells.Font.Name = "Arial"
   WksPoint.Cells.Font.Size = 10
   
   xRow = 1
   WksPoint.Range("A" & xRow).Value = "工程師"
   WksPoint.Range("B" & xRow).Value = "智權人員"
   WksPoint.Range("C" & xRow).Value = "本所案號"
   WksPoint.Range("D" & xRow).Value = "支援時數"
   WksPoint.Range("E" & xRow).Value = "支援次數"
   WksPoint.Range("A" & xRow, "E" & xRow).Font.Bold = True
   WksPoint.Range("A1:E" & xRow).Columns.AutoFit
   RsTemp.MoveFirst
   Do While Not RsTemp.EOF
      xRow = xRow + 1
      For intI = 0 To 4
         WksPoint.Range(Chr(65 + intI) & xRow).Value = RsTemp.Fields(intI)
      Next
      RsTemp.MoveNext
   Loop
   xRow = xRow + 1
   WksPoint.Range("C" & xRow).Value = "合計:"
   WksPoint.Range("D" & xRow).Formula = "=sum($D$1:$D" & (xRow - 1) & ")"
   WksPoint.Range("E" & xRow).Formula = "=sum($E$1:$E" & (xRow - 1) & ")"
   WksPoint.Range("C" & xRow, "E" & xRow).Font.Bold = True
   WksPoint.Range("A1:E" & xRow).Columns.AutoFit
   xlsPoint.ActiveWindow.SplitColumn = 0
   xlsPoint.ActiveWindow.SplitRow = 1
   xlsPoint.ActiveWindow.FreezePanes = True
   xlsPoint.Sheets(1).Select '選擇工作表

   '判斷版本
   If Val(xlsPoint.Version) < 12 Then
        xlsPoint.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=-4143
   Else
        xlsPoint.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=56
   End If
   
   
   If MsgBox("Excel檔案產生完成！檔案位置：" & strExcelPathN & vbCrLf & vbCrLf & "是否要開啟？", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
      xlsPoint.Visible = True
   Else
      xlsPoint.Workbooks.Close
      xlsPoint.Quit
   End If
   
   
   
   ProcSaveExcel2 = True
   Exit Function
   
ErrHnd1:
    If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
    If bolOpenxlsPoint = True Then
        xlsPoint.Workbooks(1).Close xlDoNotSaveChanges
        xlsPoint.Quit
    End If
End Function
