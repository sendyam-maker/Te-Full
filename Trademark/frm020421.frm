VERSION 5.00
Begin VB.Form frm020421 
   BorderStyle     =   1  '單線固定
   Caption         =   "大陸商申查名統計表"
   ClientHeight    =   2010
   ClientLeft      =   3600
   ClientTop       =   3360
   ClientWidth     =   4320
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   4320
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   540
      Left            =   240
      TabIndex        =   5
      Top             =   1380
      Width           =   3825
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   765
         Style           =   2  '單純下拉式
         TabIndex        =   6
         Top             =   180
         Width           =   2880
      End
      Begin VB.Label Label4 
         Caption         =   "印表機"
         Height          =   180
         Left            =   105
         TabIndex        =   7
         Top             =   255
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   3375
      TabIndex        =   3
      Top             =   135
      Width           =   756
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2580
      TabIndex        =   2
      Top             =   135
      Width           =   756
   End
   Begin VB.TextBox txt1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   2700
      MaxLength       =   7
      TabIndex        =   1
      Top             =   795
      Width           =   1065
   End
   Begin VB.TextBox txt1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   1500
      MaxLength       =   7
      TabIndex        =   0
      Top             =   795
      Width           =   1065
   End
   Begin VB.Line Line1 
      X1              =   2220
      X2              =   3240
      Y1              =   975
      Y2              =   975
   End
   Begin VB.Label Label1 
      Caption         =   "統計期間："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   4
      Top             =   840
      Width           =   1200
   End
End
Attribute VB_Name = "frm020421"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create by Lydia 2023/02/01
Option Explicit

Dim strPrinter As String

Private Sub cmdOK_Click(Index As Integer)
Select Case Index
Case 0
     If Len(txt1(0)) = 0 Then
         MsgBox "統計區間不可空白!!", , "USER 輸入錯誤"
         txt1(0).SetFocus
         txt1_GotFocus (0)
         Exit Sub
      End If
     If Len(txt1(1)) = 0 Then
         MsgBox "統計區間不可空白!!", , "USER 輸入錯誤"
         txt1(1).SetFocus
         txt1_GotFocus (1)
         Exit Sub
      End If
      If PUB_CheckKeyInDate(Me.txt1(0)) = -1 Then
         Me.txt1(0).SetFocus
         txt1_GotFocus 0
         Exit Sub
      End If
      If PUB_CheckKeyInDate(Me.txt1(1)) = -1 Then
         Me.txt1(1).SetFocus
         txt1_GotFocus 1
         Exit Sub
      End If
      Screen.MousePointer = vbHourglass
      Me.Enabled = False
      ClearQueryLog (Me.Name) '清除查詢印表記錄檔欄位
      PUB_SetOsDefaultPrinter Combo1 '切換Word/Excel印表機
      PUB_RestorePrinter Combo1
      Call ReadData
      PUB_SetOsDefaultPrinter strPrinter '切換Word/Excel印表機
      PUB_RestorePrinter strPrinter
      Me.Enabled = True
      Screen.MousePointer = vbDefault
Case 1
     Unload Me
Case Else
End Select
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   PUB_SetPrinter Me.Name, Combo1, strPrinter
   
   txt1(0) = (Left(CompDate(1, -1, strSrvDate(1)), 6) & "01") - 19110000
   txt1(1) = CompDate(2, -1, Left(strSrvDate(1), 6) & "01") - 19110000
End Sub

Private Sub Form_Unload(Cancel As Integer)
   '若印表機變動, 則更新列印設定
   If Me.Combo1.Text <> Me.Combo1.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
   End If
   
   Set frm020421 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub Txt1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cmdOK(0).SetFocus
End If
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txt1_LostFocus(Index As Integer)
Select Case Index
Case 0, 1
   If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
      Me.txt1(Index).SetFocus
      txt1_GotFocus Index
      Exit Sub
   End If
   If Index = 1 Then
     If RunNick(txt1(Index - 1), txt1(Index)) Then
         txt1(Index - 1).SetFocus
         txt1_GotFocus (Index - 1)
         Exit Sub
      End If
    End If

Case Else
End Select
End Sub

Private Sub ReadData()
Dim strR1 As String, rsRD As New ADODB.Recordset
Dim strFileName As String
Dim xlsReport
Dim wksReport1
Dim intRow As Integer

On Error GoTo ErrHnd
    
    If Len(txt1(0)) <> 0 Or Len(Trim(txt1(1))) <> 0 Then
        pub_QL05 = pub_QL05 & ";" & Label1 & txt1(0) & "-" & txt1(1)
    End If
      
    If Trim(txt1(0)) <> "" Then
       strR1 = strR1 & " and cp05>='" & DBDATE(txt1(0)) & "' "
    End If
    If Trim(txt1(1)) <> "" Then
       strR1 = strR1 & " and cp05<='" & DBDATE(txt1(1)) & "' "
    End If
    strSql = "select cp14, st02 ,count(*) cnt From caseprogress, Trademark, staff " & _
                "where cp01='T' and cp159=0 and cp10='101' and nvl(cp143,0)>0 " & strR1 & _
                "and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and tm10='020' " & _
                "and cp14=st01(+) group by cp14,st02 order by cp14 "
    intI = 1
    Set rsRD = ClsLawReadRstMsg(intI, strSql)
    If intI = 1 Then
       strFileName = "$$" & Me.Caption & MsgText(43)
       If Dir(App.path & "\" & strFileName) <> "" Then
          Kill App.path & "\" & strFileName
       End If
       InsertQueryLog (rsRD.RecordCount)
       
       rsRD.MoveFirst
       Do While Not rsRD.EOF
           If intRow = 0 Then
               Set xlsReport = CreateObject("Excel.Application")
               xlsReport.SheetsInNewWorkbook = 1
               xlsReport.Workbooks.add
               Set wksReport1 = xlsReport.Worksheets(1)
               wksReport1.Activate
               'xlsReport.Visible = True
               If Val(xlsReport.Version) < 12 Then
                   xlsReport.Workbooks(1).SaveAs FileName:=App.path & "\" & strFileName, FileFormat:=-4143
               Else
                   xlsReport.Workbooks(1).SaveAs FileName:=App.path & "\" & strFileName, FileFormat:=56
               End If
               With wksReport1
                  .Range("A:A").ColumnWidth = 6
                  .Range("A:E").Font.Size = 14
                  .Range("A:E").Font.Name = "標楷體"
                  .Range("B:B").ColumnWidth = 15
                  .Range("C:C").ColumnWidth = 15
                  .Range("D:D").ColumnWidth = 15
                  .Range("D:D").HorizontalAlignment = xlCenter
                  .Range("E:E").ColumnWidth = 20
                  .PageSetup.PaperSize = 9
                  .PageSetup.Orientation = xlPortrait '直印
                  .PageSetup.Zoom = 100 '縮放比例為100%,列印頁面水平置中
                  .PageSetup.CenterHorizontally = True '版面設定->邊界->水平置中
                  
                  '***表頭設定***
                  intRow = intRow + 1
                  .Range("B" & intRow).Value = ChangeTStringToTDateString(txt1(0)) & "－" & ChangeTStringToTDateString(txt1(1)) & " 大陸商申查名統計表"
                  .Range("B" & intRow).Font.Size = 18
                  .Range("B" & intRow).Font.Bold = True
                  .Range("B" & intRow & ":" & "E" & intRow).MergeCells = True
                  .Range("B" & intRow & ":" & "E" & intRow).HorizontalAlignment = xlCenter
                  .Range("B" & intRow & ":" & "E" & intRow).VerticalAlignment = xlCenter
                  .Range(intRow & ":" & intRow).RowHeight = 32
                  intRow = intRow + 1
                  .Range("E" & intRow).Value = "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
                  .Range("E" & intRow).Font.Bold = True
                  intRow = intRow + 1
                  .Range("E" & intRow).Value = "列印人員：" & strUserName
                  .Range("E" & intRow).Font.Bold = True
                  intRow = intRow + 2
                  .Range("C" & intRow).Value = "承辦人員"
                  .Range("C" & intRow).Font.Bold = True
                  .Range("D" & intRow).Value = "件數"
                  .Range("D" & intRow).Font.Bold = True
                  intRow = intRow + 1
               End With
           End If

           wksReport1.Range("C" & intRow).Value = "" & rsRD.Fields("st02")
           wksReport1.Range("D" & intRow).Value = "" & rsRD.Fields("cnt")
           intRow = intRow + 1
           rsRD.MoveNext
       Loop
       '***表尾設定***
       wksReport1.Range("C" & intRow).Value = "總件數"
       wksReport1.Range("C" & intRow).Font.Bold = True
       wksReport1.Range("D" & intRow).Formula = "=SUM(C6:D" & intRow - 1 & ")"
       wksReport1.Range("D" & intRow).Font.Bold = True
       With wksReport1.Range("C5:D" & intRow)
             .Borders(xlEdgeTop).LineStyle = xlContinuous
             .Borders(xlEdgeTop).Weight = xlThin '細線
             .Borders(xlEdgeBottom).LineStyle = xlContinuous
             .Borders(xlEdgeBottom).Weight = xlThin
             .Borders(xlEdgeLeft).LineStyle = xlContinuous
             .Borders(xlEdgeLeft).Weight = xlThin
             .Borders(xlEdgeRight).LineStyle = xlContinuous
             .Borders(xlEdgeRight).Weight = xlThin
             .Borders(xlInsideVertical).LineStyle = xlContinuous
             .Borders(xlInsideVertical).Weight = xlThin
             .Borders(xlInsideHorizontal).LineStyle = xlContinuous
             .Borders(xlInsideHorizontal).Weight = xlThin
       End With
       wksReport1.Range("2:" & intRow).RowHeight = 24
       xlsReport.Workbooks(1).Save
       wksReport1.PrintOut Copies:=1, Collate:=True
       xlsReport.Workbooks(1).Save
       xlsReport.Quit
       Set wksReport1 = Nothing
       Set xlsReport = Nothing
    Else
        InsertQueryLog (0)
        ShowNoData
    End If
    
    Set rsRD = Nothing
    Exit Sub

ErrHnd:
    If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Sub

