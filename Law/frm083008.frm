VERSION 5.00
Begin VB.Form frm083008 
   BorderStyle     =   1  '單線固定
   Caption         =   "收款業績比較表"
   ClientHeight    =   2115
   ClientLeft      =   2550
   ClientTop       =   2625
   ClientWidth     =   3960
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   3960
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   0
      Left            =   1452
      MaxLength       =   3
      TabIndex        =   0
      Top             =   948
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   1
      Left            =   1452
      MaxLength       =   3
      TabIndex        =   1
      Top             =   1308
      Width           =   375
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   400
      Left            =   2148
      TabIndex        =   2
      Top             =   120
      Width           =   800
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   2976
      TabIndex        =   3
      Top             =   120
      Width           =   760
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "第一年度：               年度"
      Height          =   180
      Index           =   0
      Left            =   492
      TabIndex        =   5
      Top             =   948
      Width           =   1992
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "第二年度：               年度"
      Height          =   180
      Index           =   1
      Left            =   492
      TabIndex        =   4
      Top             =   1308
      Width           =   2028
   End
End
Attribute VB_Name = "frm083008"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/3 日期欄已修改
Option Explicit

Dim SDay As String, EDay As String, PLeft(0 To 10) As Integer
Dim TitName(1 To 2) As String
Dim m_print As Integer


Private Sub cmdBack_Click()
   Unload Me
End Sub

Private Sub cmdPrint_Click()
   m_print = 0
   If Text1(0) = "" Or Text1(1) = "" Then
      MsgBox "兩年之年度必須皆有值 !", vbCritical
      Exit Sub
   End If
   Screen.MousePointer = 11
   GetPrintLeft
   PrintCase
   Screen.MousePointer = 0
   If m_print = 0 Then
      MsgBox "列印結束!", vbInformation
   End If
End Sub

Private Sub PrintCase()
Dim i As Integer, Page As Integer, iPrint As Integer, j As Integer
Dim TmpArea As String, Qty As String
Dim StRng(1 To 4) As String, StS(1 To 6) As String, Stt(1 To 8) As String
Dim Wo As DAO.Workspace, Db As DAO.Database, Rc As DAO.Recordset
Dim StNum(1 To 2) As String
Dim intAvg As Single
Dim intTemp As Single

'On Error GoTo err
   If CreateDatabase = False Then
      MsgBox "無法建立暫存區，列印失敗 !", vbInformation
      m_print = 1
      Exit Sub
   End If
   
   Set Wo = DBEngine.Workspaces(0)
   Set Db = Wo.OpenDatabase(App.Path & "\Case.mdb", False, False, ";PWD=taie")
   Qty = "DELETE FROM TEMP"
   Db.Execute Qty
   
   If Me.Tag = 0 Then
      StNum(1) = "'410102','411102','414102','418102'"
      StNum(2) = "'4141','414101','418101','4181','4182','4183','4184'"
      TitName(1) = "顧問"
      TitName(2) = "法務"
   Else
      StNum(1) = "'416101'"
      StNum(2) = "'416102'"
      TitName(1) = "FCL"
      TitName(2) = "CFL"
   End If
   StS(1) = "第一季": StS(2) = "第二季": StS(3) = "第三季"
   StS(4) = "第四季": StS(5) = "全年": StS(6) = "平均每月"
   
   If RsTemp.State = adStateOpen Then RsTemp.Close
   For i = 1 To 4
      Qty = "INSERT INTO TEMP (TMP01,TMP02,TMP03,TMP04,TMP05,TMP06,TMP07,TMP08,TMP09) " & _
         "VALUES ('" & StS(i) & "','0','0','0','0','0','0','0','0')"
      Db.Execute Qty
      
      Qty = "SELECT (SUM(A0407)-SUM(A0406))/1000 FROM ACC040 " & _
            " WHERE A0405 IN (" & StNum(1) & ")" & _
            " AND A0401 =" & Text1(0) & _
            " AND (A0402 BETWEEN " & Format((i - 1) * 3 + 1, "00") & _
            " AND " & Format(i * 3, "00") & ")" & _
            " AND A0403='1' AND A0404='TOT'"
      RsTemp.Open Qty, cnnConnection
      If IsNull(RsTemp.Fields(0)) = True Then
         Stt(1) = "0.00"
      Else
         Stt(1) = Format(RsTemp.Fields(0), "0.00")
      End If
      RsTemp.Close
      
      Qty = "SELECT (SUM(A0407)-SUM(A0406))/1000 FROM ACC040 " & _
            " WHERE A0405 IN (" & StNum(2) & ")" & _
            " AND A0401 =" & Text1(0) & _
            " AND (A0402 BETWEEN " & Format((i - 1) * 3 + 1, "00") & _
            " AND " & Format(i * 3, "00") & ")" & _
            " AND A0403='1' AND A0404='TOT'"
      RsTemp.Open Qty, cnnConnection
      If IsNull(RsTemp.Fields(0)) = True Then
         Stt(2) = "0.00"
      Else
         Stt(2) = Format(RsTemp.Fields(0), "0.00")
      End If
      Stt(3) = Format(Val(Stt(1)) + Val(Stt(2)), "0.00")
      RsTemp.Close
            
      Qty = "SELECT SUM(A0409)/1000 FROM ACC040 WHERE A0401=" & Text1(0).Text & _
            " AND (A0402 BETWEEN " & Format((i - 1) * 3 + 1, "00") & _
            " AND " & Format(i * 3, "00") & ")" & _
            " AND A0403='1' AND A0404='TOT'" & _
            " AND A0405 IN (" & StNum(1) & "," & StNum(2) & ")"
      RsTemp.Open Qty, cnnConnection
      If IsNull(RsTemp.Fields(0)) = True Then
         Stt(4) = "0.00"
      Else
         If RsTemp.Fields(0) <> "0" Then
            Stt(4) = Format((Stt(3) / Val(RsTemp.Fields(0))) * 100, "0.00")
         Else
            Stt(4) = "0.00"
         End If
      End If
      RsTemp.Close
      
      Qty = "SELECT (SUM(A0407)-SUM(A0406))/1000 FROM ACC040 " & _
            " WHERE A0405 IN (" & StNum(1) & ")" & _
            " AND A0401 =" & Text1(1) & _
            " AND (A0402 BETWEEN " & Format((i - 1) * 3 + 1, "00") & _
            " AND " & Format(i * 3, "00") & ")" & _
            " AND A0403='1' AND A0404='TOT'"
      RsTemp.Open Qty, cnnConnection
      If IsNull(RsTemp.Fields(0)) = True Then
         Stt(5) = "0.00"
      Else
         Stt(5) = Format(RsTemp.Fields(0), "0.00")
      End If
      RsTemp.Close
      
      Qty = "SELECT (SUM(A0407)-SUM(A0406))/1000 FROM ACC040 " & _
            " WHERE A0405 IN (" & StNum(2) & ")" & _
            " AND A0401 =" & Text1(1) & _
            " AND (A0402 BETWEEN " & Format((i - 1) * 3 + 1, "00") & _
            " AND " & Format(i * 3, "00") & ")" & _
            " AND A0403='1' AND A0404='TOT'"
      RsTemp.Open Qty, cnnConnection
      If IsNull(RsTemp.Fields(0)) = True Then
         Stt(6) = "0.00"
      Else
         Stt(6) = Format(RsTemp.Fields(0), "0.00")
      End If
      Stt(7) = Format(Val(Stt(5)) + Val(Stt(6)), "0.00")
      RsTemp.Close
      
      Qty = "SELECT SUM(A0409)/1000 FROM ACC040 WHERE A0401=" & Text1(1).Text & _
            " AND (A0402 BETWEEN " & Format((i - 1) * 3 + 1, "00") & _
            " AND " & Format(i * 3, "00") & ")" & _
            " AND A0403='1' AND A0404='TOT'" & _
            " AND A0405 IN (" & StNum(1) & "," & StNum(2) & ")"
      RsTemp.Open Qty, cnnConnection
      If IsNull(RsTemp.Fields(0)) = True Or RsTemp.Fields(0) = 0 Then
         Stt(8) = "0.00"
      Else
         Stt(8) = Format((Val(Stt(7)) / Val(RsTemp.Fields(0))) * 100, "0.00")
      End If
      RsTemp.Close
      
      Qty = "UPDATE TEMP SET TMP02='" & Stt(1) & "',TMP03='" & Stt(2) & "',TMP04='" & _
         Stt(3) & "',TMP05='" & Stt(4) & "',TMP06='" & Stt(5) & "'," & _
         "TMP07='" & Stt(6) & "',TMP08='" & Stt(7) & "',TMP09='" & _
         Stt(8) & "' WHERE TMP01='" & StS(i) & "'"
      Db.Execute Qty
   Next
   
   Qty = "SELECT SUM(VAL(TMP02)),SUM(VAL(TMP03)),SUM(VAL(TMP04))," & _
      "SUM(VAL(TMP06)),SUM(VAL(TMP07)),SUM(VAL(TMP08)) FROM TEMP"
   Set Rc = Db.OpenRecordset(Qty)
   With Rc
      Do While Not .EOF
         Qty = "INSERT INTO TEMP (TMP01,TMP02,TMP03,TMP04,TMP06,TMP07,TMP08) " & _
            "VALUES ('" & StS(5) & "','" & Format(.Fields(0), "0.00") & "','" & _
            Format(.Fields(1), "0.00") & "','" & Format(.Fields(2), "0.00") & "','" & _
            Format(.Fields(3), "0.00") & "','" & Format(.Fields(4), "0.00") & "','" & _
            Format(.Fields(5), "0.00") & "')"
         Db.Execute Qty
         Exit Do
      Loop
      .Close
   End With
   
   Qty = "SELECT TMP04,TMP08 FROM TEMP WHERE TMP01='全年'"
   Set Rc = Db.OpenRecordset(Qty)
   Do While Not Rc.EOF
      Qty = "SELECT SUM(A0409)/1000 FROM ACC040 WHERE A0401=" & Text1(0) & " AND " & _
            " A0403='1' AND A0404='TOT' " & _
            " AND A0405 IN (" & StNum(1) & "," & StNum(2) & ")"
      RsTemp.Open Qty, cnnConnection
      If IsNull(RsTemp.Fields(0)) = True Then
         Stt(5) = "0"
      Else
         If Format(RsTemp.Fields(0), "0.00") <> "0.00" Then
            Stt(5) = Format((Val(Rc.Fields(0)) / Val(RsTemp.Fields(0))) * 100, "0.00")
         Else
            Stt(5) = "0.00"
         End If
      End If
      RsTemp.Close
      
       Qty = "SELECT SUM(A0409)/1000 FROM ACC040 WHERE A0401=" & Text1(1) & " AND " & _
             "A0403 ='1' AND A0404='TOT' AND A0405 IN (" & StNum(1) & "," & StNum(2) & ")"
      RsTemp.Open Qty, cnnConnection
      If IsNull(RsTemp.Fields(0)) = True Then
         Stt(8) = "0"
      Else
         If Format(RsTemp.Fields(0), "0.00") <> "0.00" Then
            Stt(8) = Format((Val(Rc.Fields(1)) / Val(RsTemp.Fields(0))) * 100, "0.00")
         Else
            Stt(8) = "0.00"
         End If
      End If
      RsTemp.Close
      Exit Do
   Loop
   Rc.Close
   
   Qty = "UPDATE TEMP SET TMP05='" & Stt(5) & "',TMP09='" & Stt(8) & "' WHERE TMP01='全年'"
   Db.Execute Qty
   
   Qty = "INSERT INTO TEMP (TMP01,TMP02,TMP03,TMP04,TMP05,TMP06,TMP07,TMP08,TMP09) " & _
      "SELECT '平均每月',FORMAT(VAL(TMP02)/12,""0.00""),FORMAT(VAL(TMP03)/12,""0.00"")," & _
      "FORMAT(VAL(TMP04)/12,""0.00""),'---',FORMAT(VAL(TMP06)/12,""0.00"")," & _
      "FORMAT(VAL(TMP07)/12,""0.00""),FORMAT(VAL(TMP08)/12,""0.00""),'---' FROM TEMP WHERE TMP01='全年'"
   Db.Execute Qty
   
   Page = 1
   CaseTitle TmpArea, 1
   iPrint = 3100
   Qty = "SELECT TMP01,TMP02,TMP03,TMP04,TMP05,TMP06,TMP07,TMP08,TMP09 FROM TEMP"
   Set Rc = Db.OpenRecordset(Qty)
   With Rc
      Do While Not .EOF
         Printer.CurrentX = PLeft(0):      Printer.CurrentY = iPrint
         Printer.Print .Fields(0)
         Printer.CurrentX = PLeft(1) + 1000 - (Printer.TextWidth(CheckStr(.Fields(1))))
         Printer.CurrentY = iPrint
         Printer.Print .Fields(1)
         Printer.CurrentX = PLeft(2) + 1000 - (Printer.TextWidth(CheckStr(.Fields(2))))
         Printer.CurrentY = iPrint
         Printer.Print .Fields(2)
         Printer.CurrentX = PLeft(3) + 1000 - (Printer.TextWidth(CheckStr(.Fields(3))))
         Printer.CurrentY = iPrint
         Printer.Print .Fields(3)
         Printer.CurrentX = PLeft(4) + 1000 - (Printer.TextWidth(CheckStr(.Fields(4))))
         Printer.CurrentY = iPrint
         Printer.Print .Fields(4)
         Printer.CurrentX = PLeft(5) + 1000 - (Printer.TextWidth(CheckStr(.Fields(5))))
         Printer.CurrentY = iPrint
         Printer.Print .Fields(5)
         Printer.CurrentX = PLeft(6) + 1000 - (Printer.TextWidth(CheckStr(.Fields(6))))
         Printer.CurrentY = iPrint
         Printer.Print .Fields(6)
         Printer.CurrentX = PLeft(7) + 1000 - (Printer.TextWidth(CheckStr(.Fields(7))))
         Printer.CurrentY = iPrint
         Printer.Print .Fields(7)
         Printer.CurrentX = PLeft(8) + 1000 - (Printer.TextWidth(CheckStr(.Fields(8))))
         Printer.CurrentY = iPrint
         Printer.Print .Fields(8)
         iPrint = iPrint + 300
         .MoveNext
      Loop
   End With
   Rc.Close
   
   Qty = "SELECT TMP02,TMP03,TMP04,TMP06,TMP07,TMP08 FROM TEMP WHERE TMP01='全年'"
   Set Rc = Db.OpenRecordset(Qty)
   For i = 0 To 5
      Stt(i + 1) = Rc.Fields(i)
   Next
   Rc.Close
   
   Qty = "SELECT TMP02,TMP03,TMP04,TMP06,TMP07,TMP08 FROM TEMP WHERE TMP01='平均每月'"
   Set Rc = Db.OpenRecordset(Qty)
   For i = 0 To 5
      StS(i + 1) = Rc.Fields(i)
   Next
   Rc.Close
   
   Printer.CurrentX = PLeft(0):         Printer.CurrentY = iPrint
   Printer.Print String(205, "-")
   iPrint = iPrint + 300
   Printer.CurrentX = PLeft(0):         Printer.CurrentY = iPrint
   Printer.Print "分析說明 :"
   iPrint = iPrint + 300
   Printer.CurrentX = PLeft(1):         Printer.CurrentY = iPrint
   If Val(Stt(3)) - Val(Stt(6)) >= 0 Then
      Qty = "增加"
   Else
      Qty = "減少"
   End If
   intTemp = Abs(Val(Stt(3)) - Val(Stt(6)))
   If intTemp <> 0 Then
      intAvg = Round(intTemp / 12, 2)
   ElseIf intTemp = 0 Then
      intAvg = 0
   End If
   
   Printer.Print "1. " & Text1(0) & " 年度點數較 " & Text1(1) & " 年度" & Qty & _
      " " & Abs(Val(Stt(3)) - Val(Stt(6))) & " 點，平均每月" & Qty & " " & _
      intAvg & " 點"
      
   iPrint = iPrint + 300
   Printer.CurrentX = PLeft(1):         Printer.CurrentY = iPrint
   If Val(Stt(1)) - Val(Stt(4)) >= 0 Then
      Qty = "增加"
   Else
      Qty = "減少"
   End If

   intTemp = Abs(Val(Stt(1)) - Val(Stt(4)))
   If intTemp <> 0 Then
      intAvg = Round(intTemp / 12, 2)
   ElseIf intTemp = 0 Then
      intAvg = 0
   End If

   Printer.Print "2. " & TitName(1) & " 點數較 " & Text1(1) & " 年度" & Qty & " " & _
      Abs(Val(Stt(1)) - Val(Stt(4))) & " 點，平均每月" & Qty & " " & _
      intAvg & " 點"
      
   iPrint = iPrint + 300
   Printer.CurrentX = PLeft(1):         Printer.CurrentY = iPrint
   If Val(Stt(2)) - Val(Stt(5)) >= 0 Then
      Qty = "增加"
   Else
      Qty = "減少"
   End If
   intTemp = Abs(Val(Stt(2)) - Val(Stt(5)))
   If intTemp <> 0 Then
      intAvg = Round(intTemp / 12, 2)
   ElseIf intTemp = 0 Then
      intAvg = 0
   End If
   
   Printer.Print "3. " & TitName(2) & " 點數較 " & Text1(1) & " 年度" & Qty & " " & _
      Abs(Val(Stt(2)) - Val(Stt(5))) & " 點，平均每月" & Qty & " " & _
      intAvg & " 點"
   iPrint = iPrint + 300
   Printer.EndDoc
   Exit Sub
ErrHand:
   MsgBox Err.Description
End Sub

Private Sub GetPrintLeft()
   PLeft(0) = 500:      PLeft(1) = 2000
   PLeft(2) = 3500:     PLeft(3) = 5000
   PLeft(4) = 6500:     PLeft(5) = 8000
   PLeft(6) = 9500:     PLeft(7) = 11000
   PLeft(8) = 12500:    PLeft(9) = 14000
End Sub

Private Sub CaseTitle(ByVal Area As String, ByVal Page As String)
 Dim i As Integer, St As String
   i = 500
   Printer.Orientation = 2
   Printer.Font.Size = 22
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.CurrentX = 6000:         Printer.CurrentY = i
   Printer.Print "收款業績比較表"
   Printer.Font.Underline = False
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   Printer.CurrentX = 500:              Printer.CurrentY = i + 800
   Printer.Print "列印人 : " & strUserName
   Printer.CurrentX = 13000:            Printer.CurrentY = i + 800
   Printer.Print "列印日期 : " & ChangeTStringToTDateString(GetTaiwanTodayDate)
   Printer.CurrentX = 13000:            Printer.CurrentY = i + 1100
   Printer.Print "頁次 : " & Page
   
'   Printer.Font.Underline = True
   Printer.CurrentX = PLeft(2):              Printer.CurrentY = i + 1400
   Printer.Print Text1(0) & "　年度"
   Printer.CurrentX = PLeft(6):              Printer.CurrentY = i + 1400
   Printer.Print Text1(1) & "　年度"
'   Printer.Font.Underline = False
   Printer.Line (PLeft(2), i + 1700)-(PLeft(2) + 900, i + 1700)
   Printer.Line (PLeft(6), i + 1700)-(PLeft(6) + 900, i + 1700)

   Printer.CurrentX = 500:              Printer.CurrentY = i + 1700
   Printer.Print String(205, "-")
   Printer.CurrentX = PLeft(0):         Printer.CurrentY = i + 2000
   Printer.Print "季 別"
   Printer.CurrentX = PLeft(1):         Printer.CurrentY = i + 2000
   Printer.Print TitName(1) & " 點數"
   Printer.CurrentX = PLeft(2):         Printer.CurrentY = i + 2000
   Printer.Print TitName(2) & " 點數"
   Printer.CurrentX = PLeft(3):         Printer.CurrentY = i + 2000
   Printer.Print "達成點數"
   Printer.CurrentX = PLeft(4):         Printer.CurrentY = i + 2000
   Printer.Print "達成率 %"
   Printer.CurrentX = PLeft(5):         Printer.CurrentY = i + 2000
   Printer.Print TitName(1) & " 點數"
   Printer.CurrentX = PLeft(6):         Printer.CurrentY = i + 2000
   Printer.Print TitName(2) & " 點數"
   Printer.CurrentX = PLeft(7):         Printer.CurrentY = i + 2000
   Printer.Print "達成點數"
   Printer.CurrentX = PLeft(8):         Printer.CurrentY = i + 2000
   Printer.Print "達成率 %"
   Printer.CurrentX = PLeft(9):         Printer.CurrentY = i + 2000
   Printer.Print "備 註"
   Printer.CurrentX = 500:         Printer.CurrentY = i + 2300
   Printer.Print String(205, "-")
End Sub

Private Sub Form_Activate()
  Text1(0).SetFocus
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm083008 = Nothing
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   If (KeyAscii > 57 Or KeyAscii < 48) And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub
